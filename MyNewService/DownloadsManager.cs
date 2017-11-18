using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;

namespace DownloadsManager
{
    public partial class DownloadsManager : ServiceBase
    {
        // добавление структуры для состояния, которая будет использоваться при вызове неуправляемого кода
        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public long dwServiceType;
            public ServiceState dwCurrentState;
            public long dwControlsAccepted;
            public long dwWin32ExitCode;
            public long dwServiceSpecificExitCode;
            public long dwCheckPoint;
            public long dwWaitHint;
        };

        // Функция, реализованная с помощью вызова неуправляемого кода
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);
        // данный конструктор задает имя источника события и журнала в соответствии с заданными параметрами запуска или использует значения по умолчанию, если никакие аргументы не заданы.
        public DownloadsManager(string[] args)
        {
            InitializeComponent(); string eventSourceName = "MySource";
            string logName = "MyNewLog";
            if (args.Count() > 0)
            {
                eventSourceName = args[0];
            }
            if (args.Count() > 1)
            {
                logName = args[1];
            }
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists(eventSourceName))
            {
                System.Diagnostics.EventLog.CreateEventSource(eventSourceName, logName);
            }
            eventLog1.Source = eventSourceName;
            eventLog1.Log = logName;
        }

        // конструктор для определения настраиваемого журнала событий ( журнал нужен для тестирования в процессе разработки и обновления)
        public DownloadsManager()
        {
            InitializeComponent();
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "MySource", "MyNewLog");
            }
            eventLog1.Source = "MySource";
            eventLog1.Log = "MyNewLog";
        }

        int eventId = 0;
        // обработка события таймера
        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
             
            eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);
        }
        // создание массивов, содержащих расширения файлов 
        public static string[] ExtencionsOfImages = { ".png", ".jpeg", ".gif", ".jpg" };
        public static string[] ExtencionsOfMusic = { ".mp3", ".flac" };
        public static string[] ExtencionsOfVideos = { ".mp4", ".avi", ".flv", ".wmv" };
        public static string[] ExtencionsOfDocuments = { ".docx", ".doc", ".ppt", ".xls", ".txt", ".pdf" };

        protected override void OnStart(string[] args)
        {
            // реализация состояния ожидание
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            // таймер для отслеживания работы службы
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000; // 60 секунд  
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();
            // при запуске службы в журнал будет добавлена следующая запись
            eventLog1.WriteEntry("In OnStart");

            // присвоение состоянию значения "Запуск"
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // создание объекта класса FileSystemWatcher для мониторинга нужной нам папки, здесь же указывается её путь
            var fileSystemWatcher = new FileSystemWatcher(@"C:\Users\user\Desktop\Downloads1")
            {
                EnableRaisingEvents = true
            };
            // создание события Created, которое происходит при "создании" файл по заданному пути
            fileSystemWatcher.Created += (a, e) =>
            {
                FileInfo info = new FileInfo(e.FullPath); //переменная, которой будет хранится информация о "созданном" файле
                string newFullPath; // переменная, в которую будет записан путь файла, после перемещения в соответствующую папку
                bool IsException = true;// булева переменная, хранящие значение о наличии ошибки при создании файла прямо в наблюдаемой папке, и для дальнеёшей " поимки" с помощью try-catch
                bool IsMovementDone = false;// булева переменная, хранящая значение о том, был ли файл уже перемещён в нужную папку

                // цикл, для проверки расширения файла и , при совпадении,  дальнейшего перемещения в папку Images
                foreach (var extencion in ExtencionsOfImages)
                {
                    if (info.Extension == extencion)
                    {
                        newFullPath = @"C:\Users\user\Desktop\Downloads1\Images\"; // путь папки в которую будет перемещён файл, при совпадении расширений
                        newFullPath += info.Name;// новый полный путь файла
                        // "поимка" ошибки при создании файла непосредственно в наблюдаемой папке, после перемещения файла в нужную папку цикл завершится
                        while (IsException)
                        {
                            try
                            {
                                File.Move(info.FullName, newFullPath); // попытка перемещения файла по новому пути
                                IsException = false; // маркер показывающий, что ошибки нет
                            }

                            catch (IOException exception)
                            {
                            }
                        }
                        IsMovementDone = true; // маркер, показывающий, что файл уже был перемещён
                    }
                }

                if (!IsMovementDone) // проверка, был ли файл уже перемещён в нужную папку
                {
                    // цикл, для проверки расширения файла и , при совпадении,  дальнейшего перемещения в папку Videos
                    foreach (var extencion in ExtencionsOfVideos)
                    {
                        if (info.Extension == extencion)
                        {
                            newFullPath = @"C:\Users\user\Desktop\Downloads1\Videos\"; // путь папки в которую будет перемещён файл, при совпадении расширений
                            newFullPath += info.Name;// новый полный путь файла
                            // "поимка" ошибки при создании файла непосредственно в наблюдаемой папке, после перемещения файла в нужную папку цикл завершится
                            while (IsException)
                            {
                                try
                                {
                                    File.Move(info.FullName, newFullPath);// попытка перемещения файла по новому пути
                                    IsException = false; // маркер показывающий, что ошибки нет
                                }

                                catch (IOException exception)
                                {
                                }
                            }
                            IsMovementDone = true; // маркер, показывающий, что файл уже был перемещён
                        }
                    }
                }
                if (!IsMovementDone)// проверка, был ли файл уже перемещён в нужную папку
                {
                    // цикл, для проверки расширения файла и , при совпадении,  дальнейшего перемещения в папку Documents
                    foreach (var extencion in ExtencionsOfDocuments)
                    {
                        if (info.Extension == extencion)
                        {
                            newFullPath = @"C:\Users\user\Desktop\Downloads1\Documents\";// путь папки в которую будет перемещён файл, при совпадении расширений
                            newFullPath += info.Name;// новый полный путь файла
                           // "поимка" ошибки при создании файла непосредственно в наблюдаемой папке, после перемещения файла в нужную папку цикл завершится
                            while (IsException)
                            {
                                try
                                {
                                    File.Move(info.FullName, newFullPath);// попытка перемещения файла по новому пути
                                    IsException = false;// маркер показывающий, что ошибки нет
                                }

                                catch (IOException exception)
                                {
                                }
                            }
                            IsMovementDone = true;// маркер, показывающий, что файл уже был перемещён
                        }
                    }
                }
                if (!IsMovementDone)// проверка, был ли файл уже перемещён в нужную папку
                {
                    // цикл, для проверки расширения файла и , при совпадении,  дальнейшего перемещения в папку Documents
                    foreach (var extencion in ExtencionsOfMusic)
                    {
                        if (info.Extension == extencion)
                        {
                            newFullPath = @"C:\Users\user\Desktop\Downloads1\Music\";// путь папки в которую будет перемещён файл, при совпадении расширений
                            newFullPath += info.Name;// новый полный путь файла
                            // "поимка" ошибки при создании файла непосредственно в наблюдаемой папке, после перемещения файла в нужную папку цикл завершится
                            while (IsException)
                            {
                                try
                                {
                                    File.Move(info.FullName, newFullPath);// попытка перемещения файла по новому пути
                                    IsException = false;// маркер показывающий, что ошибки нет
                                }

                                catch (IOException exception)
                                {
                                }
                            }
                            IsMovementDone = true;// маркер, показывающий, что файл уже был перемещён
                        }
                    }
                }
                // если файл не подошёл ни в одну папку он будет перемещён в папку Other
                if (!IsMovementDone)// проверка, был ли файл уже перемещён в нужную папку
                {
                    newFullPath = @"C:\Users\user\Desktop\Downloads1\Other\";// путь папки в которую будет перемещён файл, при совпадении расширений
                    newFullPath += info.Name;// новый полный путь файла
                    // "поимка" ошибки при создании файла непосредственно в наблюдаемой папке, после перемещения файла в нужную папку цикл завершится
                    while (IsException)
                    {
                        try
                        {
                            File.Move(info.FullName, newFullPath);// попытка перемещения файла по новому пути
                            IsException = false;// маркер показывающий, что ошибки нет
                        }

                        catch (IOException exception)
                        {
                        }
                    }
                }

            };
           }

        protected override void OnStop()
        {
            // реализация состояния ожидание  
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            // при остановке службы в журнал будет добавлена следующая запись
            eventLog1.WriteEntry("In onStop.");

            // присвоение состоянию значения "Запуск"  
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnContinue()
        {
            // при возобновлении работы службы в журнал будет добавлена следующая запись
            eventLog1.WriteEntry("In OnContinue.");
        }

    }
}
