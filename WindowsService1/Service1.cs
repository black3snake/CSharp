using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        static string PathToLog = @"C:\Robot20\SvcDocumentDZO\logs\";
        static string LogInstector = @"C:\Robot20\SvcDocumentDZO\logs\_CommonInspectorLog.txt";
        //static string FindStrung = "System.ComponentModel.Win32Exception";
        static string FindString = "Win32Exception";
        public Service1()
        {
            InitializeComponent();
        }

        protected override async void OnStart(string[] args)
        {
            //Console.WriteLine($"Путь будет такой {NameLog(PathToLog)}");
            File.AppendAllText(LogInstector, $"Путь будет такой {NameLog(PathToLog)} \r\ntime: {DateTime.Now.ToString("dd-MM-yyyy")} \r\n");
            bool OneStart = true;

            while (true)
            {
                if (!OneStart)
                {
                    if (!File.Exists(NameLog(PathToLog)))
                    {
                        //Console.WriteLine("Такого файла нет");
                        await Task.Delay(10000);
                    }
                    else
                    {
                        //Console.WriteLine($"Файл нашелся: {NameLog(PathToLog)}");
                        File.AppendAllText(LogInstector, $"Файл наконец то появился .. дождались: {NameLog(PathToLog)} \r\n");
                        //DateTime now = DateTime.Now;
                        DateTime creation = File.GetCreationTime(NameLog(PathToLog));
                        //DateTime modification = File.GetLastWriteTime(NameLog(PathToLog));


                        using (FileStream logFileStream = new FileStream(NameLog(PathToLog), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            if (DateTime.Now.AddHours(-1) > creation)
                            {
                                logFileStream.Seek(0, SeekOrigin.End);
                            }
                            else
                            {
                                logFileStream.Seek(0, SeekOrigin.Begin);
                            }

                            using (StreamReader logFileReader = new StreamReader(logFileStream, Encoding.Default))
                            {
                                while (File.Exists(NameLog(PathToLog)))
                                {
                                    //Thread.Sleep(10000);
                                    await Task.Delay(10000);
                                    try
                                    {
                                        string line = logFileReader.ReadLine();
                                        if (line == null)
                                        {
                                            continue;
                                        }
                                        else if (Regex.IsMatch(line, FindString))
                                        {
                                            //Console.WriteLine($"Найдена Строка: {line}");
                                            File.AppendAllText(LogInstector, $"{line} {DateTime.Now.ToString("dd-MM-yyyy")} \r\n");
                                            RestartService("Наблюдатель ContractDZO");
                                            RestartService("Наблюдатель DocumentDZO");

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        //Console.WriteLine(ex.Message);
                                        File.AppendAllText(LogInstector, $"{ex.Message} \r\n");

                                    }

                                }

                            }
                        }
                    }
                    //Console.WriteLine($"Файл нашелся: {NameLog(PathToLog)}");
                }
                OneStart = false;
                await Task.Delay(1000);
            }

        }
        // Возращает имя файла с текущей датой дня
        static string NameLog(string path)
        {
            DateTime dateTime = DateTime.Now;
            string NameFileLog = $"{dateTime.ToString("yyyy-MM-dd")}_errors.log";
            string PathToLogF = path + NameFileLog;
            return PathToLogF;
        }
        // Рестарт службу
        public static void RestartService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            TimeSpan timeout = TimeSpan.FromMinutes(2);
            if (service.Status != ServiceControllerStatus.Stopped)
            {
                //Console.WriteLine("Перезапуск службы. Останавливаем службу...");
                File.AppendAllText(LogInstector, $"Перезапуск службы. Останавливаем службу...{serviceName} \r\n");
                // Останавливаем службу
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
                //Console.WriteLine("Служба была успешно остановлена!");
                File.AppendAllText(LogInstector, $"Служба была успешно остановлена! {serviceName} \r\n");
            }
            if (service.Status != ServiceControllerStatus.Running)
            {
                //Console.WriteLine("Перезапуск службы. Запускаем службу...");
                File.AppendAllText(LogInstector, $"Перезапуск службы. Запускаем службу...{serviceName}  \r\n");
                // Запускаем службу
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                //Console.WriteLine("Служба была успешно запущена!");
                File.AppendAllText(LogInstector, $"Служба была успешно запущена! {serviceName}  \r\n");
            }
        }

        // Запуск службы
        public static void StartService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            // Проверяем не запущена ли служба
            if (service.Status != ServiceControllerStatus.Running)
            {
                // Запускаем службу
                service.Start();
                // В течении минуты ждём статус от службы
                service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromMinutes(1));
                Console.WriteLine("Служба была успешно запущена!");
            }
            else
            {
                Console.WriteLine("Служба уже запущена!");
            }
        }

        // Останавливаем службу
        public static void StopService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            // Если служба не остановлена
            if (service.Status != ServiceControllerStatus.Stopped)
            {
                // Останавливаем службу
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromMinutes(1));
                //Console.WriteLine("Служба была успешно остановлена!");
                File.AppendAllText(LogInstector, $"Служба была успешно остановлена! {serviceName}  \r\n");
            }
            else
            {
                //Console.WriteLine("Служба уже остановлена!");
                File.AppendAllText(LogInstector, $"Служба уже остановлена! {serviceName}  \r\n");
            }
        }



        protected override void OnStop()
        {
        }
    }
}
