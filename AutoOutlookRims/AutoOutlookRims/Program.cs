using AutoOutlookRims.DataModel;
using Config.Net;
using Newtonsoft.Json;
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace AutoOutlookRims
{
    public partial class Program
    {
        #region AES начальные данные
        static string passPhrase = "TestPassphrase";        //Может быть любой строкой
        static string saltValue = "TestSaltValue";        // Может быть любой строкой
        static string hashAlgorithm = "SHA256";             // может быть "MD5"
        static int passwordIterations = 2;                //Может быть любым числом
        static string initVector = "!1A3g2D4s9K556g7"; // Должно быть 16 байт
        static int keySize = 256;                // Может быть 192 или 128
        #endregion

        public static IMySettings configiniP = new ConfigurationBuilder<IMySettings>()
            .UseIniFile(@"config.ini", true)
            .Build();
        static string plainText { get; set; }
        static long countDB { get; set; }
        static string elapsedTimeSaveBD { get; set; }
        static int statusAUsers { get; set; }
        static int QThred { get; set; }
        static string elapsedTimeGetStatus { get; set; }
        static string elapsedTimeSetAutoAnswer { get; set; }

        static bool statusDBTest;
        static string statusRimsTest;
        static bool statusZDataRims;
        static bool statusZDataSQL;
        static bool statusCheckAutoAnswer;
        // создадим булеву переменную для отладки если true то времязатратные методы отключены
        static bool startDebug = false;

        static Logger logger = LogManager.GetCurrentClassLogger();
        static Dictionary<string, string> DLeaveType = new Dictionary<string, string>()
        {
            {"VC","Отпуск"},
            {"SL","Больничный"},
            {"BT","Командировка"},
            {"DV","Декретный отпуск"}
        };
        static Dictionary<string, int> DLeaveType2C = new Dictionary<string, int>()
        {
            {"VC", 1},
            {"SL", 2},
            {"BT", 3},
            {"DV", 4}
        };

        static List<Datauser> Datauserslist = new List<Datauser>();
        static List<string> exlistusers = new List<string>();

        // Структура формирование эл. ответа в Exchange
        struct OutOfOffice
        {
            public string leaveName { get; set; }
            public DateTime DateStart { get; set; }
            public DateTime DateEnd { get; set; }
            public string DS { get; set; }
            public string DE;
            public string FIO { get; set; }

            public OutOfOffice(string leaveName, DateTime DateStart, DateTime DateEnd, string FIO)
            {
                this.leaveName = leaveName;
                this.DateStart = DateStart;
                this.DateEnd = DateEnd;
                this.DS = "";
                this.DE = "";
                this.FIO = FIO;
            }

            public string As()
            {

                DS = DateStart.ToString().Remove(DateStart.ToString().IndexOf(" "));
                DE = DateEnd.ToString().Remove(DateStart.ToString().IndexOf(" "));

                if (leaveName == "VC")
                {
                    return leaveName = $"<html><body><b>Добрый день.</b><br> В данный момент я нахожусь в отпуске c {DS} по {DE}<br>    С уважением {FIO}</body></html>";
                }
                else if (leaveName == "SL")
                {
                    return leaveName = $"<html><body><b>Добрый день.</b><br> В данный момент я нахожусь на больничном c {DS} по {DE}<br>    С уважением {FIO}</body></html>";
                }
                else if (leaveName == "BT")
                {
                    return leaveName = $"<html><body><b>Добрый день.</b><br> В данный момент я нахожусь в командировке c {DS} по {DE}<br>   С уважением {FIO}</body></html>";
                }
                else if (leaveName == "DV")
                {
                    return leaveName = $"<html><body><b>Добрый день.</b><br> В данный момент я нахожусь в декретном отпуске c {DS} по {DE}<br>  С уважением {FIO}</body></html>";
                }
                else
                    return leaveName;

            }
        }



        async static Task Main(string[] args)
        {

            #region NLog Initializator

            var config = new NLog.Config.LoggingConfiguration();
            LogManager.Configuration = new LoggingConfiguration();
            const string LayoutFile = @"[${date:format=yyyy-MM-dd HH\:mm\:ss}] [${logger}/${uppercase: ${level}}] [THREAD: ${threadid}] >> ${message} ${exception: format=ToString}";

            var logfile = new FileTarget();

            if (!Directory.Exists("logs"))
                Directory.CreateDirectory("logs");

            // Rules for mapping loggers to targets
            config.AddRule(LogLevel.Trace, LogLevel.Fatal, logfile);

            DateTime dtlog = DateTime.Now;

            logfile.CreateDirs = true;
            logfile.FileName = $"logs{Path.DirectorySeparatorChar}{dtlog:yyyy-MM-dd}.log";
            logfile.AutoFlush = true;
            logfile.LineEnding = LineEndingMode.CRLF;
            logfile.Layout = LayoutFile;
            logfile.FileNameKind = FilePathKind.Absolute;
            logfile.ConcurrentWrites = false;
            logfile.KeepFileOpen = false;


            // Apply config
            NLog.LogManager.Configuration = config;

            #endregion NLog Initializator

            logger.Info($"!!!---- Start Main Program {DateTime.Now}");

            plainText = RijndaelAlgorithm.Decrypt
            (
                configiniP.password,
                passPhrase,
                saltValue,
                hashAlgorithm,
                passwordIterations,
                initVector,
                keySize
            );

            // Проверка на количество потоков в файле config.ini
            if (configiniP.ThreadsInt < 1 || configiniP.ThreadsInt > 10)
            {
                Console.WriteLine("Количество потоков для работы разрешено от 1..10");
                logger.Info("Количество потоков для работы разрешено от 1..10");
                Console.ReadKey();
                System.Environment.Exit(0);
            }

            // call Testing RIMS
            statusRimsTest = await TestRims(configiniP.domen, configiniP.login, plainText);
            if (statusRimsTest == "OK")
            {
                Console.WriteLine("Провека к Римсу прошла удачно -> Летим дальше");
                logger.Info("Провека к Римсу прошла удачно -> Летим дальше");
            }
            else
            {
                Console.WriteLine("Что-то не так");
                logger.Info("Что-то не так с проверкой к РИМС");
                Console.ReadKey();
                System.Environment.Exit(0);
            }

            // call Testing MS SQL
            if (testDb())
            {
                Console.WriteLine("Связь с SQL есть! -> Летим дальше ");
                logger.Info("Связь с SQL есть! -> Летим дальше ");
            }
            else
            {
                Console.WriteLine("Что-то не так");
                logger.Info("Что-то не так с проверкой к SQL");
                Console.ReadKey();
                System.Environment.Exit(0);
            }

            // Получим данные из РИМСа
            statusZDataRims = await ZaprosRimsIn(configiniP.domen, configiniP.login, plainText);
            if (statusZDataRims)
            {
                Console.WriteLine("Данные от РИМСа получены -> Летим дальше");
                logger.Info("Данные от РИМСа получены -> Летим дальше");

            }
            else
            {
                Console.WriteLine("Что-то не так");
                logger.Info("Что-то не так, данные не получены от РИМС");
                Console.ReadKey();
                System.Environment.Exit(0);
            }

            // Записываем данные в Базу данных на сервере SQL
            if (!startDebug)
            {
                if (SaveDB())
                {
                    Console.WriteLine("Данные в сервер MS SQL записаны -> Летим дальше");
                    logger.Info("Данные в сервер MS SQL записаны -> Летим дальше");
                }
                else
                {
                    Console.WriteLine("Что-то не так");
                    logger.Info("Что-то не так, не получилось записать в базу SQL");
                    Console.ReadKey();
                    System.Environment.Exit(0);
                }
            }

            // Проверка в RIMS и установка Автоответа в SQL DB (PLINQ - несколькими потоками)
            if (!startDebug)
            {
                Console.WriteLine("Проверка статусов в RIMS и установка Автоответа в SQL");
                logger.Info("Проверка статусов в RIMS и установка Автоответа в SQL");
                if (CheckSetAutoAnswer())
                {
                    Console.WriteLine("Данные о статусе Автоответа пользователей записаны в нашу Базу -> Летим дальше");
                    logger.Info("Данные о статусе Автоответа пользователей записаны в нашу Базу -> Летим дальше");
                }
                else
                {
                    Console.WriteLine("Что-то не так");
                    logger.Info("Что-то не так, не удалось записать статусы в базу SQL");
                    Console.ReadKey();
                    System.Environment.Exit(0);
                }

            }

            // Обработаем пользователей (исключения), что в файле config.ini
            ExUser();

            // Запись в Exchange через запросы в РИМС
            if (!startDebug)
            {
                SaveInExchange();
            }

            // Простой вывод пользователей которые попали в список исключений
            string exuser = "";
            foreach (var item in exlistusers)
            {
                Console.Write($"{item} ");
                exuser += item + ",";
            }
            logger.Info("Пользователи находящиеся в исключениях");
            logger.Info($"{exuser.TrimEnd(' ', ',')}");
            Console.WriteLine();

            // Определим количество разрешенных лог файлов и возьмем его в config.ini
            string[] filePaths = Directory.GetFiles(@"logs\", "*.log");
            string[] filePaths2 = null;
            if (filePaths.Count() > configiniP.QuantityLogs)
            {
                DateTime dtDate;
                CultureInfo provider = CultureInfo.CreateSpecificCulture("en-US");
                DateTimeStyles styles = DateTimeStyles.None;
                int count = 0;
                foreach (var logFile in filePaths)
                {
                    if (DateTime.TryParse(logFile.Remove(logFile.Length - 4).Remove(0, 5), provider, styles, out dtDate))
                    {
                        if (DateTime.Today.AddDays(-configiniP.QuantityLogs) > dtDate)
                        {
                            if (filePaths2?.Count() == configiniP.QuantityLogs + 1) { break; }
                            //Console.WriteLine($"{DateTime.Today.AddDays(-configiniP.QuantityLogs)} > {dtDate}" );
                            try
                            {
                                File.Delete(logFile);
                                count++;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            //continue;
                        }
                        //Console.WriteLine($"{logFile} -> {dtDate}");
                    }
                    filePaths2 = Directory.GetFiles(@"logs\", "*.log");
                }
                Console.WriteLine($"Всего было лог файлов {filePaths.Length} из них удалено старых {count}");
                logger.Info($"Всего было лог файлов {filePaths.Length} из них удалено старых {count}");

                Console.WriteLine($"Дата устаревания {DateTime.Today.AddDays(-configiniP.QuantityLogs)} и разрешенное кол-во логов {configiniP.QuantityLogs}");
                logger.Info($"Дата устаревания {DateTime.Today.AddDays(-configiniP.QuantityLogs)} и разрешенное кол-во логов {configiniP.QuantityLogs}");

            }

            //Console.WriteLine("Конец отчета, press Any key");
            logger.Info("Конец отчета");

            // Будем формировать архивный файл отчета, для отправки в письме
            string pathFL = $"{dtlog:yyyy-MM-dd}.log";
            string pathArhivFL = $"{dtlog:yyyy-MM-dd}.zip";
            using (var fileStream = new FileStream($@"logs\{pathArhivFL}", FileMode.Create, FileAccess.ReadWrite))
            {
                using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create))
                {
                    archive.CreateEntryFromFile($@"logs\{pathFL}", pathFL);
                }
            }
            // вот тут будем формировать отчеты в письме
            //Program2 send22 = new Program2();
            //if(await send22.SendM(configiniP, countDB))
            if (!startDebug)
            {
                if (await SendM())
                {
                    Console.WriteLine($"Письмо отправлено");
                    logger.Info($"Письмо отправлено");
                }
                else
                {
                    Console.WriteLine($"Письмо НЕ отправлено");
                    logger.Info($"Письмо НЕ отправлено");
                }
            }

            // Удалим ненужный архивный файл отчета
            if (!startDebug)
            {
                try
                {
                    if (File.Exists($@"logs\{pathArhivFL}"))
                    {
                        File.Delete($@"logs\{pathArhivFL}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    logger.Info(ex.Message);
                }
            }
            //Console.ReadKey();
        }

        // Testing SQL 
        static bool testDb()
        {
            using (DataModelContext context = new DataModelContext())
            {
                IEnumerable<Leave> listsLEnum = context.Leaves;
                var countDL = listsLEnum.ToList().LongCount();

                //var listsL = context.Leaves.ToList();
                //var countDL = listsL.LongCount();

                IEnumerable<Datauser> listsDEnum = context.Datausers;
                var countDD = listsDEnum.ToList().LongCount();


                Console.WriteLine($"Quantity items in DB Datausers: {countDD}");
                logger.Info($"Quantity items in DB Datausers: {countDD}");

                if (countDL > 0)
                {
                    statusDBTest = true;
                    Console.WriteLine(statusDBTest);
                }
                else
                {
                    statusDBTest = false;
                    Console.WriteLine(statusDBTest);
                }
            }

            if (statusDBTest)
            {

                using (DataModelContext contextOld = new DataModelContext())
                {
                    var usersOld = contextOld.Datausers.Where(p => p.LeaveEnd <= DateTime.Now.AddHours(-24)).ToList();
                    foreach (var uOld in usersOld)
                    {
                        //Console.WriteLine($"{uOld.LastName} {uOld.FirstName} {uOld.MiddleName} :: {uOld.LeaveStart} -  {uOld.LeaveEnd}");
                        contextOld.Datausers.Remove(uOld);
                        contextOld.SaveChanges();

                    }
                    Console.WriteLine($"Quantity olds removed: {usersOld.Count}");
                }

                using (DataModelContext status4 = new DataModelContext())
                {
                    var listsC = status4.Datausers.Where(l => l.AnswerId == 4).ToList();
                    Console.WriteLine($"Пользователей со своим установленным статусом автоответа: {listsC.Count()}");
                    logger.Info($"Пользователей со своим установленным статусом автоответа: {listsC.Count()}");


                    // проверим на Null и при нахождении заменим на 1
                    var listsNull = status4.Datausers.Where(l => l.AnswerId == null).ToList();
                    foreach (var item in listsNull)
                    {
                        item.AnswerId = 1;
                        status4.SaveChanges();
                    }
                }

            }

            return statusDBTest;
        }

        // Method Testing RIMS
        static async Task<string> TestRims(string domen, string login, string password)
        {
            // NTLM Secured URL
            var uri = new Uri(configiniP.URI1);
            // Create a new Credential
            var credentialsCache = new CredentialCache();
            credentialsCache.Add(uri, "NTLM", new NetworkCredential(
                login, password, domen));

            var handler = new HttpClientHandler() { Credentials = credentialsCache, PreAuthenticate = true };
            var httpClient = new HttpClient(handler) { Timeout = new TimeSpan(0, 0, 10) };

            var response = await httpClient.GetAsync(uri);

            var result = await response.Content.ReadAsStringAsync();

            //txtBoxConsole.AppendText(result + Environment.NewLine);
            //Console.WriteLine(result);
            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine($"Статус проверки REST Get: {response.ReasonPhrase}");
                logger.Info($"Статус проверки REST Get: {response.ReasonPhrase}");
                //statusRimsTest = true;
            }
            else
            {
                Console.WriteLine($"Статус проверки REST Get: {response.ReasonPhrase}");
                logger.Info($"Статус проверки REST Get: {response.ReasonPhrase}");
                //statusRimsTest = false;
            }

            return response.ReasonPhrase;
        }

        // Заберем данные с РИМС
        static async Task<bool> ZaprosRimsIn(string domen, string login, string password)
        {
            // NTLM Secured URL
            var uri = new Uri(configiniP.URI2);
            // Create a new Credential
            var credentialsCache = new CredentialCache();
            credentialsCache.Add(uri, "NTLM", new NetworkCredential(
                login, password, domen));

            var handler = new HttpClientHandler() { Credentials = credentialsCache, PreAuthenticate = true };
            var httpClient = new HttpClient(handler) { Timeout = new TimeSpan(0, 0, 10) };
            var response = await httpClient.GetAsync(uri);
            var result = await response.Content.ReadAsStringAsync();
            var statusL = JsonConvert.DeserializeObject<List<RimsZaprosOutput>>(result);

            double count = 0;
            Console.Write($"Начинаем заполнять список..");
            foreach (var item in statusL)
            {
                count++;
                if ((count % 100) == 0)
                {
                    Console.Write($"..{(int)count}");
                }
                // Обнаружил Ошибку в Статус LeaveType (появился неизвестный WC - исключим такие)
                if (!DLeaveType2C.ContainsKey(item.LeaveType))
                {
                    logger.Info($"!!!--- Ошибка в {item.AccountName} -> {item.LeaveType} ---!!!");
                    continue;
                }

                Datauser userObject = new Datauser()
                {
                    FimSyncKey = item.FimSyncKey,
                    AccountId = item.AccountId,
                    AccountName = item.AccountName.Count() > 50 ? item.AccountName.Substring(49) : item.AccountName,
                    LastName = item.LastName?.Count() > 100 ? item.LastName.Substring(99) : item.LastName,
                    FirstName = item.FirstName?.Count() > 100 ? item.FirstName.Substring(99) : item.FirstName,
                    MiddleName = item.MiddleName?.Count() > 100 ? item.MiddleName.Substring(99) : item.MiddleName,
                    EmployeeNumber = item.EmployeeNumber?.Count() > 20 ? item.EmployeeNumber.Substring(19) : item.EmployeeNumber,
                    Birthday = item.Birthday,
                    CompanyName = item.CompanyName?.Count() > 300 ? item.CompanyName.Substring(299) : item.CompanyName,
                    DepartmentName = item.DepartmentName?.Count() > 200 ? item.DepartmentName.Substring(199) : item.DepartmentName,
                    JobTitle = item.JobTitle?.Count() > 200 ? item.JobTitle.Substring(199) : item.JobTitle,
                    DateIn = item.DateIn,
                    LeaveId = DLeaveType2C[item.LeaveType],
                    LeaveStart = item.LeaveStart,
                    LeaveEnd = item.LeaveEnd,
                    City = item.City?.Count() > 100 ? item.City.Substring(99) : item.City,
                    Phone = item.Phone?.Count() > 100 ? item.Phone.Substring(99) : item.Phone,
                    Email = item.Email?.Count() > 100 ? item.Email.Substring(99) : item.Email,
                    Disabled = item.DisabledDomain,
                    AnswerId = 1

                };
                Datauserslist.Add(userObject);

            }
            Console.WriteLine();
            Console.WriteLine($"Пользователей в списке: {Datauserslist.Count}");
            logger.Info($"Пользователей в списке: {Datauserslist.Count}");


            return response.IsSuccessStatusCode;
        }

        // Записываем данные в Базу данных на сервере SQL
        static bool SaveDB()
        {
            long count_db = 0;
            statusZDataSQL = false;
            using (DataModelContext countDB = new DataModelContext())
            {
                var lists = countDB.Datausers.ToList();
                count_db = lists.LongCount();
            }

            Console.WriteLine($"Начинаем запись в Базу данных {configiniP.Address}:{configiniP.BDName}");
            logger.Info($"Начинаем запись в Базу данных {configiniP.Address}:{configiniP.BDName}");
            Stopwatch stopwatchBD = new Stopwatch();
            stopwatchBD.Start();


            // Сохранение статусов 3 и 4 от перезаписи
            using (DataModelContext context = new DataModelContext())
            {
                foreach (var item in Datauserslist)
                {
                    var user = context.Datausers.FirstOrDefault(u => u.AccountName == item.AccountName);
                    if (user != null)
                    {
                        if (user.AnswerId == 3)
                        {
                            item.AnswerId = 3;
                        }
                        else if (user.AnswerId == 4)
                        {
                            item.AnswerId = 4;
                        }
                    }

                }
            } // Сохранение статусов 3 и 4 от перезаписи

            // запишем наши данные в базу
            using (DataModelContext context = new DataModelContext())
            {
                if (count_db > 1000)
                {
                    foreach (var item in Datauserslist)
                    {
                        try
                        {
                            context.Datausers.Update(item);
                            context.SaveChanges();
                        }
                        catch
                        {
                            context.Datausers.Add(item);
                            //MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            context.SaveChanges();
                        }
                    }
                }
                else
                {
                    foreach (var item in Datauserslist)
                    {
                        try
                        {
                            context.Datausers.Add(item);
                            context.SaveChanges();
                        }
                        catch
                        {
                            context.Datausers.Update(item);
                            //MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            context.SaveChanges();
                        }
                    }

                }
                stopwatchBD.Stop();
                TimeSpan stopwatchElapsed = stopwatchBD.Elapsed;
                var milsec = Convert.ToInt32(stopwatchElapsed.TotalMilliseconds);
                var sec = milsec / 1000;
                var ts = TimeSpan.FromSeconds(sec);
                elapsedTimeSaveBD = $"{ts.Hours}ч:{ts.Minutes}м:{ts.Seconds}с";
                Console.WriteLine($"На запись ушло времени: {elapsedTimeSaveBD}");
                logger.Info($"На запись ушло времени: {elapsedTimeSaveBD}");
                statusZDataSQL = true;


                #region AddRange List no work AddAndUpdate
                /*try
                {
                    //context.Configuration.AutoDetectChangesEnabled = false;
                    context.Datausers.  Range(Datauserslist);
                    context.SaveChanges();
                }
                catch
                {
                    context.Datausers.UpdateRange(Datauserslist);
                    context.SaveChanges();
                }
                finally
                {
                    //context.Configuration.AutoDetectChangesEnabled = true;
                }*/
                #endregion
            };

            using (DataModelContext countDB2 = new DataModelContext())
            {
                var lists = countDB2.Datausers.ToList();
                countDB = lists.LongCount();
                Console.WriteLine($"Quantity items after Save in DB: {countDB}");
                logger.Info($"Quantity items after Save in DB: {countDB}");
            }

            return statusZDataSQL;
        }

        // Проверка в RIMS и установка Автоответа в SQL DB (PLINQ - несколькими потоками)
        static bool CheckSetAutoAnswer()
        {
            statusCheckAutoAnswer = false;
            List<string> listsID = new List<string>();

            using (DataModelContext context = new DataModelContext())
            {
                // не выбираються записи в базе с автоответами 3 уже постав. ранее и 4 установлен свой
                var listsIdDB = context.Datausers.Where(l => l.AnswerId != 3 && l.AnswerId != 4).Select(l => l.AccountName).ToList();
                foreach (var item in listsIdDB)
                {
                    listsID.Add(item);
                    //if (listsID.Count > 40) break;   // Ограничение на Внесение стутуса 3 кнопка
                }

            }
            int threadId = Thread.CurrentThread.ManagedThreadId;
            Console.WriteLine($"Main: запущен в потоке # {threadId}");
            logger.Info($"Main: запущен потоке: {threadId}");

            Thread.Sleep(1000);

            ParallelOptions options = new ParallelOptions();
            // Выделить определенное количество процессорных ядер.
            //options.MaxDegreeOfParallelism = Environment.ProcessorCount > 4 ? Environment.ProcessorCount - 1 : 1;
            /*if (Environment.ProcessorCount > 4)
                options.MaxDegreeOfParallelism = Environment.ProcessorCount < 10 ? 4 : 5;
            else
                options.MaxDegreeOfParallelism = 1;*/
            // Данные о количестве потоков мы возьмем из config.ini
            options.MaxDegreeOfParallelism = configiniP.ThreadsInt;


            QThred = options.MaxDegreeOfParallelism;
            Console.WriteLine($"Количество логических ядер CPU: {Environment.ProcessorCount}");
            Console.WriteLine($"Будем использовать: {options.MaxDegreeOfParallelism} потоков");

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            listsID.AsParallel().WithDegreeOfParallelism(options.MaxDegreeOfParallelism).ForAll(ls => { MyTask(ls); });

            stopwatch.Stop();
            TimeSpan stopwatchElapsed = stopwatch.Elapsed;
            var milsec = Convert.ToInt32(stopwatchElapsed.TotalMilliseconds);
            var sec = milsec / 1000;
            var ts = TimeSpan.FromSeconds(sec);

            elapsedTimeGetStatus = $"{ts.Hours}ч:{ts.Minutes}м:{ts.Seconds}с";
            Console.WriteLine($"Затраченное время: {ts.Hours}ч:{ts.Minutes}м:{ts.Seconds}с");
            Console.WriteLine($"Всего AcountName попавших в обработку: {listsID.Count}");
            Console.WriteLine($"Основной поток завершен.");
            logger.Info($"Затраченное время на обработку получения Статуса пользователей из РИМСа: {ts.Hours}ч:{ts.Minutes}м:{ts.Seconds}с");
            logger.Info($"Всего AcountName попавших в обработку: {listsID.Count}");
            logger.Info($"Основной поток завершен.");

            using (DataModelContext contextC = new DataModelContext())
            {
                var listsC = contextC.Datausers.Where(l => l.AnswerId == 4).ToList();
                statusAUsers = listsC.Count();
                Console.WriteLine($"Итого добавлено {statusAUsers} статусов в DBase о наличии пользовательских автоответов, которые мы сохраним");
                logger.Info($"Итого добавлено {statusAUsers} статусов в DBase о наличии пользовательских автоответов, которые мы сохраним");
            }

            statusCheckAutoAnswer = true;

            return statusCheckAutoAnswer;
        }

        #region Методы по получению Статуса из РИМС
        // Task for Plinq.
        public static void MyTask(object arg)
        {
            string ak = (string)arg;
            logger.Info($"MyTask: CurrentId {Task.CurrentId} with ManagedThreadId {Thread.CurrentThread.ManagedThreadId} запущен, ак пользователя {ak}");
            GetOutState2(ak);
            logger.Info($"MyTask: CurrentId {Task.CurrentId} завершен.");
        }

        // доступ к RIMS в потоках
        static void GetOutState2(string acount)
        {
            // NTLM Secured URL
            var uri = new Uri(configiniP.URI3);

            var httpRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpRequest.Timeout = 30000;
            httpRequest.Method = "POST";
            httpRequest.Accept = "application/json";

            if (!string.IsNullOrEmpty(configiniP.login) & !string.IsNullOrEmpty(configiniP.password))
            {
                NetworkCredential credential =
                    new NetworkCredential(configiniP.login, plainText, configiniP.domen);
                httpRequest.Credentials = credential;
            }
            else
            {
                httpRequest.UseDefaultCredentials = true;
            }
            httpRequest.ContentType = "application/json";

            GetOutOfOffice getOutOf = new GetOutOfOffice()
            {
                Account = acount,
                Domain = "IIE.CP"
            };

            string jsonz = JsonConvert.SerializeObject(getOutOf);
            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(jsonz);
            }
            try
            {
                var result = "";
                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }

                //Console.WriteLine(httpResponse.StatusCode);
                //Console.WriteLine(result);
                var status = JsonConvert.DeserializeObject<GetOutOfResult>(result);

                logger.Info($"Статус проверки: CurrentId:{Task.CurrentId}, ManagedThreadId:{Thread.CurrentThread.ManagedThreadId}");
                logger.Info($"Статус автоответа {acount}: {status.Data?.Enabled} DateS&E:{status.Data?.DateStart}-{status.Data?.DateEnd} ");

                bool? stE = status.Data?.Enabled;
                if (stE.HasValue)
                {
                    if (stE.Value)
                    {
                        using (DataModelContext context = new DataModelContext())
                        {
                            var user1 = context.Datausers.FirstOrDefault(u => u.AccountName == acount);
                            if (user1 != null)
                            {
                                user1.AnswerId = 4;
                                context.SaveChanges();
                                logger.Info($"Статус автоответа {acount}: 4 и он прописан в DB");
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                logger.Info($"CurrentId:{Task.CurrentId}: " + ex.Message);
            }
            //return groupID_tmp;
        }

        #endregion Методы по получению Статуса из РИМС


        // Обработаем пользователей (исключения ) , что в файле config.ini
        static void ExUser()
        {
            exlistusers = ExUserF();
            if (configiniP.UserNameEx.Count() > 0)
            {
                foreach (var item in configiniP.UserNameEx)
                {
                    string item2 = item.Trim();
                    exlistusers.Add(item2.Count() > 50 ? item2.Substring(49) : item2);
                }
            }

            using (DataModelContext context = new DataModelContext())
            {
                foreach (string acount in exlistusers)
                {
                    var euser = context.Datausers.FirstOrDefault(u => u.AccountName == acount);
                    if (euser != null)
                    {
                        euser.AnswerId = 4;
                        context.SaveChanges();
                        Console.WriteLine($"** Исключение из INI ** установим статус автоответа для {acount}: 4 и запишем в DB");
                        logger.Info($"** Исключение из INI ** установим статус автоответа для {acount}: 4 и запишем в DB");
                    }

                }
            }
        }
        // Обработаем пользователей (исключения ) , что в файле exusersfile.txt
        static List<string> ExUserF()
        {
            string s;
            string[] words;
            List<string> exlius = new List<string>();
            if (File.Exists(configiniP.exUsersFile))
            {
                using (StreamReader streamReader = new StreamReader(configiniP.exUsersFile, Encoding.GetEncoding(1251)))
                {
                    while ((s = streamReader.ReadLine()) != null)
                    {
                        if (Regex.IsMatch(s, @"^;") | string.IsNullOrEmpty(s))
                            continue;
                        s = s.Trim();
                        if (Regex.IsMatch(s, " "))
                        {
                            words = s.Split(new char[] { ' ' });
                            foreach (var w in words)
                            {
                                exlius.Add(w.Count() > 50 ? w.Substring(49) : w);

                            }
                            continue;
                        }

                        exlius.Add(s.Count() > 50 ? s.Substring(49) : s);
                    }
                }
            }
            return exlius;
        }




        // Запись в Exchange через запросы в РИМС
        static void SaveInExchange()
        {
            List<Qtable> qtables = new List<Qtable>();

            // Join new table
            using (DataModelContext context = new DataModelContext())
            {
                var datausers = from d in context.Datausers
                                join l in context.Leaves on d.LeaveId equals l.Id
                                join a in context.Answers on d.AnswerId equals a.Id
                                select new Qtable()
                                {
                                    AccountName = d.AccountName,
                                    LeaveName = l.LeaveType,
                                    LeaveStart = d.LeaveStart,
                                    LeaveEnd = d.LeaveEnd,
                                    Uid = d.FimSyncKey,
                                    LastName = d.LastName,
                                    FirstName = d.FirstName,
                                    MiddleName = d.MiddleName,
                                    AnswerName = a.AnswerType
                                };

                int count0 = 1;
                foreach (var ds in datausers)
                {
                    // для проверки только 2 записи
                    //if (count0 > 4) break;

                    qtables.Add(ds);
                    count0++;
                }
            }

            // Ограничение по core
            ParallelOptions options = new ParallelOptions();
            /*if (Environment.ProcessorCount > 4)
                options.MaxDegreeOfParallelism = Environment.ProcessorCount < 10 ? 4 : 5;
            else
                options.MaxDegreeOfParallelism = 1;*/
            // Данные о количестве потоков мы возьмем из config.ini
            options.MaxDegreeOfParallelism = configiniP.ThreadsInt;

            Console.WriteLine($"Начинаем запись статусов для пользователей в Exchange: {DateTime.Now}");
            Console.WriteLine($"Количество используемых потоков: {options.MaxDegreeOfParallelism}.");
            logger.Info($"\r\n !!!! -------- Начинаем запись статусов для пользователей в Exchange ------------!!!!:{DateTime.Now}");
            logger.Info($"Количество используемых потоков: {options.MaxDegreeOfParallelism}.");
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            qtables.Where(q => q.AnswerName != "AU" && q.AnswerName != "AG").AsParallel().WithDegreeOfParallelism(options.MaxDegreeOfParallelism).ForAll(q => { MyTaskSet(q); });
            // | q.AnswerName != "AG"
            stopwatch.Stop();
            TimeSpan stopwatchElapsed = stopwatch.Elapsed;
            var milsec = Convert.ToInt32(stopwatchElapsed.TotalMilliseconds);
            var sec = milsec / 1000;
            var ts = TimeSpan.FromSeconds(sec);

            elapsedTimeSetAutoAnswer = $"{ts.Hours}ч:{ts.Minutes}м:{ts.Seconds}с";
            Console.WriteLine($"Затраченное время: {ts.Hours}ч:{ts.Minutes}м:{ts.Seconds}с");
            Console.WriteLine($"Всего пользователей попавших в обработку: {qtables.Count}");
            Console.WriteLine($"Основной поток завершен.{DateTime.Now}\r\n");
            logger.Info($"Затраченное время: {ts.Hours}ч:{ts.Minutes}м:{ts.Seconds}с");
            logger.Info($"Всего пользователей попавших в обработку: {qtables.Count}");
            logger.Info($"Основной поток завершен.{DateTime.Now}");

        }

        #region Метод отправки запросов в РИМС (потоки)
        static void MyTaskSet(object arg)
        {
            Qtable qtable = (Qtable)arg;

            var uri = new Uri(configiniP.URI4);

            var httpRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpRequest.Timeout = 30000;
            httpRequest.Method = "POST";
            httpRequest.Accept = "application/json";

            if (!string.IsNullOrEmpty(configiniP.login) & !string.IsNullOrEmpty(plainText))
            {
                NetworkCredential credential =
                    new NetworkCredential(configiniP.login, plainText, configiniP.domen);
                httpRequest.Credentials = credential;
            }
            else
            {
                httpRequest.UseDefaultCredentials = true;
            }
            httpRequest.ContentType = "application/json";


            // Структура формирования автоОтвета
            OutOfOffice ofOffice = new OutOfOffice
            {
                leaveName = qtable.LeaveName,
                DateStart = (DateTime)qtable.LeaveStart,
                DateEnd = (DateTime)qtable.LeaveEnd,
                FIO = $"{qtable.LastName} {qtable.FirstName} {qtable.MiddleName}"
            };
            logger.Info($"Автоответ для пользователя: {ofOffice.As()}");
            string setStatusUser = ofOffice.As();

            // JSON формирование
            SetOutOfOffice setOutOf = new SetOutOfOffice()
            {
                InternalContent = setStatusUser,
                ExternalContent = setStatusUser,
                DateStart = (DateTime)qtable.LeaveStart,
                DateEnd = (DateTime)qtable.LeaveEnd,
                Uid = qtable.Uid,
                Account = qtable.AccountName,
                Disable = false

            };

            string jsonz = JsonConvert.SerializeObject(setOutOf);
            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(jsonz);
            }

            // проверочный статус
            bool sts = false;
            int count = 0;
            do
            {
                try
                {
                    count++;
                    var result = "";
                    var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        result = streamReader.ReadToEnd();
                    }

                    var status = JsonConvert.DeserializeObject<SetOutOffResult>(result);
                    sts = status.Status;
                    if (status.Status)
                    {
                        //logger.Info($"Для {qtable.AccountName} добавлен автоматический ответ в почте" + Environment.NewLine);
                        using (DataModelContext context = new DataModelContext())
                        {
                            var user1 = context.Datausers.FirstOrDefault(u => u.AccountName == qtable.AccountName);
                            if (user1 != null & user1.AnswerId != 3)
                            {
                                user1.AnswerId = 3;
                                context.SaveChanges();
                            }
                        }
                    }
                    else
                    {
                        using (DataModelContext context = new DataModelContext())
                        {
                            var user1 = context.Datausers.FirstOrDefault(u => u.AccountName == qtable.AccountName);
                            if (user1 != null & user1.AnswerId != 2)
                            {
                                user1.AnswerId = 2;
                                context.SaveChanges();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Info($"Исключение: {ex.Message} - попытка N{count} для пользователя {qtable.AccountName}. CurrentId {Task.CurrentId}, ManagedThreadId:{Thread.CurrentThread.ManagedThreadId}");
                    using (DataModelContext context = new DataModelContext())
                    {
                        var user1 = context.Datausers.FirstOrDefault(u => u.AccountName == qtable.AccountName);
                        if (user1 != null & user1.AnswerId != 2)
                        {
                            user1.AnswerId = 2;
                            context.SaveChanges();
                        }
                    }
                }
            } while (!sts & count < 3);

            logger.Info($"Пользователю {qtable.AccountName} прописан автоответ: {sts}.");
            logger.Info($"MyTaskSet: CurrentId {Task.CurrentId}, ManagedThreadId:{Thread.CurrentThread.ManagedThreadId} завершен.");
        }

        #endregion

    }



    // интерфейс для конигурационного файла config.ini
    public interface IMySettings
    {
        string BDName { get; }
        string Address { get; }
        // имена пользователей которым статус не будет изменен
        IEnumerable<string> UserNameEx { get; }
        string domen { get; }
        string login { get; }
        string password { get; }
        // URI1 - для проверки связи
        string URI1 { get; }
        //URI2 - для получения данных о временно отсутсвующих
        string URI2 { get; }
        //URI3 - для получения статуса Автоответа от РИМСа
        string URI3 { get; }
        // URI4 - для отправки запросов к РИМС на установку автоответов в Exchange
        string URI4 { get; }
        // количество потоков мы будем использовать
        int ThreadsInt { get; }
        // Email
        string UserFrom { get; }
        string UserTo { get; }
        string UserCc { get; }
        string ServerPost { get; }
        int Port { get; }
        // названия файла для исключений пользователей
        string exUsersFile { get; }
        // количество лог файлов в директории
        int QuantityLogs { get; }

    }

}
