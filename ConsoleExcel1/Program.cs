using Microsoft.Office.Interop.Excel;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;


namespace ConsoleExcel1
{
    internal class Program
    {
        private static string ReplaceString { get; set; }
        private static string FindString { get; set; }
        private static string Path { get; set; }

        static Logger logger = LogManager.GetCurrentClassLogger();
        class fileT
        {
            public string name;
            public long size;
            public fileT(string n, long s) { this.name = n; this.size = s; }
        }
        static void Main(string[] args)
        {

            Console.WriteLine("Укажите путь до каталога, в котором лежит файлы:");
            Console.Write(">");
            Path = Console.ReadLine().Trim();
            //Path = @"D:\2021\xls\test_файл1.xlsx";

            Console.WriteLine("Введите значение, которое нужно найти:");
            Console.Write(">");
            FindString = Console.ReadLine().Replace("\\", "/").Trim();

            Console.WriteLine("Введите значение, на которое нужно заменить:");
            Console.Write(">");
            ReplaceString = Console.ReadLine().Replace("\\", "/").Trim();

            Console.WriteLine($@"Путь к каталогу: {Path}{"\n"}
            значение, которое найти: {FindString}{"\n"}
            значение, на которое заменить: {ReplaceString}{"\n"}
            Потвердите введеные входные параметры:(y - да/ n - нет)");
            Console.Write(">");
            string result = Console.ReadLine();

            Console.WriteLine("\nОсновной поток запущен.");
            logger.Info($"Основной поток запущен.");

            List<fileT> listF = new List<fileT>();
            listF = GetRecursFiles(Path);

            ParallelOptions options = new ParallelOptions();
            // Выделить определенное количество процессорных ядер.
            options.MaxDegreeOfParallelism = Environment.ProcessorCount > 2 ? Environment.ProcessorCount - 1 : 1;
            //
            //options.MaxDegreeOfParallelism = 4; // Попробывать 1 или 2
            Console.WriteLine("Количество логических ядер CPU: {0}", Environment.ProcessorCount);



            if (result.ToLower() != "y")
            {
                logger.Warn("Не получено разрешение на продолжение работы программы -> exit");
                Environment.Exit(0);
            }
            else
            {
                Console.WriteLine("Идет обработка...");
                logger.Info($"Путь к каталогу: {Path}{"\n"} значение ищем: {FindString}{"\n"} значение на которое меняем: {ReplaceString}{"\n"}");

                //Parallel.ForEach(people, options, person => { MyTask(person); });
                //Parallel.ForEach(listF, options, ls => { MyTask(ls); });

                // Вариант 2 используя PLINQ - быстрее и экономней
                listF.AsParallel().WithDegreeOfParallelism(options.MaxDegreeOfParallelism).ForAll(ls => { MyTask(ls);});

                //ReplaceExcel(Path);
            }


            Console.WriteLine(listF.Count);
            Console.WriteLine("\nОсновной поток завершен.");
            logger.Info($"Основной поток завершен\n");

            // Delay
            Console.ReadLine();

        }
        static void MyTask(object arg)
        {
            fileT ft = (fileT)arg;

            Console.WriteLine("MyTask: CurrentId {0} with ManagedThreadId {1} запущен, имя файла {2}", Task.CurrentId,
                Thread.CurrentThread.ManagedThreadId, ft.name);
            logger.Info("MyTask: CurrentId {0} with ManagedThreadId {1} запущен, имя файла {2}", Task.CurrentId,
                Thread.CurrentThread.ManagedThreadId, ft.name);

            ReplaceExcel(ft.name);

            #region Random            
            /*var random = new Random();
            var lowerBound = 2000;
            var upperBound = 20000;
            var rNum = random.Next(lowerBound, upperBound);

            Thread.Sleep(rNum);*/
            #endregion

            Console.WriteLine("MyTask: CurrentId " + Task.CurrentId + " завершен.");
            logger.Info("MyTask: CurrentId " + Task.CurrentId + " завершен.");
        }

        static List<fileT> GetRecursFiles(string path)
        {
            List<fileT> ls22 = new List<fileT>();
            try
            {
                string[] filePaths = Directory.GetFiles(path, "*.xl*", SearchOption.AllDirectories);
                foreach (string filePath in filePaths)
                {
                    FileInfo fileInfo = new FileInfo(filePath);
                    fileT fT = new fileT(filePath, fileInfo.Length);
                    ls22.Add(fT);
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                logger.Error(e.Message);
            }
            return ls22;
        }

        public static void ReplaceExcel(string path)
        {
            logger.Info($"Начал работу с {path}");

            Application xlApp = new Application(); //Excel
            Workbook xlWB = null; //рабочая книга


            int i = 0;
            try
            {
                xlWB = xlApp.Workbooks.OpenXML(path);

                xlApp.DisplayAlerts = false;
                xlApp.AskToUpdateLinks = false;
                Array array = xlWB.LinkSources(XlLink.xlExcelLinks) as Array;

                if (array != null)
                {
                    foreach (var item in array)
                    {
                        if (Regex.IsMatch(item.ToString().Replace("\\", "/"), FindString, RegexOptions.IgnoreCase))
                        {
                            string newPath = item.ToString().Replace("\\", "/").Replace(FindString, ReplaceString).Replace("/", "\\");
                            try
                            {
                                xlWB.ChangeLink(item.ToString(), newPath, XlLinkType.xlLinkTypeExcelLinks);
                                i++;
                            }
                            catch (COMException ex)
                            {
                                //Console.WriteLine(ex.Message);
                                //xlSht.Cells[iRow, iCol].Formula = newPath;
                                xlWB.ChangeLink(item.ToString(), newPath, XlLinkType.xlLinkTypeExcelLinks);
                                i++;
                                logger.Error($"COMException {ex} \n {newPath}");
                            }
                            catch (Exception ex)
                            {
                                //Console.WriteLine(ex.Message);
                                logger.Error($"Exception {ex} \n {newPath}");
                            }

                        }

                    }
                }

                //Console.Clear();
                Console.WriteLine("Сохраняем изменения {0}", path);
                logger.Info("Сохраняем изменения для {0}", path);

            }
            catch (COMException e)
            {
                //Console.WriteLine($"Файл уже открыт. {path} \n Закройте файл и повторите");
                logger.Info($"The file is already open.{path} \n Close the file and repeat");
                logger.Error($"Exception: {e.Message}");
                //Console.WriteLine(e.Message);

            }
            catch (Exception e)
            {
                //Console.WriteLine($"Произошла ошибка ");
                logger.Error($"Exception: {e.Message}");
                //Console.WriteLine(e.Message);
            }
            finally
            {

                if (xlWB != null)
                    xlWB.Save();
                xlApp.AskToUpdateLinks = true;
                xlApp.DisplayAlerts = true;
                xlApp.Workbooks.Close();
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                xlApp = null;
                xlWB = null;
                //_workSheet = null;

                GC.Collect();

                Console.WriteLine($"Файл обработан, Заменено {i} в {path}");
                logger.Info($"Заменено {i} в {path}");
            }
        }
    }
}
