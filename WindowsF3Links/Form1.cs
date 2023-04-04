using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsF3Links.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace WindowsF3Links
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
            if (string.IsNullOrEmpty(textBox1.Text)) { textBox1.Text = "Empty"; }
            textBox1.GotFocus += (s, a) => { if (textBox1.Text == "Empty") { textBox1.Text = ""; } };
            textBox1.LostFocus += (s, a) =>
            {
                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    textBox1.Text = "Empty";
                }
                else
                {
                    if (File.Exists(textBox1.Text) && Regex.IsMatch(textBox1.Text, @"^.*.xls.?$"))
                    {
                        pictureBox1.BackgroundImage = Resources.Ok_270071;
                    }
                    else
                    {
                        pictureBox1.BackgroundImage = Resources.Close_2_26986;
                    }
                }

            };

            if (string.IsNullOrEmpty(textBox2.Text)) { textBox2.Text = "Empty"; }
            textBox2.GotFocus += (s, a) => { if (textBox2.Text == "Empty") { textBox2.Text = ""; } };
            textBox2.LostFocus += (s, a) => { if (string.IsNullOrEmpty(textBox2.Text)) { textBox2.Text = "Empty"; } };

            if (string.IsNullOrEmpty(textBox3.Text)) { textBox3.Text = "Empty"; }
            textBox3.GotFocus += (s, a) => { if (textBox3.Text == "Empty") { textBox3.Text = ""; } };
            textBox3.LostFocus += (s, a) => { if (string.IsNullOrEmpty(textBox3.Text)) { textBox3.Text = "Empty"; } };

            if (string.IsNullOrEmpty(textBox5.Text)) { textBox5.Text = "Empty"; }
            textBox5.GotFocus += (s, a) => { if (textBox5.Text == "Empty") { textBox5.Text = ""; } };
            textBox5.LostFocus += (s, a) => { if (string.IsNullOrEmpty(textBox5.Text)) { textBox5.Text = "Empty"; } };

            if (string.IsNullOrEmpty(textBox4.Text)) { textBox4.Text = "Empty"; }
            textBox4.GotFocus += (s, a) => { if (textBox4.Text == "Empty") { textBox4.Text = ""; } };
            textBox4.LostFocus += (s, a) => { if (string.IsNullOrEmpty(textBox4.Text)) { textBox4.Text = "Empty"; } };

            if (string.IsNullOrEmpty(textBox7.Text)) { textBox7.Text = "Empty"; }
            textBox7.GotFocus += (s, a) => { if (textBox7.Text == "Empty") { textBox7.Text = ""; } };
            textBox7.LostFocus += (s, a) => { if (string.IsNullOrEmpty(textBox7.Text)) { textBox7.Text = "Empty"; } };

            if (string.IsNullOrEmpty(textBox6.Text)) { textBox6.Text = "Empty"; }
            textBox6.GotFocus += (s, a) => { if (textBox6.Text == "Empty") { textBox6.Text = ""; } };
            textBox6.LostFocus += (s, a) => { if (string.IsNullOrEmpty(textBox6.Text)) { textBox6.Text = "Empty"; } };

            // Проверка реестра
            RegistryKey keyS = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Security", true);
            if(Convert.ToInt32(keyS.GetValue("AccessVBOM")) == 0)
            {
                pictureBox2.BackgroundImage = Resources.Close_2_26986;
                pictureBox2.Click += (s, a) => { 
                    if(MessageBox.Show("Внеси изменения в реестр?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        keyS.SetValue("AccessVBOM", 1);
                        RegistryKey keySM = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Security", false);
                        if (Convert.ToInt32(keySM.GetValue("AccessVBOM")) == 1)
                        {
                            pictureBox2.BackgroundImage = Resources.Ok_270071;
                        }
                    } 
                };
            } else
            {
                pictureBox2.BackgroundImage = Resources.Ok_270071;
            }
          /*  pictureBox2.MouseLeave += (s, a) =>
            {
                RegistryKey keySM = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Security", false);
                if(Convert.ToInt32(keySM.GetValue("AccessVBOM")) == 1) {
                    pictureBox2.BackgroundImage = Resources.Ok_270071;
                }
            };*/


        }
        // кнопка Clear
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "Empty") { textBox1.Text = "Empty"; }
            if (textBox2.Text != "Empty") { textBox2.Text = "Empty"; }
            if (textBox3.Text != "Empty") { textBox3.Text = "Empty"; }
            if (textBox5.Text != "Empty") { textBox5.Text = "Empty"; }
            if (textBox4.Text != "Empty") { textBox4.Text = "Empty"; }
            if (textBox7.Text != "Empty") { textBox7.Text = "Empty"; }
            if (textBox6.Text != "Empty") { textBox6.Text = "Empty"; }

            if (!string.IsNullOrEmpty(textBox8.Text)) { textBox8.Text = ""; }

            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;

            pictureBox1.BackgroundImage = null;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            if(!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked)
            {
                MessageBox.Show("Необходимо поставить галочку где заменять", "Забыли?", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
                return;
            }

            WorkAs();

        }

        private async void WorkAs()
        {
            await Task.Run(() =>
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                try
                {
                    if (textBox2.Text != "Empty" && textBox3.Text != "Empty")
                    {
                        dic.Add(ZamenaSL(textBox2.Text.Trim()), ZamenaSL(textBox3.Text.Trim()));
                    }
                    if (textBox5.Text != "Empty" && textBox4.Text != "Empty")
                    {
                        dic.Add(ZamenaSL(textBox5.Text.Trim()), ZamenaSL(textBox4.Text.Trim()));
                    }
                    if (textBox7.Text != "Empty" && textBox6.Text != "Empty")
                    {
                        dic.Add(ZamenaSL(textBox7.Text.Trim()), ZamenaSL(textBox6.Text.Trim()));
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Исключение", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                // Main
                if (checkBox1.Checked || checkBox2.Checked)
                {
                    //Объявляем переменную приложения Excel 
                    Excel.Application app = null;
                    Excel.Workbooks workbooks = null;
                    Excel.Workbook workbook = null;
                    Excel.Sheets worksheets = null;

                    try
                    {
                        app = new Excel.Application();
                        workbooks = app.Workbooks;
                        workbook = workbooks.Open(textBox1.Text);
                        worksheets = workbook.Sheets;

                        app.AskToUpdateLinks = false;
                        app.DisplayAlerts = false;

                        // 1. Первая часть обработки xls именно в макросах
                        if (checkBox1.Checked)
                        {
                            Invoke(new System.Action(() =>
                            {
                                textBox8.AppendText($"Начинаеи обработку Макросов в файле {textBox1.Text}" + Environment.NewLine);
                            }));
                            //textBox8.Text += $"Начинаеи обработку Макросов в файле {textBox1.Text}" + Environment.NewLine;
                            WorkMacro(workbook, dic);

                            Invoke(new System.Action(() =>
                            {
                                textBox8.AppendText("Конец обработки Макросов" + Environment.NewLine);
                            }));
                        }

                        // 2. Это будет вторая часть обработки в самих листах
                        //Переберем в цикле все листы которые присутсуют в рабочей книге
                        if (checkBox2.Checked)
                        {
                            Invoke(new System.Action(() =>
                            {
                                textBox8.AppendText($"Начинаеи обработку Листов кол:{worksheets.Count} в файле {textBox1.Text}" + Environment.NewLine);
                            }));
                            //textBox8.Text += $"Начинаеи обработку Листов кол:{worksheets.Count} в файле {textBox1.Text}" + Environment.NewLine;
                            for (int i = 1; i <= worksheets.Count; i++)
                            {
                                WorkSheetM(worksheets[i], dic);
                                Marshal.ReleaseComObject(worksheets[i]);
                            }
                            Invoke(new System.Action(() =>
                            {
                                textBox8.AppendText("Конец обработки Листов" + Environment.NewLine);
                            }));
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Исключение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {

                        // всегда выполняемы блок
                        workbook.Save();
                        // Закроем Excel
                        workbooks.Close();
                        app.Quit();

                        Marshal.ReleaseComObject(worksheets);
                        Marshal.ReleaseComObject(workbook);
                        Marshal.ReleaseComObject(workbooks);
                        Marshal.ReleaseComObject(app);
                        GC.Collect();

                    }
                }

                // 3. Третья часть обработки - внешние ссылки
                // надо закрыть xls и открыть заного только как OpenXML
                if (checkBox3.Checked)
                {
                    Invoke(new System.Action(() =>
                    {
                        textBox8.AppendText($"Начинаеи обработку Внешних ссылок в файле {textBox1.Text}" + Environment.NewLine);
                    }));

                    //textBox8.Text += $"Начинаеи обработку Внешних ссылок в файле {textBox1.Text}" + Environment.NewLine;
                    ReplaceExcel(textBox1.Text, dic);
                    
                    Invoke(new System.Action(() =>
                    {
                        textBox8.AppendText("Конец обработки Внешних ссылок" + Environment.NewLine);
                    }));
                }


            });

        }


        // замена слешей в пути
        string ZamenaSL(string str)
        {
            return str.Replace("\\", "/");
        }

        // замена в макросах
        private void WorkMacro(Excel.Workbook xlBook, Dictionary<string,string> dicM)
        {
            int count = 0;

            if (xlBook.HasVBProject)
            {
                VBProject project = xlBook.VBProject;

                foreach (VBComponent component in project.VBComponents)
                {
                    if (component.Type == vbext_ComponentType.vbext_ct_StdModule ||
                        component.Type == vbext_ComponentType.vbext_ct_ClassModule ||
                            component.Type == vbext_ComponentType.vbext_ct_MSForm)
                    {
                        CodeModule module = component.CodeModule;

                        string[] lines = module.get_Lines(1, module.CountOfLines).Split(
                            new string[] { "\r\n" }, StringSplitOptions.None);

                        for (int i = 0; i < lines.Length; i++)
                        {

                            foreach (var dM in dicM)
                            {
                                if (lines[i].Replace("\\", "/").Contains(dM.Key))
                                {
                                    lines[i] = lines[i].Replace(dM.Key.Replace("/", "\\"), dM.Value.Replace("/", "\\"));
                                    module.ReplaceLine(i + 1, lines[i]);
                                    count++;
                                }
                            }
                        }
                    }
                }
            }
            if (count > 0)
            {
                //Console.WriteLine($"Произведено замен {count} строк");
                Invoke(new System.Action(() =>
                {
                    textBox8.AppendText($"Произведено замен {count} строк" + Environment.NewLine);
                }));
                //textBox8.Text += $"Произведено замен {count} строк" + Environment.NewLine;
                xlBook.Save();
            }
        }

        // замена в листах
        private void WorkSheetM(Worksheet xlSht, Dictionary<string, string> dicM)
        {
            // Создадим Словарь
            var dict = new Dictionary<string, string>();
            int count = 0;

            int iLastRow = xlSht.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            int iLastCol = xlSht.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;


            Range range1 = xlSht.Range[xlSht.Cells[1, 1], xlSht.Cells[iLastRow, iLastCol]];
            if (range1.Count <= 1)
            {
                if (String.IsNullOrEmpty(xlSht.Cells[1, 1].FormulaLocal))
                    //Console.WriteLine($"Пустой Лист по имени: {xlSht.Name}");
                    Invoke(new System.Action(() =>
                    {
                        textBox8.AppendText($"Пустой Лист по имени: {xlSht.Name}" + Environment.NewLine);
                    }));
                return;
            }

            var arrData = (object[,])xlSht.Range[xlSht.Range["A1"], xlSht.Cells[iLastRow, iLastCol]].Formula;
            int iTotalRows = arrData.GetUpperBound(0);
            int iTotalColumns = arrData.GetUpperBound(1);

            for (int i = 1; i <= iTotalRows; i++)
            {
                for (int j = 1; j <= iTotalColumns; j++)
                {
                    //Console.Write($"{arrData[i, j].ToString()} \t");
                    if (!String.IsNullOrEmpty(arrData[i, j].ToString()))
                    {
                        //Console.Write($"{i}{j} : {arrData[i, j].ToString()}");
                        string key = $"{i},{j}";
                        dict.Add(key, arrData[i, j].ToString());
                    }
                }
            }

            foreach (var d in dict)
            {
                foreach (var dM in dicM)
                {
                    if (Regex.IsMatch(d.Value.Replace('\\', '/'), dM.Key, RegexOptions.IgnoreCase))
                    {
                        string newStr = d.Value.Replace("\\", "/").Replace(dM.Key, dM.Value);
                        string[] array_tmp = d.Key.Split(',');
                        int i = Convert.ToInt32(array_tmp[0]);
                        int j = Convert.ToInt32(array_tmp[1]);
                        xlSht.Cells[i, j] = newStr.Replace("/", "\\");
                        count++;
                    }
                }
            }

            if (count > 0)
            {
                //Console.WriteLine($"Лист: {xlSht.Name}");
                //Console.WriteLine($"Количество найденных не пустых элементов: {dict.Count}");
                //Console.WriteLine($"Количество изменненых ссылок: {count}");

                Invoke(new System.Action(() =>
                {
                    textBox8.AppendText($"Лист: {xlSht.Name}" + Environment.NewLine);
                    textBox8.AppendText($"Количество найденных не пустых элементов: {dict.Count}" + Environment.NewLine);
                    textBox8.AppendText($"Количество изменненых ссылок: {count}" + Environment.NewLine);

                }));

            }
            else
            {
                //Console.WriteLine($"Обработка Листа: {xlSht.Name}");
                Invoke(new System.Action(() =>
                {
                    textBox8.AppendText($"Обработка Листа: {xlSht.Name}" + Environment.NewLine);
                }));
                //textBox8.Text += $"Обработка Листа: {xlSht.Name}" + Environment.NewLine;
            }

        }

        // замена во внешних ссылках
        private void ReplaceExcel(string path, Dictionary<string, string> dicM)
        {
            //   logger.Info($"Начал работу с {path}");
            Excel.Application xlApp = new Excel.Application(); //Excel
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
                        foreach (var dM in dicM)
                        {
                            if (Regex.IsMatch(item.ToString().Replace("\\", "/"), dM.Key, RegexOptions.IgnoreCase))
                            {
                                string newPath = item.ToString().Replace("\\", "/").Replace(dM.Key, dM.Value).Replace("/", "\\");
                                try
                                {
                                    xlWB.ChangeLink(item.ToString(), newPath, XlLinkType.xlLinkTypeExcelLinks);
                                    i++;
                                    xlWB.BreakLink((string)item, XlLinkType.xlLinkTypeExcelLinks);
                                }
                                catch (COMException ex)
                                {
                                    //Console.WriteLine(ex.Message);
                                    //xlSht.Cells[iRow, iCol].Formula = newPath;
                                    xlWB.ChangeLink(item.ToString(), newPath, XlLinkType.xlLinkTypeExcelLinks);
                                    i++;
                                    MessageBox.Show(ex.Message, "Исключение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    //                     logger.Error($"COMException {ex} \n {newPath}");
                                }
                                catch (Exception ex)
                                {
                                    //Console.WriteLine(ex.Message);
                                    //                   logger.Error($"Exception {ex} \n {newPath}");
                                    MessageBox.Show(ex.Message, "Исключение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                            }
                        }

                    }
                }
                // Обновляем и разрываем лишние связи
                //WorkbookUpdateLink(xlWB,dicM);

                //Console.Clear();
                //Console.WriteLine("Сохраняем изменения {0}", path);
                Invoke(new System.Action(() =>
                {
                    textBox8.AppendText($"Сохраняем изменения, {textBox1.Text}" + Environment.NewLine);
                }));
                // logger.Info("Сохраняем изменения для {0}", path);

            }
            catch (COMException e)
            {
                //Console.WriteLine($"Файл уже открыт. {path} \n Закройте файл и повторите");
                //logger.Info($"The file is already open.{path} \n Close the file and repeat");
                //logger.Error($"Exception: {e.Message}");
                //Console.WriteLine(e.Message);
                MessageBox.Show(e.Message, "Исключение", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch (Exception e)
            {
                //Console.WriteLine($"Произошла ошибка ");
                //logger.Error($"Exception: {e.Message}");
                //Console.WriteLine(e.Message);
                MessageBox.Show(e.Message, "Исключение", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {

                if (xlWB != null)
                    xlWB.Save();
                xlApp.AskToUpdateLinks = true;
                xlApp.DisplayAlerts = true;
                xlApp.Workbooks.Close();
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWB);
                Marshal.ReleaseComObject(xlApp);

                xlApp = null;
                xlWB = null;
                //_workSheet = null;

                GC.Collect();

                //Console.WriteLine($"Файл обработан, Заменено {i} в {path}");
                Invoke(new System.Action(() =>
                {
                    textBox8.AppendText($"Файл обработан, Заменено {i} в {textBox1.Text}" + Environment.NewLine);
                }));
                //logger.Info($"Заменено {i} в {path}");
            }

        }
        
        private void WorkbookUpdateLink(Workbook xlWB, Dictionary<string, string> dicMA)
        {
            Array links = (Array)xlWB.LinkSources(XlLink.xlExcelLinks);
            if (links != null)
            {
                foreach(var link in links)
                {
                    foreach (var dMA in dicMA)
                    {
                        if(Regex.IsMatch(link.ToString().Replace("\\", "/"), dMA.Key, RegexOptions.IgnoreCase))
                        {
                            xlWB.UpdateLink((string)link, XlLinkType.xlLinkTypeExcelLinks);
                        } else
                        {
                            xlWB.BreakLink((string)link, XlLinkType.xlLinkTypeExcelLinks);
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;

                if (File.Exists(textBox1.Text) && Regex.IsMatch(textBox1.Text, @"^.*.xls.?$"))
                {
                    pictureBox1.BackgroundImage = Resources.Ok_270071;
                }
                else
                {
                    pictureBox1.BackgroundImage = Resources.Close_2_26986;
                }
            }

        }

        
    }
}
