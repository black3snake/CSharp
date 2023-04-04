using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WForRestGet.DataModel;
using WForRestGet.Properties;
using Microsoft.EntityFrameworkCore;

namespace WForRestGet
{
    public partial class Form1 : Form
    {
        Point lastPoint;
        public string sServiceUser { get; set; }
        public string sServicePassword { get; set; }
        public string sServiceDomain { get; set; }
        public bool statusRimsTest = false;
        public bool statusDBTest = false;

        Dictionary<string, string> DLeaveType = new Dictionary<string, string>()
        {
            {"VC","Отпуск"},
            {"SL","Больничный"},
            {"BT","Командировка"},
            {"DV","Декретный отпуск"}
        };
        Dictionary<string, int> DLeaveType2C = new Dictionary<string, int>()
        {
            {"VC", 1},
            {"SL", 2},
            {"BT", 3},
            {"DV", 4}
        };

        List<Datauser> Datauserslist = new List<Datauser>();

        // Структура формирование эл. ответа в Exchange
        struct OutOfOffice
        {
            public string leaveName { get; set;}
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
            
            public string As() {
                
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
        


    public Form1()
        {
            InitializeComponent();

            // Описание событий
            picBoxXclose.MouseEnter += (s, a) => { picBoxXclose.BackgroundImage = Resources.x_red; };
            picBoxXclose.MouseLeave += (s, a) => { picBoxXclose.BackgroundImage = Resources.square_x_icon_215388; };
            picBoxXclose.Click += (s, a) => { this.Close(); };

            panelTop.MouseDown += (s, a) => { lastPoint = new Point(a.X, a.Y); };
            panelTop.MouseMove += (s, a) => { 
                if(a.Button == MouseButtons.Left)
                {
                    this.Left += a.X - lastPoint.X;
                    this.Top += a.Y - lastPoint.Y;
                }
            };
           
            #region ListView - Инициализация
           
            ColumnHeader header1, header2, header3, header4, header5, header6;
            header1 = new ColumnHeader();  // 
            header2 = new ColumnHeader();
            header3 = new ColumnHeader();
            header4 = new ColumnHeader();
            header5 = new ColumnHeader();
            header6 = new ColumnHeader();

            header1.Text = "AccountName";
            header1.TextAlign = HorizontalAlignment.Left;
            header1.Width = 90;

            header2.Text = "FullName";
            header2.TextAlign = HorizontalAlignment.Left;
            header2.Width = 180;

            header3.Text = "City";
            header3.TextAlign = HorizontalAlignment.Left;
            header3.Width = 110;
            
            header4.Text = "LT";
            header4.TextAlign = HorizontalAlignment.Left;
            header4.Width = 110;
            
            header5.Text = "LStart";
            header5.TextAlign = HorizontalAlignment.Left;
            header5.Width = 70;

            header6.Text = "LEnd";
            header6.TextAlign = HorizontalAlignment.Left;
            header6.Width = 70;


            listView1.Columns.Add(header1);
            listView1.Columns.Add(header2);
            listView1.Columns.Add(header3);
            listView1.Columns.Add(header4);
            listView1.Columns.Add(header5);
            listView1.Columns.Add(header6);
            //listView1.Columns[0].Width = 200;
            listView1.View = View.Details;

            #endregion

            #region ListView - События выпадающего меню
            listView1.MouseUp += (s, a) => {
                if (a.Button == MouseButtons.Right)
                {
                    contextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right);
                }
            };

            toolStripMenuItem1.Click +=(s, a) => {
                foreach (ListViewItem item in listView1.Items)
                {
                    item.Selected = true;
                }
            };

            toolStripMenuItem2.Click += (s, a) =>
            {
                if (listView1.SelectedItems.Count > 1)
                {
                    string manyCl = "";
                    foreach (ListViewItem lV in listView1.SelectedItems)
                    {
                        if (lV.SubItems.Count > 1)
                        {
                            manyCl += lV.Text + " : " + lV.SubItems[1].Text + " : " +
                            lV.SubItems[2].Text + " : " + lV.SubItems[3].Text + " : " + lV.SubItems[4].Text + " - " +
                            lV.SubItems[5].Text + Environment.NewLine;
                        }
                        else
                        {
                            manyCl += lV.Text + " : " + lV.SubItems[1].Text + " : " +
                            lV.SubItems[2].Text + " : " + lV.SubItems[3].Text + " : " + lV.SubItems[4].Text + " - " +
                            lV.SubItems[5].Text + Environment.NewLine;
                        }
                    }
                    Clipboard.SetText(manyCl);

                }
                else if (listView1.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Ничего не выбрано", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    Clipboard.SetText(listView1.SelectedItems[0].Text);
                }
            };
            #endregion

        }

        // TEST connection RIMS
        private async void btnTestRims_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBLogin.Text) | string.IsNullOrEmpty(txtBPass.Text) | string.IsNullOrEmpty(txtBDomen.Text))
            {
                MessageBox.Show("Заполни все поля Login, Password, Domain!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sServiceUser = txtBLogin.Text.Trim();
            sServicePassword = txtBPass.Text.Trim();
            sServiceDomain = txtBDomen.Text.Trim();
            // NTLM Secured URL
            var uri = new Uri("https://xx/Export/api/Person/EmplAO?%24filter=Domain%20eq%20'IIE.Cp'%20&%24top=1");
            // Create a new Credential
            var credentialsCache = new CredentialCache();
            credentialsCache.Add(uri, "NTLM", new NetworkCredential(
                sServiceUser, sServicePassword, sServiceDomain));

            var handler = new HttpClientHandler() { Credentials = credentialsCache, PreAuthenticate = true };
            var httpClient = new HttpClient(handler) { Timeout = new TimeSpan(0, 0, 10) };

            var response = await httpClient.GetAsync(uri);

            var result = await response.Content.ReadAsStringAsync();

            //txtBoxConsole.AppendText(result + Environment.NewLine);
            //Console.WriteLine(result);
            if(response.IsSuccessStatusCode)
            {
                txtBoxConsole.AppendText($"Статус проверки REST Get: {response.ReasonPhrase}" + Environment.NewLine);
                picBoxTestRIMS.BackgroundImage = Resources.Ok_27007;
                statusRimsTest = true;
            } else
            {
                txtBoxConsole.AppendText($"Статус проверки REST Get: {response.ReasonPhrase}" + Environment.NewLine);
                picBoxTestRIMS.BackgroundImage = Resources.Close_2_26986;
                statusRimsTest = false;
            }

        }

        // TEST connection MS SQL
        private void btnTestDB_Click(object sender, EventArgs e)
        {
            using (DataModelContext context = new DataModelContext())
            {
                var lists = context.Datausers.ToList();
                var countD = lists.LongCount();
                txtBoxConsole.AppendText($"Quantity items in DB Datausers: {countD}" + Environment.NewLine);
                if(countD > 0)
                {
                    picBoxTestDB.BackgroundImage = Resources.Ok_27007;
                    statusDBTest = true;
                } else
                {
                    picBoxTestDB.BackgroundImage = Resources.Close_2_26986;
                    statusDBTest= false;
                }
            }

            using (DataModelContext contextOld = new DataModelContext())
            {
                var usersOld = contextOld.Datausers.Where(p => p.LeaveEnd <= DateTime.Now.AddHours(-24)).ToList();
                foreach (var uOld in usersOld)
                {
                    txtBoxConsole.AppendText($"{uOld.LastName} {uOld.FirstName} {uOld.MiddleName} :: {uOld.LeaveStart} -  {uOld.LeaveEnd}" + Environment.NewLine);
                    contextOld.Datausers.Remove(uOld);
                    contextOld.SaveChanges();

                }
                txtBoxConsole.AppendText($"Quantity olds: {usersOld.Count}" + Environment.NewLine);
            }


        }

        //Eye hidden
        private void picBoxEye_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtBPass.Text) && txtBPass.UseSystemPasswordChar)
            {
                txtBPass.UseSystemPasswordChar = false;
                picBoxEye.BackgroundImage = Resources.eye_icon_224636;
                txtBPass.Focus();
            }
            else
            {
                txtBPass.UseSystemPasswordChar = true;
                picBoxEye.BackgroundImage = Resources.slash_eye_icon_224538;
                txtBPass.Focus();
            }
        }

        // Заберем данные с РИМС
        private async void btnZRims_Click(object sender, EventArgs e)
        {
            if(!statusRimsTest)
            {
                MessageBox.Show("Нажми на кнопку Test RIMS\r\nПроверь Логин и Пароль\r\n Проверь связь", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            sServiceUser = txtBLogin.Text.Trim();
            sServicePassword = txtBPass.Text.Trim();
            sServiceDomain = txtBDomen.Text.Trim();
            // NTLM Secured URL
            // без ограничения
            var uri = new Uri("https://xx/Export/api/Person/Odata?%24filter=Domain%20eq%20'IIE.Cp'%20and%20Disabled%20eq%20false%20and%20LeaveType%20ne%20null");
            // Create a new Credential
            var credentialsCache = new CredentialCache();
            credentialsCache.Add(uri, "NTLM", new NetworkCredential(
                sServiceUser, sServicePassword, sServiceDomain));

            var handler = new HttpClientHandler() { Credentials = credentialsCache, PreAuthenticate = true };
            var httpClient = new HttpClient(handler) { Timeout = new TimeSpan(0, 0, 10) };
            var response = await httpClient.GetAsync(uri);
            var result = await response.Content.ReadAsStringAsync();
            #region Test array
            // Test array
            //JArray jArray = JArray.Parse(result);
            #endregion
            var statusL = JsonConvert.DeserializeObject<List<RimsZaprosOutput>>(result);

            double count = 0;
            txtBoxConsole.AppendText($"Начинаем заполнять список..");
            foreach (var item in statusL)
            {
                count++;
                if ( (count % 100) == 0)
                {
                    txtBoxConsole.AppendText($"..{(int)count}");
                }
                ListViewItem viewItem = new ListViewItem(item.AccountName);
                string FIO = $"{item.LastName} {item.FirstName} {item.MiddleName}";
                viewItem.SubItems.Add(FIO);
                viewItem.SubItems.Add(item.City);
                #region old Dict
                /*string LeaveType = "";
                foreach (var DLT in DLeaveType)
                {
                    if(DLT.Key == item.LeaveType)
                    {
                        LeaveType = DLT.Value;
                        break;
                    }
                }*/
                
                //viewItem.SubItems.Add(LeaveType);
                #endregion
                viewItem.SubItems.Add(DLeaveType[item.LeaveType]);
                viewItem.SubItems.Add(item.LeaveStart.ToString().Remove(item.LeaveStart.ToString().IndexOf(" ")));
                viewItem.SubItems.Add(item.LeaveEnd.ToString().Remove(item.LeaveEnd.ToString().IndexOf(" ")));
                listView1.Items.Add(viewItem);
               
                Datauser userObject = new Datauser()
                {
                    FimSyncKey = item.FimSyncKey,
                    AccountId = item.AccountId,
                    AccountName = item.AccountName,
                    LastName = item.LastName,
                    FirstName = item.FirstName,
                    MiddleName = item.MiddleName,
                    EmployeeNumber = item.EmployeeNumber,
                    Birthday = item.Birthday,
                    CompanyName = item.CompanyName,
                    DepartmentName = item.DepartmentName,
                    JobTitle = item.JobTitle,
                    DateIn = item.DateIn,
                    LeaveId = DLeaveType2C[item.LeaveType],
                    LeaveStart = item.LeaveStart,
                    LeaveEnd = item.LeaveEnd,
                    City = item.City,
                    Phone = item.Phone,
                    Email = item.Email,
                    Disabled = item.DisabledDomain
                    
                };
                Datauserslist.Add(userObject);

            }

            txtBoxConsole.AppendText(Environment.NewLine);
            txtBoxConsole.AppendText($"Пользователей в списке: {listView1.Items.Count}" + Environment.NewLine);


        }

        // Save in DataBase
        private void btnSaveDB_Click(object sender, EventArgs e)
        {
            if (!statusDBTest & listView1.Items.Count<1)
            { 
                MessageBox.Show("Нажми на кнопку Test DB\r\nИ убедись , что появилась зеленая галочка\r\nПроверь связь", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            long count_db = 0;
            using (DataModelContext countDB = new DataModelContext())
            {
                var lists = countDB.Datausers.ToList();
                count_db = lists.LongCount();
            }


            using (DataModelContext context = new DataModelContext())
            {

                if(count_db > 1000)
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
                            context.SaveChanges();
                        }

                    }

                } else
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
                            context.SaveChanges();
                        }
                    }

                }

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
                var countD = lists.LongCount();
                txtBoxConsole.AppendText($"Quantity items after Save in DB: {countD}" + Environment.NewLine);
            }


        }

        // Отправка запросов REST POST в РИМС для установки статуса в EXchange
        private async void btnExToRims_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBLogin.Text) | string.IsNullOrEmpty(txtBPass.Text) | string.IsNullOrEmpty(txtBDomen.Text)) {
                MessageBox.Show("Заполни все поля Login, Password, Domain!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            #region start for RIMS
            sServiceUser = txtBLogin.Text.Trim();
            sServicePassword = txtBPass.Text.Trim();
            sServiceDomain = txtBDomen.Text.Trim();
            
            List<Qtable> qtables = new List<Qtable>();
            var uri = new Uri("https://k.hq.root.ad/api/Exchange/SetOutOfOffice");

            var credentialsCache = new CredentialCache();
            credentialsCache.Add(uri, "NTLM", new NetworkCredential(
                sServiceUser, sServicePassword, sServiceDomain));

            var handler = new HttpClientHandler() { Credentials = credentialsCache, PreAuthenticate = true };
            var httpClient = new HttpClient(handler) { Timeout = new TimeSpan(0, 0, 30) };
            #endregion

            // Join new table
            using (DataModelContext context = new DataModelContext())
            {
                var datausers = context.Datausers.Join(context.Leaves,
                    d => d.LeaveId,
                    l => l.Id,
                    (d, l) => new Qtable()
                    {
                        AccountName = d.AccountName,
                        LeaveName = l.LeaveType,
                        LeaveStart = d.LeaveStart,
                        LeaveEnd = d.LeaveEnd,
                        Uid = d.FimSyncKey,
                        LastName = d.LastName,
                        FirstName = d.FirstName,
                        MiddleName = d.MiddleName
                    });

                int count0 = 1;
                foreach (var ds in datausers)
                {
                    if (count0 > 2)
                    {
                        break;
                    }
                    qtables.Add(ds);
                    count0++;
                }
            }
            
            foreach (var qt in qtables)
            {
                // Структура формирования автоОтвета
                OutOfOffice ofOffice = new OutOfOffice
                {
                    leaveName = qt.LeaveName,
                    DateStart = (DateTime)qt.LeaveStart,
                    DateEnd = (DateTime)qt.LeaveEnd,
                    FIO = $"{qt.LastName} {qt.FirstName} {qt.MiddleName}"
                };
                txtBoxConsole.AppendText($"Статус: {ofOffice.As()}" + Environment.NewLine);
                string setStatusUser = ofOffice.As();

                #region OLDz
                // REST POST
                //var uri = new Uri("https://k.hq.root.ad/api/Exchange/SetOutOfOffice");
                // Create a new Credential
                /*var credentialsCache = new CredentialCache();
                credentialsCache.Add(uri, "NTLM", new NetworkCredential(
                    sServiceUser, sServicePassword, sServiceDomain));

                var handler = new HttpClientHandler() { Credentials = credentialsCache, PreAuthenticate = true };
                var httpClient = new HttpClient(handler); //{ Timeout = new TimeSpan(0, 0, 10) };
*/
                #endregion

                SetOutOfOffice setOutOf = new SetOutOfOffice()
                {
                    InternalContent = setStatusUser,
                    ExternalContent = setStatusUser,
                    DateStart = (DateTime)qt.LeaveStart,
                    DateEnd = (DateTime)qt.LeaveEnd,
                    Uid = qt.Uid,
                    Account = qt.AccountName,
                    Disable = false

                };
                var content = JsonConvert.SerializeObject(setOutOf);
                var data = new StringContent(content, Encoding.UTF8, "application/json");

                // проверочный статус
                bool sts = false;
                int count = 0;
                do
                {
                    try
                    {
                        count++;
                        var response = await httpClient.PostAsync(uri, data);
                        response.EnsureSuccessStatusCode();
                        var json = await response.Content.ReadAsStringAsync();
                        var status = JsonConvert.DeserializeObject<SetOutOffResult>(json);
                        txtBoxConsole.AppendText($"Статус добавления: {status.Status} Message:{status.Message}" + Environment.NewLine);
                        sts = status.Status;
                        if(status.Status)
                        {
                            txtBoxConsole.AppendText($"Для {qt.AccountName} добавлен автоматический ответ в почте" + Environment.NewLine);
                        }
                    }
                    catch (Exception ex)
                    {
                        txtBoxConsole.AppendText($"Исключение: {ex.Message} - попытка N{count}" + Environment.NewLine);
                        await Task.Delay(2000);
                    }
                } while (!sts & count < 3);
                

            }







        }


    }
}
