using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using WFAddAndRemoveUsr.Properties;

namespace WFAddAndRemoveUsr
{
    public partial class Form1 : Form
    {
        List<string> users;
        bool fileB = false;
        public string grID { get; set; }
        public string sServiceUser { get; set; }
        public string sServicePassword { get; set; }
        public const string adressRims = "https://k.hq.root.ad";


        public Form1()
        {
            InitializeComponent();

            pictureBox1.LostFocus += (s, a) => {
                if (File.Exists(textBox1.Text) && Regex.IsMatch(textBox1.Text, @"^.*.txt.?$"))
                {
                    pictureBox1.BackgroundImage = Resources.Ok_27007;
                    fileB = true;
                } else
                {
                    pictureBox1.BackgroundImage = Resources.Close_2_26986;
                    fileB = false;
                }
            };

            listView1.MouseUp += (s, a) => {
                if(a.Button == MouseButtons.Right)
                {
                    contextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right);
                }
            };

            btnRIMSzapros.GotFocus += (s, a) =>
            {
                txtBPass.UseSystemPasswordChar = true;
                pictureBox2.BackgroundImage = Resources.slash_eye_icon_224538;

            };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;

                if (File.Exists(textBox1.Text) && Regex.IsMatch(textBox1.Text, @"^.*.txt.?$"))
                {
                    pictureBox1.BackgroundImage = Resources.Ok_27007;
                    fileB = true;
                }
                else
                {
                    pictureBox1.BackgroundImage = Resources.Close_2_26986;
                    fileB = false;
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!fileB) {
                MessageBox.Show("Выбери файл!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //List<Person> users = new List<Person>();
            //List<string> users = new List<string>();
            listView1.Items.Clear();

            if (listView1.Columns.Count > 0)
            {
                foreach (ColumnHeader ch in listView1.Columns)
                {
                    listView1.Columns.Remove(ch);
                }
            }

            ColumnHeader columnHeader1;
            columnHeader1 = new ColumnHeader();

            columnHeader1.Text = "Name";
            columnHeader1.TextAlign = HorizontalAlignment.Left;
            columnHeader1.Width = 460;

            listView1.Columns.Add(columnHeader1);
            listView1.View = View.Details;


            foreach (string str in File.ReadLines(textBox1.Text,Encoding.GetEncoding(1251)))
            {
                ListViewItem itemL = new ListViewItem(str.Trim());
                listView1.Items.Add(itemL);

            }
            label3.Text = $"Count: {listView1.Items.Count}";


        }

        private void btnADtest_Click(object sender, EventArgs e)
        {
            users = new List<string>();
            if (listView1.Columns.Count == 0)
            {
                MessageBox.Show("Сначало нужно нажать \r\n кнопку \"Заполнить\"", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


                foreach (ListViewItem str in listView1.Items)
            {
                users.Add(str.Text);
            }
            
            listView1.Items.Clear();
            
            if(listView1.Columns.Count > 0)
            {
                foreach(ColumnHeader ch in listView1.Columns)
                {
                    listView1.Columns.Remove(ch);
                }
            }

            ColumnHeader header1, header2;
            header1 = new ColumnHeader();
            header2 = new ColumnHeader();

            header1.Text = "ID";
            header1.TextAlign = HorizontalAlignment.Left;
            header1.Width = 180;

            header2.Text = "Name";
            header2.TextAlign = HorizontalAlignment.Left;
            header2.Width = 320;

            listView1.Columns.Add(header1);
            listView1.Columns.Add(header2);
            //listView1.Columns[0].Width = 200;
            listView1.View = View.Details;

            foreach (string item in users)
            {
                Dictionary<string, string> tempD = new Dictionary<string, string>();
                tempD = FindGuid(item);
                
                foreach (var tD in tempD)
                {
                    ListViewItem itemL2 = new ListViewItem(tD.Key);
                    itemL2.SubItems.Add(tD.Value);
                    listView1.Items.Add(itemL2);
                    
                }

            }
            label3.Text = $"Count: {listView1.Items.Count}";

        }
        static Dictionary<string,string> FindGuid(string str)
        {
            Dictionary<string,string> dict= new Dictionary<string,string>();
            // create LDAP connection object  
            DirectoryEntry myLdapConnection = createDirectoryEntry();
            DirectorySearcher search = new DirectorySearcher(myLdapConnection);

          
            if (Regex.IsMatch(str, "^[A-Za-z0-9\\.\\s_]+$"))
            {
                search.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" + str + "))";
            } 
            else if (Regex.IsMatch(str, "^.*@.*$")) 
            {
                search.Filter = "(&(objectClass=user)(objectCategory=person)(cn=" + str.Split('@')[0] + "*" + "))";
            }
            else
            {
                search.Filter = "(&(objectClass=user)(objectCategory=person)(cn=" + str + "*" + "))";
            }

            var resultM = search.FindAll();
            if (resultM.Count > 1)
            {
                foreach (SearchResult rM in resultM)
                {
                    ResultPropertyCollection fields = rM.Properties;
                    string nameCN = fields["cn"][0].ToString();
                    //string nameCN = fields["samaccountname"][0].ToString();
                    
                    string objectGUID = new Guid((System.Byte[])fields["objectguid"][0]).ToString();
                    dict.Add(objectGUID,nameCN);
                    //ValueTuple<string, string> tuple = (objectGUID, nameCN);
                    
                }
            }
            else
            {
                SearchResult result = search.FindOne();  // только один элемент
                if (result != null)
                {
                    ResultPropertyCollection fields = result.Properties;
                    string nameCN = fields["cn"][0].ToString();
                    //string nameCN = fields["samaccountname"][0].ToString();
                    
                    string objectGUID = new Guid((System.Byte[])fields["objectguid"][0]).ToString();
                    dict.Add(objectGUID, nameCN);
                }

            }
            return dict;
        }
        private void btnRemove_Click(object sender, EventArgs e)
        {
            if(listView1.Items.Count > 0)
            {
                try
                {
                    listView1.Items.Remove(listView1.SelectedItems[0]);
                } catch
                {
                    MessageBox.Show("Не могу удалить пустое поле!!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            label3.Text = $"Count: {listView1.Items.Count}";
        }

        static DirectoryEntry createDirectoryEntry()
        {
            // create and return new LDAP connection with desired settings  
            DirectoryEntry ldapConnection = new DirectoryEntry("ms01.iie.cp");
            ldapConnection.Path = "LDAP://DC=iie,DC=cp";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            return ldapConnection;
        }

        private void btnRIMSzapros_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBGroupRims.Text) | string.IsNullOrEmpty(grID))
            {
                MessageBox.Show("Заполни поле с именем группы \r\n И проверь группу в RIMS, чтобы получить ID group", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (listView1.Columns.Count < 2)
            {
                MessageBox.Show("Вы не нажали кнопку \"Проверить в AD\" \r\n и пытаетесь передать в RIMS необработанные данные", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // Переберем все УЗ, что есть в списке и отправим запрос в RIMS
            foreach (ListViewItem item in listView1.Items)
            {
                RimsAddUsersGroup(adressRims, grID, item.Text, item.SubItems[1].Text, sServiceUser, sServicePassword);


            }



        }

        #region Работа с буфером в ListView
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            string[] InsCl = Clipboard.GetText().Split(new char[] {'\r','\n'}, StringSplitOptions.RemoveEmptyEntries) ;
            if(InsCl.Length == 0 )
            {
                MessageBox.Show("Вставляемый буфер пустой :(", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            users = new List<string>();
            listView1.Items.Clear();

            if (listView1.Columns.Count > 0)
            {
                foreach (ColumnHeader ch in listView1.Columns)
                {
                    listView1.Columns.Remove(ch);
                }
            }

            ColumnHeader columnHeader1;
            columnHeader1 = new ColumnHeader();

            columnHeader1.Text = "Name";
            columnHeader1.TextAlign = HorizontalAlignment.Left;
            columnHeader1.Width = 317;

            listView1.Columns.Add(columnHeader1);
            listView1.View = View.Details;


            foreach (string str in InsCl)
            {
                ListViewItem itemL = new ListViewItem(str.Trim());
                listView1.Items.Add(itemL);
                users.Add(str);

            }
            label3.Text = $"Count: {listView1.Items.Count}";

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if(listView1.SelectedItems.Count>1)
            {
                string manyCl = "";
                foreach(ListViewItem lV in listView1.SelectedItems)
                {
                    if(lV.SubItems.Count >1)
                    {
                        manyCl += lV.Text + " : " + lV.SubItems[1].Text + Environment.NewLine;
                    } else
                    {
                        manyCl += lV.Text + Environment.NewLine;
                    }
                }
                Clipboard.SetText(manyCl);

            } else if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Ничего не выбрано", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else
            {
                Clipboard.SetText(listView1.SelectedItems[0].Text);
            }
        }
        // Clear table ListView
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            if (listView1.Columns.Count > 0)
            {
                foreach (ColumnHeader ch in listView1.Columns)
                {
                    listView1.Columns.Remove(ch);
                }
            }
            label3.Text = $"Count: {listView1.Items.Count}";

        }
        #endregion


        #region RIMS методы
        private void btnTestGroup_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBGroupRims.Text))
            {
                MessageBox.Show("Заполни поле с именем группы", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sServiceUser = txtBLogin.Text.Trim();
            sServicePassword = txtBPass.Text.Trim();

            grID = ZaprosToRims(adressRims, txtBGroupRims.Text.Trim(), sServiceUser, sServicePassword);
            if (!string.IsNullOrEmpty(grID))
            {
                pictureBox3.BackgroundImage = Resources.Ok_27007;
            } else
            {
                pictureBox3.BackgroundImage = Resources.Close_2_26986;
            }





        }
        // Получение ID группы от RIMS
        public string ZaprosToRims(string adressRims, string group, string login, string password)
        {
            string url = $"{adressRims}/api/ActiveDirectory/CheckExistsGroup";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";

            httpRequest.Accept = "application/json";

            if (!string.IsNullOrEmpty(sServiceUser) & !string.IsNullOrEmpty(sServicePassword))
            {
                NetworkCredential credential =
                    new NetworkCredential(login, password, "iie.cp");
                httpRequest.Credentials = credential;
            }
            else
            {
                httpRequest.UseDefaultCredentials = true;
            }

            httpRequest.ContentType = "application/json";

            JSONZC jSONZC = new JSONZC
            {
                GroupName = group,
                Domain = "IE.CORP"
            };

            string jsonz = JsonConvert.SerializeObject(jSONZC);

            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(jsonz);
            }

            //await Task.Delay(2000);
            string groupID_tmp = "";
            try
            {
                var result = "";
                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }

                Console.WriteLine(httpResponse.StatusCode);
                Console.WriteLine(result);
                JObject json_result = JsonConvert.DeserializeObject<JObject>(result);

                foreach (KeyValuePair<string, JToken> sourceProperty in json_result)
                {
                    textBox2.AppendText($"{sourceProperty.Key} : {sourceProperty.Value}" + Environment.NewLine);
                    if (Regex.IsMatch(sourceProperty.Value.ToString(), "^.*-.*-.*"))
                    {
                        groupID_tmp = sourceProperty.Value.ToString();
                    }

                }
            }
            catch (Exception ex)
            {
                textBox2.AppendText(ex.Message + Environment.NewLine);
            }
            return groupID_tmp;
        }

        //Добавление УЗ в группу (в RIMS)
        public void RimsAddUsersGroup(string adressRims, string groupId, string objGuid, string userSam, string login, string password)
        {

            string userSamAc = Regex.Replace(userSam, @"^.*\(|\).*$","");

            // Метод получения UID пользователя из RIMS
            string UID = RimsGetID(adressRims, userSamAc,  sServiceUser, sServicePassword);

            if(string.IsNullOrEmpty(UID))
            {
                textBox2.AppendText($"{userSam} не получен UID" + Environment.NewLine);
                return;
            }

            string url = $"{adressRims}/api/Permission/v2/ApplyUsersPermissionsChanges";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";
            httpRequest.Accept = "application/json";

            if (!string.IsNullOrEmpty(sServiceUser) & !string.IsNullOrEmpty(sServicePassword))
            {
                NetworkCredential credential =
                    new NetworkCredential(sServiceUser, sServicePassword, "iie.cp");
                httpRequest.Credentials = credential;
            }
            else
            {
                httpRequest.UseDefaultCredentials = true;
            }
            httpRequest.ContentType = "application/json";


            Chs chsobj = new Chs()
            {
                // guid group
                PermissionId = groupId,
                // ObjectGuid user
                AccountId = objGuid,
                ChangeType = "add",
                Values = new Values() { }
            };
            List<Chs> chslist = new List<Chs>();
            chslist.Add(chsobj);

            UsChanges<Chs> usChangesobj = new UsChanges<Chs>()
            {
                Entity = "Person",
                Uid = UID,
                Changes = chslist
            };
            List<UsChanges<Chs>> usChangeslst = new List<UsChanges<Chs>>();
            usChangeslst.Add(usChangesobj);

            JSONaddUsers<UsChanges<Chs>> jsonaddusers = new JSONaddUsers<UsChanges<Chs>>()
            {
                Culture = "ru",
                RequestNumber = "",
                Explanation = "Add users",
                UserChanges = usChangeslst
            };

            string content = JsonConvert.SerializeObject(jsonaddusers);

            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(content);
            }

            try
            {
                var result = "";
                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }
                //textBox5.AppendText($"{httpResponse.StatusCode}");
                //Console.WriteLine($"httpResponse.StatusCode: {httpResponse.StatusCode}");

                var resultJS = JsonConvert.DeserializeObject<HttpResponseModel>(result);
                //textBox5.AppendText($". Количество УЗ: {status.Data?.Count}" + Environment.NewLine);
                //Console.WriteLine($"Status operation: {resultJS.Status}");

                //status.Data.Sort();

                if (resultJS.Data != null)
                {
                    foreach (var item in resultJS.Data)
                    {
                        //test1.Add(item);
                        //Console.WriteLine($"Status ispol: {item.Status}");
                        //Console.WriteLine($"Message: {item.Message}");
                        //Console.WriteLine($"Uid: {item.Uid}");
                        //string Ok = "(OK)";
                        //string No = "(NO)";
                        //item.Status ? Ok : No;
                        textBox2.AppendText($"{userSam}: {item.Status} " + Environment.NewLine);
                    }
                }
                else
                {
                    textBox2.AppendText($"УЗ в запросе не получено" + Environment.NewLine);
                    //Console.WriteLine($"УЗ в запросе не получено" + Environment.NewLine);
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                textBox2.AppendText(ex.Message + Environment.NewLine);
                //Console.WriteLine(ex.Message + Environment.NewLine);
            }










        }



        // Получение UID пользователя
        public string RimsGetID(string adressRims, string userSam, string login, string password)
        {
            string id = "";
            string url = $"{adressRims}/api/Search/Iterate";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";
            httpRequest.Accept = "application/json";
            
            if (!string.IsNullOrEmpty(sServiceUser) & !string.IsNullOrEmpty(sServicePassword))
            {
                NetworkCredential credential =
                    new NetworkCredential(sServiceUser, sServicePassword, "ie.corp");
                httpRequest.Credentials = credential;
            }
            else
            {
                httpRequest.UseDefaultCredentials = true;
            }
            httpRequest.ContentType = "application/json";




            // Create Objects for Serialization
            Flags flags = new Flags()
            {
                only_with_email = false
            };

            SearchParams ser = new SearchParams()
            {
                Search = $"{userSam} iie.cp",
                Culture = "ru",
                PageSize = 20,
                PageNumber = 1,
                Flags = flags
            };

            JSONUidUser jSONUidUser = new JSONUidUser
            {
                SearchParams = ser,
                Entities = new string[] { "ADAccount" }
            };

            string content = JsonConvert.SerializeObject(jSONUidUser);
            
            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(content);
            }

            try
            {
                var result = "";
                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }

                var resultJS = JsonConvert.DeserializeObject<JSONUidUserOutput<DataAd<Result<Itms>>>>(result);
                //textBox5.AppendText($". Количество УЗ: {status.Data?.Count}" + Environment.NewLine);

                //id = resultJS.Data[0].Result.Items[0].Uid;

                //var test = from c in resultJS.Data[0].Result.Items where c.AccountName == userSam select c;
                // OR
                IEnumerable<Itms> id2;
                id2 = Enumerable.Where(resultJS.Data[0].Result.Items, n => n.AccountName == userSam).Select(n => n);
                
                foreach(Itms i2 in id2)
                {
                    id = i2.Uid;
                }


            } catch (Exception ex)
            {
                textBox2.AppendText(ex.Message + Environment.NewLine);
            }



            return id;
        } // Получение UID пользователя




        #endregion


        // Скрыть пароль или показать
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(txtBPass.Text) && txtBPass.UseSystemPasswordChar) {
                txtBPass.UseSystemPasswordChar = false;
                pictureBox2.BackgroundImage = Resources.eye_icon_224636;
                txtBPass.Focus();
            } else
            {
                txtBPass.UseSystemPasswordChar = true;
                pictureBox2.BackgroundImage = Resources.slash_eye_icon_224538;
                txtBPass.Focus();
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
        }

        
    }



    // Serialzation json Get group in RiMS
    #region class for group to RiMS
    public class JSONZC
    {
        public string GroupName { get; set; }
        public string Domain { get; set; }
    }
    #endregion

    //Deserialization json group from RIMS
    #region Главный классы расшифровки JSON - start
    public class HttpResponseModel
    {
        public bool Status { get; set; }
        public string Message { get; set; }
        public List<Dat> Data { get; set; }
    }

    public class Dat //: IComparable<Dat>
    {
        public string Uid { get; set; }
        public string Entity { get; set; }
        public bool Status { get; set; }
        public string Message { get; set; }
        public Change Change { get; set; }
    }
    public class Change
    {
        public string PermissionId { get; set; }
        public string AccountId { get; set; }
        public string ChangeType { get; set; }
        public Values Values { get; set; }

    }
    public class Values
    {
        public string RequestNumber { get; set; }
        public string Explanation { get; set; }
        public string Uid { get; set; }
        public string AccountId { get; set; }
        public string Entity { get; set; }
    }
    #endregion
    


    // Serialzation json add user to Group
    #region Serialzation json add to Group
    public class JSONaddUsers<T>
    {
        public string Culture { get; set; }
        public string RequestNumber { get; set; }
        public string Explanation { get; set; }
        public List<T> UserChanges { get; set; }

    }
    public class UsChanges<T>
    {
        public string Entity { get; set; }
        public string Uid { get; set; }
        public List<T> Changes { get; set; }
    }
    public class Chs
    {
        public string PermissionId { get; set; }
        public string AccountId { get; set; }
        public string ChangeType { get; set; }
        public Values Values { get; set; }
    }
    /*public class Values
    {
    }*/

    #endregion
    // Serialzaton json



    
    //Serialization json for get UID user
    #region Serialization json for get UID user
    public class JSONUidUser
    {
        public SearchParams SearchParams { get; set; }
        public string[] Entities = { "ADAccount" };
    }
    public class SearchParams
    {
        public string Search { get; set; }
        public string Culture { get; set; }
        public int PageSize { get; set; }
        public int PageNumber { get; set; }
        public Flags Flags { get; set; }
    }
    public class Flags
    {
        public bool only_with_email { get; set; }
    }

    #endregion



    //DeSerialization json for get UID user
    #region DeSerialization json for get UID user
    public class JSONUidUserOutput<T>
    {
        public bool Status { get; set; }
        public string Message { get; set; }
        public List<T> Data { get; set; }
    }
    public class DataAd<T>
    {
        public string Entity { get; set; }
        public Result<T> Result { get; set; }
    }
    public class Result<T>
    {
        public string PageSize { get; set; }
        public string PageNumber { get; set; }
        public string Total { get; set; }
        public List<Itms> Items { get; set; }
        
    }
    public class Itms
    {
        public string AccountId { get; set; }
        public string Uid { get; set; }
        public string AccountName { get; set; }
        public string Caption { get; set; }
        public string Company { get; set; }
        public string JobTitle { get; set; }
        public string City { get; set; }
        public string AppName { get; set; }
        public string Email { get; set; }

    }
    #endregion





    }
