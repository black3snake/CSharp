using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using WFRemoveFromGroup.Properties;

namespace WFRemoveFromGroup
{
    public partial class Form1 : Form
    {
        //List<string> users;
        //bool fileB = false;
        public string grID { get; set; }
        public string sServiceUser { get; set; }
        public string sServicePassword { get; set; }
        public const string adressRims = "https://k.hq.root.ad";

        public Form1()
        {
            InitializeComponent();

            listView1.MouseUp += (s, a) => {
                if (a.Button == MouseButtons.Right)
                {
                    contextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right);
                }
            };

        }

        private void btnGroup_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBGroup.Text))
            {
                MessageBox.Show("Заполни поле с именем группы", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sServiceUser = txtBLogin.Text.Trim();
            sServicePassword = txtBPass.Text.Trim();

            grID = ZaprosToRims(adressRims, txtBGroup.Text.Trim(), sServiceUser, sServicePassword);
            if (!string.IsNullOrEmpty(grID))
            {
                pBox2.BackgroundImage = Resources.Ok_27007;
            }
            else
            {
                pBox2.BackgroundImage = Resources.Close_2_26986;
            }


        }
        private void btnGetAccouns_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(grID))
            {
                MessageBox.Show("Сначало нужно нажать \r\n кнопку \"Get Info\" и проверить группу", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            listView1.Items.Clear();

            if (listView1.Columns.Count > 0)
            {
                foreach (ColumnHeader ch in listView1.Columns)
                {
                    listView1.Columns.Remove(ch);
                }
            }

            ColumnHeader header1, header2;
            header1 = new ColumnHeader();
            header2 = new ColumnHeader();

            header1.Text = "ID";
            header1.TextAlign = HorizontalAlignment.Left;
            header1.Width = 100;

            header2.Text = "Name";
            header2.TextAlign = HorizontalAlignment.Left;
            header2.Width = 250;

            listView1.Columns.Add(header1);
            listView1.Columns.Add(header2);
            //listView1.Columns[0].Width = 200;
            listView1.View = View.Details;

            ZaprosToRims2(adressRims, grID, sServiceUser, sServicePassword);



        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtBConsole.Clear();
        }

        private void btnRemoveAll_Click(object sender, EventArgs e)
        {
            if(!checkBox1.Checked)
            {
                MessageBox.Show("Сначало нужно нажать \r\n на чекбокc, чтобы запустить процедуру \r\n удаления Учетных записей из группы в RIMS", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Удаление пользователей из группы RIMS
            // Переберем все УЗ, что есть в списке и отправим запрос в RIMS
            foreach (ListViewItem item in listView1.Items)
            {
                RimsDeleteUsersGroup(adressRims, grID, item.Text, sServiceUser, sServicePassword);


            }
            listView1.Items.Clear();
            txtBConsole.AppendText("Работа выполнена" + Environment.NewLine);

        }

        // Скрытие пароля
        private void pBox1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtBPass.Text) && txtBPass.UseSystemPasswordChar)
            {
                txtBPass.UseSystemPasswordChar = false;
                pBox1.BackgroundImage = Resources.eye_icon_224636;
                txtBPass.Focus();
            }
            else
            {
                txtBPass.UseSystemPasswordChar = true;
                pBox1.BackgroundImage = Resources.slash_eye_icon_224538;
                txtBPass.Focus();
            }

        }



        #region Method for RIMS
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
                Domain = "IIE.CP"
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
                    txtBConsole.AppendText($"{sourceProperty.Key} : {sourceProperty.Value}" + Environment.NewLine);
                    if (Regex.IsMatch(sourceProperty.Value.ToString(), "^.*-.*-.*"))
                    {
                        groupID_tmp = sourceProperty.Value.ToString();
                    }

                }
            }
            catch (Exception ex)
            {
                txtBConsole.AppendText(ex.Message + Environment.NewLine);
            }
            return groupID_tmp;
        }

        // получим список пользователей в группе
        public async void ZaprosToRims2(string adressRims, string groupId, string login, string password)
        {
            string url = $"{adressRims}/api/Gather/GetEntityExtensionData";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";
            httpRequest.Accept = "application/json";

            if (sServiceUser != "empty" & !string.IsNullOrEmpty(sServicePassword))
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


            JSONZCR jSONZCR = new JSONZCR
            {
                Uid = groupId,
                Entity = "Permission",
                Culture = "ru",
                Data = new Data { Key = "AD_GROUP_USERS", Scope = "permission" }
            };

            string content = JsonConvert.SerializeObject(jSONZCR);

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

                await Task.Delay(2000);
                txtBConsole.AppendText($"{httpResponse.StatusCode}");

                var status = JsonConvert.DeserializeObject<HttpResponseModel>(result);
                txtBConsole.AppendText($". Количество УЗ: {status.Data?.Count}" + Environment.NewLine);

                status.Data.Sort();

                if (status.Data != null)
                {
                    foreach (var item in status.Data)
                    {
                        //txtBConsole.AppendText($"{item.Caption} : ({item.SamAccount})" + Environment.NewLine);

                        ListViewItem itemL1 = new ListViewItem(item.SamAccount);
                        itemL1.SubItems.Add(item.Caption);
                        listView1.Items.Add(itemL1);

                    }
                    label4.Text = $"Count: {listView1.Items.Count}";
                }
                else
                {
                    txtBConsole.AppendText($"УЗ в запросе не получено" + Environment.NewLine);
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                txtBConsole.AppendText(ex.Message + Environment.NewLine);
            }

        }

        // Получение UID пользователя
        public ValueTuple<string, string> RimsGetID(string adressRims, string userSam, string login, string password)
        {
            //string id = "";
            //ValueTuple<string Uid, string AccountId> 
                var vtup = (Uid:"", AccountId:"");
            
            string url = $"{adressRims}/api/Search/Iterate";

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

                //id = resultJS.Data[0].Result.Items[0].Uid;
                //var test = from c in resultJS.Data[0].Result.Items where c.AccountName == userSam select c;
                // OR
                IEnumerable<Itms> id2;
                id2 = Enumerable.Where(resultJS.Data[0].Result.Items, n => n.AccountName == userSam).Select(n => n);

                foreach (Itms i2 in id2)
                {
                    //id = i2.Uid;
                    //i2.AccountId;
                    // vtup.item
                    vtup = (Uid:i2.Uid, AccountId:i2.AccountId);
                }


            }
            catch (Exception ex)
            {
                txtBConsole.AppendText(ex.Message + Environment.NewLine);
            }

            return vtup;
        } // Получение UID пользователя


        //Добавление УЗ в группу (в RIMS)
        public void RimsDeleteUsersGroup(string adressRims, string groupId, string userSam, string login, string password)
        {

            //string userSamAc = Regex.Replace(userSam, @"^.*\(|\).*$", "");
            var vtup2 = (Uid:"", AccountId:"");
            // Метод получения UID пользователя из RIMS
            vtup2 = RimsGetID(adressRims, userSam, sServiceUser, sServicePassword);

            if (string.IsNullOrEmpty(vtup2.Uid))
            {
                txtBConsole.AppendText($"{userSam} не получен UID" + Environment.NewLine);
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
                AccountId = vtup2.AccountId,
                ChangeType = "delete",
                Values = new Values() { }
            };
            List<Chs> chslist = new List<Chs>();
            chslist.Add(chsobj);

            UsChanges<Chs> usChangesobj = new UsChanges<Chs>()
            {
                Entity = "Person",
                Uid = vtup2.Uid,
                Changes = chslist
            };
            List<UsChanges<Chs>> usChangeslst = new List<UsChanges<Chs>>();
            usChangeslst.Add(usChangesobj);

            JSONdeleteUsers<UsChanges<Chs>> jsonaddusers = new JSONdeleteUsers<UsChanges<Chs>>()
            {
                Culture = "ru",
                RequestNumber = "",
                Explanation = "Delete users",
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

                var resultJS = JsonConvert.DeserializeObject<HttpResponseModel2>(result);
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
                        txtBConsole.AppendText($"{userSam}: {item.Status} " + Environment.NewLine);
                    }
                }
                else
                {
                    txtBConsole.AppendText($"УЗ в запросе не получено" + Environment.NewLine);
                    //Console.WriteLine($"УЗ в запросе не получено" + Environment.NewLine);
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                txtBConsole.AppendText(ex.Message + Environment.NewLine);
                //Console.WriteLine(ex.Message + Environment.NewLine);
            }






        }



        #endregion


        // Обработка StripMenu ListView 
        #region StripMenu
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count > 0)
            {
                try
                {
                    if(listView1.SelectedItems.Count > 1)
                    {
                        foreach (ListViewItem item in listView1.SelectedItems)
                        {
                            listView1.Items.Remove(item);
                        }

                    } else
                    {
                        listView1.Items.Remove(listView1.SelectedItems[0]);
                    }
                    
                }
                catch
                {
                    MessageBox.Show("Не могу удалить пустое поле!!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            label4.Text = $"Count: {listView1.Items.Count}";


        }


        #endregion
    }

    // Serialzation json Get group in RiMS
    // and
    //Deserialization json group from RIMS
    #region Get Group guID from RIMS
    public class JSONZC
    {
        public string GroupName { get; set; }
        public string Domain { get; set; }
    }
    //Deserialization json group from RIMS
    // Не использую классы - забираю просто
    #endregion

    #region Get Accounts from RIMS
    // Serilization
    public class JSONZCR
    {
        public string Uid { get; set; }
        public string Entity { get; set; }
        public string Culture { get; set; }
        public Data Data { get; set; }

    }
    public class Data
    {
        public string Key = "AD_GROUP_USERS";
        public string Scope = "permission";
    }

    //DeSerilazation
    public class HttpResponseModel
    {
        public bool Status { get; set; }
        public string Message { get; set; }
        public List<Dat> Data { get; set; }
    }
    public class Dat : IComparable<Dat>
    {
        public string AppId { get; set; }
        public string AppName { get; set; }
        public string SyncKey { get; set; }
        public string Type { get; set; }
        public string City { get; set; }
        public string Owner { get; set; }
        public string Caption { get; set; }
        public string Company { get; set; }
        public string CN { get; set; }
        public string ObjectClass { get; set; }
        public string SamAccount { get; set; }
        public string DNSHostName { get; set; }
        public string Department { get; set; }
        public string JobTitle { get; set; }
        public int CompareTo(Dat o)
        {
            if (o is Dat dat) return Caption.CompareTo(dat.Caption);
            else throw new ArgumentException("Некорректное значение параметра");
        }
    }

    #endregion


    // --------------------------------
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

    
    // --------------------------------
    // Serialzation json Delete user to Group
    #region Serialzation json add to Group
    public class JSONdeleteUsers<T>
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

    //Deserialization json group from RIMS
    #region Главный классы расшифровки JSON - start
    public class HttpResponseModel2
    {
        public bool Status { get; set; }
        public string Message { get; set; }
        public List<Dat2> Data { get; set; }
    }

    public class Dat2 //: IComparable<Dat>
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

}
