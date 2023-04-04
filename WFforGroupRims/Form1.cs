using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Security.Principal;
using ActiveDs;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace WFforGroupRims
{
    public partial class Form1 : Form
    {
        static string addrRims = "https://k.hq.root.ad";

        #region Переменные AD
        public static string sDomainDefault { get; set; } //куда подключаемся по умолчанию
        public static string sDomain = "iie.cp"; //куда подключаемся
        public static string sDefaultRootOU = "DC=iie,DC=cp"; //где ищем по умолчанию
        public static string sServiceUser { get; set; }  //пользователь от кого делаем
        public static string sServicePassword { get; set; } // пароль
        
        //public static string sPDCname { get; set; }  // Имя PDC

        //private static bool enabl = true;
        public enum LdapFilter { UsersSAN, UsersCN, Computers, Groups, OU, UsersName };
        public struct GroupProperty
        {
            public string samaccountname;
            public string description;
            public string name;
            public string cn;
            public string distinguishedname;
            public int grouptype;
            public string coment;
            public string managedby;
        }

        #endregion


        public Form1()
        {
            InitializeComponent();
            // Проверка реестра
            RegistryKey keyS = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\"+ Text, false);
            if(keyS == null)
            {
                textBox3.Text = "default address, click on the inscription";
            } else {
                textBox3.Text = keyS.GetValue("address").ToString();
            }
            textBox3.Click += (s, a) => { textBox3.Text = addrRims; };

            // Получим ИП адрес PDC ассинхронно
            pdcIP();

            if (string.IsNullOrEmpty(textBox1.Text)) { textBox1.Text = "empty"; } // else { sServiceUser = textBox1.Text.Trim(); }
            textBox1.GotFocus += (s, a) => { if (textBox1.Text == "empty") { textBox1.Text = ""; } };
            textBox1.LostFocus += (s, a) => 
            { 
                if (string.IsNullOrEmpty(textBox1.Text)) 
                { 
                    textBox1.Text = "empty"; 
                } else
                {
                    sServiceUser = textBox1.Text.Trim();
                }
            };

            textBox2.LostFocus += (s, a) =>
            {
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    sServicePassword = textBox2.Text.Trim();
                }
            };

            //if (!string.IsNullOrEmpty(textBox2.Text)) { sServicePassword = textBox2.Text.Trim(); }

            //label9.TextChanged += (s, a) => { if (label9.Text != "IP..") { sPDCname = label9.Text.Trim(); } };
        }



        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            Pen pen = new Pen(Color.Black, 1);
            e.Graphics.DrawLine(pen, 10, 73, 500, 73);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RegistryKey keyS = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft", true);
            using (var subkey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\"+Text))
            {
                subkey.SetValue("address", addrRims);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            RegistryKey keyS = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft", true);
            using (RegistryKey keySub = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\"+Text, true))
            {
                if (keySub != null)
                {
                    keySub.DeleteValue("address");
                    keyS.DeleteSubKey(Text);
                }
            }
            textBox3.Text = null;
        }

        // Запрос
        private void button3_Click(object sender, EventArgs e)
        {

            List<string> UsersGrs = new List<string>();
            UsersGrs = ListMembersGroup1500(textBox4.Text.Trim());
            foreach (var UG in UsersGrs)
            {
                Match match = Regex.Match(UG, "CN=(.*?),OU=");
                textBox5.AppendText(match.Groups[1].Value + Environment.NewLine);
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox5.Text = null;
        }

        // Получение асинхронно IP PDC
        private async void pdcIP()
        {
            await Task.Run(() =>
            {
                Process process = Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell",
                    Arguments = "/command (Get-ADDomainController -Discover -Service PrimaryDC).HostName",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                });

                sDomainDefault = process.StandardOutput.ReadToEnd().Trim();
                Invoke(new System.Action(() =>
                {
                    label9.Text = sDomainDefault;
                }));


                //return process.StandardOutput.ReadToEnd();
            });
        }

        #region Методы получения Users из Group

        /// <summary>
        /// Возвращает пользователей находящихся в группе (более 1500 членов)
        /// </summary>
        /// <param name="sGroupName">Имя группы</param>
        /// <returns>Возвращает List со всеми пользователями группы</returns>
        public static List<string> ListMembersGroup1500(string sGroupName)
        {
            List<string> myItems = new List<string>();
            SearchResult group = LDAPFindOne("", sGroupName, LdapFilter.Groups);

            DirectoryEntry objGroupEntry = group.GetDirectoryEntry();

            foreach (object objMember in objGroupEntry.Properties["member"])
            {
                myItems.Add(objMember.ToString());
            }

            return myItems;
        }
        /// <summary>
        /// Возвращает найденный обьект из АД согласно выбора
        /// </summary>
        /// <param name="ou">Место поиска</param>
        /// <param name="obj">имя объекта</param>
        /// <param name="ldf">LdapFilter:  выбор обьекта</param>
        /// <returns>Возвращает SearchResult</returns>
        public static SearchResult LDAPFindOne(string ou, string obj, LdapFilter ldf, string user = null, string password = null)
        {
            string filter = "";


            switch (ldf)
            {
                case LdapFilter.Computers: filter = "(&(objectCategory=computer)(name=" + obj + "))"; break;
                case LdapFilter.OU: filter = "(objectCategory=organizationalUnit)"; break;//----------!!!!!!!!
                case LdapFilter.UsersSAN: filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" + obj + "))"; break;
                case LdapFilter.UsersCN: filter = "(&(objectCategory=person)(objectClass=user)(CN=" + obj + "))"; break;
                case LdapFilter.Groups: filter = "(&(objectCategory=group)(name=" + obj + ")) "; break;
            }

            return LDAPFindOne(ou, filter, user, password);
        }

        /// <summary>
        /// Возвращает найденный обьект из АД согласно фильтру
        /// </summary>
        /// <param name="ou">Место поиска</param>
        /// <param name="Filter">Параметры фильтра</param>
        /// <returns>Возвращает SearchResult</returns>
        public static SearchResult LDAPFindOne(string ou, string Filter, string user = null, string password = null)
        {

            if (ou == "")
            {
                ou = sDefaultRootOU;
            }

            //var domainPath = @"LDAP://" + sDomain + "/" + ou;
            DirectoryEntry myLdapConnection = createDirectoryEntry();

            /*if (String.IsNullOrEmpty(user) || String.IsNullOrEmpty(password))
            {
                directoryEntry = new DirectoryEntry(domainPath, sServiceUser, sServicePassword);
            }
            else
            {
                directoryEntry = new DirectoryEntry(domainPath, user, password);
            }*/

            var dirSearcher = new DirectorySearcher(myLdapConnection);
            dirSearcher.SearchScope = SearchScope.Subtree;
            dirSearcher.PageSize = 100;
            dirSearcher.SizeLimit = 5000;
            dirSearcher.Filter = Filter;

            try
            {

                return dirSearcher.FindOne();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Исключение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }

        #endregion

        #region Методы проверки

        /// <summary>
        /// Проверка имени пользователя и пароля
        /// </summary>
        /// <param name="sUserName">Имя пользователя</param>
        /// <param name="sPassword">Пароль</param>
        /// <returns>Возвращает true, если имя и пароль верны</returns>
        public static bool OldValidateCredentials(string sUserName, string sPassword)
        {
            return GetPrincipalContext().ValidateCredentials(sUserName, sPassword);
        }


        /// <summary>
        /// Проверка имени пользователя и пароля
        /// </summary>
        /// <param name="sUserName">Имя пользователя</param>
        /// <param name="sPassword">Пароль</param>
        /// <returns>Возвращает true, если имя и пароль верны</returns>
        public static bool LDAPValidateCredentials(string sUserName, string sPassword)
        {

            var pc = LDAPFindOne("", sUserName, LdapFilter.UsersSAN, sUserName, sPassword);

            if (pc == null) return false;
            else { return true; }
        }

        /// <summary>
        /// Проверка имени пользователя и пароля
        /// </summary>
        /// <param name="sUserName">Имя пользователя</param>
        /// <param name="sPassword">Пароль</param>
        /// <returns>Возвращает true, если имя и пароль верны</returns>
        public static bool LDAPValidateCredentials()
        {

            var pc = LDAPFindOne("", sServiceUser, LdapFilter.UsersSAN);

            if (pc == null) return false;
            else { return true; }
        }
        #endregion

        /// <summary>
        /// Получить основной контекст указанного OU
        /// </summary>
        /// <param name="sOU">OU для которого нужно получить основной контекст</param>
        /// <returns>Возвращает объект PrincipalContext</returns>
        public static PrincipalContext GetPrincipalContext(string sOU = "")
        {
            if (string.IsNullOrEmpty(sOU)) return new PrincipalContext(ContextType.Domain, sDomain, sServiceUser, sServicePassword);
            else
                return new PrincipalContext(ContextType.Domain, sDomain, sOU, sServiceUser, sServicePassword);
        }



        static DirectoryEntry createDirectoryEntry()
        {
            // create and return new LDAP connection with desired settings  
            DirectoryEntry ldapConnection = new DirectoryEntry(sDomainDefault, sServiceUser, sServicePassword);
            ldapConnection.Path = "LDAP://DC=iie,DC=cp";
            //ldapConnection.Path = "LDAP://OU=staffusers,DC=leeds-art,DC=ac,DC=uk";

            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

            return ldapConnection;
        }

    }
}
