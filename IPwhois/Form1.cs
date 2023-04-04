using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IPwhois
{
    public partial class Form1 : Form
    {
        string addressG = "www.google.ru/maps/place/";
        public Form1()
        {
            InitializeComponent();
            // Быстрая обработка нажатия клавиши через Лямбду (внизу описано с методом, а здесь его создавать отдельно не надо) 
            KeyPreview = true;
            KeyDown += (s, e) => { if (e.KeyValue == (char)Keys.Enter) button1_Click(button1, null); };
            textBox1.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string line = "", lineLat = "", lineLon = "",address = "";
            if (textBox1.Text != "")
            {

                if (!Regex.IsMatch(textBox1.Text.Substring(textBox1.Text.Length-1), @"\d"))
                {
                    MessageBox.Show("В конце должна стоять только цифра", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string[] TB = textBox1.Text.Split('.');
                foreach(string t1 in TB)
                {
                    int t1int = Convert.ToInt32(t1);
                    if(t1int >= 255)
                    {
                        MessageBox.Show("Неправельный диапазон", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                using (WebClient wc = new WebClient())
                {
                    //JSON
                    line = wc.DownloadString($"https://ipwho.is/{textBox1.Text}?output=json");
                }
                //JSON
                Match match = Regex.Match(line, "\"country\":\"(.*?)\",(.*?)region\":\"(.*?)\",(.*?)city\":\"(.*?)\",\"latitude\":(.*?),\"longitude\":(.*?),\"(.*?)isp\":\"(.*?)\",(.*?)utc\":\"(.*?)\",");
                label10.Text = match.Groups[1].Value;
                label11.Text = match.Groups[3].Value;
                label12.Text = match.Groups[5].Value;
                label13.Text = match.Groups[9].Value;
                label14.Text = match.Groups[11].Value;
                label15.Text = match.Groups[6].Value;
                label16.Text = match.Groups[7].Value;

                lineLat = match.Groups[6].Value;
                lineLon = match.Groups[7].Value;

                address = $"https://{addressG}{lineLat}+{lineLon}";


                //webBrowser1.Navigate($"https://www.google.ru/maps/place/{lineLat}+{lineLon}"); //52%C2%B017'13.1%22N+104%C2%B018'18.1%22E

                webBrowser1.Navigate(new Uri(address));
                
                if (Size.Width <= 1000)
                {
                    Size = new Size(1164, 708);
                }
            } else
            {
                MessageBox.Show("Введите ИП адрес", Text, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (Regex.IsMatch(textBox1.Text, "[^0-9.]"))
            {
                MessageBox.Show("Только цифры", Text, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                textBox1.Text = textBox1.Text.Remove(textBox1.Text.Length - 1);
                textBox1.SelectionStart = textBox1.Text.Length;

            }
        }
    }
}
