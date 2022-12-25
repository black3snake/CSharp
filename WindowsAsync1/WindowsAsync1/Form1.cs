using Config.Net;
using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WindowsAsync1
{
    public partial class Form1 : Form
    {
        public IMySettings settings = new ConfigurationBuilder<IMySettings>()
            .UseIniFile(@"config.ini", true)
            .Build();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public Form1()
        {
            InitializeComponent();

        }

        public void button1_Click(object sender, EventArgs e)
        {
            textBox1.AppendText($"Main ThreadID {Thread.CurrentThread.ManagedThreadId}\r\n");
            logger.Info($"Main ThreadID {Thread.CurrentThread.ManagedThreadId}\r\n");
            MyClass my = new MyClass();
            my.OperationAsync();

            // Delay
            //Console.ReadKey();

            logger.Info("Основной поток закончил работу");
            textBox1.AppendText("Основной поток закончил работу\r\n");

            textBox1.AppendText($"MyOption: {settings.MyOption} \r\n");
            textBox1.AppendText($"Adress: {settings.Adress} \r\n");
            foreach (var item in settings.UserNameEx)
            {
                textBox1.AppendText($"{item} \r\n");
            }


        }


        private void button2_Click(object sender, EventArgs e)
        {
            WorkAs();
        }

        private async void WorkAs()
        {
            await Task.Run(() =>
            {

                Thread.Sleep(8000);
                textBox1.Invoke(new MethodInvoker(() =>
                    {
                        textBox1.AppendText($"Operation ThreadID {Thread.CurrentThread.ManagedThreadId} + {settings.Adress}\r\n");
                    }));


            });
            Thread.Sleep(2000);
        }


    }

    public interface IMySettings
    {
        string MyOption { get; }
        string Adress { get; }
        IEnumerable<string> UserNameEx { get; }
    }
}
