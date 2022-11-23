using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WindowsAsync1
{
    public partial class Form1 : Form
    {
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

        }
    }
    class MyClass:Form1
    {
        public void Operation()
        {
            //textBox1.AppendText($"Operation ThreadID {Thread.CurrentThread.ManagedThreadId}\r\n");
            logger.Info($"Operation ThreadID {Thread.CurrentThread.ManagedThreadId}\r\n");

            logger.Info("Begin");
            Thread.Sleep(2000);
            logger.Info("End");
        }

        public async void OperationAsync()
        {
            // Id потока совпадает с Id первичного потока. Это значит, что
            // данный метод начинает выполняться в контексте первичного потока.
            logger.Info($"OperationAsync (Part I) ThreadID {Thread.CurrentThread.ManagedThreadId}\r\n");

            Task task = new Task(Operation);
            task.Start();
            await task;

            // Id потока совпадает с Id вторичного потока. Это значит, что
            // данный метод заканчивает выполняться в контексте вторичного потока.
            logger.Info($"OperationAsync (Part II) ThreadID {Thread.CurrentThread.ManagedThreadId}");
        }
    }
}
