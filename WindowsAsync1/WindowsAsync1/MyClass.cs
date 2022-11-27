using System.Threading;
using System.Threading.Tasks;

namespace WindowsAsync1
{
    internal class MyClass : Form1
    {
        public void Operation()
        {
            //Invoke(new Action(() =>
            logger.Info($"Operation ThreadID {Thread.CurrentThread.ManagedThreadId}\r\n");

            logger.Info("Begin");
            Thread.Sleep(2000);
            logger.Info("End");

            /*BeginInvoke(new System.Action(() =>
            {
                textBox1.AppendText($"End Sync thread" + Environment.NewLine);
            }));*/
        }

        public async void OperationAsync()
        {
            // Id потока совпадает с Id первичного потока. Это значит, что
            // данный метод начинает выполняться в контексте первичного потока.
            logger.Info($"OperationAsync (Part I) ThreadID {Thread.CurrentThread.ManagedThreadId}\r\n");


            Task task = new Task(Operation);
            task.Start();
            await task;

            //this.textBox1.Text = "End Sync thread";
            // Id потока совпадает с Id вторичного потока. Это значит, что
            // данный метод заканчивает выполняться в контексте вторичного потока.
            logger.Info($"OperationAsync (Part II) ThreadID {Thread.CurrentThread.ManagedThreadId}");
        }
    }
}
