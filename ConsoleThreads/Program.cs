using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConsoleThreds
{
    internal class Program
    {

        // Метод выполняющийся во вторичном потоке
        static void Function()
        {
            Console.WriteLine($"ID {Thread.CurrentThread.Name} запуск потока: {Thread.CurrentThread.ManagedThreadId}");
            Console.ForegroundColor = ConsoleColor.Yellow;

            for (int i = 0; i < 160; i++)
            {
                Thread.Sleep(20);
                Console.Write(".");
            }
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine($"{Thread.CurrentThread.Name} поток завершился");
        }

        static void Dowork(object obj)
        {
            Obj1 obj_tmp = (Obj1)obj;
            
            Console.WriteLine($"ID {Thread.CurrentThread.ManagedThreadId} запускаем поток");
            Console.ForegroundColor = ConsoleColor.Blue;

            for (int i = 0; i < obj_tmp.max; i++)
            {
                Thread.Sleep(20);
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.Write("+");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine($"Имя принимаемого файла {obj_tmp.nameF}");
            Console.WriteLine($" Поток {Thread.CurrentThread.ManagedThreadId} завершился");
            obj_tmp.Print();
        }
        static void Main(string[] args)
        {
            Console.WriteLine("ID Первичного потока: {0}", Thread.CurrentThread.GetHashCode());

            //Thread threadR = Thread.CurrentThread;
            //Console.WriteLine($" имя род потока {threadR.Name}");

            int kol_thred = 8;
            var threads = new Thread[kol_thred];
            //Создание нового потока
            for (int t = 1; t < kol_thred; t++) {
                Thread threadP = new Thread(new ThreadStart(Function));
                threadP.Name = $"Thread{t}";

                threads[t] = threadP;
                threadP.Start();
            }
            Obj1 obj = new Obj1(260, @"C:\Users\snake\Desktop\test\1.txt");
            Obj1 obj2 = new Obj1(180, @"C:\Users\snake\Desktop\test\2.txt");
            //obj.max = 260;
            //obj.nameF = @"C:\Users\snake\Desktop\test\1.txt";

            Thread thread2 = new Thread(new ParameterizedThreadStart(Dowork));
            thread2.Start(obj);

            Thread thread3 = new Thread(Dowork);
            thread3.Start(obj2);

            // Ожидание Первычным потоком, завершение работы вторичного потока
            //thread.Join(); //TODO Снять или установить комментарий
            for (int t = 1; t < kol_thred; t++)
            {
                threads[t].Join();
            }

            thread2.Join();
            thread3.Join();

            Console.ForegroundColor = ConsoleColor.Green;

            for (int i = 0; i < 160; i++)
            {
                Thread.Sleep(20);
                Console.Write("-");
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nПервичный поток завершился.");

            //Delay
            Console.ReadKey();

        }
    
    }
    public class Obj1 
    {
        public int max;
        public string nameF;
        public Obj1(int m, string s) { max = m; nameF = s; }
        
        public void Print()
        {
            Console.WriteLine($"int max = {max}, string nameF = {nameF}");
        }
    }
    
}
