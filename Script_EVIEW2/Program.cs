using System;
using System.Runtime.InteropServices;
using System.Threading;
using AutoIt;

namespace Script_EVIEW2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");

            string path_open = "";
            string path_save = "";

            path_open = args[0];
            path_save = args[1];

            AutoItX.Run(@path_open, "", AutoItX.SW_SHOW);
            Thread.Sleep(1000);

            AutoItX.WinWait("EView - V2.2.9.138");
            AutoItX.WinWaitActive("EView - V2.2.9.138");

            //Нажатие на кнопку: "Add Network"
            //AutoItX.ControlClick("[CLASS:TEview2_MainForm]", "Add Network", "TJvPanel1","left",1,70,14);
            AutoItX.Send("{LALT}+{N}");
            Thread.Sleep(100);
            AutoItX.Send("{ENTER}");
            Thread.Sleep(500);
            //Нажатие на кнопку: "OK" на панели выбора подключения COM порта
            AutoItX.ControlClick("[CLASS:TAddNetworkForm]", "", "TJvBitBtn1", "left", 1, 50, 13);
            Thread.Sleep(500);
            //Нажатие на кнопку: "View" на панели меню
            AutoItX.ControlClick("[CLASS:TEview2_MainForm]", "&View", "");
            //Нажатие на кнопку: "View" на панели меню
            AutoItX.Send("{LALT}+{V}");
            Thread.Sleep(100);
            AutoItX.Send("{DOWN}");
            Thread.Sleep(100);
            AutoItX.Send("{RIGHT}");
            Thread.Sleep(100);
            AutoItX.Send("{ENTER}");
            Thread.Sleep(100);
            //Нажатие на кнопку: "Detect"
            AutoItX.ControlClick("[CLASS:TEview2_MainForm]", "", "TJvPanel2","left",1,44,13);
            Thread.Sleep(100);

            AutoItX.WinWaitClose("[CLASS:TScanningNetworkForm]");
            Thread.Sleep(100);
            //Сохранение данных в файл
            AutoItX.ControlClick("[CLASS:TEview2_MainForm]", "", "TJvListView1", "right", 1, 465, 53);
            Thread.Sleep(100);
            AutoItX.Send("{DOWN}");
            Thread.Sleep(100);
            Thread.Sleep(100);
            AutoItX.Send("{DOWN}");
            Thread.Sleep(100);
            Thread.Sleep(100);
            AutoItX.Send("{DOWN}");
            Thread.Sleep(100);
            Thread.Sleep(100);
            AutoItX.Send("{DOWN}");
            Thread.Sleep(100);
            AutoItX.Send("{ENTER}");
            //Ждем загрузки окна сохранения
            AutoItX.WinWaitActive("[CLASS:#32770]");
            //Выбираем строку ввода пути расположения файла
            AutoItX.ControlClick("[CLASS:#32770]", "", "ToolbarWindow323", "left", 1, 600, 10);
            Thread.Sleep(100);
            //Вставляем путь расположения лог файла
            AutoItX.ControlSetText("[CLASS:#32770]", "", "Edit2", @path_save);
            Thread.Sleep(100);
            AutoItX.Send("{ENTER}");
            Thread.Sleep(100);
            //Вставляем имя файла
            AutoItX.ControlSetText("[CLASS:#32770]", "", "Edit1", "LogOut.txt");
            Thread.Sleep(100);
            //Жмем кнопку сохранить
            AutoItX.ControlClick("[CLASS:#32770]", "", "Button1", "left", 1, 50, 13);
            Thread.Sleep(500);
            AutoItX.WinClose();
            Thread.Sleep(500);
            AutoItX.WinClose();

            Console.WriteLine("...Finish");
        }
    }
}
