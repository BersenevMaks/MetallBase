using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MetalBase
{
    public delegate void MyDelegate(string data);
    public delegate void Action();
    public delegate void settingDelegate(List<string> sett);
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
