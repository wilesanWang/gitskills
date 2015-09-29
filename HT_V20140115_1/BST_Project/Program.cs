using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace BST_Project
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        private static System.Threading.Mutex mutex;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            mutex = new System.Threading.Mutex(true, "OnlyRun");
            if (mutex.WaitOne(0, false))
            {

                Application.Run(new MainForm());
            }
            else
            {
                MessageBox.Show("程序已经在运行！", "提示");
                Application.Exit();
            }
        }
    }
}
