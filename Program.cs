using System;
using System.Windows.Forms;

namespace NetThrottle
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            // standard winforms setup stuff
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
