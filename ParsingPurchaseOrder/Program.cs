using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace ParsingPurchaseOrder
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (Process.GetProcessesByName("ParsingPurchaseOrder").Length > 1)
            {
                Application.Exit();
            }
            else
                Application.Run(new ParsingPO());
        }
    }
}
