using System;
using System.Collections.Generic;

using System.Windows.Forms;

namespace Kd_Test
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            frmLogin frmLogin = new frmLogin();

            if (frmLogin.ShowDialog() == DialogResult.OK)
            {
                Application.Run(new KDApiOder());
            }
        }
    }
}
