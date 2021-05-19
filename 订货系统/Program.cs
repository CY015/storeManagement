using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 订货系统
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
            confirmID cID = new confirmID();
            cID.ShowDialog();
            if (cID.DialogResult == DialogResult.OK)
            {
                Application.Run(new mainForm());
                //string MySqlCon = "Data Source=DESKTOP-5NV5EFJ;Initial Catalog=OrderSystem;Integrated Security=True";
            }
        }
    }
}
