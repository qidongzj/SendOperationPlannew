using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendOperationPlan
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //危急值
            //Application.Run(new FormWjz());

            //手术
            //Application.Run(new Form1());

            //中医优势病种
            Application.Run(new FormZybz());
        }
    }
}
