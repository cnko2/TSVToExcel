using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TSVToExcel
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string[] args = System.Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                if (System.IO.File.Exists(args[1]))
                {
                    string filePath = args[1];    //包含路徑的檔案名稱

                    var obj = new TSVToExcelHelper();
                    obj.executeTsvToCsv(filePath);
                }
            }
            else
                Application.Run(new Form1());
        }
    }
}
