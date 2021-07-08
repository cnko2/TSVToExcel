using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CsvHelper;
using Microsoft.Win32;

namespace TSVToExcel
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// APP 執行檔案位置
        /// </summary>
        private string _app_exec_path = "";

        public Form1()
        {
            InitializeComponent();

            

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //this.Visible = false;
            string[] args = System.Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                _app_exec_path = args[0];

                if (System.IO.File.Exists(args[1]))
                {
                    string filePath = args[1];    //包含路徑的檔案名稱
                    executeTsvToCsv(filePath);
                }
                Application.Exit();

            }
            else
            {
               // this.Visible = true;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filename = @"d:\temp\temp.tsv";
            string csvfilename = @"d:\temp\temp.tsv"+ ".csv";
            DataTable dt =  ReadCsvToDatatable(filename);
            ReadCsvToExcel(dt, csvfilename);

            OpenCSVFile(csvfilename);
        }

        private void executeTsvToCsv(string tsv_filename)
        {
            string csv_filename = tsv_filename + DateTime.Now.Ticks + ".csv";
            DataTable dt = ReadCsvToDatatable(tsv_filename);
            ReadCsvToExcel(dt, csv_filename);
            OpenCSVFile(csv_filename);
        }

        private void ReadCsv()
        {
            string filename = @"d:\temp\temp.tsv";

            var myconfig = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture);
            myconfig.Delimiter = "\t";
            myconfig.HasHeaderRecord = true;

            using (var stream = new StreamReader(filename, Encoding.Default))
            using (var csv = new CsvReader(stream, myconfig))
            {
                //CsvHelper.Configuration.CsvConfiguration myconfig = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture);
                //myconfig.Delimiter = "\t";
                //myconfig.HasHeaderRecord = true;
                //stream.Configuration.Delimiter = "\t";

                DataTable t = new DataTable();
                var dataReader = new CsvDataReader(csv, t);

                var schemaTable = dataReader.GetSchemaTable();
            }
        }

        private DataTable ReadCsvToDatatable(string filename)
        {
            //string filename = @"d:\temp\temp.tsv";

            var myconfig = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture);
            myconfig.Delimiter = "\t";
            myconfig.HasHeaderRecord = true;

            using (var stream = new StreamReader(filename, Encoding.Default))
            using (var csv = new CsvReader(stream, myconfig))
            using (var datareader = new CsvDataReader(csv))
            using (var dt = new DataTable())
            {
                dt.Load(datareader);
                return dt;
            }

             
        }

        private void ReadCsvToExcel(DataTable dt, string output_filename)
        {
            string filename = @"d:\temp\temp.csv";

            
                var records = new List<dynamic>();

            //dynamic record = new ExpandoObject();
            //record.Id = 1;
            //record.Name = "one";
            //records.Add(record);

            dynamic record = DataTableObject.ToExpandoObjectList(dt);


            using (var writer = new StreamWriter(output_filename, false, Encoding.Default))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(record);


            }
        }

        private void OpenCSVFile(string filename)
        {
            FileInfo fi = new FileInfo(filename);
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename);
            }
            else
            {
                //file doesn't exist
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Registry.ClassesRoot.CreateSubKey(".tsv").SetValue("", "tsv_to_csv", RegistryValueKind.String); //步驟1,2
            Registry.ClassesRoot.CreateSubKey("tsv_to_csv\\shell\\open\\command").SetValue("", Application.ExecutablePath + " %1", RegistryValueKind.ExpandString); //步驟3,4,5
            MessageBox.Show("tsv 檔案已設定關聯至本程式");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Registry.ClassesRoot.DeleteSubKey(".tsv");//.SetValue("", "tsv_to_csv", RegistryValueKind.String); //步驟1,2
            Registry.ClassesRoot.DeleteSubKey("tsv_to_csv\\shell\\open\\command");//.SetValue("", Application.ExecutablePath + " %1", RegistryValueKind.ExpandString); //步驟3,4,5
            MessageBox.Show("tsv 已解除關聮至本程式");
        }
    }
}
