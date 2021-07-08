using CsvHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSVToExcel
{
    public class TSVToExcelHelper
    {
        public void executeTsvToCsv(string tsv_filename)
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

    }
}
