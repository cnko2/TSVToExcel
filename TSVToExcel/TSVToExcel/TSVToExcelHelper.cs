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
        #region POCO
        public string col_1 { get; set; }
        public string col_2 { get; set; }
        public string col_3 { get; set; }
        public string col_4 { get; set; }
        public string col_5 { get; set; }
        public string col_6 { get; set; }
        public string col_7 { get; set; }
        public string col_8 { get; set; }
        public string col_9 { get; set; }
        public string col_10 { get; set; }
        public string col_11 { get; set; }
        public string col_12 { get; set; }
        public string col_13 { get; set; }
        public string col_14 { get; set; }
        public string col_15 { get; set; }
        public string col_16 { get; set; }
        public string col_17 { get; set; }
        public string col_18 { get; set; }
        public string col_19 { get; set; }
        public string col_20 { get; set; }
        public string col_21 { get; set; }
        public string col_22 { get; set; }
        public string col_23 { get; set; }
        public string col_24 { get; set; }
        public string col_25 { get; set; }
        public string col_26 { get; set; }
        public string col_27 { get; set; }
        public string col_28 { get; set; }
        public string col_29 { get; set; }
        public string col_30 { get; set; }
        public string col_31 { get; set; }
        public string col_32 { get; set; }
        public string col_33 { get; set; }
        public string col_34 { get; set; }
        public string col_35 { get; set; }
        public string col_36 { get; set; }
        public string col_37 { get; set; }
        public string col_38 { get; set; }
        public string col_39 { get; set; }
        public string col_40 { get; set; }
        public string col_41 { get; set; }
        public string col_42 { get; set; }
        public string col_43 { get; set; }
        public string col_44 { get; set; }
        public string col_45 { get; set; }
        public string col_46 { get; set; }
        public string col_47 { get; set; }
        public string col_48 { get; set; }
        public string col_49 { get; set; }
        public string col_50 { get; set; }
        public string col_51 { get; set; }
        public string col_52 { get; set; }
        public string col_53 { get; set; }
        public string col_54 { get; set; }
        public string col_55 { get; set; }
        public string col_56 { get; set; }
        public string col_57 { get; set; }
        public string col_58 { get; set; }
        public string col_59 { get; set; }
        public string col_60 { get; set; }
        public string col_61 { get; set; }
        public string col_62 { get; set; }
        public string col_63 { get; set; }
        public string col_64 { get; set; }
        public string col_65 { get; set; }
        public string col_66 { get; set; }
        public string col_67 { get; set; }
        public string col_68 { get; set; }
        public string col_69 { get; set; }
        public string col_70 { get; set; }
        public string col_71 { get; set; }
        public string col_72 { get; set; }
        public string col_73 { get; set; }
        public string col_74 { get; set; }
        public string col_75 { get; set; }
        public string col_76 { get; set; }
        public string col_77 { get; set; }
        public string col_78 { get; set; }
        public string col_79 { get; set; }
        public string col_80 { get; set; }
        public string col_81 { get; set; }
        public string col_82 { get; set; }
        public string col_83 { get; set; }
        public string col_84 { get; set; }
        public string col_85 { get; set; }
        public string col_86 { get; set; }
        public string col_87 { get; set; }
        public string col_88 { get; set; }
        public string col_89 { get; set; }
        public string col_90 { get; set; }
        public string col_91 { get; set; }
        public string col_92 { get; set; }
        public string col_93 { get; set; }
        public string col_94 { get; set; }
        public string col_95 { get; set; }
        public string col_96 { get; set; }
        public string col_97 { get; set; }
        public string col_98 { get; set; }
        public string col_99 { get; set; }
        public string col_100 { get; set; }
        public string col_101 { get; set; }
        public string col_102 { get; set; }
        public string col_103 { get; set; }
        public string col_104 { get; set; }
        public string col_105 { get; set; }
        public string col_106 { get; set; }
        public string col_107 { get; set; }
        public string col_108 { get; set; }
        public string col_109 { get; set; }
        public string col_110 { get; set; }
        public string col_111 { get; set; }
        public string col_112 { get; set; }
        public string col_113 { get; set; }
        public string col_114 { get; set; }
        public string col_115 { get; set; }
        public string col_116 { get; set; }
        public string col_117 { get; set; }
        public string col_118 { get; set; }
        public string col_119 { get; set; }
        public string col_120 { get; set; }
        public string col_121 { get; set; }
        #endregion

        public void executeTsvToCsv(string tsv_filename)
        {
            string csv_filename = tsv_filename + DateTime.Now.Ticks + ".csv";
            DataTable dt = ReadCsvToDatatable(tsv_filename);
            ReadCsvToExcel(dt, csv_filename);

            //var records = ReadCsvToList(tsv_filename);
            //ReadCsvToExcel(records, tsv_filename);
            OpenCSVFile(csv_filename);
        }

       

        private DataTable ReadCsvToDatatable(string filename)
        {
            //string filename = @"d:\temp\temp.tsv";

            var myconfig = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture);
            myconfig.Delimiter = "\t";
            myconfig.HasHeaderRecord = true;
            myconfig.MissingFieldFound = null;
            myconfig.BadDataFound = null;

            //using (var stream = new StreamReader(filename, Encoding.Default))
            using (var stream = new StreamReader(filename, Encoding.UTF8))
            using (var csv = new CsvReader(stream, myconfig))
            using (var datareader = new CsvDataReader(csv))
            using (var dt = new DataTable())
            {
                dt.Load(datareader);
                return dt;
            }


        }

        private List<TSVToExcelHelper> ReadCsvToList(string filename)
        {
            //string filename = @"d:\temp\temp.tsv";

            var myconfig = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture);
            myconfig.Delimiter = "\t";
            myconfig.HasHeaderRecord =false;
            myconfig.MissingFieldFound = null;
            //myconfig.HeaderValidated = null;
            myconfig.BadDataFound = null;

            var records = new List<TSVToExcelHelper>();

            using (var stream = new StreamReader(filename, Encoding.Default))
            using (var csv = new CsvReader(stream, myconfig))
            {
                while (csv.Read())
                {
                    var record = csv.GetRecord<TSVToExcelHelper>();
                    records.Add(record);
                }
            };
            return records;
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

        private void ReadCsvToExcel(List<TSVToExcelHelper> records, string output_filename)
        {
            string filename = @"d:\temp\temp.csv";


            //var records = new List<dynamic>();

            ////dynamic record = new ExpandoObject();
            ////record.Id = 1;
            ////record.Name = "one";
            ////records.Add(record);

            //dynamic record = DataTableObject.ToExpandoObjectList(dt);


            using (var writer = new StreamWriter(output_filename, false, Encoding.Default))
            {
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(records);
                }
            }

            var oldLines = System.IO.File.ReadAllLines(output_filename);
            var newLines = oldLines.Where(line => !(line.IndexOf("col_1,col_2,_col_3,col_4,col_5,col_6,_col_7,col_8")>-1));
            System.IO.File.WriteAllLines(output_filename, newLines);
            FileStream obj = new FileStream(output_filename, FileMode.Append);
            obj.Close();
            // once deleted the selected line and once again read the text file and diplay the new text file in listBox  
            FileInfo fi = new FileInfo(output_filename);
    
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
