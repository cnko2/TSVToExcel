using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Reflection;
using System.Dynamic;

namespace TSVToExcel
{
    public static class DataTableObject
    {
        private static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    else
                        continue;
                }
            }
            return obj;
        }

        public static IEnumerable<dynamic> ToExpandoObjectList(this DataTable self)
        {
            var result = new List<dynamic>(self.Rows.Count);
            foreach (var row in self.Rows.OfType<DataRow>())
            {
                var expando = new ExpandoObject() as IDictionary<string, object>;
                foreach (var col in row.Table.Columns.OfType<DataColumn>())
                {
                    expando.Add(col.ColumnName, row[col]);
                }
                result.Add(expando);
            }
            return result;
        }
    }
}
