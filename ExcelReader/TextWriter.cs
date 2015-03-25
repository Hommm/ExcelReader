using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public static class TextWriter
    {
        public static void Writer(DataTable table,string path)
        {
            StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
            String DataRow = "";
            for (int i = 0; i < table.Columns.Count; i++) //获取列名 
            {
                DataRow += table.Columns[i].ColumnName;
                if (i < table.Columns.Count - 1)
                    DataRow += " ";
            }
            sw.WriteLine(DataRow);
            for (int i = 0; i < table.Rows.Count; i++) //获取数据 
            {
                DataRow = "";
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    DataRow += table.Rows[i][j].ToString();
                    if (j < table.Columns.Count - 1) DataRow += " ";
                }
                sw.WriteLine(DataRow);
            }
            sw.Close();
        }
    }
}
