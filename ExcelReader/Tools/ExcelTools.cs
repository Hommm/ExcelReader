using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ex = Microsoft.Office.Interop.Excel;
using Data = System.Data;
using ExcelReader.Model;
using System.Reflection;

namespace ExcelReader.Tools
{
    // 定义事件的参数类
    public class ValueEventArgs : EventArgs
    {
        public int Value { set; get; }
    }

    // 定义事件使用的委托
    public delegate void ValueChangedEventHandler(object sender, ValueEventArgs e);

    public class ExcelTools
    {
        // 定义一个事件来提示界面工作的进度
        public event ValueChangedEventHandler ValueChanged;

        // 触发事件的方法
        protected void OnValueChanged(ValueEventArgs e)
        {
            if (this.ValueChanged != null)
            {
                this.ValueChanged(this, e);
            }
        }

        //加载Excel 
        public DataSet LoadDataFromExcel(string filePath, string sheetName, string beginColumn, string endColumn)
        {
            try
            {
                var scope = "";
                if (!string.IsNullOrEmpty(beginColumn) && !string.IsNullOrEmpty(endColumn))
                    scope = beginColumn + ":" + endColumn;

                string strConn;
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath
                    + ";Extended Properties='Excel 8.0;HDR=False;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [" + sheetName + "$" + scope + "]";    //可是更改Sheet名称，比如sheet2，等等 

                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);

                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle);
                OleConn.Close();
                return OleDsExcle;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataSet LoadDataFromExcel(List<string> filePathList, string sheetName, string beginColumn, string endColumn)
        {
            int count = filePathList.Count;
            int i = 0;
            DataSet MergeDataSet = new DataSet();
            foreach (var filePath in filePathList)
            {
                DataSet dataSet = LoadDataFromExcel(filePath,sheetName,beginColumn,endColumn);
                MergeDataSet.Merge(dataSet);
                // 计算进度
                i++;
                int processValue = (int)((1 * 1.0) / count * 100);
                double weight = 0.5;
                // 触发事件
                ValueEventArgs e = new ValueEventArgs() { Value = (int)(processValue * weight) };
                this.OnValueChanged(e);
            }

            return MergeDataSet;
        }

        public void SaveDataTableToExcel(System.Data.DataTable excelTable,string sheetName, string filePath)
        {
            var fileName = filePath.Split('\\').LastOrDefault().Split('.').FirstOrDefault();

            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                wSheet.Name = sheetName;
                if (excelTable.Rows.Count > 0)
                {
                    int rowCount = excelTable.Rows.Count;
                    int colCount = excelTable.Columns.Count;
                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            string str = excelTable.Rows[i][j].ToString();
                            wSheet.Cells[i + 2, j + 1] = str;
                        }

                        // 计算进度
                        int processValue = (int)((1 * 1.0) / rowCount * 100);
                        double weight = 0.5;
                        // 触发事件
                        ValueEventArgs e = new ValueEventArgs() { Value = (int)(processValue * weight) };
                        this.OnValueChanged(e);
                    }
                }

                int size = excelTable.Columns.Count;
                for (int i = 0; i < size; i++)
                {
                    wSheet.Cells[1, 1 + i] = excelTable.Columns[i].ColumnName;
                }
                //设置禁止弹出保存和覆盖的询问提示框 
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                //保存工作簿 
                wBook.SaveAs(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value, Missing.Value);
                //保存excel文件 
                app.Save(filePath);
                app.SaveWorkspace(filePath);
                app.Quit();
                app = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }
        }
    }
}
