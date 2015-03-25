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

namespace ExcelReader.Tools
{
    public static class ExcelTools
    {
        //加载Excel 
        public static DataSet LoadDataFromExcel(string filePath, string startNode, string endNode,string saveSheetName)
        {
            try
            {
                var scope = "";
                if (!string.IsNullOrEmpty(startNode) && !string.IsNullOrEmpty(endNode))
                    scope = startNode + ":" + endNode;

                string strConn;
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath
                    + ";Extended Properties='Excel 8.0;HDR=False;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [党员基础信息$" + scope + "]";//可是更改Sheet名称，比如sheet2，等等 

                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);

                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, saveSheetName);
                OleConn.Close();
                
                return OleDsExcle;
            }
            catch (Exception err)
            {
                throw new Exception("数据绑定Excel失败!失败原因：" + err.Message);
            }
        }

        public static bool SaveDataTableToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int col = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = excelTable.Rows[i][j].ToString();
                            wSheet.Cells[i + 2, j + 1] = str;
                        }
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
                wBook.Save();
                //保存excel文件 
                app.Save(filePath);
                app.SaveWorkspace(filePath);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                Console.WriteLine("导出Excel出错！错误原因：" + err.Message);
                return false;
            }
            finally
            {
            }
        }
    }
}
