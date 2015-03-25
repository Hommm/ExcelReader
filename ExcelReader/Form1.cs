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

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var excelFolderpath = @"D:\嘉善县教育系统云平台模板20150325";
            var excelSavePath = @"D:\excel.xls";

            var pathList = GetFiles(excelFolderpath);

            var tabeList = new List<Data.DataTable>();
            foreach (var path in pathList)
            {
                DataSet data = LoadDataFromExcel(path,"J","K");
                if(data == null)
                {
                    Console.WriteLine("data is null");
                    continue;
                }
                tabeList.Add(data.Tables[0]);
            }

            SaveGroupExcel(tabeList, excelSavePath);
        }

        //加载Excel 
        public static DataSet LoadDataFromExcel(string filePath,string startNode,string endNode)
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
                OleDaExcel.Fill(OleDsExcle, "党组织信息");
                OleConn.Close();
                return OleDsExcle;
            }
            catch (Exception err)
            {
                MessageBox.Show("数据绑定Excel失败!失败原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
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
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
            }
        }

        public static bool SaveGroupExcel(List<Data.DataTable> excelTableList, string filePath)
        {
            var newTable = new Data.DataTable();
            newTable.Columns.Add("序号", Type.GetType("System.Int32"));
            newTable.Columns[0].AutoIncrement = true;
            newTable.Columns[0].AutoIncrementSeed = 1;
            newTable.Columns[0].AutoIncrementStep = 1;

            newTable.Columns.Add("组织名称", Type.GetType("System.String"));
            newTable.Columns.Add("党组织类型", Type.GetType("System.String"));
            newTable.Columns.Add("隶属组织", Type.GetType("System.String"));

            var parentGroup = "嘉善县教育局";

            var groupSet = new HashSet<Group>();

            foreach (Data.DataTable dataTable in excelTableList)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    var name = row[0].ToString();
                    var type = row[1].ToString();
                    var belongs = parentGroup;

                    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(type))
                    {
                        Console.WriteLine("名称: " + name + " 类型:" + type);
                        continue;
                    }

                    groupSet.Add(new Group(name, type, belongs));
                }

            }
            foreach (var item in groupSet)
            {
                newTable.Rows.Add(new object[] { null, item.Name, item.Type, item.Belongs });
            }

            return SaveDataTableToExcel(newTable, filePath);
        }

        private string[] GetFiles(string folder)//传入参数是文件夹路径
        {
            if (Directory.Exists(folder))
            {
                //文件夹及子文件夹下的所有文件的全路径
                string[] files = Directory.GetFiles(folder, "*.xls", SearchOption.AllDirectories);
                for (int i = 0; i < files.Length; i++)
                {
                    files[i] = Path.GetFullPath(files[i]);//.GetFileNameWithoutExtension(files[i]);//只取后缀
                }
                return files;
            }
            else
            {
                return null;
            }
        }
    }
}
