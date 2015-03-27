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

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.txtPath.Text = path.SelectedPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //var excelFolderpath = @"D:\嘉善县教育局";
            //var excelFolderpath = @"D:\嘉善县大云镇";
            var excelFolderpath = this.txtPath.Text;

            var lev1Name = excelFolderpath.Split('\\').LastOrDefault();
            var excelSavePath = excelFolderpath + "\\党组织\\" ;
            
            if (!Directory.Exists(excelSavePath))
            {
                Directory.CreateDirectory(excelSavePath);
            }
            
            var pathList = GetFiles(excelFolderpath);

            var tabeList = new List<Data.DataTable>();
            var groupSet = new HashSet<Group>();
            var groupSet2 = new HashSet<Group>();

            // 遍历所有的Excel文件
            foreach (var path in pathList)
            {
                DataSet data = LoadDataFromExcel(path, "J", "K");
                if (data == null)
                {
                    Console.WriteLine("data is null");
                    continue;
                }
                //tabeList.Add(data.Tables[0]);
                var dataTable = data.Tables[0];
                foreach (DataRow row in dataTable.Rows)
                {
                    var name = row[0].ToString().Replace("浙江省", "").Replace("嘉善县", "");
                    var type = row[1].ToString();
                    var belongs = lev1Name.Replace("浙江省", "").Replace("嘉善县", "");

                    if (name.Equals("所属党支部"))
                        continue;
                    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(type))
                    {
                        Console.WriteLine("名称: " + name + " 类型:" + type);
                        continue;
                    }

                    if (!(name.Contains("一支部") || name.Contains("二支部") || name.Contains("三支部") || name.Contains("四支部")
                            || name.Contains("五支部") || name.Contains("六支部") || name.Contains("七支部")
                            || name.Contains("八支部") || name.Contains("九支部") || name.Contains("十支部")
                            || name.Contains("十一支部")))
                    {
                        groupSet.Add(new Group(name, type, belongs));
                    }
                    else
                    {
                        var index = name.IndexOf("支部") - 1;
                        var filter = name.Substring(index);

                        var parentName = name.Replace(filter, "党总支");

                        groupSet.Add(new Group(parentName, type, belongs));
                        groupSet2.Add(new Group(name, type, parentName));
                    }
                }
            }

            //SaveGroupExcel(tabeList,lev1Name, excelSavePath + lev1Name + "二级.xls");
            SaveGroupExcel(groupSet, excelSavePath + "二级.xls");
            SaveGroupExcel(groupSet2, excelSavePath + "三级.xls");

            MessageBox.Show("整理完成", "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var excelFolderpath = this.txtPath.Text;
            var fileName = excelFolderpath.Split('\\').LastOrDefault();
            var excelSavePath = excelFolderpath + "\\合并\\";

            if (!Directory.Exists(excelSavePath))
            {
                Directory.CreateDirectory(excelSavePath);
            }

            var pathList = GetFiles(excelFolderpath);

            Data.DataTable tableMerge = null;
            foreach (var path in pathList)
            {
                DataSet data = LoadDataFromExcel(path, "", "");
                if (data == null)
                {
                    Console.WriteLine("data is null");
                    continue;
                }

                if (tableMerge == null)
                {
                    tableMerge = data.Tables[0].Clone();
                }
                else
                {
                    object[] objArray = new object[tableMerge.Columns.Count];
                    int count = 1;
                    for (int i = 0; i < data.Tables.Count; i++)
                    {
                        for (int j = 0; j < data.Tables[i].Rows.Count; j++)
                        {
                            var rowItem = data.Tables[i].Rows[j].ItemArray;
                            if (string.IsNullOrEmpty(rowItem.FirstOrDefault().ToString()))
                                continue;

                            rowItem[9] = rowItem[9].ToString().Replace("浙江省", "").Replace("嘉善县", "");

                            rowItem.CopyTo(objArray, 0);   //将表的一行的值存放数组中。
                            tableMerge.Rows.Add(objArray);                          //将数组的值添加到新表中。
                        }
                    }
                }
            }

            SaveDataTableToExcel(tableMerge, excelSavePath + fileName + "党员信息合并.xls");

            MessageBox.Show("合并完成", "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                OleDaExcel.Fill(OleDsExcle);
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
            var fileName = filePath.Split('\\').LastOrDefault().Split('.').FirstOrDefault();

            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                wSheet.Name = "党组织信息";
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
                wBook.SaveAs(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, 
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, 
                        Missing.Value, Missing.Value, Missing.Value);
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

        public static bool SaveGroupExcel(HashSet<Group> groupSet, string filePath)
        {
            var newTable = new Data.DataTable();
            newTable.Columns.Add("序号", Type.GetType("System.Int32"));
            newTable.Columns[0].AutoIncrement = true;
            newTable.Columns[0].AutoIncrementSeed = 1;
            newTable.Columns[0].AutoIncrementStep = 1;

            newTable.Columns.Add("组织名称", Type.GetType("System.String"));
            newTable.Columns.Add("党组织类型", Type.GetType("System.String"));
            newTable.Columns.Add("隶属组织", Type.GetType("System.String"));

            foreach (var item in groupSet)
            {
                newTable.Rows.Add(new object[] { null, item.Name, item.Type, item.Belongs });
            }

            return SaveDataTableToExcel(newTable, filePath);
        }

        public static bool SaveGroupExcel(List<Data.DataTable> excelTableList,string parentGroup, string filePath)
        {
            var groupSet = new HashSet<Group>();
            foreach (Data.DataTable dataTable in excelTableList)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    var name = row[0].ToString().Replace("浙江省","").Replace("嘉善县","");
                    var type = row[1].ToString();
                    var belongs = parentGroup.Replace("浙江省", "").Replace("嘉善县", "");

                    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(type))
                    {
                        Console.WriteLine("名称: " + name + " 类型:" + type);
                        continue;
                    }

                    groupSet.Add(new Group(name, type, belongs));
                }

            }

            return SaveGroupExcel(groupSet, filePath);
        }

        private string[] GetFileNames(string folder)//传入参数是文件夹路径
        {
            if (Directory.Exists(folder))
            {
                //文件夹及子文件夹下的所有文件的全路径
                string[] files = Directory.GetFiles(folder, "*.xls", SearchOption.TopDirectoryOnly);
                for (int i = 0; i < files.Length; i++)
                {
                    files[i] = Path.GetFileNameWithoutExtension(files[i]);//只取后缀
                }
                return files;
            }
            else
            {
                return null;
            }
        }

        private string[] GetFiles(string folder)//传入参数是文件夹路径
        {
            if (Directory.Exists(folder))
            {
                //文件夹及子文件夹下的所有文件的全路径
                string[] files = Directory.GetFiles(folder, "*.xls", SearchOption.TopDirectoryOnly);
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
