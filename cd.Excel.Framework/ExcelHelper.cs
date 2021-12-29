using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace cd.Excel.Framework
{
    public class ExcelHelper : IExcelHelper
    {
        public string DataSetToExcel(DataSet tmDataSet)
        {
            //申明保存对话框
            SaveFileDialog dlg = new SaveFileDialog();
            //默然文件后缀
            dlg.DefaultExt = "xlsx";
            //文件后缀列表
            dlg.Filter = "EXCEL文件(*.XLSX)|*.xlsx";
            //默然路径是系统当前路径
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //打开保存对话框
            dlg.ShowDialog();
            //返回文件路径
            string ex_file = dlg.FileName;
            //验证strFileName是否为空或值无效
            if (ex_file.Trim() == "") return "";

            if (tmDataSet.Tables.Count == 0)
            {
                return "";
            }

            return DataSetToExcel(ex_file, tmDataSet);
        }

        public string DataSetToExcel(string ex_file, DataSet tmDataSet)
        {
            //创建一个workbook对象，默认创建03版的Excel
            Workbook workbook = new Workbook();

            //指定版本信息，07及以上版本最多可以插入1048576行数据
            workbook.Version = ExcelVersion.Version2013;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
                workbook.Worksheets.Remove(0);
            //获取第一张sheet
            // Worksheet sheet = workbook.Worksheets[0];

            for (int i = 0; i < tmDataSet.Tables.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets.Add(tmDataSet.Tables[i].TableName);
                //得到在datatable里的数据
                DataTable dt = tmDataSet.Tables[i];

                //从第一行第一列开始插入数据，true代表数据包含列名
                sheet.InsertDataTable(dt, true, 1, 1);
                sheet.AllocatedRange.AutoFitColumns();
            }
            workbook.Worksheets.Remove(0);
            //保存文件
            workbook.SaveToFile(ex_file, ExcelVersion.Version2013);

            return ex_file;
        }

        public string DataTableToExcel(DataTable tmDataTable)
        {
            DataSet dataSet=new DataSet();
            dataSet.Tables.Add(tmDataTable);
            return DataSetToExcel(dataSet);
        }

        public string DataTableToExcel(string ex_file, DataTable tmDataTable)
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tmDataTable);
            return DataSetToExcel(ex_file,dataSet);
        }

        public string DgvToExcel(DataGridView tmDgv)
        {
            return DgvToExcel(tmDgv, true);
        }

        public string DgvToExcel(DataGridView tmDgv, bool headText = true)
        {
            return DataTableToExcel(SaveDgvToTable(tmDgv, headText));
        }

        public string DgvToExcel(string ex_file, DataGridView tmDgv, bool headText = true)
        {
            return DataTableToExcel(ex_file,SaveDgvToTable(tmDgv, headText));
        }

        public DataTable ExcelToTable()
        {
            string fileName = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "选择文件";
            ofd.Filter = "*.xlsx|*.xlsx|*.xls|*.xls";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = ofd.FileName;
                return ExcelToTable(fileName, null);
            }
            else
                return null;

          
        }

        public DataTable ExcelToTable(string ex_file)
        {
            return ExcelToTable(ex_file, null);
        }

        public DataTable ExcelToTable(string ex_file, string sheetName)
        {
            DataTable dt = new DataTable();
            Workbook workbook = new Workbook();
            try
            {
                if (ex_file.Contains(".xlsx"))
                    workbook.LoadFromFile(ex_file, ExcelVersion.Version2007);
                else
                    workbook.LoadFromFile(ex_file, ExcelVersion.Version97to2003);
                //获取第N张sheet
                Worksheet sheet = null;
                if (sheetName == null)
                    sheet = workbook.Worksheets[0];
                else
                    sheet = workbook.Worksheets[sheetName];
                if (sheet.LastRow == -1)
                    return null;
                //设置range范围
                CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];

                //输出数据, 同时输出列名以及公式值
                dt = sheet.ExportDataTable(range, true, true);
                RemoveEmpty(dt);
            }
            catch (Exception ex)
            {
                throw ex;
            }



            return dt;
        }

        protected DataTable SaveDgvToTable(DataGridView dgv ,bool headText = true)
        {
            DataTable dt = new DataTable();

            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                if (dgv.Columns[count].Visible == true)
                {
                    string columnName = "";
                    if(headText == true)
                    {
                        columnName = dgv.Columns[count].HeaderText.ToString();
                    }
                    else
                    {
                        columnName = dgv.Columns[count].Name.ToString();
                    }
                    DataColumn dc = new DataColumn(columnName);
                    dt.Columns.Add(dc);
                }

            }

            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                int r = 0;
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    if (dgv.Columns[countsub].Visible == true)
                    {
                        dr[r] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                        r++;
                    }

                }
                dt.Rows.Add(dr);
            }
            return dt;
        }


        protected void RemoveEmpty(DataTable dt)
        {
            List<DataRow> removeList = new List<DataRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {
                        rowdataisnull = false;
                    }
                }
                if (rowdataisnull)
                {
                    removeList.Add(dt.Rows[i]);
                }
            }
            for (int i = 0; i < removeList.Count; i++)
            {
                dt.Rows.Remove(removeList[i]);
            }
        }
    }
}
