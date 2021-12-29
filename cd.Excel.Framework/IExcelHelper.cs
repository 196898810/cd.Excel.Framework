using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace cd.Excel.Framework
{
    public interface IExcelHelper
    {
        #region 读取

        /// <summary>
        /// 读取Excel文件,默认第一个Sheet
        /// </summary>
        /// <returns>DataTable</returns>
        DataTable ExcelToTable();

        /// <summary>
        /// 读取Excel文件,默认第一个Sheet
        /// </summary>
        /// <param name="ex_file">Excel文件路径</param>
        /// <returns>DataTable</returns>
        DataTable ExcelToTable(string ex_file);

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="ex_file">Excel文件路径</param>
        /// <param name="sheetName">Sheet名称</param>
        /// <returns>DataTable</returns>
        DataTable ExcelToTable(string ex_file,string sheetName);


        #endregion

        #region 输出

        /// <summary>
        /// DataSet输出Excel
        /// </summary>
        /// <param name="tmDataSet">DataSet</param>
        /// <returns>输出文件路径</returns>
        string DataSetToExcel(DataSet tmDataSet);

        /// <summary>
        /// DataSet输出Excel
        /// </summary>
        /// <param name="ex_file">输出文件路径</param>
        /// <param name="tmDataSet">DataSet</param>
        /// <returns>输出文件路径</returns>
        string DataSetToExcel(string ex_file, DataSet tmDataSet);

        /// <summary>
        /// DataTable输出Excel
        /// </summary>
        /// <param name="tmDataTable">DataTable</param>
        /// <returns>输出文件路径</returns>
        string DataTableToExcel(DataTable tmDataTable);

        /// <summary>
        /// DataTable输出Excel
        /// </summary>
        /// <param name="ex_file">输出文件路径</param>
        /// <param name="tmDgv">DataTable</param>
        /// <returns>输出文件路径</returns>
        string DataTableToExcel(string ex_file, DataTable tmDataTable);

        /// <summary>
        /// DataGridView输出Excel
        /// </summary>
        /// <param name="tmDataSet">DataGridView</param>
        /// <returns>输出文件路径</returns>
        string DgvToExcel(DataGridView tmDgv);

        /// <summary>
        /// DataGridView输出Excel
        /// </summary>
        /// <param name="tmDataSet">DataGridView</param>
        /// <param name="headText">取HeadText文本</param>
        /// <returns>输出文件路径</returns>
        string DgvToExcel(DataGridView tmDgv,bool headText = true);

        /// <summary>
        /// DataGridView输出Excel
        /// </summary>
        /// <param name="ex_file">输出文件路径</param>
        /// <param name="tmDgv">DataGridView</param>
        /// <param name="headText">取HeadText文本</param>
        /// <returns>输出文件路径</returns>
        string DgvToExcel(string ex_file, DataGridView tmDgv,bool headText = true);

        #endregion
    }
}
