using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace HeiheWaterScenario
{
    class ReadExcel
    {
        /// <summary>
        /// 读取excel表
        /// </summary>
        /// <param name="strExcelFileName">excel表路径</param>
        /// <param name="sheetNumber">sheet编号</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string strExcelFileName, int sheetNumber)
        {
            string strConn = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + strExcelFileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'";//NO:第一行为数据，YES:第一行为标题
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            //返回Excel的架构，包括各个sheet表的名称,类型，创建时间和修改时间等　  
            DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            //包含excel中表名的字符串数组  
            string[] strTableNames = new string[dtSheetName.Rows.Count];
            for (int k = 0; k < dtSheetName.Rows.Count; k++)
            {
                strTableNames[k] = dtSheetName.Rows[k]["TABLE_NAME"].ToString();
            }
            OleDbDataAdapter myCommand = null;
            DataTable dt = new DataTable();
            //从指定的表明查询数据,可先把所有表明列出来供用户选择  
            string strExcel = "select*from[" + strTableNames[sheetNumber] + "]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            dt = new DataTable();
            myCommand.Fill(dt);
            return dt;   
        }
    }
}
