using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace EStudio
{
    class DataAccess
    {
        public static string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + 
                                                    System.IO.Directory.GetCurrentDirectory() + "\\Word.mdb";
  
        /// <summary>
        ///读取错误列表数据
        /// </summary>
        /// <param name="sql">连接语句</param>
        /// <returns></returns>
        public static DataTable ReadAllData(string sql)
        {
                using (OleDbConnection odcConnection = new OleDbConnection(connectionString))
                {
                    //打开连接   
                    odcConnection.Open();  
        
                    OleDbDataAdapter ad = new OleDbDataAdapter(sql, odcConnection);

                    DataTable dt = new DataTable();

                    ad.Fill(dt);
                    //关闭连接
                    odcConnection.Close();

                    return dt; 
                }
                
        }
        
        /// <summary>
        ///连接Excel
        /// </summary>
        /// <param name="fileName">Excel表名</param>
        /// <returns>dt</returns>
        public static DataTable ExcelToDataSet(string fileName)
        {
            string strOdbcCon = @"Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;
                        Data Source="+ System.IO.Directory.GetCurrentDirectory() + "\\WORDlibrary\\" + fileName +".xls"+ 
                                     ";extended properties='Excel 8.0;HDR=YES;IMEX=1;'";

            OleDbConnection OleDB = new OleDbConnection(strOdbcCon);

            OleDbDataAdapter OleDat = new OleDbDataAdapter("select * from [Sheet1$]", OleDB);

            DataTable dt = new DataTable();

            OleDat.Fill(dt);

            return dt;
        }

        /// <summary>
        ///插入数据到数据库
        /// </summary>
        /// <param name="word,explain,date">单词，解释，错误日期</param>
        /// <returns></returns>
        public static void InsertData(string word,string explain,string date) 
        {
            String sql = "insert into WrongList(Words,Explain,WrongDate)values( '" + word + "' , '" + explain + "','" + date + "')";

            using (OleDbConnection odcConnection = new OleDbConnection(connectionString))
            {
                OleDbCommand cmd = new OleDbCommand(sql, odcConnection);

                odcConnection.Open();

                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        ///删除数据
        /// </summary>
        /// <param name="ID">删除行的ID</param>
        /// <returns></returns>
        public static void DeleteData(string ID)
        {
            String sql = "delete from WrongList where[ID] =" + ID;

            using (OleDbConnection odcConnection = new OleDbConnection(connectionString))
            {
                OleDbCommand cmd = new OleDbCommand(sql, odcConnection);

                odcConnection.Open();

                cmd.ExecuteNonQuery();
            }
        }
        
    }
    
}
