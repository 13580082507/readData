﻿using System;
using System.Data;
using System.Data.OleDb;
using System.Timers;

namespace practiceReadExcel
{
    class Program
    {
        private static int count = 1;
        private static DataTable table;
        //private static string excelPath1 = "C:/Users/DELL/Desktop/readData/readExcel/practiceReadExcel/zb.xlsx";
        //private static string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + excelPath1 + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
        private static string excelPath1 = "C:/Users/DELL/Desktop/readData/readExcel/practiceReadExcel/zb.xls";
        private static string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelPath1 + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
        private static string sql = "select * from[Sheet1$]";
        static Timer timer = new Timer(); //实例化Timer类，设置间隔时间为10000毫秒；

        static void Main(string[] args)
        {
            timer.Interval = 500;
            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Elapsed += timer_Elapsed;
            Console.Read();
        }
        public static void timer_Elapsed(object source, System.Timers.ElapsedEventArgs e)
          {
            try {
                //timer.Stop();
            //将要读取的表格放进connection对象
               
                OleDbConnection connection = new OleDbConnection(strConn);
                try 
                { 
                    connection.Open();
                }
                catch(Exception ex)
                {
                    Console.Write(ex);
                }
                
                OleDbDataAdapter adapter = new OleDbDataAdapter(sql, connection);
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet);
                connection.Close();

                DataTableCollection dataTableCollection = dataSet.Tables;
                table = dataTableCollection[0];
                
                DataRow row;
                row = table.Rows[0];
                Console.Write(count + ":" + "(" + row[0] + "," + row[1] + ")");
                Console.WriteLine();
                count++;
                //timer.Start();
            }
            catch (Exception ex)
            {
                Console.Write(ex);
            }
        }
    }
}
