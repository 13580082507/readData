using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
 读取excel数据的流程
    1.建立连接
    2.打开连接
    3.读取内容
    4.打印内容
 */
namespace readExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            /*1.建立连接*/
            string excelPath = "C:/Users/DELL/Desktop/readData/readExcel/readExcel/zuobaio.xls";
            //string excelPath = "C:/Users/DELL/Desktop/readData/readExcel/readExcel/zuobiao.xlsx";
            //excelPath读取一个excel的文件，把excel文件复制粘贴到项目里面【设置属性如果较新者复制】
            //excelPath是数据的源头，直接写路径和名字
            //Data Source：数据源（可以是数据库的地址）
            //解析方式
            string strConn;
            
            if (excelPath == ".xls")
            {
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelPath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            }
            else
            {
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + excelPath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            }
            
            //使用OleDbConnection和数据源建立连接,对象建立连接
            OleDbConnection connection = new OleDbConnection(strConn);
            //连接字符串，excel有两个版本，旧版本用第一个解析，新版本用第二个
 
                /*2.打开连接*/
                connection.Open();     

            /*3.查询*/
            //查询结果是一个表格
            //使用OleDbDataAdapter适配器做查询
            //OleDbDataAdapter需要传递两个对象:查询命令 + 连接对象
            string sql = "select * from[Sheet1$]";        //这个是查询命令，表示从哪个表查询.从Sheet1查询
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql,connection);
            
            //查询完把数据放到dataset里面
            DataSet dataSet = new DataSet();        //用来存放数据，用来存放dataTable
           
            adapter.Fill(dataSet);          //把查询结果（datatable）放到（填充）到dataset里面
            connection.Close();

            /*4.取数据*/
            DataTableCollection tableCollection =  dataSet.Tables;   //获取当前集合中所有的表格
            //通过索引器访问记录
            DataTable table =  tableCollection[0];         //只有张表格，这里的索引为刚刚查询到的表格
            
            //取的table中所有的行
            DataRowCollection rowCollection = table.Rows;       //返回行的集合

            //取得表中的数据

            //使用循环遍历row集合，取得每一个行的row对象
            foreach (DataRow row in rowCollection)
            {
                //取出row中列的数据，索引0-n；               
                Console.Write("(" + row[0] + " " + row[1] + ")");
                //Console.Write("(" + ((Double)row[0] - 1) + "," + row[1] + ")");

                Console.WriteLine();
            }       
            Console.ReadKey();
        }
    }
}
