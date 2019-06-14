using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {


            string relativePath = "\\example.xlsx";
            string baseDirectory = "C:\\Users\\eladn\\Desktop\\excel";
            string absolutePath = Path.GetFullPath(baseDirectory + relativePath);

            String strConnection = @"Data Source=ELAD\SQLEXPRESS;Initial Catalog=elad;Integrated Security=True;";

            String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", absolutePath);
            using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
            {
                using (System.Data.OleDb.OleDbCommand cmdd = new OleDbCommand("Select [Name] from [1$]", excelConnection))
                {
                    excelConnection.Open();
                    using (OleDbDataReader dReader = cmdd.ExecuteReader())
                    {


                        SqlConnection conn = new SqlConnection(@"Data Source=ELAD\SQLEXPRESS;Initial Catalog=elad;Integrated Security=True;");
                        SqlCommand cmd = new SqlCommand("create table #MyTempTable(SomeColumn varchar(50))", conn);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        SqlBulkCopy bulkCopy = new SqlBulkCopy(conn);
                        bulkCopy.DestinationTableName = "#MyTempTable";
                        bulkCopy.WriteToServer(dReader);
                        conn.Close();
                    }
                }
            }
        }
    }
}
//            using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
//            {
//                Create OleDbCommand to fetch data from Excel
//                using (System.Data.OleDb.OleDbCommand cmd = new OleDbCommand("Select [Name] from [1$]", excelConnection))
//                {
//                    excelConnection.Open();
//                    using (OleDbDataReader dReader = cmd.ExecuteReader())
//                    {
//                        using (SqlBulkCopy sqlBulk = new SqlBulkCopy(strConnection))
//                        {
//                            Give your Destination table name
//                            sqlBulk.DestinationTableName = "department";
//                            sqlBulk.WriteToServer(dReader);
//                        }
//                    }
//                }
//            }
//        }
//    }
//}

