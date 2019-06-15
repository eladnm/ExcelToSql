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


                        SqlConnection conn = new SqlConnection(@"Data Source=ELAD\SQLEXPRESS;Initial Catalog=elad;Integrated Security=True;MultipleActiveResultSets = True;");
                        SqlCommand cmd = new SqlCommand("create table ##MyTable(Name varchar(50))", conn);
                        SqlCommand cmd2 = new SqlCommand("select * from tempdb.dbo.##MyTable join elad.dbo.employees on elad.dbo.employees.first_name=tempdb.dbo.##MyTable.Name", conn);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        SqlBulkCopy bulkCopy = new SqlBulkCopy(conn);
                        bulkCopy.DestinationTableName = "##MyTable";
                        bulkCopy.WriteToServer(dReader);
                        //conn.Close();
                        //conn.Open();
                        SqlDataReader dr2 =cmd2.ExecuteReader();
                        while (dr2.Read())
                        {
                            Console.WriteLine(String.Format("{0}", dr2[0]));
                        }
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

