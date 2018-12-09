using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace OLEDBConnTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;

            DirectoryInfo di = new DirectoryInfo(path);
            //string DbName = "PharmcogeneCombine.mdb";
            string DbName = "clinicalVariant.mdb";
            string connectionString =string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}",Path.Combine(di.Parent.FullName,DbName));


            string targetGeneName = "CYP2C8";
            string targetGenoType = "*1A";
       



            //string SQLcommand = string.Format(@"SELECT {0} FROM TablePharmacogene WHERE GeneName='{1}' AND GenoType='{2}'","Key,GeneName,GenoType,DrugName",targetGeneName,targetGenoType);
            string SQLcommand = string.Format(@"SELECT * FROM DataTable WHERE GeneName='{0}'",targetGeneName);



            using (OleDbConnection connection = new OleDbConnection(connectionString)){

                // Create a command and set its connection  
                OleDbCommand command = new OleDbCommand(SQLcommand, connection);
                // Open the connection and execute the select command.  
                try{
                    // Open connecton  
                    connection.Open();

                    // Execute command  
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        Console.WriteLine("------------Original data----------------");
                        while (reader.Read()){
                            Console.WriteLine("{0}\t{1}\t{2}\t{3}",reader["Key"], reader["GeneName"].ToString(), reader["GenoType"].ToString(), reader["DrugName"].ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                // The connection is automatically closed becasuse of using block.  
            }
            Console.ReadKey();
        }
    }
}