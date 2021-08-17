using System;
using System.Data;
using System.Data.SqlClient;



namespace emailAlert
{
    class Program
    {
        static void Main(string[] args)
        {
            var datasource = "vmdatabase1";
            var database = "HealthyMaterial";
            var username = "vmdatabaseuser";
            var password = "P3pp3r123!";/* MultipleActiveResultSets=true"*/
            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;
            //SqlConnection conn = new SqlConnection(connString);

            try
            {
                //Console.WriteLine("Openning Connection ...");

                //open connection
                //conn.Open();
                //write the sql query
                ReadData(connString);

                //Console.WriteLine("Finish");
                //ReadData(connString);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }

            Console.Read();

        }

        private static void ReadData(string connectionString)
        {
            //string queryString = "select * from dbo.ExpirationTable;";
            string queryString = "dbo.HMexpireDate;";

            using SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand(queryString, connection);
            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            //Start building html table header
            string expiredHtmlTable =
                //"Expires Status: There are total expired"+ "will expired"+
                "<table>" +
                    "<tr>" +
                        "<th>CSI</th>" +
                        "<th>Manufacturer Name</th>" +
                        "<th>Product Name</th>" +
                        "<th>Material Description</th>" +
                        "<th>Expires Status</th>" +
                        "<th>EPD Status</th>" +
                        "<th>EPD Option Status</th>" +
                        "<th>MT Expiration Status</th>" +
                        "<th>Mat Opt Expires Status</th>" +
                        "<th>LEM Expires Status</th>" +
                    "</tr>";
            while (reader.Read())
            {
                expiredHtmlTable += ReadSingleRow((IDataRecord)reader);
            }
            // close html table
            expiredHtmlTable += "</table>";

            // email
            LocalEmailService LEM = new LocalEmailService();
            //string ListEmails = "jgarcia@pepperconstruction.com, kbritton@pepperconstruction.com";
            string ListEmails = "lzhang@pepperconstruction.com";
            LEM.SendEmail(ListEmails, "HM Expire Material Monthly Alert", expiredHtmlTable);
            reader.Close();
        }

        private static string ReadSingleRow(IDataRecord record)
        {
            //Console.WriteLine(String.Format("{0}, {1},{2}, {3}, {4}, {5}, {6}, {7},{8}",
            //record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7], record[8]));

            string expiredHtmlRow =
                "<tr>" +
                    $"<td>{record[0]}</td>" +
                    $"<td>{record[1]}</td>" +
                    $"<td>{record[2]}</td>" +
                    $"<td>{record[3]}</td>" +
                    $"<td>{record[4]}</td>" +
                    $"<td>{record[5]}</td>" +
                    $"<td>{record[6]}</td>" +
                    $"<td>{record[7]}</td>" +
                    $"<td>{record[8]}</td>" +
                "</tr>";
            return expiredHtmlRow;
        }

    }
}