using System;
using System.Data;
using System.Data.SqlClient;
using ClosedXML.Excel;


namespace emailAlert
{
    class Program
    {
        static void Main(string[] args)
        {
            var datasource = "vmdatabase1";
            var database = "HealthyMaterial_V2";
            var username = "vmdatabaseuser";
            var password = "P3pp3r123!";
            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;

            try
            {
                ReadData(connString);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
            Console.ReadKey();

        }

        private static void ReadData(string connectionString)
        {
            string queryString = "dbo.HMexpireDate;";

            using SqlConnection connection = new(connectionString);
            SqlCommand command = new(queryString, connection);
            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Expirations");
                int row = 2;
                string[] columns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I" };
                string cell = "";

                worksheet.Cell("A1").Value = "CSI";
                worksheet.Cell("B1").Value = "Manufacturer Name";
                worksheet.Cell("C1").Value = "Product Name";
                worksheet.Cell("D1").Value = "Material Description";
                worksheet.Cell("E1").Value = "EPD Status";
                worksheet.Cell("F1").Value = "EPD Option Status";
                worksheet.Cell("G1").Value = "MT Expiration Status";
                worksheet.Cell("H1").Value = "Mat Opt Expires Status";
                worksheet.Cell("I1").Value = "LEM Expires Status";

                while (reader.Read())
                {
                    for (int j = 0; j < 9; j++)
                    {
                        cell = $"{columns[j]}{row}";
                        worksheet.Cell(cell).Value = reader[j];
                        if (worksheet.Cell(cell).Value is not null) {
                            worksheet.Cell(cell).Value = worksheet.Cell(cell).Value.ToString();
                        }
                    }
                    row++;
                }
                
                workbook.SaveAs(@"C:\PepperPepper\HM_Expired.xlsx");
            }
            Console.WriteLine("Done writing to excel");

            // email
            string[] ListEmails = new string[] { "lzhang@pepperconstruction.com", "jgarcia@pepperconstruction.com", "kbritton@pepperconstruction.com", "DDo@pepperconstruction.com" };
            //string[] ListEmails = new string[] { "lzhang@pepperconstruction.com"};
            LocalEmailService.SendEmail(ListEmails, "HM Expire Material Monthly Alert", "Please see attached Expire Status.");
        }
    }
}