using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkInsertExcelToMSSqlServer
{
    class Program
    {
        static void Main(string[] args)
        {
            string connetionString;
            SqlConnection cnn;
            connetionString = @"Data Source=192.168.10.2,1433;Initial Catalog=Demodb;User ID=sa;Password=Passw0rd";
            cnn = new SqlConnection(connetionString);

            //create object of SqlBulkCopy which help to insert  
            SqlBulkCopy objbulk = new SqlBulkCopy(cnn);

            //assign Destination table name  
            objbulk.DestinationTableName = "DemoTable";
            objbulk.ColumnMappings.Add("Id", "Id");
            objbulk.ColumnMappings.Add("Name", "Name");
            objbulk.ColumnMappings.Add("DateOfBirth", "DateOfBirth");
            objbulk.ColumnMappings.Add("Age", "Age");


            Workbook wb = new Workbook("Sample.xlsx");
            Worksheet worksheet = wb.Worksheets[0];
            // Exporting the contents into DataTable
            //worksheet.Cells.MaxDataRow
            DataTable dt = worksheet.Cells.ExportDataTable(0, 0, worksheet.Cells.MaxDataRow + 1, worksheet.Cells.MaxDataColumn + 1, true);

            cnn.Open();
            //insert bulk Records into DataBase.  
            objbulk.WriteToServer(dt);
            cnn.Close();
            Console.WriteLine("Records added");
            Console.ReadKey();

        }
    }
}
