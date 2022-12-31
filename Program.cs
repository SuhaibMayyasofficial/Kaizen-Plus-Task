using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Data;

namespace ReadExcelSheet
{
    public class User
    {
        public int Id
        {
            get;
            set;
        }

        public string FullName
        {
            get;
            set;
        }

        public string Email
        {
            get;
            set;
        }

        public bool IsActive
        {
            get;
            set;
        }
    }
    class UserContext : DbContext
    {

        public DbSet<User> Users
        {
            get;
            set;
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Data Source = Users.db");

        }

    }
    class Program
    {

        static void Main(string[] args)
        {
            try
            {
                string path;
                int counter = 0;

                DataTable dtExcelFile = new DataTable();
                dtExcelFile.Columns.Add("FullName");
                dtExcelFile.Columns.Add("Email");
                dtExcelFile.Columns.Add("IsActive");

                Console.WriteLine("Welcome ******************* to our program");

                Console.Write("PLease insert path for the excel file: ");

                path = Console.ReadLine();

                dtExcelFile = ChangeExcelFileToDatatable(path);

                foreach (DataRow dr in dtExcelFile.Rows)
                {

                    if (dr["IsActive"].ToString() == "Yes")

                        dr["IsActive"] = 1;

                    else

                        dr["IsActive"] = 0;
                }

                string connectionString = "Data Source = C:\\Users\\suhai\\source\\repos\\ReadExcelSheet\\ReadExcelSheet\\Users.db";

                using (SqliteConnection con = new SqliteConnection(connectionString))
                {

                    SqliteCommand cmd = new SqliteCommand();

                    foreach (DataRow dr in dtExcelFile.Rows)
                    {

                        cmd.Connection = con;

                        con.Open();
                        cmd.CommandText = "INSERT INTO Users (FullName, Email, IsActive) VALUES (\"" + dr["FullName"] + "\",\"" + dr["Email"] + "\",\"" + dr["IsActive"] + "\")";
                        cmd.ExecuteNonQuery();

                        counter++;

                    }

                    Console.WriteLine("it's done and the total of users were inserted is: " + counter);
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine("Unhandled error: " + ex.Message);
            }

        }

        public static DataTable ChangeExcelFileToDatatable(string path, bool hasHeader = true)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets[0];
                DataTable dtExcelFile = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    dtExcelFile.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = dtExcelFile.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return dtExcelFile;
            }
        }

    }

}