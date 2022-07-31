
using System.Data.SqlClient;

namespace ExcelReader
{
    public class DatabaseManager
    {
        const string connectionString = "Server=(localdb)\\Local;Integrated Security=True";

        const string excelReaderConnectionString = "Server=(localdb)\\Local;Initial Catalog=ExcelReaderDb; Integrated Security=True";

        public static void CheckDatabase()
        {
            // This is an application that will read data from an Excel spreadsheet into a database
            // When the application starts, it should delete the database if it exists,
            // create a new one, create all tables, read from Excel, seed into the database.
            // You shouldn't read into Json first.
            // Once the database is populated, you'll fetch data from it and show it in the console.
            // You should tell the user what it's doing: (i.e. reading from excel; creating tables, etc)
            // The application will be written for a known table, you don't need to make it dynamic.
            // You should tell the user what it's doing: (i.e. reading from excel; creating tables, etc)


            //var input = UserInputs.GetStringInput("Do you want to delete database, perss y: ");
            //if (input.ToLower() == "y")
            using (var connection = new SqlConnection(connectionString))
            {
                using (var tableCmd = connection.CreateCommand())
                {
                    connection.Open();
                    tableCmd.CommandText =
                    $@"DROP DATABASE IF EXISTS ExcelReaderDb";
                    tableCmd.ExecuteNonQuery();
                }
            }
            CreateDatabase();
        }

        private static void CreateDatabase()
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    using (var tableCmd = connection.CreateCommand())
                    {
                        connection.Open();
                        tableCmd.CommandText =
                        $@"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'ExcelReaderDb')
                           BEGIN
                             CREATE DATABASE ExcelReaderDb;
                           END;
                         ";
                        tableCmd.ExecuteNonQuery();
                    }
                }
                CreateTable();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void CreateTable()
        {
            try
            {
                using (var connection = new SqlConnection(excelReaderConnectionString))
                {
                    using (var tableCmd = connection.CreateCommand())
                    {
                        connection.Open();
                        tableCmd.CommandText =
                        $@" IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Person')
                        CREATE TABLE Person (
                          Id int IDENTITY(1,1) NOT NULL PRIMARY KEY,
                          FirstName varchar(50) NOT NULL,
                          LastName varchar(50) NOT NULL,
                          Age int NOT NULL,
                          City varchar(50) NOT NULL,
                            );";
                        tableCmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public void WriteExcelToDatabase(List<Person> person)
        {
            try
            {
                foreach (var p in person)
                {
                    using (var connection = new SqlConnection(excelReaderConnectionString))
                    {
                        using (var tableCmd = connection.CreateCommand())
                        {
                            connection.Open();
                            tableCmd.CommandText = $@"INSERT INTO PERSON(FirstName, LastName, Age, City) 
                            Values('{p.FirstName}', '{p.LastName}', '{p.Age}', '{p.City}')";
                            tableCmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static List<PersonDto> Get()
        {
            List<PersonDto> people = new List<PersonDto>();
            using (var connection = new SqlConnection(excelReaderConnectionString))
            {
                using (var tableCmd = connection.CreateCommand())
                {
                    connection.Open();
                    tableCmd.CommandText = $@"SELECT * FROM Person";
                    SqlDataReader reader = tableCmd.ExecuteReader();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            people.Add(new PersonDto()
                            {
                                Id = reader.GetInt32(0),
                                FirstName = reader.GetString(1),
                                LastName = reader.GetString(2),
                                Age = reader.GetInt32(3),
                                City = reader.GetString(4),
                            });
                        }
                    }
                    else
                    {
                        Console.WriteLine("No rows");
                    }
                }
            }
            return people;
        }
    }
}
