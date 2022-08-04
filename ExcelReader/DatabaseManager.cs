
using System.Data.SqlClient;

namespace ExcelReader
{
    public class DatabaseManager
    {
        const string connectionString = "Server=localhost;Database=master; Integrated Security=true";

        const string excelReaderConnectionString = "Server=localhost;Database=ExcelReaderDb; Integrated Security=true";

        public static void CheckDatabase()
        {
            var dbExists = CheckDatabaseExists();
            if (!dbExists)
            {
                Console.WriteLine("Database does not exist. Creating database.....");
                CreateDatabase();
            }
            else
            {
                var input = UserInputs.GetStringInput("Do you want to delete and recreate your database?. Press: \"y\" to delete the database. Press any other key to continue");
                if (input.ToLower() == "y")
                {
                    try
                    {
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
                    catch (Exception e)
                    {

                        Console.WriteLine(e.Message);
                    }
                }
            }

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
                        Console.WriteLine("Database \"ExcelReaderDb\" created");
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
                        Console.WriteLine("\"Person\" Table created");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public static void WriteExcelToDatabase(List<Person> person)
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

        public static List<Person> Get()
        {
            List<Person> people = new List<Person>();
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
                            people.Add(new Person()
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

        public static bool CheckDatabaseExists()
        {
            using (var connection = new SqlConnection(connectionString))
            {
                using (var command = new SqlCommand($"SELECT db_id('ExcelReaderDb')", connection))
                {
                    connection.Open();
                    return (command.ExecuteScalar() != DBNull.Value);
                }
            }
        }

    }

}
