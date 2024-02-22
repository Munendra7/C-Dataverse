using System;
using System.Data;
using System.Data.SqlClient;

class Program
{
    static void Main(string[] args)
    {
        // Connection string for connecting to SQL Server
        string connectionString = "Server=your_server;Database=your_database;Integrated Security=True;";

        // User to impersonate (Windows account)
        string userName = "domain\\username"; // Replace with the desired Windows user

        // Execute CRUD operations under the security context of the specified user
        using (ImpersonationContext context = new ImpersonationContext(userName))
        {
            // Example CRUD operations
            try
            {
                // Insert operation
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand insertCommand = new SqlCommand("INSERT INTO YourTable (Column1, Column2) VALUES (@Value1, @Value2)", connection);
                    insertCommand.Parameters.AddWithValue("@Value1", "Value1");
                    insertCommand.Parameters.AddWithValue("@Value2", "Value2");
                    insertCommand.ExecuteNonQuery();
                    Console.WriteLine("Record inserted successfully.");
                }

                // Read operation
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand selectCommand = new SqlCommand("SELECT * FROM YourTable", connection);
                    SqlDataReader reader = selectCommand.ExecuteReader();
                    while (reader.Read())
                    {
                        Console.WriteLine($"Column1: {reader["Column1"]}, Column2: {reader["Column2"]}");
                    }
                }

                // Update operation
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand updateCommand = new SqlCommand("UPDATE YourTable SET Column1 = @UpdatedValue WHERE Column2 = @Value2", connection);
                    updateCommand.Parameters.AddWithValue("@UpdatedValue", "UpdatedValue");
                    updateCommand.Parameters.AddWithValue("@Value2", "Value2");
                    updateCommand.ExecuteNonQuery();
                    Console.WriteLine("Record updated successfully.");
                }

                // Delete operation
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand deleteCommand = new SqlCommand("DELETE FROM YourTable WHERE Column1 = @Value1", connection);
                    deleteCommand.Parameters.AddWithValue("@Value1", "Value1");
                    deleteCommand.ExecuteNonQuery();
                    Console.WriteLine("Record deleted successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}

// Helper class for impersonation
public class ImpersonationContext : IDisposable
{
    private readonly System.Security.Principal.WindowsImpersonationContext _impersonationContext;

    public ImpersonationContext(string userName)
    {
        _impersonationContext = new System.Security.Principal.WindowsIdentity(userName).Impersonate();
    }

    public void Dispose()
    {
        if (_impersonationContext != null)
        {
            _impersonationContext.Undo();
            _impersonationContext.Dispose();
        }
    }
}
