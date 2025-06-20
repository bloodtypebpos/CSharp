using System;
using ClosedXML.Excel;
using System;
using System.ComponentModel;
using System.IO;
using System.Collections.Generic;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Office2010.CustomUI;

class Program
{
    static void Main()
    {
        // Sets current directory to where files are stored
        //Directory.SetCurrentDirectory(@"C:\Users\Matt Tigrett\Desktop\csharp");
        Directory.SetCurrentDirectory(@"C:\Users\Sad_Matt\OneDrive\Desktop\C Sharp\1");

        // SQL local database

        /*
        // This is if the database doesn't exist yet...
        var masterConnectionString = @"Server=(localdb)\MSSQLLocalDB;Integrated Security=true;Initial Catalog=master";
        using (var cnx = new SqlConnection(masterConnectionString))
        {
            cnx.Open();
            var createDbCmd = new SqlCommand(
                "IF DB_ID('InventoryDB') IS NULL CREATE DATABASE InventoryDB;", cnx);
            createDbCmd.ExecuteNonQuery();
        }
        */

        var connectionString = @"Server=(localdb)\MSSQLLocalDB;Integrated Security=true;Initial Catalog=InventoryDB";
        using var connection = new SqlConnection(connectionString);
        connection.Open();


        var cmd = new SqlCommand("SELECT SUSER_NAME()", connection);
        var username = (string)cmd.ExecuteScalar();
        Console.WriteLine("Your SQL login is: " + username);

        /*
        var cmd = new SqlCommand("USE InventoryDB;", connection);
        cmd.ExecuteNonQuery();
        cmd = new SqlCommand("SELECT name FROM sys.tables WHERE name = 'Inventory';", connection);
        using var reader = cmd.ExecuteReader();
        if (reader.HasRows)
        {
            while (reader.Read())
            {
                string tableName = reader.GetString(0);
                Console.WriteLine($"Table found: {tableName}");
            }
        }
        else
        {
            Console.WriteLine("Table 'Inventory' not found...");
        }
        */

        //Dropping the Inventory Table
        var dropSql = "IF OBJECT_ID('dbo.Inventory', 'U') IS NOT NULL DROP TABLE dbo.Inventory;";
        using var dropCmd = new SqlCommand(dropSql, connection);
        dropCmd.ExecuteNonQuery();

        var items = new List<Item>();
        var fields = new List<string>();
        // Open the inventory excel file
        using (var workbook = new XLWorkbook(@"Inventory.xlsx"))
        {
            var worksheet = workbook.Worksheet(1); // First worksheet
            // Get the fields from the header of the worksheet
            items = new List<Item>();
            fields = new List<string>();
            var header = worksheet.FirstRowUsed();
            foreach (var field in header.Cells())
            {
                fields.Add(field.GetString().Trim());
            }
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);
            foreach (var row in rows)
            {
                var item = new Item();
                for (int i = 0; i < fields.Count; i++)
                {
                    var field = fields[i];
                    var rawText = row.Cell(i + 1).GetValue<string>().Trim();
                    var value = ParseValueByField(field, rawText);
                    item.SetAttr(field, value);
                }
                items.Add(item);
            }
        }

        string query = BuildCreateTableSql("Inventory", fields);
        var createCmd = new SqlCommand(query, connection);
        createCmd.ExecuteNonQuery();
        Console.WriteLine("Inventory Table created Successfully.");

        foreach (var item in items)
        {
            using var command = new SqlCommand(BuildInsertCommand(fields), connection);

            for (int i = 0; i < fields.Count; ++i)
            {
                string field = fields[i];
                object value = item.GetAttr(field) ?? DBNull.Value;

                string paramName = "@" + SanitizeParamName(field);
                command.Parameters.AddWithValue(paramName, value);
            }

            command.ExecuteNonQuery();
        }
        Console.WriteLine("Inventory Table Populated Successfully.");

    }

    class Item
    {
        private Dictionary<string, object> _attributes = new();
        public void SetAttr(string key, object value) => _attributes[key] = value;
        public object GetAttr(string key) => _attributes.TryGetValue(key, out var value) ? value : null;
    }

    static string GetSqlType(string fieldName)
    {
        // Basic mapping — you can get fancy later
        return fieldName switch
        {
            "Item ID" => "NVARCHAR(50)",
            "Item Description" => "NVARCHAR(255)",
            "Stocking U/M" => "NVARCHAR(50)",
            "Last Unit Cost" => "DECIMAL(18, 2)",
            "Qty on Hand" => "DECIMAL(18, 2)",
            "Count" => "INT",
            _ => "NVARCHAR(100)" // fallback
        };
    }

    static object ParseValueByField(string fieldName, string rawText)
    {
        return fieldName switch
        {
            "Item ID" => string.IsNullOrWhiteSpace(rawText) ? DBNull.Value : rawText.Trim(),
            "Item Description" => string.IsNullOrWhiteSpace(rawText) ? DBNull.Value : rawText.Trim(),
            "Stocking U/M" => string.IsNullOrWhiteSpace(rawText) ? DBNull.Value : rawText.Trim(),
            "Last Unit Cost" => decimal.TryParse(rawText, out var d1) ? d1 : DBNull.Value,
            "Qty on Hand" => decimal.TryParse(rawText, out var d2) ? d2 : DBNull.Value,
            "Count" => int.TryParse(rawText, out var i) ? i : DBNull.Value,
            _ => string.IsNullOrWhiteSpace(rawText) ? DBNull.Value : rawText
        };
    }

    static string BuildCreateTableSql(string tableName, List<string> fields)
    {
        var query = $"CREATE TABLE {tableName} (";
        foreach (var field in fields)
        {
            query += $"[{field}] {GetSqlType(field)}, ";
        }
        query = query.TrimEnd(',', ' ') + ");";
        return query;
    }

    static string BuildInsertCommand(List<string> fields)
    {
        var columnNames = string.Join(", ", fields.ConvertAll(f => $"[{f}]"));      // Escapes [Item ID]
        var paramNames = string.Join(", ", fields.ConvertAll(f => $"@{SanitizeParamName(f)}"));  // Safe @ItemID

        return $"INSERT INTO Inventory ({columnNames}) VALUES ({paramNames})";
    }

    static string SanitizeParamName(string field)
    {
        return field.Replace(" ", "").Replace("/", "").Replace("-", ""); // tweak as needed
    }


}
