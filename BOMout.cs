using System;
using ClosedXML.Excel;
using System;
using System.ComponentModel;
using System.IO;
using System.Collections.Generic;
using System.Data.SqlClient;

class Program
{
    static void Main()
    {
        // Sets current directory to where files are stored
        Directory.SetCurrentDirectory(@"C:\Users\Matt Tigrett\Desktop\csharp");

        // SQL local database
        var masterConnectionString = @"Server=(localdb)\MSSQLLocalDB;Integrated Security=true;Initial Catalog=master";
        using (var cnx = new SqlConnection(masterConnectionString))
        {
            cnx.Open();
            var createDbCmd = new SqlCommand(
                "IF DB_ID('InventoryDB') IS NULL CREATE DATABASE InventoryDB;", cnx);
            createDbCmd.ExecuteNonQuery();
        }

        var connectionString = @"Server=(localdb)\MSSQLLocalDB;Integrated Security=true;Initial Catalog=InventoryDB";
        using var connection = new SqlConnection(connectionString);
        connection.Open();
        var dropCmd = new SqlCommand("IF OBJECT_ID('dbo.Inventory', 'U') IS NOT NULL DROP TABLE dbo.Inventoryl", connection);
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
                for (var i = 0; i < fields.Count; i++)
                {
                    var field = fields[i];
                    var value = row.Cell(i + 1).Value;
                    item.SetAttr(field, value);
                }
                items.Add(item);
            }
        }

        string query = BuildCreateTableSql("Inventory", fields);
        var createCmd = new SqlCommand(query, connection);
        createCmd.ExecuteNonQuery();
        Console.WriteLine("Inventory Table created Successfully.");

        foreach(var item in items)
        {
            foreach (var field in fields)
            {
                Console.WriteLine($"{field}: {item.GetAttr(field)}");
            }
            Console.WriteLine("========================================================");
        }


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
}

