using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Xml.Linq;

class Program
{
    

    static void Main()
    {
        // Sets current directory to where files are stored
        Directory.SetCurrentDirectory(@"C:\Users\Matt Tigrett\Desktop\csharp");
        //Directory.SetCurrentDirectory(@"C:\Users\Sad_Matt\OneDrive\Desktop\C Sharp\1");

        var connectionString = @"Server=(localdb)\MSSQLLocalDB;Integrated Security=true;Initial Catalog=InventoryDB";
        using var connection = new SqlConnection(connectionString);
        connection.Open();

        // Erase and Repopulate Tables
        ImportExcelToSqlTable("Inventory.xlsx", "Inventory", connection);
        ImportExcelToSqlTable("PurchaseOrderReport.xlsx", "PurchaseOrderReport", connection);
        ImportExcelToSqlTable("OpenSalesOrders.xlsx", "OpenSalesOrders", connection);
        ImportExcelToSqlTable("BOM.xlsx", "BOM", connection);
        ImportExcelToSqlTable("PartLocations.xlsx", "PartLocations", connection);


        Console.WriteLine("====================================================");
        Console.WriteLine("====   GETTING PROCUREMENT DETAILS FROM ORDERS  ====");

        // Get information on what needs to be built
        var query = "SELECT * FROM OpenSalesOrders";
        var openOrders = LoadItemsFromQuery(query, connection);
        foreach (var item in openOrders)
        {
            foreach(var kvp in item.GetAllAttributes())
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }
            Console.WriteLine("-------------------------------------");
        }
    }

    class Item
    {
        private Dictionary<string, object> _attributes = new();
        public void SetAttr(string key, object value) => _attributes[key] = value;
        public object GetAttr(string key) => _attributes.TryGetValue(key, out var value) ? value : null;
        public Dictionary<string, object> GetAllAttributes()
        {
            return new Dictionary<string, object>(_attributes); // copies for safety
        }
    }

    static readonly Dictionary<string, Type> FieldTypes = new()
    {
        ["Item ID"] = typeof(string),
        ["Assembly"] = typeof(string),
        ["Assembly Description"] = typeof(string),
        ["Item Description"] = typeof(string),
        ["Line Description"] = typeof(string),
        ["Stocking U/M"] = typeof(string),
        ["U/M ID"] = typeof(string),
        ["Last Unit Cost"] = typeof(decimal),
        ["Est Cost"] = typeof(decimal),
        ["Qty on Hand"] = typeof(decimal),
        ["Qty Needed"] = typeof(decimal),
        ["Count"] = typeof(int),
        ["SO Date"] = typeof(DateTime),
        ["PO Date"] = typeof(DateTime),
        ["Ship By"] = typeof(DateTime),
        ["Unit Price"] = typeof(decimal),
        ["Qty Ordered"] = typeof(decimal),
        ["Qty Shipped"] = typeof(decimal),
        ["Qty Received"] = typeof(decimal),
        ["Qty Remaining"] = typeof(decimal),
        ["Remaining Amt"] = typeof(decimal),
        ["SO No"] = typeof(string),
        ["PO No"] = typeof(string),
        ["PO State"] = typeof(string),
        ["Vendor ID"] = typeof(string),
        ["Vendor Name"] = typeof(string),
        ["Preferred Vendor"] = typeof(string),
        ["Customer Name"] = typeof(string),
        ["Ship To Name"] = typeof(string),
        ["Ship To City"] = typeof(string),
        ["Ship To State"] = typeof(string),
        ["Location"] = typeof(string),
        ["Code"] = typeof(string),
        ["Type"] = typeof(string),
        ["MTL"] = typeof(string),
        ["Note"] = typeof(string),
        ["Thickness"] = typeof(decimal),
        ["Width"] = typeof(decimal),
        ["Length"] = typeof(decimal),
        ["OD"] = typeof(decimal),
        ["Created On"] = typeof(DateTime)
    };

    static string GetSqlType(string fieldName)
    {
        if (!FieldTypes.TryGetValue(fieldName, out var type))
            type = typeof(string);

        return type switch
        {
            var t when t == typeof(string) => "NVARCHAR(255)",
            var t when t == typeof(int) => "INT",
            var t when t == typeof(decimal) => "DECIMAL(18, 2)",
            var t when t == typeof(DateTime) => "DATETIME",
            _ => "NVARCHAR(100)"
        };
    }

    static object ParseValueByField(string fieldName, string rawText)
    {
        if (!FieldTypes.TryGetValue(fieldName, out var type))
            type = typeof(string);

        if (string.IsNullOrWhiteSpace(rawText))
            return DBNull.Value;

        return type switch
        {
            var t when t == typeof(string) => rawText.Trim(),
            var t when t == typeof(int) => int.TryParse(rawText, out var i) ? i : DBNull.Value,
            var t when t == typeof(decimal) => decimal.TryParse(rawText, out var d) ? d : DBNull.Value,
            var t when t == typeof(DateTime) => DateTime.TryParse(rawText, out var dt) ? dt : DBNull.Value,
            _ => rawText
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

    static string BuildInsertCommand(List<string> fields, string tableName)
    {
        var columnNames = string.Join(", ", fields.ConvertAll(f => $"[{f}]"));
        var paramNames = string.Join(", ", fields.ConvertAll(f => $"@{SanitizeParamName(f)}"));
        return $"INSERT INTO [{tableName}] ({columnNames}) VALUES ({paramNames})";
    }

    static string SanitizeParamName(string field)
    {
        return field.Replace(" ", "").Replace("/", "").Replace("-", "");
    }

    static void ImportExcelToSqlTable(string excelPath, string tableName, SqlConnection connection)
    {
        var dropSql = $"IF OBJECT_ID('dbo.{tableName}', 'U') IS NOT NULL DROP TABLE dbo.{tableName};";
        using var dropCmd = new SqlCommand(dropSql, connection);
        dropCmd.ExecuteNonQuery();

        var items = new List<Item>();
        var fields = new List<string>();

        using (var workbook = new XLWorkbook(excelPath))
        {
            var worksheet = workbook.Worksheet(1);
            var header = worksheet.FirstRowUsed();

            foreach (var field in header.Cells())
                fields.Add(field.GetString().Trim());

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

        var createSql = BuildCreateTableSql(tableName, fields);
        using var createCmd = new SqlCommand(createSql, connection);
        createCmd.ExecuteNonQuery();
        Console.WriteLine($"{tableName} table created successfully.");

        foreach (var item in items)
        {
            using var insertCmd = new SqlCommand(BuildInsertCommand(fields, tableName), connection);

            for (int i = 0; i < fields.Count; i++)
            {
                string field = fields[i];
                object value = item.GetAttr(field) ?? DBNull.Value;
                insertCmd.Parameters.AddWithValue("@" + SanitizeParamName(field), value);
            }

            insertCmd.ExecuteNonQuery();
        }

        Console.WriteLine($"{tableName} table populated successfully.");
    }

    static List<Item> LoadItemsFromQuery(string query, SqlConnection connection)
    {
        var items = new List<Item>();
        using var command = new SqlCommand(query, connection);
        using var reader = command.ExecuteReader();
        var fieldCount = reader.FieldCount;
        while (reader.Read())
        {
            var item = new Item();
            for (int i = 0; i < fieldCount; i++)
            {
                string fieldName = reader.GetName(i);
                object value = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
                item.SetAttr(fieldName, value);
            }
            items.Add(item);
        }
        return items;
    }
}
