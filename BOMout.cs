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
        //ImportExcelToSqlTable("BOM.xlsx", "BOM", connection); // Should only have to do this once a month or so
        ImportExcelToSqlTable("PartLocations.xlsx", "PartLocations", connection);


        Console.WriteLine("====================================================");
        Console.WriteLine("====   GETTING PROCUREMENT DETAILS FROM ORDERS  ====");

        // Get information on what needs to be built
        var partsDict = new Dictionary<string, decimal>();

        var openOrders = LoadItemsFromQuery(
            "SELECT [SO No], [Item ID], [Qty Remaining] FROM [OpenSalesOrders]", connection);

        foreach (var orderItem in openOrders)
        {
            string itemId = orderItem.GetAttr("Item ID")?.ToString();
            decimal qtyRemaining = orderItem.GetAttr("Qty Remaining") is decimal d ? d : 0;

            if (string.IsNullOrWhiteSpace(itemId) || qtyRemaining <= 0)
                continue;

            // Check BOM for subparts of this item (i.e., see if it's an assembly)
            string bomQuery = $"SELECT [Item ID], [Qty Needed] FROM [BOM] WHERE [Assembly] = @assembly";
            using var cmd = new SqlCommand(bomQuery, connection);
            cmd.Parameters.AddWithValue("@assembly", itemId);
            using var reader = cmd.ExecuteReader();
            if (!reader.HasRows)
            {
                // It's a standalone part
                if (!partsDict.ContainsKey(itemId))
                    partsDict[itemId] = 0;

                partsDict[itemId] += qtyRemaining;
            }
            else
            {
                // It's an assembly — explode into parts
                while (reader.Read())
                {
                    string subItemId = reader["Item ID"].ToString();
                    decimal qtyNeeded = reader["Qty Needed"] is decimal qn ? qn : 0;
                    if (string.IsNullOrWhiteSpace(subItemId) || qtyNeeded <= 0)
                        continue;
                    decimal totalQty = qtyNeeded * qtyRemaining;
                    if (!partsDict.ContainsKey(subItemId))
                        partsDict[subItemId] = 0;
                    partsDict[subItemId] += totalQty;
                }
            }
            reader.Close(); // make sure reader is closed before the next iteration
        }

        // Convert dictionary into List<Item>
        var parts = new List<Item>();

        foreach (var kvp in partsDict)
        {
            string itemId = kvp.Key;
            decimal qtyNeeded = kvp.Value;
            var item = new Item();
            item.SetAttr("Item ID", itemId);
            item.SetAttr("Qty Needed", qtyNeeded);
            string sql = @"
            SELECT 
                i.[Item Description],
                i.[Qty on Hand],
                pl.[Location],
                pl.[Code],
                pl.[Preferred Vendor],
                po.[PO No],
                po.[Vendor Name],
                po.[Qty Remaining]
            FROM Inventory i
            LEFT JOIN PartLocations pl ON i.[Item ID] = pl.[Item ID]
            LEFT JOIN PurchaseOrderReport po ON i.[Item ID] = po.[Item ID]
            WHERE i.[Item ID] = @id";
            using var cmd = new SqlCommand(sql, connection);
            cmd.Parameters.AddWithValue("@id", itemId);
            using var reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                // Inventory
                item.SetAttr("Item Description", reader["Item Description"]?.ToString());
                decimal qtyOnHand = reader["Qty on Hand"] is decimal q ? q : 0;
                item.SetAttr("Qty on Hand", qtyOnHand);
                item.SetAttr("Qty Difference", qtyOnHand - qtyNeeded);

                // PartLocations
                item.SetAttr("Location", reader["Location"]?.ToString());
                item.SetAttr("Code", reader["Code"]?.ToString());
                item.SetAttr("Preferred Vendor", reader["Preferred Vendor"]?.ToString());

                // PO Report
                item.SetAttr("PO No", reader["PO No"]?.ToString());
                item.SetAttr("Vendor Name", reader["Vendor Name"]?.ToString());
                decimal poQty = reader["Qty Remaining"] is decimal p ? p : 0;
                item.SetAttr("PO Qty Remaining", poQty);

                parts.Add(item);  // ✅ only add if Inventory match was found
            }
        }

        ExportPartsToExcel(parts);
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

    static void ExportPartsToExcel(List<Item> parts)
    {
        string templatePath = "BOMoutTemplate.xlsx";
        string outputPath = "BOMout.xlsx";
        // Header mapping
        var headerMap = new Dictionary<string, string>
    {
        { "Item ID", "PART" },
        { "Item Description", "DESCRIPTION" },
        { "Qty on Hand", "HAVE" },
        { "Qty Needed", "NEED" },
        { "Qty Difference", "DIFF" },
        { "Location", "LOCATION" },
        { "PO No", "PO No" },
        { "Preferred Vendor", "VENDOR" },
        { "PO Qty Remaining", "QTY" },
        { "Code", "CODE" }
    };
        using var workbook = new XLWorkbook(templatePath);
        var sheet = workbook.Worksheet(1);
        // Write header row
        int row = 1;
        int col = 1;
        foreach (var label in headerMap.Values)
        {
            sheet.Cell(row, col++).Value = label;
        }
        // Write part data
        foreach (var item in parts)
        {
            row++;
            col = 1;
            foreach (var field in headerMap.Keys)
            {
                var value = item.GetAttr(field);
                sheet.Cell(row, col++).Value = value?.ToString() ?? "";
            }
        }
        workbook.SaveAs(outputPath);
        Console.WriteLine($"Exported {parts.Count} parts to '{outputPath}'.");
    }
}
