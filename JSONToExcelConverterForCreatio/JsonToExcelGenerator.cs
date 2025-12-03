using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Data;
public class JsonToExcelGenerator
{
    public static void GenerateExcel(string json, string outputPath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        JObject root = JObject.Parse(json);
        JArray rows = (JArray)root["rows"];
        HashSet<string> columns = new HashSet<string>();
        foreach (JObject row in rows)
        {
            foreach (var prop in row.Properties())
            {
                columns.Add(prop.Name);
            }
        }
        DataTable dt = new DataTable("Result");
        foreach (string col in columns)
        {
            dt.Columns.Add(col);
        }
        foreach (JObject row in rows)
        {
            DataRow dr = dt.NewRow();
            foreach (string col in columns)
            {
                JToken value = row[col];

                if (value == null)
                {
                    dr[col] = DBNull.Value;
                }
                else if (value.Type == JTokenType.Object)
                {
                    var displayValue = value["displayValue"]?.ToString();
                    dr[col] = displayValue ?? "";
                }
                else
                {
                    dr[col] = value.ToString();
                }
            }
            dt.Rows.Add(dr);
        }
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            worksheet.Cells["A1"].LoadFromDataTable(dt, true);
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            package.SaveAs(new FileInfo(outputPath));
        }
    }
}
