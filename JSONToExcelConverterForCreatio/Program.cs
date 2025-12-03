namespace JSONToExcelConverterForCreatio
{
    using System;
    using System.IO;
    class Program
    {
        static int Main(string[] args)
        {
            string inputJsonPath;
            string outputExcelPath;
            if (args.Length == 2)
            {
                inputJsonPath = args[0];
                outputExcelPath = args[1];
            }
            else
            {
                string configFile = Path.Combine(AppContext.BaseDirectory, "appconfig.txt");
                if (!File.Exists(configFile))
                {
                    Console.WriteLine("Missing config.txt");
                    Console.WriteLine("Create config.txt with:");
                    Console.WriteLine("input=PATH");
                    Console.WriteLine("output=PATH");
                    return 1;
                }

                var lines = File.ReadAllLines(configFile);
                inputJsonPath = null;
                outputExcelPath = null;
                foreach (var line in lines)
                {
                    if (line.StartsWith("input="))
                        inputJsonPath = line.Substring("input=".Length).Trim();

                    if (line.StartsWith("output="))
                        outputExcelPath = line.Substring("output=".Length).Trim();
                }
                if (string.IsNullOrWhiteSpace(inputJsonPath) ||
                    string.IsNullOrWhiteSpace(outputExcelPath))
                {
                    Console.WriteLine("config.txt is incorrect. Example:");
                    Console.WriteLine("input=C:\\temp\\response.json");
                    Console.WriteLine("output=C:\\temp\\result.xlsx");
                    return 2;
                }
            }
            try
            {
                if (!File.Exists(inputJsonPath))
                {
                    Console.WriteLine($"Input file not found: {inputJsonPath}");
                    return 3;
                }
                string json = File.ReadAllText(inputJsonPath);
                JsonToExcelGenerator.GenerateExcel(json, outputExcelPath);
                Console.WriteLine("Excel file successfully generated:");
                Console.WriteLine(outputExcelPath);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:");
                Console.WriteLine(ex.Message);
                return 4;
            }
        }
    }
}
