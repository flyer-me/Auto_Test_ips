using OfficeOpenXml;
class Program
{
    static void Main()
    {
        Console.WriteLine("街乡平台设备数据情况比对程序\n\n数据表文件名必须以'街乡平台设备在线'开头，以'.xlsx'结尾,且在线原数据位于第二个sheet;\n\n统计结果文件名为'小区统计表1775.xlsx',不能更改");
        var file_list = Directory.GetFiles(".", "街乡平台设备在线*.xlsx");
        if (file_list.Length == 0)
        {
            Console.WriteLine("未找到'街乡平台设备在线'开头的文件，按enter退出...");
            Console.ReadLine();
            Environment.Exit(1);
        }

        var sum = new Dictionary<string, int>();
        var online = new Dictionary<string, int>();
        var not_online = new Dictionary<string, int>();

        foreach (var file_name in file_list)
        {
            using (var package = new ExcelPackage(new FileInfo(file_name)))
            {
                var workbook = package.Workbook;
                var sheet = workbook.Worksheets[1];
                for (int row = 2; row <= sheet.Dimension.End.Row; row++)
                {
                    var category = sheet.Cells[row, 2].Text.Split('/').Last();
                    var status = sheet.Cells[row, 4].Text;
                    if (!sum.ContainsKey(category))
                    {
                        sum[category] = 0;
                    }
                    sum[category] += 1;
                    if (!online.ContainsKey(category))
                    {
                        online[category] = 0;
                    }
                    if (!not_online.ContainsKey(category))
                    {
                        not_online[category] = 0;
                    }
                    if (status == "在线" || status == "未检测")
                    {
                        online[category] = online.GetValueOrDefault(category, 0) + 1;
                    }
                    else if (status == "离线" && sheet.Cells[row, 7].Text != "0")
                    {
                        not_online[category] = not_online.GetValueOrDefault(category, 0) + 1;
                    }
                }
            }
        }
        Console.WriteLine("统计完成..");

        using (var file = new ExcelPackage(new FileInfo("小区统计表1775.xlsx")))
        {
            var sheet = file.Workbook.Worksheets[0];

            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                var value = sheet.Cells[row, 5].Text;
                if (online.ContainsKey(value))
                {
                    sheet.Cells[row, 11].Value = online[value];
                    sheet.Cells[row, 12].Value = not_online[value];
                    sheet.Cells[row, 13].Value = sum[value];
                }
                else
                {
                    sheet.Cells[row, 11].Value = "平台无设备";
                }
            }
            try
            {
                file.Save();
            }
            catch (IOException)
            {
                Console.WriteLine("结果写入失败,请关闭文件后重试");
                Environment.Exit(1);
            }
        }
        Console.WriteLine("结果写入成功,按enter退出...");
        Console.ReadLine();
    }
}

