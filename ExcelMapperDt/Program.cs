using System.Dynamic;
using Ganss.Excel;

class Program
{
    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);
        var app = builder.Build();

        app.MapGet("/", async () =>
        {
            var list = Creator.CreateTestData();
            await new ExcelMapper().SaveAsync("Sheet.xlsx", list, "Sheet 1");
        });
        app.Run();
    }
}

public static class Creator
{
    public static List<object> CreateTestData()
    {
        var list = new List<object>();
        for (int i = 1; i <= 100; i++)
        {
            var obj = new ExpandoObject();
            var code = i.ToString().PadLeft(3, '0');
            obj.TryAdd("Sn", i);
            obj.TryAdd("Name", $"Employee #{code}");
            obj.TryAdd("Code", $"EMP{code}");
            list.Add(obj);
        }

        return list;
    }
}