# dynamic-excel-mapper

Map dynamic values using the excellend Excel Mapper library.

Uses ExpandoObject to dynamically create properties and the list of such objects are sent to ExcelMapper which does the heavy lifting.

> Generating Dynamic Data

```csharp
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
```

> Generating Excel

```csharp
app.MapGet("/", async () =>
{
    var list = Creator.CreateTestData();
    await new ExcelMapper().SaveAsync("Sheet.xlsx", list, "Sheet 1");
});
```

### Library Used

https://github.com/mganss/ExcelMapper
