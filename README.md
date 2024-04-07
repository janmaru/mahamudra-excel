# Introduction

The repository for reading and writing Excel File Dynamically.

# Getting Started

Download solutions and tests
 
# Build

in order to execute tests:
```shell
dotnet test 
``` 

## Features

- First you need to decorate your class/viewmodel with a "HeaderAttribute". Those properties that aren't decorated will be ignored.

```csharp 
[Table("products")]
public class Product
{
    [Column("product_id")]
    [Header("Product Id", Order = 2, Style = XCellStyle.Wrapper)]
    [Required]
    public int Id { get; set; }

    [Column("product_name")]
    [Header("Product Name", Order = 1, Style = XCellStyle.None)]
    [Required(AllowEmptyStrings = false)]
    public string? Name { get; set; }

    [Column("brand_id")]
    [Header("Brand Id", Order = 3, Style = XCellStyle.None)]
    [Required]
    public int BrandId { get; set; }

    [Column("category_id")]
    [Required]
    [Header("Category Id", Order = 4, Style = XCellStyle.None)]
    public int CategoryId { get; set; }

    [Column("model_year")]
    [Required]
    public short ModelYear { get; set; }

    [Column("list_price")]
    [Header("Very Long List Price Caption", Order = 5, Style = XCellStyle.None)]
    [Required]
    public decimal ListPrice { get; set; }

    [Column("Date")]
    [HeaderAttribute("Date Time", Order = 6, Style = XCellStyle.None)]
    [Required]
    public DateTime Date { get; set; }
}
```
- The simpler use case would be a list or a collection of entities (i.e. Product) that are lined up in rows inside the Excel File.
- While the columns inherit from the attributes.
```csharp
        var expectedListOfProducts = new List<Product>()
        {
            new()
            {
                 BrandId = CustomExtensions.GetRandomInteger(),
                 CategoryId = CustomExtensions.GetRandomInteger(),
                 Id =  CustomExtensions.GetRandomInteger(),
                 ListPrice = CustomExtensions.GetRandomDecimal(),
                 ModelYear =(short) CustomExtensions.GetRandomInteger(DateTime.Now.Year),
                 Name = CustomExtensions.GetRandomString(),
            }
        };

        var dataSet = expectedListOfProducts.FillOneSheet();
        using var streamFile = dataSet.ToExcel(); 
        var array = streamFile.ToArray();
``` 

- You can save the byte array on the file system
```csharp
         File.WriteAllBytes($"D:\\{new Random().Next()}.xlsx", array);
``` 
- Or you can use memory stream to download the file in a minimal api
```csharp
        var filename = $"{Guid.NewGuid()}.xlsx";
        var content = expectedListOfProducts.FillOneSheet().ToExcel();
        var mime = "application/vnd.ms-excel";
        return Results.File(content!, mime, filename);
``` 
- You can load an excel file into a dataset
```csharp
        var filePath = Path.Combine([_root, "temp", $"read_test.xlsx"]); 
        var streamFile = filePath.Read(); 
        var dt = streamFile.ReadExcel();
``` 