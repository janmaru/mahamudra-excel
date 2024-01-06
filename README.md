# Introduction

The repository for reading and writing Excel File Dynamically.

# Getting Started

Install docker, docker compose and dotnet sdk in order to build and tests the solution. 
 
# Build

in order to build and launch locally the application at port 57000 you need to do this:
```shell
docker compose -f docker-compose.yml -f docker-compose.override.yml build --no-cache
docker compose -f docker-compose.yml -f docker-compose.override.yml up
```
**NB:** you can change the defaults configuration in a `docker-compose.override.yml` file.

## Features

- First you need to decorate your class with a "HeaderAttribute". Those properties that aren't decorated will be ignored.
```csharp 
[Table("products")]
public class Product
{
    [Column("product_id")]
    [HeaderAttribute("Product Id") ]
    [Required]
    public int Id { get; set; }

    [Column("product_name")]
    [HeaderAttribute("Product Name")]
    [Required(AllowEmptyStrings = false)]
    public string? Name { get; set; }

    [Column("brand_id")]
    [HeaderAttribute("Brand Id")]
    [Required]
    public int BrandId { get; set; }

    [Column("category_id")]
    [Required]
    [HeaderAttribute("Category Id")]
    public int CategoryId { get; set; }

    [Column("model_year")]
    [Required]
    public short ModelYear { get; set; }

    [Column("list_price")]
    [HeaderAttribute("ListPrice")]
    [Required]
    public decimal ListPrice { get; set; } 
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