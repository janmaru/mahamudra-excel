# Mahamudra.Excel

A lightweight .NET library for reading and writing Excel files (.xlsx) using attribute-based column mapping.

## Features

- **Attribute-based mapping** - Use `[Header]` attribute to define column headers, order, and styling
- **Write Excel files** - Convert collections to Excel with automatic column configuration
- **Read Excel files** - Load Excel data into DataTables
- **Cell styling** - Support for text wrapping, date/number formatting
- **Auto-sized columns** - Columns automatically sized based on content

## Requirements

- .NET Standard 2.1+
- Dependencies: DocumentFormat.OpenXml 3.0.0

## Installation

Add a reference to the `Mahamudra.Excel` project or build and reference the DLL.

## Quick Start

### Running Tests

```shell
dotnet test
```

## Usage

### 1. Define Your Model

Decorate your class properties with the `[Header]` attribute. Properties without this attribute will be ignored.

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
    [Header("Date Time", Order = 6, Style = XCellStyle.None)]
    [Required]
    public DateTime Date { get; set; }
}
```

### 2. Write Excel File

Convert a collection to an Excel file:

```csharp
var products = new List<Product>
{
    new()
    {
        Id = 1,
        Name = "Product A",
        BrandId = 100,
        CategoryId = 10,
        ModelYear = 2024,
        ListPrice = 99.99m,
        Date = DateTime.Now
    }
};

// Create Excel as byte array
var dataSet = products.FillOneSheet();
using var stream = dataSet.ToExcel();
var bytes = stream.ToArray();

// Save to file
File.WriteAllBytes("products.xlsx", bytes);
```

### 3. Download in Minimal API

Return Excel file as HTTP response:

```csharp
app.MapGet("/export", () =>
{
    var products = GetProducts();
    var content = products.FillOneSheet().ToExcel();
    var filename = $"{Guid.NewGuid()}.xlsx";
    var mime = "application/vnd.ms-excel";
    return Results.File(content!, mime, filename);
});
```

### 4. Read Excel File

Load an Excel file into a DataTable:

```csharp
var filePath = "data.xlsx";
var stream = filePath.Read();
var dataTable = stream.ReadExcel();
```

## Header Attribute Options

| Property | Type | Description |
|----------|------|-------------|
| `Caption` | string | Column header text (required) |
| `Order` | short | Column position (default: 0) |
| `Style` | XCellStyle | Cell formatting style |
| `ReadOnly` | bool | Mark column as read-only |

## Cell Styles

- `XCellStyle.None` - Default formatting
- `XCellStyle.Wrapper` - Text wrapping enabled
- `XCellStyle.Header` - Header cell formatting

## Benchmarks

Performance benchmarks using BenchmarkDotNet (.NET 9.0):

| Rows | Mean Time | Memory Allocated | Memory/Row |
|------|-----------|------------------|------------|
| 50,000 | 1.70 s | ~280 MB | ~5.6 KB |
| 100,000 | 3.52 s | ~559 MB | ~5.6 KB |
| 200,000 | ~6.8 s | ~1.1 GB | ~5.6 KB |

**File Size Estimation:**
- ~50 MB file: 100-150k rows
- ~100 MB file: 200-300k rows

### Running Benchmarks

```shell
dotnet run -c Release --project Mahamudra.Excel.Benchmarks
```

## License

MIT
