using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using Mahamudra.Excel.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Mahamudra.Excel.Domain;

namespace Mahamudra.Excel.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net90)]
public class ExcelBenchmarks
{
    private List<BenchmarkProduct> _products = null!;

    [Params(50_000, 100_000, 200_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _products = new List<BenchmarkProduct>(RowCount);
        var random = new Random(42);

        for (int i = 0; i < RowCount; i++)
        {
            _products.Add(new BenchmarkProduct
            {
                Id = i,
                Name = $"Product {i} with some additional text to increase size",
                BrandId = random.Next(1, 1000),
                CategoryId = random.Next(1, 100),
                ModelYear = (short)random.Next(2000, 2025),
                ListPrice = (decimal)(random.NextDouble() * 10000),
                Date = DateTime.Now.AddDays(-random.Next(0, 365))
            });
        }
    }

    [Benchmark]
    public byte[] CreateExcelFile()
    {
        var dataSet = _products.FillOneSheet("Products");
        using var stream = dataSet.ToExcel();
        return stream.ToArray();
    }
}

[Table("benchmark_products")]
public class BenchmarkProduct
{
    [Column("product_id")]
    [Header("Product Id", Order = 1, Style = XCellStyle.None)]
    [Required]
    public int Id { get; set; }

    [Column("product_name")]
    [Header("Product Name", Order = 2, Style = XCellStyle.None)]
    [Required]
    public string? Name { get; set; }

    [Column("brand_id")]
    [Header("Brand Id", Order = 3, Style = XCellStyle.None)]
    [Required]
    public int BrandId { get; set; }

    [Column("category_id")]
    [Header("Category Id", Order = 4, Style = XCellStyle.None)]
    [Required]
    public int CategoryId { get; set; }

    [Column("model_year")]
    [Header("Model Year", Order = 5, Style = XCellStyle.None)]
    [Required]
    public short ModelYear { get; set; }

    [Column("list_price")]
    [Header("List Price", Order = 6, Style = XCellStyle.None)]
    [Required]
    public decimal ListPrice { get; set; }

    [Column("date")]
    [Header("Date", Order = 7, Style = XCellStyle.None)]
    [Required]
    public DateTime Date { get; set; }
}
