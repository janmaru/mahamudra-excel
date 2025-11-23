using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Mahamudra.Excel.Domain;

namespace Mahamudra.Excel.Tests.Products;

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

    [Column("Calculation")]
    [Header("Calculation", Order = 7, Style = XCellStyle.None)]
    [Required]
    public long Calculation { get; set; }
}