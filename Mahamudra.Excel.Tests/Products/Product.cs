using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Mahamudra.Excel.Common;

namespace Mahamudra.Excel.Tests.Products;

[Table("products")]
public class Product
{
    [Column("product_id")]
    [HeaderAttribute("Product Id", Order = 2, Style = XCellStyle.Wrapper)]
    [Required]
    public int Id { get; set; }

    [Column("product_name")]
    [HeaderAttribute("Product Name", Order = 1, Style = XCellStyle.None)]
    [Required(AllowEmptyStrings = false)]
    public string? Name { get; set; }

    [Column("brand_id")]
    [HeaderAttribute("Brand Id", Order = 3, Style = XCellStyle.Wrapper)]
    [Required]
    public int BrandId { get; set; }

    [Column("category_id")]
    [Required]
    [HeaderAttribute("Category Id", Order = 4, Style = XCellStyle.Wrapper)]
    public int CategoryId { get; set; }

    [Column("model_year")]
    [Required]
    public short ModelYear { get; set; }

    [Column("list_price")]
    [HeaderAttribute("ListPrice", Order = 5, Style = XCellStyle.Wrapper)]
    [Required]
    public decimal ListPrice { get; set; }

    [Column("Date")]
    [HeaderAttribute("Date Time", Order = 6, Style = XCellStyle.Wrapper)]
    [Required]
    public DateTime Date { get; set; }
}