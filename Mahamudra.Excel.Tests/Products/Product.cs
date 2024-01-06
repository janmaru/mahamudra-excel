using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Mahamudra.Excel.Tests.Products;

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