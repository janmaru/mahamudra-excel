using Mahamudra.Excel.Common;
using Mahamudra.Excel.Tests.Common;
using Mahamudra.Excel.Tests.Products;
using System.Data;

namespace Mahamudra.Excel.Tests;
public class Tests
{
    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public void ToTable_ShouldRetrieveAllHeaders_ShouldSucceed()
    {
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

        var singleProduct = expectedListOfProducts[0];
        var (table, headers) = ReflectionExtensions.ToTable<Product>();
        Assert.NotNull(table);
        var columnListPrice = table.Columns.Cast<DataColumn>().Where(x => x.ColumnName == nameof(Product.ListPrice)).FirstOrDefault();
        Assert.NotNull(columnListPrice);
        Assert.That(columnListPrice.Caption, Is.EqualTo(nameof(Product.ListPrice)));
    }

    [Test]
    public void Fill_ShouldRetrieveDataTable_ShouldSucceed()
    {
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

        var singleProduct = expectedListOfProducts[0];

        var table = expectedListOfProducts.FillOneSheet().Tables[0];

        Assert.That(table.Rows.Count, Is.EqualTo(1));
        var row = table.Rows[0];
        var listPrice = row[nameof(Product.ListPrice)];
        Assert.That(listPrice, Is.EqualTo(singleProduct.ListPrice));
    }

    [Test]
    public void ToExcel_ShoulCreateExcelFile_ShouldSucceed()
    {
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

        Assert.That(array, Is.Not.Null);
    }
}