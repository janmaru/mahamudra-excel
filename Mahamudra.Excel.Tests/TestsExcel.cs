using System.Data;
using System.Reflection;
using Mahamudra.Excel.Extensions;
using Mahamudra.Excel.Tests.Common;
using Mahamudra.Excel.Tests.Products;

namespace Mahamudra.Excel.Tests;
public class Tests
{
    private static string _root = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;

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
        var (table, headers, _) = DataSetExtensions.CreateTable<Product>();
        Assert.That(table, Is.Not.Null);
        var columnListPrice = table.Columns.Cast<DataColumn>().Where(x => x.ColumnName == nameof(Product.ListPrice)).FirstOrDefault();
        Assert.That(columnListPrice, Is.Not.Null);
        Assert.That(columnListPrice.Caption, Is.EqualTo("Very Long List Price Caption"));
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

        Assert.That(table.Rows, Has.Count.EqualTo(1));
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
                 Date = DateTime.Now
            },
           new()
            {
                 BrandId = CustomExtensions.GetRandomInteger(),
                 CategoryId = CustomExtensions.GetRandomInteger(),
                 Id =  CustomExtensions.GetRandomInteger(),
                 ListPrice = CustomExtensions.GetRandomDecimal(),
                 ModelYear =(short) CustomExtensions.GetRandomInteger(DateTime.Now.Year),
                 Name = CustomExtensions.GetRandomString(),
                 Date = DateTime.Now
            },
           new()
            {
                 BrandId = CustomExtensions.GetRandomInteger(),
                 CategoryId = CustomExtensions.GetRandomInteger(),
                 Id =  CustomExtensions.GetRandomInteger(),
                 ListPrice = CustomExtensions.GetRandomDecimal(),
                 ModelYear =(short) CustomExtensions.GetRandomInteger(DateTime.Now.Year),
                 Name = CustomExtensions.GetRandomString(),
                 Date = DateTime.Now
            }
        };

        var dataSet = expectedListOfProducts.FillOneSheet();
        var streamFile = dataSet.ToExcel();  
        Assert.That(streamFile, Is.Not.Null);

        var filePath = Path.Combine([_root, "temp", $"{ new Random().Next()}.xlsx"]);
        streamFile.Write(filePath);
        Assert.That(filePath.Exists(), Is.EqualTo(true));
    }

    [Test]
    public void ReadExcel_ShoulReadExcelFile_ShouldSucceed()
    { 
        var filePath = Path.Combine([_root, "temp", $"read_test.xlsx"]); 
        var streamFile = filePath.Read();
        Assert.That(streamFile, Is.Not.Null); 
        var dt = streamFile.ReadExcel();
        Assert.That(dt, Is.Not.Null);
        Assert.That(dt.Rows[0][1], Is.EqualTo(790.ToString()));
    }
}