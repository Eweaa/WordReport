using WordReport.Models;

namespace WordReport.Services;

public static class DataService
{
    public static List<Product> GetProducts()
    {
        return new List<Product>
        {
            new Product
            {
                ProductId = 1,
                NameAr = "منتج 1",
                NameEn = "Product 1",
                Description = "وصف المنتج 1",
                Category = "الفئة 1",
                Price = 100.00m,
                PriceAfterDiscount = 90.00m,
                QuantityInStock = 50,
                ExpiryDate = DateTime.Now.AddMonths(6)
            },
            new Product
            {
                ProductId = 2,
                NameAr = "منتج 2",
                NameEn = "Product 2",
                Description = "وصف المنتج 2",
                Category = "الفئة 2",
                Price = 200.00m,
                PriceAfterDiscount = 180.00m,
                QuantityInStock = 30,
                ExpiryDate = DateTime.Now.AddMonths(12)
            }
        };
    }
    public static Quotation GetQuotations()
    {
        return new Quotation
        {
            Items = new List<QuotationItem>
            {
                new QuotationItem { Deliverable = "Sample analysis", Unit = "Sample", UnitCost = 2500m, Quantity = 1 },
                new QuotationItem { Deliverable = "Mobilization", Unit = "Trip", UnitCost = 1000m, Quantity = 1 }
            },
            FinalTotal = 3500m,
            Vat = 150m
        };
    }

    public static List<Test> GetTests()
    {
        return new List<Test>
        {
            new Test { TestId = 1, Parameter = "Aluminum", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 2, Parameter = "Copper (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 3, Parameter = "Zinc", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 4, Parameter = "Antimony (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 5, Parameter = "Arsenic (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 6, Parameter = "Barium (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 7, Parameter = "Boron (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 8, Parameter = "Cadmium (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 9, Parameter = "Chromium (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 10, Parameter = "Lead (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 11, Parameter = "Mercury (Inorganic) (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 12, Parameter = "Molybdenum", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 13, Parameter = "Nickel (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 14, Parameter = "Selenium (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 15, Parameter = "Uranium (H)", Unit = "mg/l", Method = "EPA 200.8", LOD = "" },
            new Test { TestId = 16, Parameter = "Acrylamide (H)", Unit = "mg/l", Method = "LC/MS-MS", LOD = "" },
            new Test { TestId = 17, Parameter = "Cyanide", Unit = "mg/l", Method = "APHA 4500 CN", LOD = "" },
            new Test { TestId = 18, Parameter = "Microcystins (H)", Unit = "mg/l", Method = "LC/MS-MS", LOD = "" },
            new Test { TestId = 19, Parameter = "Perchlorate (H)", Unit = "mg/l", Method = "EPA 6850", LOD = "" },
            new Test { TestId = 20, Parameter = "Benzene (H)", Unit = "mg/l", Method = "EPA 8260 D", LOD = "" },
            new Test { TestId = 21, Parameter = "Benzo (a) – Pyrene (H)", Unit = "mg/l", Method = "EPA 8270 E", LOD = "" },
            new Test { TestId = 22, Parameter = "Chlorobenzene (monochlorobenzene MCB)", Unit = "mg/l", Method = "EPA 8260 D", LOD = "" },
            new Test { TestId = 23, Parameter = "1,2-dichlorobenzene (H)", Unit = "mg/l", Method = "EPA 8270 E", LOD = "" },
            new Test { TestId = 24, Parameter = "1,4-dichlorobenzene (H)", Unit = "mg/l", Method = "EPA 8270 E", LOD = "" },
            new Test { TestId = 25, Parameter = "1,2-dichloroethane (H)", Unit = "mg/l", Method = "EPA 8270 E", LOD = "" },
            new Test { TestId = 26, Parameter = "1,2-dichloroethene (H)", Unit = "mg/l", Method = "EPA 8270 E", LOD = "" }
        };
    }
}
