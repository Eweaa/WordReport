namespace WordReport.Models;

public class Product
{
    public int? ProductId { get; set; }
    public string? NameEn { get; set; }
    public string? NameAr { get; set; }
    public decimal? Price { get; set; }
    public decimal? PriceAfterDiscount { get; set; }
}
