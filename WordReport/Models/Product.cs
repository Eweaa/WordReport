namespace WordReport.Models;

public class Product
{
    public int? ProductId { get; set; }
    public string? NameAr { get; set; }
    public string? NameEn { get; set; }
    public string? Description { get; set; }
    public string? Category { get; set; }
    public decimal? Price { get; set; }
    public decimal? PriceAfterDiscount { get; set; }
    public int? QuantityInStock { get; set; }
    public DateTime? ExpiryDate { get; set; }
}
