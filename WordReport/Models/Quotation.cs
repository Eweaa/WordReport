namespace WordReport.Models;

public class Quotation
{
    public List<QuotationItem> Items { get; set; } = [];
    public decimal? Vat { get; set; }
    public decimal? FinalTotal { get; set; }
}

public class QuotationItem
{
    public int QuotationId { get; set; }
    public string? Deliverable { get; set; }
    public string? Unit { get; set; }
    public decimal? UnitCost { get; set; }
    public int? Quantity { get; set; }
    public decimal? TotalCost { get; set; }
}