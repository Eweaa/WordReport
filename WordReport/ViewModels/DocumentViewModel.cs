namespace WordReport.ViewModels;

public class DocumentViewModel
{
    public string? Title { get; set; }
    public string? Date { get; set; }
    public string? prop1 { get; set; }
    public int? prop2 { get; set; }
    public DateTime? prop3 { get; set; }
    public double? prop4 { get; set; }
    public long? prop5 { get; set; }

    public List<TableItemGroupViewModel> Tables { get; set; } = [];
    public IFormFile? Logo { get; set; } 

}


public class DocumentViewModel2<T>
{
    public string? Title { get; set; }
    public string? Date { get; set; }
    public string? prop1 { get; set; }
    public int? prop2 { get; set; }
    public DateTime? prop3 { get; set; }
    public double? prop4 { get; set; }
    public long? prop5 { get; set; }

    public List<T> Table { get; set; } = [];
    public IFormFile? Logo { get; set; }

}
