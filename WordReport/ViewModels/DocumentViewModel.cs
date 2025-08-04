namespace WordReport.ViewModels;

public class DocumentViewModel
{
    public string Title { get; set; }
    public string Date { get; set; }

    public List<TableItemViewModel> Items { get; set; } = [];
}
