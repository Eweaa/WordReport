using Microsoft.AspNetCore.Mvc;
using WordReport.Services;
using WordReport.ViewModels;

namespace WordReport.Controllers;

public class DocumentController : Controller
{
    private readonly IWebHostEnvironment _env;
    private readonly WordService _wordService;

    public DocumentController(IWebHostEnvironment env)
    {
        _env = env;
        _wordService = new WordService();
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult GenerateWord(DocumentViewModel model)
    {
        string templatePath = Path.Combine(_env.WebRootPath, "templates", "Report.docx");

        var wordBytes = _wordService.GenerateDocument(model, templatePath);

        return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Generated.docx");
    }
}
