using Microsoft.AspNetCore.Mvc;
using WordReport.Services;
using WordReport.ViewModels;

namespace WordReport.Controllers;

public class DocumentController : Controller
{
    private readonly IWebHostEnvironment _env;
    private readonly WordService _wordService;
    private readonly CleanWordService _cleanwordService;
    private readonly DirectWordXmlService _wordXmlService;

    public DocumentController(IWebHostEnvironment env)
    {
        _env = env;
        _wordService = new WordService();
        _cleanwordService = new CleanWordService();
        _wordXmlService = new DirectWordXmlService();
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult GenerateWord(DocumentViewModel model)
    {
        string templatePath = Path.Combine(_env.WebRootPath, "templates", "Test.docx");

        var wordBytes = _wordService.GenerateDocument(model, templatePath);
        //var wordBytes2 = _wordXmlService.GenerateDocumentManual(model, templatePath);
        //var wordBytes3 = _cleanwordService.GenerateDocument(model, templatePath);

        return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Generated.docx");
    }
}
