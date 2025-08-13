using Microsoft.AspNetCore.Mvc;
using WordReport.Services;
using WordReport.ViewModels;

namespace WordReport.Controllers;

public class DocumentController : Controller
{
    private readonly IWebHostEnvironment _env;
    private readonly WordService _wordService;
    private readonly CleanWordService _cleanwordService;
    private readonly CleanWordPdfService _cleanWordPdfService;
    private readonly DirectWordXmlService _wordXmlService;

    public DocumentController(IWebHostEnvironment env)
    {
        _env = env;
        _wordService = new WordService();
        _cleanwordService = new CleanWordService();
        _cleanWordPdfService = new CleanWordPdfService();
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

        //var wordBytes = _wordService.GenerateDocument(model, templatePath);
        //var wordBytes2 = _wordXmlService.GenerateDocumentManual(model, templatePath);
        //var wordBytes3 = _cleanwordService.GenerateDocument(model, templatePath);
        //var pdfBytes = _wordService.GenerateProposalPdfWithLibreOffice(model, templatePath);

        //return File(wordBytes, "application/vnd./*openxmlformats*/-officedocument.wordprocessingml.document", "Generated.docx");
        //return File(wordBytes2, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Generated.docx");
        //return File(wordBytes3, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Generated.docx");
        //return File(pdfBytes, "application/pdf", "Proposal.pdf");


        //string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "templates", "Test.docx");

        //var pdfBytes = _wordService.GenerateProposalPdfWithLibreOffice(model, templatePath);

        var pdf = _cleanWordPdfService.GenerateProposalPdfWithLibreOffice(model, templatePath);

        return File(pdf, "application/pdf", "Report.pdf");
    }
}
