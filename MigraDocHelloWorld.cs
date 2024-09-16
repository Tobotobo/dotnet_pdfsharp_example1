// https://github.com/empira/PDFsharp.Samples/blob/master/src/samples/src/MigraDoc/src/HelloWorld/Program.cs

using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.Rendering;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Quality;
using PdfSharp.Snippets.Font;
using System.Globalization;

public static class MigraDocHelloWorld {
    public static void Run() {
        // Create a MigraDoc document.
        var document = CreateDocument();

        var style = document.Styles[StyleNames.Normal]!;
        style.Font.Name = "NotoSansJP-Regular";

        // Create a renderer for the MigraDoc document.
        var pdfRenderer = new PdfDocumentRenderer
        {
            // Associate the MigraDoc document with a renderer.
            Document = document,
            PdfDocument =
            {
                // Change some settings before rendering the MigraDoc document.
                PageLayout = PdfPageLayout.SinglePage,
                ViewerPreferences =
                {
                    FitWindow = true
                }
            }
        };

        // Layout and render document to PDF.
        pdfRenderer.RenderDocument();

        // Save the document...
        var filename = PdfFileUtility.GetTempPdfFullFileName("samples-MigraDoc/HelloWorldMigraDoc");
        pdfRenderer.PdfDocument.Save(filename);
        // ...and start a viewer.
        // PdfFileUtility.ShowDocument(filename);

        // Creates a minimalistic document.
        static Document CreateDocument()
        {
            // Create a new MigraDoc document.
            var document = new Document();

            // Add a section to the document.
            var section = document.AddSection();

            // Add a paragraph to the section.
            var paragraph = section.AddParagraph();

            // Set font color.
            paragraph.Format.Font.Color = Colors.DarkBlue;

            // Add some text to the paragraph.
            // paragraph.AddFormattedText("Hello, World!", TextFormat.Bold);
            paragraph.AddFormattedText("こんにちは、MigraDoc！", TextFormat.Bold);

            // Create the primary footer.
            var footer = section.Footers.Primary;

            // Add content to footer.
            paragraph = footer.AddParagraph();
            // paragraph.Add(new DateField { Format = "yyyy/MM/dd HH:mm:ss" });
            var culture = new CultureInfo("ja-JP"); 
            culture.DateTimeFormat.Calendar = new JapaneseCalendar();
            paragraph.AddText(DateTime.Now.ToString("ggy年M月d日 H時m分s秒", culture));

            // paragraph.Add(new DateField { Format = "ggyy年M月d日 H時m分s秒" });
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Add MigraDoc logo.
            // string imagePath = IOUtility.GetAssetsPath(@"migradoc/images/MigraDoc-128x128.png")!;
            // document.LastSection.AddImage(imagePath);

            return document;
        }
    }
}