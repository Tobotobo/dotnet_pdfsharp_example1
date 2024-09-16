// https://github.com/empira/PDFsharp.Samples/blob/master/src/samples/src/PDFsharp/src/HelloWorld/Program.cs

using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Quality;

public static class PdfsharpHelloWorld {
    public static void Run() {

        // GlobalFontSettings.FontResolver = PdfSharp.Fonts.PlatformFontResolver.

        // Create a new PDF document.
        var document = new PdfDocument();
        document.Info.Title = "Created with PDFsharp";
        document.Info.Subject = "Just a simple Hello-World program.";

        // Create an empty page in this document.
        var page = document.AddPage();
        //page.Size = PageSize.Letter;

        // Get an XGraphics object for drawing on this page.
        var gfx = XGraphics.FromPdfPage(page);

        // Draw two lines with a red default pen.
        var width = page.Width.Point;
        var height = page.Height.Point;
        gfx.DrawLine(XPens.Red, 0, 0, width, height);
        gfx.DrawLine(XPens.Red, width, 0, 0, height);

        // Draw a circle with a red pen which is 1.5 point thick.
        var r = width / 5;
        gfx.DrawEllipse(new XPen(XColors.Red, 1.5), XBrushes.White, new XRect(width / 2 - r, height / 2 - r, 2 * r, 2 * r));

        // Create a font.
        var font = new XFont("NotoSansJP-Regular", 20, XFontStyleEx.BoldItalic);

        // Draw the text.
        // gfx.DrawString("Hello, PDFsharp!", font, XBrushes.Black,
        //     new XRect(0, 0, page.Width.Point, page.Height.Point), XStringFormats.Center);
        gfx.DrawString("こんにちは、PDFsharp！", font, XBrushes.Black,
            new XRect(0, 0, page.Width.Point, page.Height.Point), XStringFormats.Center);

        // Save the document...
        var filename = PdfFileUtility.GetTempPdfFullFileName("samples/HelloWorldSample");
        document.Save(filename);
        // ...and start a viewer.
        // PdfFileUtility.ShowDocument(filename); // たぶん Linux 環境だと System.ComponentModel.Win32Exception
    }
}