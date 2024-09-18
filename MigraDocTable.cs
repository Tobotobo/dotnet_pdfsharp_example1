// https://github.com/empira/PDFsharp.Samples/blob/master/src/samples/src/MigraDoc/src/HelloMigraDoc/Tables.cs

using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Quality;
using PdfSharp.Snippets.Font;
using System.Globalization;

public static class MigraDocTable
{
    public static void Run()
    {
        // Create a MigraDoc document.
        var document = CreateDocument();

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

            DefineTables(document);

            return document;
        }
    }

    public static void DefineTables(Document document)
    {
        var paragraph = document.LastSection!.AddParagraph("Table Overview", "Heading1");
        paragraph.AddBookmark("Tables");

        DemonstrateSimpleTable(document);
        DemonstrateAlignment(document);
        DemonstrateCellMerge(document);
    }

    public static void DemonstrateSimpleTable(Document document)
    {
        document.LastSection.AddParagraph("Simple Tables", "Heading2");

        var table = new Table
        {
            Borders =
                {
                    Width = 0.75
                }
        };

        var column = table.AddColumn(Unit.FromCentimeter(2));
        column.Format.Alignment = ParagraphAlignment.Center;

        table.AddColumn(Unit.FromCentimeter(5));

        var row = table.AddRow();
        row.Shading.Color = Colors.PaleGoldenrod;
        var cell = row.Cells[0];
        cell.AddParagraph("Itemus");
        cell = row.Cells[1];
        cell.AddParagraph("Descriptum");

        row = table.AddRow();
        cell = row.Cells[0];
        cell.AddParagraph("1");
        cell = row.Cells[1];
        cell.AddParagraph(
            "Andigna cons nonsectem accummo diamet nis diat.");

        row = table.AddRow();
        cell = row.Cells[0];
        cell.AddParagraph("2");
        cell = row.Cells[1];
        cell.AddParagraph(
            "Loboreet autpat, quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet, consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. " +
            "Nisci tat laor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.");

        table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

        document.LastSection.Add(table);
    }

    public static void DemonstrateAlignment(Document document)
    {
        document.LastSection.AddParagraph("Cell Alignment", "Heading2");

        var table = document.LastSection.AddTable();
        table.Borders.Visible = true;
        table.Format.Shading.Color = Colors.LavenderBlush;
        table.Shading.Color = Colors.Salmon;
        table.TopPadding = 5;
        table.BottomPadding = 5;

        var column = table.AddColumn();
        column.Format.Alignment = ParagraphAlignment.Left;

        column = table.AddColumn();
        column.Format.Alignment = ParagraphAlignment.Center;

        column = table.AddColumn();
        column.Format.Alignment = ParagraphAlignment.Right;

        table.Rows.Height = 35;

        var row = table.AddRow();
        row.VerticalAlignment = VerticalAlignment.Top;
        row.Cells[0].AddParagraph("Text");
        row.Cells[1].AddParagraph("Text");
        row.Cells[2].AddParagraph("Text");

        row = table.AddRow();
        row.VerticalAlignment = VerticalAlignment.Center;
        row.Cells[0].AddParagraph("Text");
        row.Cells[1].AddParagraph("Text");
        row.Cells[2].AddParagraph("Text");

        row = table.AddRow();
        row.VerticalAlignment = VerticalAlignment.Bottom;
        row.Cells[0].AddParagraph("Text");
        row.Cells[1].AddParagraph("Text");
        row.Cells[2].AddParagraph("Text");
    }

    public static void DemonstrateCellMerge(Document document)
    {
        document.LastSection.AddParagraph("Cell Merge", "Heading2");

        var table = document.LastSection.AddTable();
        table.Borders.Visible = true;
        table.TopPadding = 5;
        table.BottomPadding = 5;

        var column = table.AddColumn();
        column.Format.Alignment = ParagraphAlignment.Left;

        column = table.AddColumn();
        column.Format.Alignment = ParagraphAlignment.Center;

        column = table.AddColumn();
        column.Format.Alignment = ParagraphAlignment.Right;

        table.Rows.Height = 35;

        var row = table.AddRow();
        row.Cells[0].AddParagraph("Merge Right");
        row.Cells[0].MergeRight = 1;

        row = table.AddRow();
        row.VerticalAlignment = VerticalAlignment.Bottom;
        row.Cells[0].MergeDown = 1;
        row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
        row.Cells[0].AddParagraph("Merge Down");

        table.AddRow();
    }
}