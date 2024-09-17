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

public static class MigraDocTable2
{
    public static void Run()
    {
        // Create a MigraDoc document.
        var document = CreateDocument();

        var section = document.LastSection;
        var pageSetup = section.PageSetup;
        pageSetup.PageFormat = PageFormat.A4;
        pageSetup.Orientation = Orientation.Landscape;

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
        // var paragraph = document.LastSection!.AddParagraph("Table Overview", "Heading1");
        // paragraph.AddBookmark("Tables");

        DemonstrateSimpleTable(document);
        // DemonstrateAlignment(document);
        // DemonstrateCellMerge(document);
    }

    public static void DemonstrateSimpleTable(Document document)
    {
        document.LastSection.AddParagraph("テーブルサンプル", "Heading1");

        var table = new Table
        {
            Borders =
                {
                    Width = 0.75
                }
        };

        // 列を作成
        var column = table.AddColumn(Unit.FromCentimeter(1));
        // column.Format.Alignment = ParagraphAlignment.Center;
        column.Format.Alignment = ParagraphAlignment.Right;
        column = table.AddColumn(Unit.FromCentimeter(10));
        column = table.AddColumn(Unit.FromCentimeter(10));
        // column.Format.　TODO: 自動改行
        var row = table.AddRow();
        row.Shading.Color = Colors.PaleGoldenrod;
        var cell = row.Cells[0];
        cell.AddParagraph("#");
        cell = row.Cells[1];
        cell.AddParagraph("内容");
        cell = row.Cells[2];
        cell.AddParagraph("内容　※自動改行位置調整");

        // 1行目
        row = table.AddRow();
        cell = row.Cells[0];
        cell.AddParagraph("1");
        cell = row.Cells[1];

        var text = @"Andigna cons nonsectem accummo diamet nis diat.Andigna cons nonsectem accummo diamet nis diat.";
        // const char SOFT_HYPHEN = '\u00AD';
        const char SOFT_HYPHEN = '\u200B'; // ゼロ幅スペース（Zero Width Space, ZWSP, U+200B）
        var text2 = String.Join(SOFT_HYPHEN, text.ToCharArray());


        cell.AddParagraph(text2);

        cell = row.Cells[2];
        cell.AddParagraph(text);

        text = @"吾輩は猫である。名前はまだない。 どこで生れたか頓（とん）と見当がつかぬ。何でも薄暗いじめじめした所でニャーニャー泣いていた事だけは記憶している。吾輩はここで始めて人間というものを見た。しかもあとで聞くとそれは書生という人間中で一番獰悪（どうあく）な種族であったそうだ。
この書生というのは時々我々を捕（つかま）えて煮て食うという話である。
しかしその当時は何という考（かんがえ）もなかったから別段恐しいとも思わなかった。
ただ彼の掌（てのひら）に載せられてスーと持ち上げられた時何だかフワフワした感じがあったばかりである。
掌の上で少し落ち付いて書生の顔を見たのがいわゆる人間というものの見始（みはじめ）であろう。
この時妙なものだと思った感じが今でも残っている。第一毛を以て装飾されべきはずの顔がつるつるしてまるで薬缶（やかん）だ。
その後猫にも大分逢（あ）ったがこんな片輪には一度も出会（でく）わした事がない。のみならず顔の真中が余りに突起している。
そうしてその穴の中から時々ぷうぷうと烟（けむり）を吹く。どうも咽（む）せぽくて実に弱った。
これが人間の飲む烟草（タバコ）というものである事は漸（ようや）くこの頃（ごろ）知った。";
        text2 = String.Join(SOFT_HYPHEN, text.ToCharArray());

        TinySegmenterDotNet.TinySegmenter seg = new TinySegmenterDotNet.TinySegmenter();
//分かち書きを行う
string[] words = seg.Segment(text);
        var text3 = String.Join(SOFT_HYPHEN, words);

        // 2行目
        row = table.AddRow();
        cell = row.Cells[0];
        cell.AddParagraph("2");
        cell = row.Cells[1];
        cell.AddParagraph(text2);
        cell = row.Cells[2];
        cell.AddParagraph(text3);

        // table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

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