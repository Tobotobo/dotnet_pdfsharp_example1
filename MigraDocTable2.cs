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
        var pageSetup = document.DefaultPageSetup.Clone(); // DefaultPageSetup.Clone()を使わないとPageWidthなどが返ってこない→常に0
        section.PageSetup = pageSetup;
        // var pageSetup = document.DefaultPageSetup; // NG
        pageSetup.PageFormat = PageFormat.A4;
        pageSetup.Orientation = Orientation.Landscape;
        // pageSetup.OddAndEvenPagesHeaderFooter = true;
        // pageSetup.StartingNumber = 1;

        // ページの幅を取得　※Orientationの指定が無視され常に縦の場合でサイズが返ってくるので自分で入れ替える
        Unit pageWidth;
        Unit leftMargin;
        Unit rightMargin;
        if (pageSetup.Orientation == Orientation.Portrait) {
            // 縦
            pageWidth = section.PageSetup.PageWidth;
            leftMargin = section.PageSetup.LeftMargin;
            rightMargin = section.PageSetup.RightMargin;
        } else {
            // 横
            pageWidth = section.PageSetup.PageHeight;
            leftMargin = section.PageSetup.TopMargin;
            rightMargin = section.PageSetup.BottomMargin;
        }

        // 有効なページ幅（マージンを除く）
        Unit effectivePageWidth = pageWidth - leftMargin - rightMargin;

        // ヘッダー作成
        {
            HeaderFooter header = section.Headers.Primary;

            // ヘッダー用のテーブルを作成
            Table table = header.AddTable();
            table.Borders.Width = 0;

            // テーブルの列を割合で定義（例として左: 30%、中央: 40%、右: 30%）
            Column columnLeft = table.AddColumn(effectivePageWidth * 0.3);
            Column columnCenter = table.AddColumn(effectivePageWidth * 0.4);
            Column columnRight = table.AddColumn(effectivePageWidth * 0.3);

            // テーブルの列を定義
            // Column columnLeft = table.AddColumn(Unit.FromCentimeter(7));
            // Column columnCenter = table.AddColumn(Unit.FromCentimeter(7));
            // Column columnRight = table.AddColumn(Unit.FromCentimeter(7));

            // テーブルの行を作成
            Row row = table.AddRow();

            // 左、中央、右にテキストを配置
            // row.Cells[0].AddParagraph("ヘッダー左");
            row.Cells[0].AddParagraph("テーブルサンプル");
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;

            row.Cells[1].AddParagraph("ヘッダー中央");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Center;

            // row.Cells[2].AddParagraph("ヘッダー右");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Right;

            var culture = new CultureInfo("ja-JP"); 
            culture.DateTimeFormat.Calendar = new JapaneseCalendar();
            Paragraph paragraph = new();
            // paragraph.AddText("作成日時: ");
            // paragraph.AddText(DateTime.Now.ToString("ggy年M月d日 H時m分s秒", culture));
            paragraph.AddText("作成日: ");
            paragraph.AddText(DateTime.Now.ToString("ggy年M月d日", culture));
            row.Cells[2].Add(paragraph);
        }

        // フッター作成
        {
            HeaderFooter footer = section.Footers.Primary;

            // フッター用のテーブルを作成
            Table table = footer.AddTable();
            table.Borders.Width = 0;

            // テーブルの列を割合で定義（例として左: 30%、中央: 40%、右: 30%）
            Column columnLeft = table.AddColumn(effectivePageWidth * 0.3);
            Column columnCenter = table.AddColumn(effectivePageWidth * 0.4);
            Column columnRight = table.AddColumn(effectivePageWidth * 0.3);

            // テーブルの列を定義
            // Column columnLeft = table.AddColumn(Unit.FromCentimeter(7));
            // Column columnCenter = table.AddColumn(Unit.FromCentimeter(7));
            // Column columnRight = table.AddColumn(Unit.FromCentimeter(7));

            // テーブルの行を作成
            Row row = table.AddRow();

            // 左、中央、右にテキストを配置
            row.Cells[0].AddParagraph("フッター左");
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;

            // row.Cells[1].AddParagraph("フッター中央");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Center;

            row.Cells[2].AddParagraph("フッター右");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Right;

            // フッターの中央にページ番号とページ数を表示
            Paragraph paragraph = new();
            
            // Page X of Y
            // paragraph.AddText("Page ");
            // paragraph.AddPageField();
            // paragraph.AddText(" of ");
            // paragraph.AddNumPagesField();

            // X / Y ページ
            paragraph.AddPageField();
            paragraph.AddText(" / ");
            paragraph.AddNumPagesField();
            paragraph.AddText(" ページ");

            row.Cells[1].Add(paragraph);
        }
        





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
        // document.LastSection.AddParagraph("テーブルサンプル", "Heading1");

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
        column = table.AddColumn(Unit.FromCentimeter(12));
        column = table.AddColumn(Unit.FromCentimeter(12));
        // column.Format.　TODO: 自動改行

        // ヘッダー行
        var row = table.AddRow();
        row.HeadingFormat = true; // ヘッダー行を固定し次ページ以降も表示
        // row.Shading.Color = Colors.PaleGoldenrod;
        // row.Shading.Color = new Color(0xFFE699u); // NG 色がつかない。なんで？
        row.Shading.Color = new Color(255,230,153);
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

        TinySegmenterDotNet.TinySegmenter seg = new();
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

        // 3行目
        row = table.AddRow();
        cell = row.Cells[0];
        cell.AddParagraph("3");
        cell = row.Cells[1];
        cell.AddParagraph(text2);
        cell = row.Cells[2];
        cell.AddParagraph(text3);

        // table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

        document.LastSection.Add(table);
    }

    // public static void DemonstrateAlignment(Document document)
    // {
    //     document.LastSection.AddParagraph("Cell Alignment", "Heading2");

    //     var table = document.LastSection.AddTable();
    //     table.Borders.Visible = true;
    //     table.Format.Shading.Color = Colors.LavenderBlush;
    //     table.Shading.Color = Colors.Salmon;
    //     table.TopPadding = 5;
    //     table.BottomPadding = 5;

    //     var column = table.AddColumn();
    //     column.Format.Alignment = ParagraphAlignment.Left;

    //     column = table.AddColumn();
    //     column.Format.Alignment = ParagraphAlignment.Center;

    //     column = table.AddColumn();
    //     column.Format.Alignment = ParagraphAlignment.Right;

    //     table.Rows.Height = 35;

    //     var row = table.AddRow();
    //     row.VerticalAlignment = VerticalAlignment.Top;
    //     row.Cells[0].AddParagraph("Text");
    //     row.Cells[1].AddParagraph("Text");
    //     row.Cells[2].AddParagraph("Text");

    //     row = table.AddRow();
    //     row.VerticalAlignment = VerticalAlignment.Center;
    //     row.Cells[0].AddParagraph("Text");
    //     row.Cells[1].AddParagraph("Text");
    //     row.Cells[2].AddParagraph("Text");

    //     row = table.AddRow();
    //     row.VerticalAlignment = VerticalAlignment.Bottom;
    //     row.Cells[0].AddParagraph("Text");
    //     row.Cells[1].AddParagraph("Text");
    //     row.Cells[2].AddParagraph("Text");
    // }

    // public static void DemonstrateCellMerge(Document document)
    // {
    //     document.LastSection.AddParagraph("Cell Merge", "Heading2");

    //     var table = document.LastSection.AddTable();
    //     table.Borders.Visible = true;
    //     table.TopPadding = 5;
    //     table.BottomPadding = 5;

    //     var column = table.AddColumn();
    //     column.Format.Alignment = ParagraphAlignment.Left;

    //     column = table.AddColumn();
    //     column.Format.Alignment = ParagraphAlignment.Center;

    //     column = table.AddColumn();
    //     column.Format.Alignment = ParagraphAlignment.Right;

    //     table.Rows.Height = 35;

    //     var row = table.AddRow();
    //     row.Cells[0].AddParagraph("Merge Right");
    //     row.Cells[0].MergeRight = 1;

    //     row = table.AddRow();
    //     row.VerticalAlignment = VerticalAlignment.Bottom;
    //     row.Cells[0].MergeDown = 1;
    //     row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
    //     row.Cells[0].AddParagraph("Merge Down");

    //     table.AddRow();
    // }
}