// https://github.com/empira/PDFsharp.Samples/blob/master/src/samples/src/MigraDoc/src/HelloMigraDoc/Tables.cs

using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Quality;
using PdfSharp.Snippets.Font;
using System.Globalization;

public static class MigraDocReport
{
    public static void Run()
    {
        var document = new Document();

        // ページ設定
        var section = document.LastSection;
        section.PageSetup = document.DefaultPageSetup.Clone(); // DefaultPageSetup.Clone()を使わないとPageWidthなどが返ってこない→常に0
        section.PageSetup.PageFormat = PageFormat.A4;
        section.PageSetup.Orientation = Orientation.Portrait;
        // section.PageSetup.OddAndEvenPagesHeaderFooter = true;
        // section.PageSetup.StartingNumber = 1;

        // 有効なページ幅（マージンを除く）
        Unit effectivePageWidth = GetEffectivePageWidth(section.PageSetup);

        // ヘッダー作成
        SetupHeader(section.Headers.Primary, effectivePageWidth);

        // フッター作成
        SetupFooter(section.Footers.Primary, effectivePageWidth);

        // 透かし追加
        // SetupWatermark(section);

        // 本文
        WriteBody(document, effectivePageWidth);

        // 保存
        SavePdfDocument(document);
    }

    private static void SetupWatermark(Section section) {
        // // 透かし用のTextFrameを作成
        // TextFrame watermarkFrame = section.Headers.Primary.AddTextFrame();
        // watermarkFrame.Orientation = TextOrientation.;
        // watermarkFrame.Width = Unit.FromCentimeter(20); // フレームの幅
        // watermarkFrame.Height = Unit.FromCentimeter(20); // フレームの高さ
        // watermarkFrame.RelativeVertical = RelativeVertical.Page;
        // watermarkFrame.RelativeHorizontal = RelativeHorizontal.Page;
        // watermarkFrame.Top = ShapePosition.Center;
        // watermarkFrame.Left = ShapePosition.Center;

        // // 透かしテキストを追加
        // var watermark = watermarkFrame.AddParagraph("透かしテキスト");
        // watermark.Format.Font.Size = 80; // 大きなフォントサイズ
        // watermark.Format.Font.Color = Colors.LightGray; // 薄いグレー色
        // watermark.Format.Alignment = ParagraphAlignment.Center; // 中央揃え
        // watermark.Format.Font.Bold = true;

        // // 傾ける場合（45度の角度をつける）
        // watermarkFrame.WrapFormat.Style = WrapStyle.None;
        // watermarkFrame.LineFormat.Color = Colors.LightGray;
        // watermarkFrame.RelativeVertical = RelativeVertical.Page;
        // watermarkFrame.RelativeHorizontal = RelativeHorizontal.Page;
        // watermarkFrame.R
        // // watermarkFrame.FillFormat.Opacity = 50; // 透過性を設定


        // // 透かしテキストの作成
        // var watermark = section.Headers.Primary.AddParagraph();
        // watermark.AddFormattedText("開発中", TextFormat.Bold);
        // watermark.Format.Font.Size = 80; // フォントサイズ
        // watermark.Format.Font.Color = Colors.LightGray; // 薄いグレー色
        // watermark.Format.Alignment = ParagraphAlignment.Center; // 中央揃え
        // watermark.Format.Font.Bold = true;

        // // 透かしテキストの位置と回転を調整
        // watermark.Format.SpaceBefore = Unit.FromCentimeter(10); // 上部にスペースを設定してセンターに配置
        // watermark.Format.Font.Color = Colors.LightGray;
        // watermark.Format.Font.Size = 100;
        // watermark.Format.Font.Bold = true;
        // watermark.Format.LeftIndent = Unit.FromCentimeter(0);
        // watermark.Format.RightIndent = Unit.FromCentimeter(0);
        // watermark.Format.Borders.DistanceFromTop = 0;
        // watermark.Format.Borders.DistanceFromBottom = 0;
    }

    private static void WriteBody(Document document, Unit effectivePageWidth) {
        var section = document.LastSection;


        // 発行日、No
        {
            // ↓使いまわし反対
            Table table;
            Column column; 
            Row row;

            table = new();
            table.Borders.Width = 0;
            
            column = table.AddColumn(Unit.FromCentimeter(1));
            column = table.AddColumn(Unit.FromCentimeter(4));
            column.Format.Alignment = ParagraphAlignment.Right;
            
            // 発行日
            row = table.AddRow();
            row.Cells[0].AddParagraph("発行日");
            row.Cells[1].AddParagraph("9999/12/30");
            row.Borders.Bottom.Width = 0.75;
            
            // No
            row = table.AddRow();
            row.Cells[0].AddParagraph("No");
            row.Cells[1].AddParagraph("1234567890");
            row.Borders.Bottom.Width = 0.75;

            // テーブル自体を右寄せ
            // table.Rows.LeftIndent = effectivePageWidth - table.Columns.Width; // table.Columns.Widthの結果が合ってない？
            table.Rows.LeftIndent = effectivePageWidth - GetColumnsWidth(table.Columns);
            section.Add(table);
        }

        // タイトル
        {
            Paragraph paragraph = new();
            paragraph.AddFormattedText("御見積書");
            paragraph.Format.Font.Size = 20;
            paragraph.Format.SpaceBefore = 10;
            paragraph.Format.SpaceAfter = 25; // 下に余白を作る
            section.Add(paragraph);
        }

        // 御中
        {
            // ↓使いまわし反対
            Table table;
            Column column; 
            Row row;

            table = new();
            table.Borders.Width = 0;
            
            column = table.AddColumn(Unit.FromCentimeter(9));
            column = table.AddColumn(Unit.FromCentimeter(1));
            
            // 御中
            row = table.AddRow();
            row.Cells[0].AddParagraph("東京　太郎");
            row.Cells[0].Format.Font.Size = 14;
            row.Cells[1].AddParagraph("御中");
            row.Cells[1].VerticalAlignment = VerticalAlignment.Bottom;
            row.Borders.Bottom.Width = 0.75;

            // テーブル自体を右寄せ
            // table.Rows.LeftIndent = effectivePageWidth - table.Columns.Width; // table.Columns.Widthの結果が合ってない？
            // table.Rows.LeftIndent = effectivePageWidth - GetColumnsWidth(table.Columns);
            section.Add(table);
        }

        // 余白
        section.Add(CreateSpace(10));

        // 件名、納期、納品場所、支払条件、見積期限
        {
            // ↓使いまわし反対
            Table table;
            Column column; 
            Row row;

            table = new();
            table.Borders.Width = 0;
            
            column = table.AddColumn(Unit.FromCentimeter(2));
            column = table.AddColumn(Unit.FromCentimeter(8));
            // column.Format.Alignment = ParagraphAlignment.Right;
            
            // 件名
            row = table.AddRow();
            row.Cells[0].AddParagraph("件　　名:");
            row.Cells[1].AddParagraph("サンプル製品納品依頼");
            row.Borders.Bottom.Width = 0.75;
            
            // 納期
            row = table.AddRow();
            row.Cells[0].AddParagraph("納　　期:");
            row.Cells[1].AddParagraph("2024年10月15日");
            row.Borders.Bottom.Width = 0.75;

            // 納品場所
            row = table.AddRow();
            row.Cells[0].AddParagraph("納品場所:");
            row.Cells[1].AddParagraph("東京都千代田区丸の内1-1-1 サンプルビル3階");
            row.Borders.Bottom.Width = 0.75;

            // 支払条件
            row = table.AddRow();
            row.Cells[0].AddParagraph("支払条件:");
            row.Cells[1].AddParagraph("30日以内に銀行振込");
            row.Borders.Bottom.Width = 0.75;

            // 見積期限
            row = table.AddRow();
            row.Cells[0].AddParagraph("見積期限:");
            row.Cells[1].AddParagraph("2024年9月30日");
            row.Borders.Bottom.Width = 0.75;

            // テーブル自体を右寄せ
            // table.Rows.LeftIndent = effectivePageWidth - table.Columns.Width; // table.Columns.Widthの結果が合ってない？
            // table.Rows.LeftIndent = effectivePageWidth - GetColumnsWidth(table.Columns);
            section.Add(table);
        }

        // 文言
        {
            Paragraph paragraph = new();
            paragraph.AddFormattedText("下記の通り、お見積申し上げます。");
            // paragraph.Format.Font.Size = 20;
            paragraph.Format.SpaceBefore = 10;
            paragraph.Format.SpaceAfter = 12; // 下に余白を作る
            section.Add(paragraph);
        }

        // 金額
        {
            // ↓使いまわし反対
            Table table;
            Column column; 
            Row row;

            table = new();
            table.Borders.Width = 0;
            
            column = table.AddColumn(Unit.FromCentimeter(2));
            column = table.AddColumn(Unit.FromCentimeter(6.8));
            column.Format.Alignment = ParagraphAlignment.Right;
            column = table.AddColumn(Unit.FromCentimeter(1.2));
            // column.Format.Alignment = ParagraphAlignment.Right;
            
            // 件名
            row = table.AddRow();
            row.Cells[0].AddParagraph("金額");
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[1].AddParagraph("\u00A5999,999,999");
            row.Cells[1].Format.Font.Size = 20;
            row.Cells[2].AddParagraph("(税込)");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].VerticalAlignment = VerticalAlignment.Bottom;
            row.Borders.Bottom.Width = 0.75;

            // 二重下線用
            row = table.AddRow();
            row.Format.Font.Size = 0.75;
            row.Borders.Bottom.Width = 0.75;



            // テーブル自体を右寄せ
            // table.Rows.LeftIndent = effectivePageWidth - table.Columns.Width; // table.Columns.Widthの結果が合ってない？
            // table.Rows.LeftIndent = effectivePageWidth - GetColumnsWidth(table.Columns);
            section.Add(table);
        }



        // document.LastSection.AddParagraph("テーブルサンプル", "Heading1");

//         var table = new Table
//         {
//             Borders =
//                 {
//                     Width = 0.75
//                 }
//         };

//         // 列を作成
//         var column = table.AddColumn(Unit.FromCentimeter(1));
//         // column.Format.Alignment = ParagraphAlignment.Center;
//         column.Format.Alignment = ParagraphAlignment.Right;
//         column = table.AddColumn(Unit.FromCentimeter(7));
//         column = table.AddColumn(Unit.FromCentimeter(7));
//         // column.Format.　TODO: 自動改行

//         // ヘッダー行
//         var row = table.AddRow();
//         row.HeadingFormat = true; // ヘッダー行を固定し次ページ以降も表示
//         // row.Shading.Color = Colors.PaleGoldenrod;
//         // row.Shading.Color = new Color(0xFFE699u); // NG 色がつかない。なんで？
//         row.Shading.Color = new Color(255,230,153);
//         var cell = row.Cells[0];
//         cell.AddParagraph("#");
//         cell = row.Cells[1];
//         cell.AddParagraph("内容");
//         cell = row.Cells[2];
//         cell.AddParagraph("内容　※自動改行位置調整");

//         // 1行目
//         row = table.AddRow();
//         cell = row.Cells[0];
//         cell.AddParagraph("1");
//         cell = row.Cells[1];

//         var text = @"Andigna cons nonsectem accummo diamet nis diat.Andigna cons nonsectem accummo diamet nis diat.";
//         // const char SOFT_HYPHEN = '\u00AD';
//         const char SOFT_HYPHEN = '\u200B'; // ゼロ幅スペース（Zero Width Space, ZWSP, U+200B）
//         var text2 = String.Join(SOFT_HYPHEN, text.ToCharArray());


//         cell.AddParagraph(text2);

//         cell = row.Cells[2];
//         cell.AddParagraph(text);

//         text = @"吾輩は猫である。名前はまだない。 どこで生れたか頓（とん）と見当がつかぬ。何でも薄暗いじめじめした所でニャーニャー泣いていた事だけは記憶している。吾輩はここで始めて人間というものを見た。しかもあとで聞くとそれは書生という人間中で一番獰悪（どうあく）な種族であったそうだ。
// この書生というのは時々我々を捕（つかま）えて煮て食うという話である。
// しかしその当時は何という考（かんがえ）もなかったから別段恐しいとも思わなかった。
// ただ彼の掌（てのひら）に載せられてスーと持ち上げられた時何だかフワフワした感じがあったばかりである。
// 掌の上で少し落ち付いて書生の顔を見たのがいわゆる人間というものの見始（みはじめ）であろう。
// この時妙なものだと思った感じが今でも残っている。第一毛を以て装飾されべきはずの顔がつるつるしてまるで薬缶（やかん）だ。
// その後猫にも大分逢（あ）ったがこんな片輪には一度も出会（でく）わした事がない。のみならず顔の真中が余りに突起している。
// そうしてその穴の中から時々ぷうぷうと烟（けむり）を吹く。どうも咽（む）せぽくて実に弱った。
// これが人間の飲む烟草（タバコ）というものである事は漸（ようや）くこの頃（ごろ）知った。";
//         text2 = String.Join(SOFT_HYPHEN, text.ToCharArray());

//         TinySegmenterDotNet.TinySegmenter seg = new();
// //分かち書きを行う
// string[] words = seg.Segment(text);
//         var text3 = String.Join(SOFT_HYPHEN, words);

//         // 2行目
//         row = table.AddRow();
//         cell = row.Cells[0];
//         cell.AddParagraph("2");
//         cell = row.Cells[1];
//         cell.AddParagraph(text2);
//         cell = row.Cells[2];
//         cell.AddParagraph(text3);

//         // 3行目
//         row = table.AddRow();
//         cell = row.Cells[0];
//         cell.AddParagraph("3");
//         cell = row.Cells[1];
//         cell.AddParagraph(text2);
//         cell = row.Cells[2];
//         cell.AddParagraph(text3);

//         // table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

//         document.LastSection.Add(table);
    }

    private static Paragraph CreateSpace(Unit height) {
        Paragraph paragraph = new();
        paragraph.Format.Font.Size = 0;
        paragraph.Format.SpaceAfter = height;
        return paragraph;
    }

    private static Unit GetColumnsWidth(Columns columns) {
        Unit totalColumnWidth = Unit.Zero;
        foreach (var column in columns)
        {
            totalColumnWidth += ((Column)column!).Width;
        }
        return totalColumnWidth;
    }

    private static void SavePdfDocument(Document document) {
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

        // 透かし
        {
            var watermark = "透かしサンプル";
            var font = new XFont("NotoSansJP-Regular", 100, XFontStyleEx.Bold);
            foreach(var page in pdfRenderer.PdfDocument.Pages) {
                using (var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend))
                {
                    // Get the size (in points) of the text.
                    var size = gfx.MeasureString(watermark, font);
 
                    // Define a rotation transformation at the center of the page.
                    gfx.TranslateTransform(page.Width / 2, page.Height / 2);
                    gfx.RotateTransform(-Math.Atan(page.Height / page.Width) * 180 / Math.PI);
                    gfx.TranslateTransform(-page.Width / 2, -page.Height / 2);
 
                    // Create a string format.
                    var format = new XStringFormat();
                    format.Alignment = XStringAlignment.Near;
                    format.LineAlignment = XLineAlignment.Near;
 
                    // Create a dimmed red brush.
                    XBrush brush = new XSolidBrush(XColor.FromArgb(80, 128, 128, 128));
 
                    // Draw the string.
                    gfx.DrawString(watermark, font, brush,
                        new XPoint((page.Width - size.Width) / 2, (page.Height - size.Height) / 2),
                        format);

                
                    // var pen = new XPen(XColors.Red, 2);
                    // gfx.DrawLine(pen, 50, 50, 200, 200);
                    // gfx.DrawString("Custom Text", new XFont("NotoSansJP-Regular", 20), XBrushes.Black, new XPoint(100, 100));
                }
            } 
        }
        

        // Save the document...
        var filename = PdfFileUtility.GetTempPdfFullFileName("samples-MigraDoc/HelloWorldMigraDoc");
        pdfRenderer.PdfDocument.Save(filename);
    }

    private static Unit GetEffectivePageWidth(PageSetup pageSetup) {
        // ページの幅を取得　※Orientationの指定が無視され常に縦の場合でサイズが返ってくるので自分で入れ替える
        Unit pageWidth;
        Unit leftMargin;
        Unit rightMargin;
        if (pageSetup.Orientation == Orientation.Portrait) {
            // 縦
            pageWidth = pageSetup.PageWidth;
            leftMargin = pageSetup.LeftMargin;
            rightMargin = pageSetup.RightMargin;
        } else {
            // 横
            pageWidth = pageSetup.PageHeight;
            leftMargin = pageSetup.TopMargin;
            rightMargin = pageSetup.BottomMargin;
        }

        // 有効なページ幅（マージンを除く）
        Unit effectivePageWidth = pageWidth - leftMargin - rightMargin;
        return effectivePageWidth;
    }

    private static void SetupFooter(HeaderFooter footer, Unit effectivePageWidth)
    {
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

    private static void SetupHeader(HeaderFooter header, Unit effectivePageWidth)
    {
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
        row.Cells[0].AddParagraph("見積書サンプル");
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
}