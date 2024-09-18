// https://github.com/empira/PDFsharp.Samples/blob/master/src/samples/src/MigraDoc/src/HelloMigraDoc/Charts.cs

using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
// using PdfSharp.Charting;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Quality;
using PdfSharp.Snippets.Font;
using System.Globalization;

public static class MigraDocChart2
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

        // 本文
        WriteBody(document, effectivePageWidth);

        // 保存
        SavePdfDocument(document);
    }

    private static void WriteLineChart(Section section) {
        section.AddParagraph("折れ線グラフ: ChartType.Line");

        var chart = section.AddChart();
        chart.Left = 0;
        chart.Width = Unit.FromCentimeter(16);
        chart.Height = Unit.FromCentimeter(12);
        
        var series = chart.SeriesCollection.AddSeries();
        series.ChartType = ChartType.Line;
        series.Add(41, 7, 5, 45, 13, 10, 21, 13, 18, 9);
        series.HasDataLabel = true; // TODO: ラベルが表示されない
        series.DataLabel.Type = DataLabelType.Value;

        var xSeries = chart.XValues.AddXSeries();
        xSeries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        chart.XAxis.MajorTickMark = TickMarkType.Outside;
        chart.XAxis.Title.Caption = "横軸タイトル";

        chart.YAxis.MajorTickMark = TickMarkType.Outside;
        chart.YAxis.HasMajorGridlines = true;
        chart.YAxis.Title.Caption = "縦軸タイトル";
        chart.YAxis.Title.Orientation = 90;
        chart.YAxis.Title.VerticalAlignment = VerticalAlignment.Center;

        chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        chart.PlotArea.LineFormat.Width = 1;
    }

    private static void WriteColumn2DChart(Section section) {
        section.AddParagraph("縦棒グラフ: ChartType.Column2D");

        var chart = section.AddChart();
        chart.Left = 0;
        chart.Width = Unit.FromCentimeter(16);
        chart.Height = Unit.FromCentimeter(12);
        
        var series = chart.SeriesCollection.AddSeries();
        series.ChartType = ChartType.Column2D;
        series.Add(1, 17, 45, 5, 3, 20, 11, 23, 8, 19);
        series.HasDataLabel = true;

        var xSeries = chart.XValues.AddXSeries();
        xSeries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        chart.XAxis.MajorTickMark = TickMarkType.Outside;
        chart.XAxis.Title.Caption = "横軸タイトル";

        chart.YAxis.MajorTickMark = TickMarkType.Outside;
        chart.YAxis.HasMajorGridlines = true;
        chart.YAxis.Title.Caption = "縦軸タイトル";
        chart.YAxis.Title.Orientation = 90;
        chart.YAxis.Title.VerticalAlignment = VerticalAlignment.Center;

        chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        chart.PlotArea.LineFormat.Width = 1;
    }

    private static void WriteColumnStacked2DChart(Section section) {
        section.AddParagraph("積み上げ縦棒グラフ: ChartType.ColumnStacked2D");

        var chart = section.AddChart(ChartType.ColumnStacked2D);
        chart.Left = 0;
        chart.Width = Unit.FromCentimeter(16);
        chart.Height = Unit.FromCentimeter(12);
        
        var series = chart.SeriesCollection.AddSeries();
        series.Name = "Series 1";
        series.Add(1, 17, 45, 5, 3, 20, 11, 23, 8, 19);
        series.HasDataLabel = true;

        series = chart.SeriesCollection.AddSeries();
        series.Name = "Series 2";
        series.Add(31, 34, 15, 32, 43, 37, 5, 49, 44, 13);
        series.HasDataLabel = true;

        var xSeries = chart.XValues.AddXSeries();
        xSeries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        chart.XAxis.MajorTickMark = TickMarkType.Outside;
        chart.XAxis.Title.Caption = "横軸タイトル";

        chart.YAxis.MajorTickMark = TickMarkType.Outside;
        chart.YAxis.HasMajorGridlines = true;
        chart.YAxis.Title.Caption = "縦軸タイトル";
        chart.YAxis.Title.Orientation = 90;
        chart.YAxis.Title.VerticalAlignment = VerticalAlignment.Center;

        chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        chart.PlotArea.LineFormat.Width = 1;

        // TODO: 凡例を表示
    }

    private static void WriteArea2DChart(Section section) {
        section.AddParagraph("面グラフ: ChartType.Area2D");

        var chart = section.AddChart(ChartType.Area2D);
        chart.Left = 0;
        chart.Width = Unit.FromCentimeter(16);
        chart.Height = Unit.FromCentimeter(12);
        
        var series = chart.SeriesCollection.AddSeries();
        series.Name = "Series 1";
        series.Add(1, 17, 45, 5, 3, 20, 11, 23, 8, 19);
        series.HasDataLabel = true;

        series = chart.SeriesCollection.AddSeries();
        series.Name = "Series 2";
        series.Add(31, 34, 15, 32, 43, 37, 5, 49, 44, 13);
        series.HasDataLabel = true;

        var xSeries = chart.XValues.AddXSeries();
        xSeries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        chart.XAxis.MajorTickMark = TickMarkType.Outside;
        chart.XAxis.Title.Caption = "横軸タイトル";

        chart.YAxis.MajorTickMark = TickMarkType.Outside;
        chart.YAxis.HasMajorGridlines = true;
        chart.YAxis.Title.Caption = "縦軸タイトル";
        chart.YAxis.Title.Orientation = 90;
        chart.YAxis.Title.VerticalAlignment = VerticalAlignment.Center;

        chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        chart.PlotArea.LineFormat.Width = 1;

        // TODO: 凡例を表示
    }

    private static void WriteBar2DChart(Section section) {
        section.AddParagraph("横棒グラフ: ChartType.Bar2D");

        var chart = section.AddChart(ChartType.Bar2D);
        chart.Left = 0;
        chart.Width = Unit.FromCentimeter(16);
        chart.Height = Unit.FromCentimeter(12);
        
        var series = chart.SeriesCollection.AddSeries();
        series.Name = "Series 1";
        series.Add(1, 17, 45, 5, 3, 20, 11, 23, 8, 19);
        series.HasDataLabel = true;

        // XとYがひっくり返る？ X = 縦軸、y = 横軸
        var xSeries = chart.XValues.AddXSeries();
        xSeries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        chart.XAxis.MajorTickMark = TickMarkType.Outside;
        // chart.XAxis.Title.Caption = "縦軸タイトル";
        // chart.XAxis.Title.Orientation = 90; // 効果が無い
        // chart.XAxis.Title.Alignment = HorizontalAlignment.Center;
        // chart.XAxis.Title.VerticalAlignment = VerticalAlignment.Center;

        chart.YAxis.MajorTickMark = TickMarkType.Outside;
        chart.YAxis.HasMajorGridlines = true;
        chart.YAxis.Title.Caption = "横軸タイトル";
        // chart.YAxis.Title.Orientation = 90;
        chart.YAxis.Title.Alignment = HorizontalAlignment.Center;
        // chart.YAxis.Title.VerticalAlignment = VerticalAlignment.Center;

        chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        chart.PlotArea.LineFormat.Width = 1;
    }

    private static void WritePie2DChart(Section section) {
        section.AddParagraph("円グラフ: ChartType.Pie2D");

        var chart = section.AddChart(ChartType.Pie2D);
        chart.Left = 0;
        chart.Width = Unit.FromCentimeter(16);
        chart.Height = Unit.FromCentimeter(12);
        
        var series = chart.SeriesCollection.AddSeries();
        series.Name = "Series 1";
        series.Add(1, 17, 45, 5, 3, 20, 11, 23, 8, 19);
        series.HasDataLabel = true;

        // series = chart.SeriesCollection.AddSeries();
        // series.Name = "Series 2";
        // series.Add(31, 34, 15, 32, 43, 37, 5, 49, 44, 13);
        // series.HasDataLabel = true;

        // var xSeries = chart.XValues.AddXSeries();
        // xSeries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        // chart.XAxis.MajorTickMark = TickMarkType.Outside;
        // chart.XAxis.Title.Caption = "横軸タイトル";

        // chart.YAxis.MajorTickMark = TickMarkType.Outside;
        // chart.YAxis.HasMajorGridlines = true;
        // chart.YAxis.Title.Caption = "縦軸タイトル";
        // chart.YAxis.Title.Orientation = 90;
        // chart.YAxis.Title.VerticalAlignment = VerticalAlignment.Center;

        // chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        // chart.PlotArea.LineFormat.Width = 1;

        // TODO: 凡例を表示
    }

    private static void WriteBody(Document document, Unit effectivePageWidth) {
        var section = document.LastSection;
        
        // Line 　　　　　　: 折れ線グラフ
        WriteLineChart(section);

        // Column2D 　　　　: 縦棒グラフ
        WriteColumn2DChart(section);

        // ColumnStacked2D : 積み上げ縦棒グラフ
        WriteColumnStacked2DChart(section);

        // Area2D 　　　　　: 面グラフ
        WriteArea2DChart(section);

        // Bar2D 　　　　　　: 横棒グラフ
        WriteBar2DChart(section);

        // BarStacked2D 　　: 2D積み上げ横棒グラフ
        // Pie2D 　　　　　: 2D円グラフ
        WritePie2DChart(section);

        // PieExploded2D 　: 2D分割円グラフ
    



        // var paragraph = document.LastSection.AddParagraph("Chart Overview", "Heading1");
        // paragraph.AddBookmark("Charts");

        // document.LastSection.AddParagraph("Sample Chart", "Heading2");

        // var chart = new Chart();
        // chart.Left = 0;

        // chart.Width = Unit.FromCentimeter(16);
        // chart.Height = Unit.FromCentimeter(12);
        // var series = chart.SeriesCollection.AddSeries();
        // series.ChartType = ChartType.Column2D;
        // series.Add(1, 17, 45, 5, 3, 20, 11, 23, 8, 19);
        // series.HasDataLabel = true;

        // series = chart.SeriesCollection.AddSeries();
        // series.ChartType = ChartType.Line;
        // series.Add(41, 7, 5, 45, 13, 10, 21, 13, 18, 9);

        // var xSeries = chart.XValues.AddXSeries();
        // xSeries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        // chart.XAxis.MajorTickMark = TickMarkType.Outside;
        // chart.XAxis.Title.Caption = "X-Axis";

        // chart.YAxis.MajorTickMark = TickMarkType.Outside;
        // chart.YAxis.HasMajorGridlines = true;

        // chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        // chart.PlotArea.LineFormat.Width = 1;

        // document.LastSection.Add(chart);
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

        // // 透かし
        // {
        //     var watermark = "透かしサンプル";
        //     var font = new XFont("NotoSansJP-Regular", 100, XFontStyleEx.Bold);
        //     foreach(var page in pdfRenderer.PdfDocument.Pages) {
        //         using (var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend))
        //         {
        //             // Get the size (in points) of the text.
        //             var size = gfx.MeasureString(watermark, font);
 
        //             // Define a rotation transformation at the center of the page.
        //             gfx.TranslateTransform(page.Width.Point / 2, page.Height.Point / 2);
        //             gfx.RotateTransform(-Math.Atan(page.Height.Point / page.Width.Point) * 180 / Math.PI);
        //             gfx.TranslateTransform(-page.Width.Point / 2, -page.Height.Point / 2);
 
        //             // Create a string format.
        //             var format = new XStringFormat();
        //             format.Alignment = XStringAlignment.Near;
        //             format.LineAlignment = XLineAlignment.Near;
 
        //             // Create a dimmed red brush.
        //             XBrush brush = new XSolidBrush(XColor.FromArgb(80, 128, 128, 128));
 
        //             // Draw the string.
        //             gfx.DrawString(watermark, font, brush,
        //                 new XPoint((page.Width.Point - size.Width) / 2, (page.Height.Point - size.Height) / 2),
        //                 format);

                
        //             // var pen = new XPen(XColors.Red, 2);
        //             // gfx.DrawLine(pen, 50, 50, 200, 200);
        //             // gfx.DrawString("Custom Text", new XFont("NotoSansJP-Regular", 20), XBrushes.Black, new XPoint(100, 100));
        //         }
        //     } 
        // }
        

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
        row.Cells[0].AddParagraph("チャートサンプル");
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