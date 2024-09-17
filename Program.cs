using PdfSharp.Fonts;
using TinySegmenterDotNet;

GlobalFontSettings.FontResolver = new MyFontResolver();

// PdfsharpHelloWorld.Run();
// MigraDocHelloWorld.Run();
// MigraDocTable.Run();
MigraDocTable2.Run();

// const char SOFT_HYPHEN = '\u00AD';
// string text = "aaabbb";

// string text2 = String.Join(SOFT_HYPHEN, text.ToCharArray());


// Console.WriteLine(text2);

// var text = "吾輩は猫である。名前はまだない。 どこで生れたか頓（とん）と見当がつかぬ。";

// //TinySegmenterオブジェクトを作成する
// TinySegmenterDotNet.TinySegmenter seg = new TinySegmenterDotNet.TinySegmenter();
// //分かち書きを行う
// string[] words = seg.Segment(text);
// //"|"で区切って表示する
// Console.WriteLine(string.Join("|", words));