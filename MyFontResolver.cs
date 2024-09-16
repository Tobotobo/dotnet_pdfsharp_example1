using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Quality;

public class MyFontResolver : IFontResolver
{
    public string DefaultFontName => "Arial";

    public byte[] GetFont(string faceName)
    {
        switch (faceName)
        {
            case "NotoSansJP-Regular":
                return File.ReadAllBytes("fonts/NotoSansJP-Regular.ttf");
            // 他のフォントも追加可能
        }
        throw new Exception("Font not found.");
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (familyName == "NotoSansJP-Regular")
            return new FontResolverInfo("NotoSansJP-Regular");
        return new FontResolverInfo("NotoSansJP-Regular"); // デフォルトでArialを使用
    }
}