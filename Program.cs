using PdfSharp.Fonts;

GlobalFontSettings.FontResolver = new MyFontResolver();

PdfsharpHelloWorld.Run();
MigraDocHelloWorld.Run();