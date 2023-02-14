using LibreOfficeLibrary;
using System.Diagnostics;

Console.WriteLine("Hello, World!");

var converter = new DocumentConverter();

Stopwatch sw = new Stopwatch();
sw.Start();
const int count = 100;

for (int i = 0; i < count; i++)
{
    converter.ConvertToPdf("../../../input.docx", $"../../../output{i}.pdf");
}

sw.Stop();

Console.WriteLine($"Done! Time per file: {sw.Elapsed.TotalSeconds / count} s");