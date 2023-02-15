using LibreOfficeLibrary;
using NPOI.XWPF.UserModel;
using System.Diagnostics;

Console.WriteLine("Hello, World!");

var converter = new DocumentConverter();

Stopwatch sw = new();
sw.Start();
const int count = 100;

Parallel.For(0, 100, GenerateDocument);

sw.Stop();
Console.WriteLine($"Done! Time per file: {sw.Elapsed.TotalSeconds / count} s");

void GenerateDocument(int i)
{
    var tempPath = $"../../../temp{i}.docx";

    using (var input = File.OpenRead("../../../input.docx"))
    {
        using var temp = File.Create(tempPath);
        var doc = new XWPFDocument(input);

        var placeHolderDictionary = new Dictionary<string, string> {
        { "{FirstName}", "Иван" },
        { "{LastName}", "Иванов" } };

        foreach (var para in doc.Paragraphs)
        {
            foreach (var placeholder in placeHolderDictionary)
            {
                // Примеры редактирования docx-файла: https://github.com/nissl-lab/npoi-examples/tree/main/xwpf
                para.ReplaceText(placeholder.Key, placeholder.Value);
            }
        }

        doc.Write(temp);
    }
    converter!.ConvertToPdf(tempPath, $"../../../output{i}.pdf");
    File.Delete(tempPath);
}