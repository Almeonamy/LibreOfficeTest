using MediaFileProcessor.Models.Common;
using MediaFileProcessor.Processors;
using NPOI.XWPF.UserModel;
using System.Diagnostics;

Console.WriteLine("Hello, World!");

var processor = new DocumentFileProcessor();

Stopwatch sw = new();
sw.Start();
const int count = 100;

Parallel.For(0, 100, async i => await GenerateDocument(i));

sw.Stop();
Console.WriteLine($"Done! Time per file: {sw.Elapsed.TotalSeconds / count} s");


async Task GenerateDocument(int i)
{
    var tempPath = $"../../../temp{i}.docx";

    using var input = File.OpenRead("../../../input.docx");
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

    using var stream = new MemoryStream();
    doc.Write(stream);
    stream.Position = 0;
    var mediaFile = new MediaFile(stream);
    var outputStream = await processor!.ConvertDocxToPdfAsStream(mediaFile);
    outputStream.Position = 0;
    var fileStream = new FileStream($"../../../output{i}.pdf", FileMode.Create);
    await outputStream.CopyToAsync(fileStream);
}