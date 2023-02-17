using LibreOfficeLibrary;
using NPOI.XWPF.UserModel;
using System.Diagnostics;
using LibreOffice.Packages.Api.Clients;
using LibreOffice.Packages.Api.Options;

Console.WriteLine("Hello, World!");

//var converter = new DocumentConverter();

var converter = new ConversionApiService(new ConversionServiceOptions() { Url = "https://libreoffice-api.smile-tech.study" });

Stopwatch sw = new();
sw.Start();
const int count = 100;

//Parallel.For(0, 100, GenerateDocument);

var tasks = new List<Task>();
for (var i = 0; i < count; i++)
{
    tasks.Add(GenerateViaApi(i));
}
//Parallel.For(0, 100, i => tasks.Add(GenerateViaApi(i)));
await Task.WhenAll(tasks);

sw.Stop();
Console.WriteLine($"Done! Time per file: {sw.Elapsed.TotalSeconds / count} s");

async Task GenerateViaApi(int i)
{
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

    using var memoryStream = new MemoryStream();
    doc.Write(memoryStream);

    var result = await converter!.Convert($"{Guid.NewGuid()}.docx", memoryStream.ToArray(), "pdf");
    using var temp = File.Create($"../../../output{i}.pdf");
    temp.Write(result, 0, result.Length);
}

//void GenerateDocument(int i)
//{
//    var tempPath = $"../../../temp{i}.docx";

//    using (var input = File.OpenRead("../../../input.docx"))
//    {
//        using var temp = File.Create(tempPath);
//        var doc = new XWPFDocument(input);

//        var placeHolderDictionary = new Dictionary<string, string> {
//        { "{FirstName}", "Иван" },
//        { "{LastName}", "Иванов" } };

//        foreach (var para in doc.Paragraphs)
//        {
//            foreach (var placeholder in placeHolderDictionary)
//            {
//                // Примеры редактирования docx-файла: https://github.com/nissl-lab/npoi-examples/tree/main/xwpf
//                para.ReplaceText(placeholder.Key, placeholder.Value);
//            }
//        }

//        doc.Write(temp);
//    }
//    converter!.ConvertToPdf(tempPath, $"../../../output{i}.pdf");
//    File.Delete(tempPath);
//}