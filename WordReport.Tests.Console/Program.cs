using System.Collections.Generic;
using System.IO;

namespace WordReport.Tests.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = new
            {
                teacher = "Ben",
                author = "John Doe",
                students = new[]
                {
                    new {name = "Foo", age = 15},
                    new {name = "Bar", age = 16},
                }
            };
            var images = new Dictionary<string, byte[]>
            {
                ["signature_pic"] = File.ReadAllBytes("signature.png")
            };
            var reporter = WordTemplate.FromFile("Template.docx");
            var mem = new MemoryStream();
            reporter.Render(mem, data, images);
            File.WriteAllBytes("Output.docx", mem.ToArray());
        }
    }
}
