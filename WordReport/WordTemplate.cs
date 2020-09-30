using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Scriban;
using Scriban.Runtime;

namespace WordReport
{
    public class WordTemplate
    {
        private readonly Template _template;

        internal WordTemplate(Template template)
        {
            _template = template;
        }

        public static WordTemplate FromFile(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                return FromStream(fs);
            }
        }

        public static WordTemplate FromStream(Stream stream)
        {
            var mem = new MemoryStream();
            stream.CopyTo(mem);
            var doc = WordprocessingDocument.Open(mem, true);
            doc.CleanRun();
            doc.Save();
            var xml = doc.MainDocumentPart.OpenXmlPackage.ToFlatOpcString().DecodeExpression();
            return new WordTemplate(Template.Parse(xml));
        }

        public void Render(Stream stream, object data, Dictionary<string, byte[]> imageDict = null)
        {
            var scriptObject = new ScriptObject();
            if (data != null)
            {
                scriptObject.Import(data);
            }
            Render(stream, scriptObject, imageDict);
        }

        public void Render(Stream stream, ScriptObject scriptObject, Dictionary<string, byte[]> imageDict = null)
        {
            var context = new XmlTemplateContext();
            context.PushGlobal(scriptObject);
            Render(stream, context, imageDict);
        }

        public void Render(Stream stream, TemplateContext context, Dictionary<string, byte[]> imageDict = null)
        {
            var transformed = _template.Render(context);
            var doc = WordprocessingDocument.FromFlatOpcString(transformed);
            if (imageDict != null)
            {
                doc.ReplaceImages(imageDict);
            }
            doc.Clone(stream);
        }
    }
}
