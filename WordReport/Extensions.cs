using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;

namespace WordReport
{
    internal static class Extensions
    {
        public static void CleanRun(this WordprocessingDocument document)
        {
            foreach (var p in document.MainDocumentPart.Document.Descendants<Paragraph>())
            {
                p.CleanRun();
            }

            foreach (var part in document.MainDocumentPart.HeaderParts)
            {
                foreach (var p in part.Header.Descendants<Paragraph>())
                {
                    p.CleanRun();
                }
            }

            foreach (var part in document.MainDocumentPart.FooterParts)
            {
                foreach (var p in part.Footer.Descendants<Paragraph>())
                {
                    p.CleanRun();
                }
            }
        }

        public static void CleanRun(this Paragraph p)
        {
            var list = p.ChildElements.ToList();
            var cleaner = new WordCleaner();
            foreach (var item in list)
            {
                cleaner.Clean(item, p);
            }
        }

        public static void ReplaceImages(this WordprocessingDocument doc, Dictionary<string, byte[]> dict)
        {
            var docPart = doc.MainDocumentPart.Document;
            foreach (var prop in docPart.Descendants<DocProperties>())
            {
                if (prop.Name == null || !dict.TryGetValue(prop.Name, out var data))
                {
                    continue;
                }

                var picture = prop.Parent.Descendants<Picture>().FirstOrDefault();
                var id = picture?.BlipFill?.Blip?.Embed?.Value;
                if (id == null)
                {
                    continue;
                }

                var img = doc.MainDocumentPart.GetPartById(id);
                var mem = new MemoryStream(data);
                img.FeedData(mem);
            }
        }

        public static bool EndsWith(this string str, char value)
        {
            int lastPos = str.Length - 1;
            return ((uint)lastPos < (uint)str.Length) && str[str.Length - 1] == value;
        }

        public static bool StartsWith(this string str, char value)
        {
            return str.Length != 0 && str[0] == value;
        }

        static readonly Regex regex = new Regex(@"\{\{.*?\}\}", RegexOptions.Compiled | RegexOptions.Multiline);
        public static string DecodeExpression(this string xml)
        {
            return regex.Replace(xml, m => WebUtility.HtmlDecode(m.Value));
        }

    }
}
