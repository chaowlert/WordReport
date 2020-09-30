using System.Net;
using Scriban;
using Scriban.Parsing;

namespace WordReport
{
    public class XmlTemplateContext : TemplateContext
    {
        public string DefaultDecimalFormat { get; set; } = "N2";
        public string DefaultIntegerFormat { get; set; } = "N0";

        public XmlTemplateContext()
        {
            EnableRelaxedMemberAccess = true;
        }

        public override string ToString(SourceSpan span, object value)
        {
            switch (value)
            {
                case string s:
                    return WebUtility.HtmlEncode(s);
                case decimal d:
                    return d.ToString(DefaultDecimalFormat);
                case int n:
                    return n.ToString(DefaultIntegerFormat);
                default:
                    return base.ToString(span, value);
            }
        }
    }
}
