using System;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordReport
{
    class WordCleaner
    {
        private enum RunState
        {
            None,
            Starting,
            Started,
            Continuing,
        }

        private RunState _state;
        private Text _firstText;
        private string _lastText;
        private readonly StringBuilder _texts = new StringBuilder();

        public void Clean(OpenXmlElement item, Paragraph p)
        {
            if (!(item is Run r))
            {
                return;
            }

            var text = r.GetFirstChild<Text>();
            if (text == null)
            {
                return;
            }

            var str = text.Text;
            switch (_state)
            {
                case RunState.None: //none => none, starting, started
                    TryStart(text);
                    break;
                case RunState.Starting: //starting => none, started
                    _state = str.StartsWith('{') ? RunState.Started : RunState.None;
                    break;
                case RunState.Continuing:   //continuing => started
                    str = _lastText + str;
                    _lastText = null;
                    _state = RunState.Started;
                    break;
            }

            if (_state != RunState.None)
            {
                _texts.Append(text.Text);
                if (!ReferenceEquals(_firstText, text))
                {
                    p.RemoveChild(r);
                }
            }

            //started => started, continuing, ending, none
            if (_state == RunState.Started)
            {
                TryEnd(str);
            }
        }

        private void TryStart(Text text)
        {
            _state = text.Text.EndsWith("{{")
                ? RunState.Started
                : text.Text.EndsWith('{')
                ? RunState.Starting
                : text.Text.Contains("{{") 
                ? RunState.Started
                : RunState.None;
            if (_state != RunState.None)
            {
                _texts.Clear();
                _firstText = text;
            }
        }

        private void TryEnd(string str)
        {
            if (!str.EndsWith("{{") && str.EndsWith('{'))
            {
                _lastText = str;
                _state = RunState.Continuing;
                return;
            }

            var i = str.LastIndexOf("}}", StringComparison.Ordinal);
            if (i >= 0)
            {
                var j = str.LastIndexOf("{{", StringComparison.Ordinal);
                if (j <= i)
                {
                    _firstText.Text = _texts.ToString();
                    //xml:space="preserve
                    _firstText.SetAttribute(new OpenXmlAttribute("space", XNamespace.Xml.NamespaceName, "preserve"));
                    _state = RunState.None;
                    _firstText = null;
                }
            }
            else if (str.EndsWith('}'))
            {
                _lastText = str;
                _state = RunState.Continuing;
            }
        }

    }
}
