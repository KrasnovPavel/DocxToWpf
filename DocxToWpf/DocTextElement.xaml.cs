using System.Windows.Controls;
// This is an independent project of an individual developer. Dear PVS-Studio, please check it.
// PVS-Studio Static Code Analyzer for C, C++ and C#: http://www.viva64.com

using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Xml.Linq;

namespace DocxToWpf
{
    /// <summary>
    /// Логика взаимодействия для TextElement.xaml
    /// </summary>
    public partial class DocTextElement : UserControl
    {
        private readonly XElement _element;
        private readonly Popup _popup;
        private readonly TextBox _textEdit;
        private readonly TextBlock _textDisplay;

        public delegate void OnUpdateHandler(XElement el);
        public OnUpdateHandler OnUpdate;

        public DocTextElement(XElement element)
        {
            _element = element;
            _popup = new Popup();
            _textEdit = new TextBox();
            _textEdit.KeyUp += OnKeyUp;
            _textEdit.Text = GetValue();
//            _popup.Child = t;
            cont.Children.Add(_popup);
            
            _popup.StaysOpen = false;
            _textDisplay = new TextBlock();
            MouseEnter += OnMouseEnter;
            _textDisplay.Text = GetValue();
            cont.Children.Add(_textDisplay);
        }

        private void OnKeyUp(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                    _element.Element("data")?.Element("value")?.SetAttributeValue("value", _textEdit.Text);
                    _textDisplay.Text = _textEdit.Text;
                    OnUpdate?.Invoke(_element);
                    break;
            }
        }

        private string GetValue()
        {
            XElement val = _element.Element("data")?.Element("value");
            XAttribute v = val?.Attribute("value");
            string value = (v != null) ? v.Value : val?.Attribute("default")?.Value;
            return value;
        }

        private void OnMouseEnter(object sender, MouseEventArgs e)
        {
            _textEdit.Text = GetValue();
            _popup.IsOpen = false; 
            _popup.IsOpen = true;
        }
    }
}
