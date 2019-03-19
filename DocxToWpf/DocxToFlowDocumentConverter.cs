using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Xml;
using System.Xml.Linq;

namespace DocxToWpf
{
    class DocxToFlowDocumentConverter : DocxReader
    {
        /// <summary> Наименования xml тегов из которых состоит docx файл. </summary>
        private const string
            // Run properties elements
            BoldElement = "b",
            ItalicElement = "i",
            UnderlineElement = "u",
            StrikeElement = "strike",
            DoubleStrikeElement = "dstrike",
            VerticalAlignmentElement = "vertAlign",
            ColorElement = "color",
            HighlightElement = "highlight",
            FontElement = "rFonts",
            FontSizeElement = "sz",
            RightToLeftTextElement = "rtl",

            // Paragraph properties elements
            AlignmentElement = "jc",
            PageBreakBeforeElement = "pageBreakBefore",
            SpacingElement = "spacing",
            IndentationElement = "ind",
            ShadingElement = "shd",

            // Control properties elements
            AliasElement = "alias",
            TagElement = "tag",

            // Attributes
            IdAttribute = "id",
            ValueAttribute = "val",
            ColorAttribute = "color",
            AsciiFontFamily = "ascii",
            SpacingAfterAttribute = "after",
            SpacingBeforeAttribute = "before",
            LeftIndentationAttribute = "left",
            RightIndentationAttribute = "right",
            HangingIndentationAttribute = "hanging",
            FirstLineIndentationAttribute = "firstLine",
            FillAttribute = "fill";
        // Note: new members should also be added to nameTable in CreateNameTable method.

        private FlowDocument _document;
        private TextElement _current;// { set { if (!isCurrentUIElement) current = value; } get { return current; } }
        private XElement _currentLabel;
        private bool _hasAnyHyperlink;
        private bool _isCurrentUiElement;
        public XDocument labels;

        public FlowDocument Document => _document;

        public DocxToFlowDocumentConverter(Stream stream)
            : base(stream)
        {
            // WTF?
            labels = XDocument.Parse("<labels></labels>");
            //labels.Add("<labels></labels>");
        }

        protected override XmlNameTable CreateNameTable()
        {
            XmlNameTable nameTable = base.CreateNameTable();

            nameTable.Add(BoldElement);
            nameTable.Add(ItalicElement);
            nameTable.Add(UnderlineElement);
            nameTable.Add(StrikeElement);
            nameTable.Add(DoubleStrikeElement);
            nameTable.Add(VerticalAlignmentElement);
            nameTable.Add(ColorElement);
            nameTable.Add(HighlightElement);
            nameTable.Add(FontElement);
            nameTable.Add(FontSizeElement);
            nameTable.Add(RightToLeftTextElement);
            nameTable.Add(AlignmentElement);
            nameTable.Add(PageBreakBeforeElement);
            nameTable.Add(SpacingElement);
            nameTable.Add(IndentationElement);
            nameTable.Add(ShadingElement);
            nameTable.Add(IdAttribute);
            nameTable.Add(ValueAttribute);
            nameTable.Add(ColorAttribute);
            nameTable.Add(AsciiFontFamily);
            nameTable.Add(SpacingAfterAttribute);
            nameTable.Add(SpacingBeforeAttribute);
            nameTable.Add(LeftIndentationAttribute);
            nameTable.Add(RightIndentationAttribute);
            nameTable.Add(HangingIndentationAttribute);
            nameTable.Add(FirstLineIndentationAttribute);
            nameTable.Add(FillAttribute);

            nameTable.Add(AliasElement);
            nameTable.Add(TagElement);

            return nameTable;
        }

        protected override void ReadDocument(XmlReader reader)
        {
            _document = new FlowDocument();
            _document.BeginInit();
            _document.ColumnWidth = double.NaN;

            base.ReadDocument(reader);

            if (_hasAnyHyperlink)
            {
                _document.AddHandler(Hyperlink.RequestNavigateEvent, 
                                    new RequestNavigateEventHandler((sender, e) => Process.Start(e.Uri.ToString())));
            }
            _document.EndInit();
        }

        protected override void ReadParagraph(XmlReader reader)
        {
            using (SetCurrent(new Paragraph()))
            {
                base.ReadParagraph(reader);
            }
        }

        protected override void ReadBlockControl(XmlReader reader)
        {
            using (SetCurrent(new BlockUIContainer()))
            {
                _isCurrentUiElement = true;
                _currentLabel = new XElement("label");
                base.ReadBlockControl(reader);
                _isCurrentUiElement = false;
            }
        }

        protected override void ReadInlineControl(XmlReader reader)
        {
            using (SetCurrent(new InlineUIContainer()))
            {
                _isCurrentUiElement = true;
                _currentLabel = new XElement("label");
                base.ReadInlineControl(reader);
                _isCurrentUiElement = false;
            }
        }

        protected override void ReadControlProperties(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace)
                {
                    switch (reader.LocalName)
                    {
                        case AliasElement:
                            _currentLabel.SetAttributeValue("alias", GetValueAttribute(reader));
                            break;
                        case TagElement:
                            _currentLabel.SetAttributeValue("tag", GetValueAttribute(reader));
                            break;
                    }
                }
            }
            
            _current.Name = _currentLabel.Attribute("tag")?.Value ?? "";
            string[] names = _current.Name.Split(new [] { '_' }, 2);
            string ltype = names[0];
            _currentLabel.Name = ltype;
            
            switch (ltype)
            {
                case "text":
                case "list":
                case "table":
                    labels.Root?.Add(_currentLabel);
                    break;
                case "col":
                    labels.Element("table_" + names[1])?.Add(_currentLabel);
                    break;
                case "item":
                    labels.Element("list_" + names[1])?.Add(_currentLabel);
                    break;
            }
        }

        protected override void ReadTable(XmlReader reader)
        {
            // TODO: Read table properties
            using (SetCurrent(new Table()))
            {
                using (SetCurrent(new TableRowGroup()))
                {
                    base.ReadTable(reader);
                }
            }
        }

        protected override void ReadTableRow(XmlReader reader)
        {
            // TODO: Read row properties
            using (SetCurrent(new TableRow()))
            {
                base.ReadTableRow(reader);
            }
        }

        protected override void ReadTableCell(XmlReader reader)
        {
            // TODO: Read cell properties
            using (SetCurrent(new TableCell()))
            {
                TableCell cell = (TableCell) _current;
                cell.BorderBrush = new SolidColorBrush(Color.FromRgb(0,0,0));
                cell.BorderThickness = new Thickness(1);
                base.ReadTableCell(reader);
            }
        }

        protected override void ReadParagraphProperties(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace)
                {
                    Paragraph paragraph = (Paragraph)_current;
                    
                    switch (reader.LocalName)
                    {
                        case AlignmentElement:
                            TextAlignment? textAlignment = ConvertTextAlignment(GetValueAttribute(reader));
                            if (textAlignment.HasValue)
                            {
                                paragraph.TextAlignment = textAlignment.Value;
                            }
                            break;
                        case PageBreakBeforeElement:
                            paragraph.BreakPageBefore = GetOnOffValueAttribute(reader);
                            break;
                        case SpacingElement:
                            paragraph.Margin = GetSpacing(reader, paragraph.Margin);
                            break;
                        case IndentationElement:
                            SetParagraphIndent(reader, paragraph);
                            break;
                        case ShadingElement:
                            Brush background = GetShading(reader);
                            if (background != null)
                            {
                                paragraph.Background = background;
                            }
                            break;
                    }
                }
            }
        }

        protected override void ReadHyperlink(XmlReader reader)
        {
            string id = reader[IdAttribute, RelationshipsNamespace];
            if (!string.IsNullOrEmpty(id))
            {
                PackageRelationship relationship = MainDocumentPart.GetRelationship(id);
                if (relationship.TargetMode == TargetMode.External)
                {
                    _hasAnyHyperlink = true;

                    Hyperlink hyperlink = new Hyperlink { NavigateUri = relationship.TargetUri };

                    using (SetCurrent(hyperlink))
                    {
                        base.ReadHyperlink(reader);
                    }
                    return;
                }
            }

            base.ReadHyperlink(reader);
        }

        protected override void ReadRun(XmlReader reader)
        {
            using (SetCurrent(new Span()))
            {
                base.ReadRun(reader);
            }
        }

        protected override void ReadRunProperties(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace)
                {
                    Inline inline = (Inline)_current;
                    switch (reader.LocalName)
                    {
                        case BoldElement:
                            inline.FontWeight = GetOnOffValueAttribute(reader) ? FontWeights.Bold : FontWeights.Normal;
                            break;
                        case ItalicElement:
                            inline.FontStyle = GetOnOffValueAttribute(reader) ? FontStyles.Italic : FontStyles.Normal;
                            break;
                        case UnderlineElement:
                            TextDecorationCollection underlineTextDecorations = GetUnderlineTextDecorations(reader, inline);
                            if (underlineTextDecorations != null)
                            {
                                inline.TextDecorations.Add(underlineTextDecorations);
                            }
                            break;
                        case StrikeElement:
                            if (GetOnOffValueAttribute(reader))
                            {
                                inline.TextDecorations.Add(TextDecorations.Strikethrough);
                            }
                            break;
                        case DoubleStrikeElement:
                            if (GetOnOffValueAttribute(reader))
                            {
                                inline.TextDecorations.Add(new TextDecoration { Location = TextDecorationLocation.Strikethrough, 
                                                                                PenOffset = _current.FontSize * 0.015 });
                                inline.TextDecorations.Add(new TextDecoration { Location = TextDecorationLocation.Strikethrough, 
                                                                                PenOffset = _current.FontSize * -0.015 });
                            }
                            break;
                        case VerticalAlignmentElement:
                            BaselineAlignment? baselineAlignment = GetBaselineAlignment(GetValueAttribute(reader));
                            if (baselineAlignment.HasValue)
                            {
                                inline.BaselineAlignment = baselineAlignment.Value;
                                if (baselineAlignment.Value == BaselineAlignment.Subscript ||
                                    baselineAlignment.Value == BaselineAlignment.Superscript)
                                {
                                    //MS Word 2002 size: 65% http://en.wikipedia.org/wiki/Subscript_and_superscript}}
                                    inline.FontSize *= 0.65; 
                                }
                            }
                            break;
                        case ColorElement:
                            Color? color = GetColor(GetValueAttribute(reader));
                            if (color.HasValue)
                            {
                                inline.Foreground = new SolidColorBrush(color.Value);
                            }
                            break;
                        case HighlightElement:
                            Color? highlight = GetHighlightColor(GetValueAttribute(reader));
                            if (highlight.HasValue)
                            {
                                inline.Background = new SolidColorBrush(highlight.Value);
                            }
                            break;
                        case FontElement:
                            string fontFamily = reader[AsciiFontFamily, WordprocessingMLNamespace];
                            if (!string.IsNullOrEmpty(fontFamily))
                            {
                                inline.FontFamily =
                                    (FontFamily) new FontFamilyConverter().ConvertFromString(fontFamily);
                            }
                            break;
                        case FontSizeElement:
                            string fontSize = reader[ValueAttribute, WordprocessingMLNamespace];
                            if (!string.IsNullOrEmpty(fontSize))
                            {
                                // Attribute Value / 2 = Points
                                // Points * (96 / 72) = Pixels
                                inline.FontSize = uint.Parse(fontSize) * 2.0 / 3.0;
                            }
                            break;
                        case RightToLeftTextElement:
                            inline.FlowDirection = (GetOnOffValueAttribute(reader)) ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
                            break;
                    }
                }
            }
        }

        protected override void ReadBreak(XmlReader reader)
        {
            AddChild(new LineBreak());
        }

        protected override void ReadTabCharacter(XmlReader reader)
        {
            AddChild(new Run("\t"));
        }

        protected override void ReadText(XmlReader reader)
        {
            AddChild(new Run(reader.ReadString()));
        }

        private void AddChild(TextElement textElement)
        {
            if (!_isCurrentUiElement)
            {
                ((IAddChild) _current ?? _document).AddChild(textElement);
            }
        }

        private static bool GetOnOffValueAttribute(XmlReader reader)
        {
            string value = GetValueAttribute(reader);

            switch (value)
            {
                case null:
                case "1":
                case "on":
                case "true":
                    return true;
                default:
                    return false;
            }
        }

        private static string GetValueAttribute(XmlReader reader)
        {
            return reader[ValueAttribute, WordprocessingMLNamespace];
        }

        private static Color? GetColor(string colorString)
        {
            if (string.IsNullOrEmpty(colorString) || colorString == "auto")
            {
                return null;
            }

            return (Color)ColorConverter.ConvertFromString('#' + colorString);
        }

        private static Color? GetHighlightColor(string highlightString)
        {
            if (string.IsNullOrEmpty(highlightString) || highlightString == "auto")
            {
                return null;
            }

            return (Color)ColorConverter.ConvertFromString(highlightString);
        }

        private static BaselineAlignment? GetBaselineAlignment(string verticalAlignmentString)
        {
            switch (verticalAlignmentString)
            {
                case "baseline":
                    return BaselineAlignment.Baseline;
                case "subscript":
                    return BaselineAlignment.Subscript;
                case "superscript":
                    return BaselineAlignment.Superscript;
                default:
                    return null;
            }
        }

        private static double? ConvertTwipsToPixels(string twips)
        {
            if (string.IsNullOrEmpty(twips))
            {
                return null;
            }
            
            return ConvertTwipsToPixels(double.Parse(twips, CultureInfo.InvariantCulture));
        }

        private static double ConvertTwipsToPixels(double twips)
        {
            return 96d / (72 * 20) * twips;
        }

        private static TextAlignment? ConvertTextAlignment(string value)
        {
            switch (value)
            {
                case "both":
                    return TextAlignment.Justify;
                case "left":
                    return TextAlignment.Left;
                case "right":
                    return TextAlignment.Right;
                case "center":
                    return TextAlignment.Center;
                default:
                    return null;
            }
        }

        private static Thickness GetSpacing(XmlReader reader, Thickness margin)
        {
            double? after = ConvertTwipsToPixels(reader[SpacingAfterAttribute, WordprocessingMLNamespace]);
            if (after.HasValue)
            {
                margin.Bottom = after.Value;
            }

            double? before = ConvertTwipsToPixels(reader[SpacingBeforeAttribute, WordprocessingMLNamespace]);
            if (before.HasValue)
            {
                margin.Top = before.Value;
            }

            return margin;
        }

        private static void SetParagraphIndent(XmlReader reader, Paragraph paragraph)
        {
            Thickness margin = paragraph.Margin;

            double? left = ConvertTwipsToPixels(reader[LeftIndentationAttribute, WordprocessingMLNamespace]);
            if (left.HasValue)
            {
                margin.Left = left.Value;
            }

            double? right = ConvertTwipsToPixels(reader[RightIndentationAttribute, WordprocessingMLNamespace]);
            if (right.HasValue)
            {
                margin.Right = right.Value;
            }

            paragraph.Margin = margin;

            double? firstLine = ConvertTwipsToPixels(reader[FirstLineIndentationAttribute, WordprocessingMLNamespace]);
            if (firstLine.HasValue)
            {
                paragraph.TextIndent = firstLine.Value;
            }

            double? hanging = ConvertTwipsToPixels(reader[HangingIndentationAttribute, WordprocessingMLNamespace]);
            if (hanging.HasValue)
            {
                paragraph.TextIndent -= hanging.Value;
            }
        }

        private static Brush GetShading(XmlReader reader)
        {
            Color? color = GetColor(reader[FillAttribute, WordprocessingMLNamespace]);
            return color.HasValue ? new SolidColorBrush(color.Value) : null;
        }

        private static TextDecorationCollection GetUnderlineTextDecorations(XmlReader reader, Inline inline)
        {
            TextDecoration textDecoration;
            
            Color? color = GetColor(reader[ColorAttribute, WordprocessingMLNamespace]);
            Brush brush = color.HasValue ? new SolidColorBrush(color.Value) : inline.Foreground;

            TextDecorationCollection textDecorations = new TextDecorationCollection {
                (textDecoration = new TextDecoration { Location = TextDecorationLocation.Underline,
                                                       Pen = new Pen { Brush = brush } }) };

            switch (GetValueAttribute(reader))
            {
                case "single":
                    break;
                case "double":
                    textDecoration.PenOffset = inline.FontSize * 0.05;
                    textDecoration = textDecoration.Clone();
                    textDecoration.PenOffset = inline.FontSize * -0.05;
                    textDecorations.Add(textDecoration);
                    break;
                case "dotted":
                    textDecoration.Pen.DashStyle = DashStyles.Dot;
                    break;
                case "dash":
                    textDecoration.Pen.DashStyle = DashStyles.Dash;
                    break;
                case "dotDash":
                    textDecoration.Pen.DashStyle = DashStyles.DashDot;
                    break;
                case "dotDotDash":
                    textDecoration.Pen.DashStyle = DashStyles.DashDotDot;
                    break;
                case "none":
                    // fallthrough
                default:
                    // If underline type is none or unsupported then it will be ignored.
                    return null;
            }

            return textDecorations;
        }

        private IDisposable SetCurrent(TextElement currentElement)
        {
            return new CurrentHandle(this, currentElement);
        }

        private struct CurrentHandle : IDisposable
        {
            private readonly DocxToFlowDocumentConverter _converter;
            private readonly TextElement _previous;

            public CurrentHandle(DocxToFlowDocumentConverter converter, TextElement current)
            {
                _converter = converter;
                _converter.AddChild(current);
                _previous = _converter._current;
                _converter._current = current;
            }

            public void Dispose()
            {
                _converter._current = _previous;
            }
        }
    }
}