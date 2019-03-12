using System;
using System.IO;
using System.IO.Packaging;
using System.Xml;

namespace DocxToWpf
{
    class DocxReader : IDisposable
    {
        /// <summary> Наименования xml тегов из которых состоит docx файл. </summary>
        protected const string
            MainDocumentRelationshipType =
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",

            // XML namespaces
            WordprocessingMLNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

            // Miscellaneous elements
            DocumentElement = "document",
            BodyElement = "body",

            // Block-Level elements
            ParagraphElement = "p",
            TableElement = "tbl",
            ControlElement = "sdt",

            // Inline-Level elements
            SimpleFieldElement = "fldSimple",
            HyperlinkElement = "hyperlink",
            RunElement = "r",

            // Run content elements
            BreakElement = "br",
            TabCharacterElement = "tab",
            TextElement = "t",

            // Table elements
            TableRowElement = "tr",
            TableCellElement = "tc",

            //Control elements
            ControlContentElement = "sdtContent",

            // Properties elements
            ParagraphPropertiesElement = "pPr",
            RunPropertiesElement = "rPr",
            ControlPropertiesElement = "sdtPr";
        // Note: new members should also be added to nameTable in CreateNameTable method.

        protected virtual XmlNameTable CreateNameTable()
        {
            NameTable nameTable = new NameTable();

            nameTable.Add(WordprocessingMLNamespace);
            nameTable.Add(RelationshipsNamespace);
            nameTable.Add(DocumentElement);
            nameTable.Add(BodyElement);
            nameTable.Add(ParagraphElement);
            nameTable.Add(TableElement);
            nameTable.Add(ParagraphPropertiesElement);
            nameTable.Add(SimpleFieldElement);
            nameTable.Add(HyperlinkElement);
            nameTable.Add(RunElement);
            nameTable.Add(BreakElement);
            nameTable.Add(TabCharacterElement);
            nameTable.Add(TextElement);
            nameTable.Add(RunPropertiesElement);
            nameTable.Add(TableRowElement);
            nameTable.Add(TableCellElement);

            nameTable.Add(ControlElement);
            nameTable.Add(ControlContentElement);
            nameTable.Add(ControlPropertiesElement);

            return nameTable;
        }

        private readonly Package _package;
        private readonly PackagePart _mainDocumentPart;

        protected PackagePart MainDocumentPart => _mainDocumentPart;

        public DocxReader(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            _package = Package.Open(stream, FileMode.Open, FileAccess.Read);

            foreach (PackageRelationship relationship in _package.GetRelationshipsByType(MainDocumentRelationshipType))
            {
                _mainDocumentPart = _package.GetPart(PackUriHelper.CreatePartUri(relationship.TargetUri));
                break;
            }
        }

        public void Read()
        {
            using (Stream mainDocumentStream = _mainDocumentPart.GetStream(FileMode.Open, FileAccess.Read))
            {
                using (XmlReader reader = XmlReader.Create(mainDocumentStream, new XmlReaderSettings {
                                                                                    NameTable = CreateNameTable(),
                                                                                    IgnoreComments = true,
                                                                                    IgnoreProcessingInstructions = true,
                                                                                    IgnoreWhitespace = true}))
                {
                    ReadMainDocument(reader);
                }
            }
        }

        private static void ReadXmlSubtree(XmlReader reader, Action<XmlReader> action)
        {
            if (action == null) return;
            
            using (XmlReader subtreeReader = reader.ReadSubtree())
            {
                // Position on the first node.
                subtreeReader.Read();
                action(subtreeReader);
            }
        }

        private void ReadMainDocument(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == DocumentElement)
                {
                    ReadXmlSubtree(reader, ReadDocument);
                    break;
                }
            }
        }

        protected virtual void ReadDocument(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == BodyElement)
                {
                    ReadXmlSubtree(reader, ReadBody);
                    break;
                }
            }
        }

        private void ReadBody(XmlReader reader)
        {
            while (reader.Read())
            {
                ReadBlockLevelElement(reader);
            }
        }

        private void ReadBlockLevelElement(XmlReader reader)
        {
            if (reader.NodeType != XmlNodeType.Element) return;
            
            Action<XmlReader> action = null;
            if (reader.NamespaceURI == WordprocessingMLNamespace)
            {
                switch (reader.LocalName)
                {
                    case ParagraphElement:
                        action = ReadParagraph;
                        break;
                    case TableElement:
                        action = ReadTable;
                        break;
                    case ControlElement:
                        action = ReadBlockControl;
                        break;
                }
            }
            
            ReadXmlSubtree(reader, action);
        }

        protected virtual void ReadParagraph(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == ParagraphPropertiesElement)
                {
                    ReadXmlSubtree(reader, ReadParagraphProperties);
                }
                else
                {
                    ReadInlineLevelElement(reader);
                }
            }
        }

        protected virtual void ReadBlockControl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == ControlPropertiesElement)
                {
                    ReadXmlSubtree(reader, ReadControlProperties);
                }
                else if (reader.NodeType == XmlNodeType.Element 
                         && reader.NamespaceURI == WordprocessingMLNamespace 
                         && reader.LocalName == ControlContentElement)
                {
                    ReadXmlSubtree(reader, ReadBlockControlContent);
                }
                //this.ReadControlContent(reader);
            }
        }

        protected virtual void ReadInlineControl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == ControlPropertiesElement)
                {
                    ReadXmlSubtree(reader, ReadControlProperties);
                }
                else if (reader.NodeType == XmlNodeType.Element
                         && reader.NamespaceURI == WordprocessingMLNamespace 
                         && reader.LocalName == ControlContentElement)
                {
                    ReadXmlSubtree(reader, this.ReadInlineControlContent);
                }

                //this.ReadControlContent(reader);
            }
        }

        protected virtual void ReadBlockControlContent(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    ReadBlockLevelElement(reader);
                }
            }
        }

        protected virtual void ReadInlineControlContent(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    ReadInlineLevelElement(reader);
                }
            }
        }

        protected virtual void ReadControlProperties(XmlReader reader)
        {
            
        }

        protected virtual void ReadParagraphProperties(XmlReader reader)
        {
            
        }

        private void ReadInlineLevelElement(XmlReader reader)
        {
            if (reader.NodeType != XmlNodeType.Element) return;
            Action<XmlReader> action = null;

            if (reader.NamespaceURI == WordprocessingMLNamespace)
                switch (reader.LocalName)
                {
                    case SimpleFieldElement:
                        action = ReadSimpleField;
                        break;
                    case HyperlinkElement:
                        action = ReadHyperlink;
                        break;
                    case RunElement:
                        action = ReadRun;
                        break;
                    case ControlElement:
                        action = ReadInlineControl;
                        break;
                }

            ReadXmlSubtree(reader, action);
        }

        private void ReadSimpleField(XmlReader reader)
        {
            while (reader.Read())
            {
                ReadInlineLevelElement(reader);
            }
        }

        protected virtual void ReadHyperlink(XmlReader reader)
        {
            while (reader.Read())
            {
                ReadInlineLevelElement(reader);
            }
        }

        protected virtual void ReadRun(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == RunPropertiesElement)
                {
                    ReadXmlSubtree(reader, ReadRunProperties);
                }
                else
                {
                    ReadRunContentElement(reader);
                }
            }
        }

        protected virtual void ReadRunProperties(XmlReader reader)
        {
            
        }

        private void ReadRunContentElement(XmlReader reader)
        {
            if (reader.NodeType != XmlNodeType.Element) return;
            Action<XmlReader> action = null;

            if (reader.NamespaceURI == WordprocessingMLNamespace)
            {
                switch (reader.LocalName)
                {
                    case BreakElement:
                        action = ReadBreak;
                        break;
                    case TabCharacterElement:
                        action = ReadTabCharacter;
                        break;
                    case TextElement:
                        action = ReadText;
                        break;
                }
            }

            ReadXmlSubtree(reader, action);
        }

        protected virtual void ReadBreak(XmlReader reader)
        {
            
        }

        protected virtual void ReadTabCharacter(XmlReader reader)
        {
            
        }

        protected virtual void ReadText(XmlReader reader)
        {
            
        }

        protected virtual void ReadTable(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == TableRowElement)
                {
                    ReadXmlSubtree(reader, ReadTableRow);
                }
            }
        }

        protected virtual void ReadTableRow(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element 
                    && reader.NamespaceURI == WordprocessingMLNamespace 
                    && reader.LocalName == TableCellElement)
                {
                    ReadXmlSubtree(reader, ReadTableCell);
                }
            }
        }

        protected virtual void ReadTableCell(XmlReader reader)
        {
            while (reader.Read())
            {
                ReadBlockLevelElement(reader);
            }
        }

        public void Dispose()
        {
            _package.Close();
        }
    }
}