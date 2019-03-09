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

            MainDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",

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
            var nameTable = new NameTable();

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

        private readonly Package package;
        private readonly PackagePart mainDocumentPart;

        protected PackagePart MainDocumentPart
        {
            get { return this.mainDocumentPart; }
        }

        public DocxReader(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            this.package = Package.Open(stream, FileMode.Open, FileAccess.Read);

            foreach (var relationship in this.package.GetRelationshipsByType(MainDocumentRelationshipType))
            {
                this.mainDocumentPart = package.GetPart(PackUriHelper.CreatePartUri(relationship.TargetUri));
                break;
            }
        }

        public void Read()
        {
            using (var mainDocumentStream = this.mainDocumentPart.GetStream(FileMode.Open, FileAccess.Read))
            using (var reader = XmlReader.Create(mainDocumentStream, new XmlReaderSettings()
            {
                NameTable = this.CreateNameTable(),
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            }))
                this.ReadMainDocument(reader);
        }

        private static void ReadXmlSubtree(XmlReader reader, Action<XmlReader> action)
        {
            if (action != null)
                using (var subtreeReader = reader.ReadSubtree())
                {
                    // Position on the first node.
                    subtreeReader.Read();


                    action(subtreeReader);
                }
        }

        private void ReadMainDocument(XmlReader reader)
        {
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == DocumentElement)
                {
                    ReadXmlSubtree(reader, this.ReadDocument);
                    break;
                }
        }

        protected virtual void ReadDocument(XmlReader reader)
        {
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == BodyElement)
                {
                    ReadXmlSubtree(reader, this.ReadBody);
                    break;
                }
        }

        private void ReadBody(XmlReader reader)
        {
            while (reader.Read())
                this.ReadBlockLevelElement(reader);
        }

        private void ReadBlockLevelElement(XmlReader reader)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                Action<XmlReader> action = null;

                if (reader.NamespaceURI == WordprocessingMLNamespace)
                    switch (reader.LocalName)
                    {
                        case ParagraphElement:
                            action = this.ReadParagraph;
                            break;
                        case TableElement:
                            action = this.ReadTable;
                            break;
                        case ControlElement:
                            action = this.ReadBlockControl;
                            break;
                    }

                ReadXmlSubtree(reader, action);
            }
        }

        protected virtual void ReadParagraph(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ParagraphPropertiesElement)
                    ReadXmlSubtree(reader, this.ReadParagraphProperties);
                else
                    this.ReadInlineLevelElement(reader);
            }
        }

        protected virtual void ReadBlockControl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ControlPropertiesElement)
                    ReadXmlSubtree(reader, this.ReadControlProperties);
                else if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ControlContentElement)
                {
                    ReadXmlSubtree(reader, this.ReadBlockControlContent);
                }
                //this.ReadControlContent(reader);
            }
        }

        protected virtual void ReadInlineControl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ControlPropertiesElement)
                    ReadXmlSubtree(reader, this.ReadControlProperties);
                else if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ControlContentElement)
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

                    this.ReadBlockLevelElement(reader);
                }
            }
        }

        protected virtual void ReadInlineControlContent(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    this.ReadInlineLevelElement(reader);

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
            if (reader.NodeType == XmlNodeType.Element)
            {
                Action<XmlReader> action = null;

                if (reader.NamespaceURI == WordprocessingMLNamespace)
                    switch (reader.LocalName)
                    {
                        case SimpleFieldElement:
                            action = this.ReadSimpleField;
                            break;

                        case HyperlinkElement:
                            action = this.ReadHyperlink;
                            break;

                        case RunElement:
                            action = this.ReadRun;
                            break;
                        case ControlElement:
                            action = this.ReadInlineControl;
                            break;
                    }

                ReadXmlSubtree(reader, action);
            }
        }

        private void ReadSimpleField(XmlReader reader)
        {
            while (reader.Read())
                this.ReadInlineLevelElement(reader);
        }

        protected virtual void ReadHyperlink(XmlReader reader)
        {
            while (reader.Read())
                this.ReadInlineLevelElement(reader);
        }

        protected virtual void ReadRun(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == RunPropertiesElement)
                    ReadXmlSubtree(reader, this.ReadRunProperties);
                else
                    this.ReadRunContentElement(reader);
            }
        }

        protected virtual void ReadRunProperties(XmlReader reader)
        {

        }

        private void ReadRunContentElement(XmlReader reader)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                Action<XmlReader> action = null;

                if (reader.NamespaceURI == WordprocessingMLNamespace)
                    switch (reader.LocalName)
                    {
                        case BreakElement:
                            action = this.ReadBreak;
                            break;

                        case TabCharacterElement:
                            action = this.ReadTabCharacter;
                            break;

                        case TextElement:
                            action = this.ReadText;
                            break;
                    }

                ReadXmlSubtree(reader, action);
            }
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
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == TableRowElement)
                    ReadXmlSubtree(reader, this.ReadTableRow);
        }

        protected virtual void ReadTableRow(XmlReader reader)
        {
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == TableCellElement)
                    ReadXmlSubtree(reader, this.ReadTableCell);
        }

        protected virtual void ReadTableCell(XmlReader reader)
        {
            while (reader.Read())
                this.ReadBlockLevelElement(reader);
        }

        public void Dispose()
        {
            this.package.Close();
        }
    }
}