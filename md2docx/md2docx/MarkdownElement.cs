using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace md2docx
{
    internal enum MarkdownElementType
    {
        Header,
        ListItem,
        NumberingListItem,
        Image,
        CodeBlock,
        Quotation,
        Paragraph
    }

    public class UnknownMarkdownElementType : Exception
    {
        public UnknownMarkdownElementType() : base()
        {
        }
    }

    internal abstract class MarkdownElementBase
    {
        public MarkdownElementType ElementType { get; protected set; }
    }

    internal class MarkdownHeaderElement : MarkdownElementBase
    {
        public string HeaderText { get; protected set; }
        public int HeaderLevel { get; protected set; }

        public MarkdownHeaderElement(string headerText, int headerLevel)
        {
            ElementType = MarkdownElementType.Header;
            HeaderText = headerText;
            HeaderLevel = headerLevel;
        }
    }

    internal class MarkdownListItemElement : MarkdownElementBase
    {
        public string ListItemText { get; protected set; }

        public MarkdownListItemElement(string listItemText)
        {
            ElementType = MarkdownElementType.ListItem;
            ListItemText = listItemText;
        }
    }

    internal class MarkdownNumberingListItemElement : MarkdownElementBase
    {
        public string ListItemText { get; protected set; }

        public MarkdownNumberingListItemElement(string listItemText)
        {
            ElementType = MarkdownElementType.NumberingListItem;
            ListItemText = listItemText;
        }
    }

    internal class MarkdownImageElement : MarkdownElementBase
    {
        public string ImageFilePath { get; protected set; }
        public string AltText { get; protected set; }

        public MarkdownImageElement(string imageFilePath, string altText)
        {
            ElementType = MarkdownElementType.Image;
            ImageFilePath = imageFilePath;
            AltText = altText;
        }
    }

    internal class MarkdownCodeBlockElement : MarkdownElementBase
    {
        public string CodeBlockText { get; protected set; }

        public MarkdownCodeBlockElement(string codeBlockText)
        {
            ElementType = MarkdownElementType.CodeBlock;
            CodeBlockText = codeBlockText;
        }
    }

    internal class MarkdownQuotationElement : MarkdownElementBase
    {
        public MarkdownQuotationElement(string imageFilePath, string altText)
        {
            ElementType = MarkdownElementType.Quotation;
        }
    }

    internal class MarkdownParagraphElement : MarkdownElementBase
    {
        public string ParagraphText { get; protected set; }

        public MarkdownParagraphElement(string paragraphText)
        {
            ElementType = MarkdownElementType.Paragraph;
            ParagraphText = paragraphText;
        }
    }
}
