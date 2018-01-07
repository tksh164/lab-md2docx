using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using DocumentBuilder;

namespace md2docx
{
    public class MD2Docx
    {
        private const int TabWidth = 4;

        public MD2Docx()
        {
        }

        public void Convert(string markdownFilePath, string baseDocxFilePath, string outputDocxFilePath)
        {
            // Read the each line of Markdown file then into the queue.
            MarkdownClient mdClient = new MarkdownClient(markdownFilePath);

            using (var docClient = new WordDocumentClient(baseDocxFilePath, outputDocxFilePath))
            {
                while (true)
                {
                    var markdownElement = mdClient.TakeNextElement();

                    if (markdownElement == null)
                    {
                        break; // no more elements.
                    }

                    switch (markdownElement.ElementType)
                    {
                        case MarkdownElementType.Header:
                            var headerElement = (MarkdownHeaderElement)markdownElement;
                            var styleName = string.Format("Heading{0}", headerElement.HeaderLevel);
                            var header = WordDocumentElementBuilder.CreateHeader(WordDocumentElementBuilder.CreateRun(headerElement.HeaderText), styleName);
                            docClient.Append(header);
                            break;

                        case MarkdownElementType.ListItem:
                            var listItemElement = (MarkdownListItemElement)markdownElement;
                            var listItem = WordDocumentElementBuilder.CreateListItem(WordDocumentElementBuilder.CreateRun(listItemElement.ListItemText), "ListParagraph");
                            docClient.Append(listItem);
                            break;

                        case MarkdownElementType.NumberingListItem:
                            var numListItemElement = (MarkdownNumberingListItemElement)markdownElement;
                            var numListItem = WordDocumentElementBuilder.CreateNumberingListItem(WordDocumentElementBuilder.CreateRun(numListItemElement.ListItemText), "ListParagraph");
                            docClient.Append(numListItem);
                            break;

                        case MarkdownElementType.Image:
                            var imageElement = (MarkdownImageElement)markdownElement;

                            // Name for the image.
                            var imageFileName = Path.GetFileName(imageElement.ImageFilePath);

                            // Add image to the word document.
                            (var relationshipId, var widthInEmus, var heightInEmus)  = docClient.AddImagePart(imageElement.ImageFilePath);

                            var image = WordDocumentElementBuilder.CreateImage(relationshipId, widthInEmus, heightInEmus, imageFileName, "Figure");
                            docClient.Append(image);
                            break;

                        case MarkdownElementType.CodeBlock:
                            var codeBlockElement = (MarkdownCodeBlockElement)markdownElement;
                            var codeBlock = WordDocumentElementBuilder.CreateCodeBlock(codeBlockElement.CodeBlockText, "Code");
                            docClient.Append(codeBlock);
                            break;

                        case MarkdownElementType.Quotation:
                            break;

                        case MarkdownElementType.Paragraph:
                            var paragraphElement = (MarkdownParagraphElement)markdownElement;
                            var paragraph = WordDocumentElementBuilder.CreateParagraph(WordDocumentElementBuilder.CreateRun(paragraphElement.ParagraphText));
                            docClient.Append(paragraph);
                            break;

                        default:
                            throw new UnknownMarkdownElementType();
                    }
                }

                docClient.Save();
            }
        }
    }
}
