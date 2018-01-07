using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace md2docx
{
    internal class MarkdownClient
    {
        private const int TabWidth = 4;
        private Queue<string> lines;

        public string MarkdownFilePath { get; private set; }
        public string BasePath { get; private set; }

        public MarkdownClient(string markdownFilePath)
        {
            MarkdownFilePath = markdownFilePath;
            BasePath = Path.GetDirectoryName(MarkdownFilePath);

            // Read the each line of Markdown file then into the queue.
            lines = ReadMarkdownFile(MarkdownFilePath);
        }

        private static Queue<string> ReadMarkdownFile(string filePath)
        {
            var lines = File.ReadAllLines(filePath);

            for (int i = 0; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                {
                    // Replace a line consisting only of whitespaces.
                    lines[i] = string.Empty;
                }
                else
                {
                    // Convert all tabs to spaces in a line.
                    lines[i] = Detab(lines[i]);
                }
            }

            return new Queue<string>(lines);
        }

        private static string Detab(string line)
        {
            return line.Replace("\t", new string(' ', TabWidth));
        }

        public MarkdownElementBase TakeNextElement()
        {
            if (!IsExistRemainingLine())
            {
                return null;  // no more lines.
            }

            string line;
            while (true)
            {
                line = GetNextLine();
                if (!string.IsNullOrWhiteSpace(line))
                {
                    break;
                }
            }

            // Setext style header.
            // e.g. Header 1
            //      ========
            if (IsSetextStyleHeaderLineStarts(PeekNextLine()))
            {
                var nextLine = GetNextLine();  // Consume a line.
                return GetSetextStyleHeaderElement(line, nextLine);
            }

            // Atx style header.
            // e.g. # Header 1
            else if (IsAtxStyleHeaderLine(line))
            {
                return GetAtxStyleHeaderElement(line);
            }

            // List
            else if (IsListItemLine(line))
            {
                return GetListItemElement(line);
            }

            // Numbering List
            else if (IsNumberingListItemLine(line))
            {
                return GetNumberingListItemElement(line);
            }

            // Image
            else if (IsImageLine(line))
            {
                return GetImageElement(line, BasePath);
            }

            // Code block
            else if (IsCodeBlockLineMarker(line))
            {
                var codeBlockText = PackCodeBlockText();
                return GetCodeBlockElement(codeBlockText);
            }

            // Quotation
            // TODO

            // Paragraph
            else
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    return GetParagraphElement(line);
                }
            }

            return null;
        }

        private bool IsExistRemainingLine()
        {
            return lines.Count > 0;
        }

        private string GetNextLine()
        {
            return IsExistRemainingLine() ? lines.Dequeue() : null;
        }

        private string PeekNextLine()
        {
            return IsExistRemainingLine() ? lines.Peek() : null;
        }

        private static bool IsSetextStyleHeaderLineStarts(string nextLine)
        {
            if (nextLine != null && Regex.IsMatch(nextLine, @"^\s*=+\s*$")) return true;
            if (nextLine != null && Regex.IsMatch(nextLine, @"^\s*-+\s*$")) return true;
            return false;
        }

        private MarkdownHeaderElement GetSetextStyleHeaderElement(string line, string nextLine)
        {
            var headerText = line.Trim();
            var headerLevelLine = nextLine.Trim();
            int headerLevel = 0;

            if (headerLevelLine[0] == '=')
            {
                headerLevel = 1;
            }
            else if (headerLevelLine[0] == '-')
            {
                headerLevel = 2;
            }

            return new MarkdownHeaderElement(headerText, headerLevel);
        }

        private static bool IsAtxStyleHeaderLine(string line)
        {
            return line.Trim().StartsWith("#");
        }

        private static MarkdownHeaderElement GetAtxStyleHeaderElement(string line)
        {
            var headerText = line.Trim();
            int headerLevel = 0;

            // Couting '#' characters as header level.
            int cnt;
            for (cnt = 0; cnt < 7; cnt++)
            {
                if (!headerText[cnt].Equals('#')) break;
            }

            headerLevel = cnt;
            headerText = headerText.Trim(new char[] { '#', ' ' });

            return new MarkdownHeaderElement(headerText, headerLevel);
        }

        private static bool IsListItemLine(string line)
        {
            var trimedLine = line.Trim();
            return trimedLine.StartsWith("-") || trimedLine.StartsWith("*") || trimedLine.StartsWith("+");
        }

        private static MarkdownListItemElement GetListItemElement(string line)
        {
            var listItemText = line.Trim().TrimStart(new char[] { ' ', '-', '+', '*' });
            return new MarkdownListItemElement(listItemText);
        }

        private static bool IsNumberingListItemLine(string line)
        {
            var trimedLine = line.Trim();
            return Regex.IsMatch(line, @"^[0-9]+\.\s.*$");
        }

        private static MarkdownNumberingListItemElement GetNumberingListItemElement(string line)
        {
            var match = Regex.Match(line, @"^[0-9]+\.\s(?<text>.*)$", RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var listItemText = match.Groups["text"].Value.Trim();
            return new MarkdownNumberingListItemElement(listItemText);
        }

        private static bool IsImageLine(string line)
        {
            return Regex.IsMatch(line, @"^\s*!\[.*?\]\(.*?\)\s*$");
        }

        private static MarkdownImageElement GetImageElement(string line, string basePath)
        {
            Match match = Regex.Match(line, @"^\s*!\[(?<alt>.*?)\]\((?<path>.*?)\)\s*$", RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.Compiled);

            var altText = match.Groups["alt"].Value;
            var imagePath = match.Groups["path"].Value;

            if (imagePath.StartsWith("./"))
            {
                // TODO: need test
                imagePath = Path.Combine(basePath, imagePath);
            }
            else if (imagePath.StartsWith("/"))
            {
                // TODO: need test
                imagePath = Path.Combine(basePath, imagePath);
            }
            else if (!imagePath.StartsWith("http://") && !imagePath.StartsWith("https://"))
            {
                // TODO: need test
                imagePath = Path.Combine(basePath, imagePath);
            }

            return new MarkdownImageElement(imagePath, altText);
        }

        private static bool IsCodeBlockLineMarker(string line)
        {
            return line.Trim().StartsWith("```");
        }

        private string PackCodeBlockText()
        {
            StringBuilder codeBlock = new StringBuilder();

            while (IsExistRemainingLine())
            {
                var line = GetNextLine();
                if (IsCodeBlockLineMarker(line))
                {
                    break;
                }

                if (codeBlock.Length != 0)
                {
                    codeBlock.AppendLine();
                }
                codeBlock.Append(line);
            }

            return codeBlock.ToString();
        }

        private static MarkdownCodeBlockElement GetCodeBlockElement(string codeBlockText)
        {
            return new MarkdownCodeBlockElement(codeBlockText);
        }

        private static MarkdownQuotationElement GetQuotationElement(string line)
        {
            throw new NotImplementedException();
        }

        private static MarkdownParagraphElement GetParagraphElement(string line)
        {
            return new MarkdownParagraphElement(line.Trim());
        }
    }
}
