using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Windows.Media.Imaging;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;

namespace md2docx
{
    internal static class WordDocumentElementBuilder
    {
        public static Run[] CreateRun(string text)
        {
            text = ReplaceIconMarker(text);

            List<Run> runElements = new List<Run>();

            var strongTexts = GetStrongTexts(text);

            var textParts = text.Split(new string[] { "**" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var textPart in textParts)
            {
                if (strongTexts.Contains(textPart))
                {
                    runElements.Add(new Run(
                        new RunProperties(
                            new Bold()
                        ),
                        new Text(textPart)
                    ));
                }
                else
                {
                    if (textPart.StartsWith(" "))
                    {
                        runElements.Add(new Run(
                            new Text(" ") { Space = SpaceProcessingModeValues.Preserve }
                        ));
                    }

                    runElements.Add(new Run(
                        new Text(textPart)
                    ));

                    if (textPart.EndsWith(" "))
                    {
                        runElements.Add(new Run(
                            new Text(" ") { Space = SpaceProcessingModeValues.Preserve }
                        ));
                    }
                }
            }

            return runElements.ToArray();
        }

        private static string ReplaceIconMarker(string text)
        {
            //return text.Replace(":bulb:", @"💡")
            //           .Replace(":warning:", @"⚠");
            return text.Replace(":bulb:", string.Empty)
                       .Replace(":warning:", string.Empty)
                       .Trim();
        }

        private static ReadOnlyCollection<string> GetStrongTexts(string text)
        {
            string pattern = @"(\*{2}.+?\*{2})";
            var matches = Regex.Matches(text, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.Compiled);

            List<string> strongText = new List<string>();
            for (int i = 0; i < matches.Count; i++)
            {
                strongText.Add(matches[i].Value.Trim(new char[] { '*' }));
            }

            return new ReadOnlyCollection<string>(strongText);
        }

        public static Paragraph CreateHeader(Run[] runElements, string styleName = null)
        {
            if (styleName == null) styleName = string.Empty;

            var header = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = styleName }
                )
            );

            foreach (var run in runElements)
            {
                header.AppendChild(run);
            }

            return header;
        }

        public static Paragraph CreateListItem(Run[] runElements, string styleName = null)
        {
            if (styleName == null) styleName = string.Empty;

            var listItem = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = styleName },
                    new NumberingProperties(
                        new NumberingLevelReference() { Val = 0 },  // bullet list level
                        new NumberingId() { Val = 1 }               // reference number to numbering.xml
                    ),
                    new Indentation() { LeftChars = 0 }
                )
            );

            foreach (var run in runElements)
            {
                listItem.AppendChild(run);
            }

            return listItem;
        }

        public static Paragraph CreateNumberingListItem(Run[] runElements, string styleName = null)
        {
            if (styleName == null) styleName = string.Empty;

            var listItem = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = styleName },
                    new NumberingProperties(
                        new NumberingLevelReference() { Val = 0 },  // bullet list level
                        new NumberingId() { Val = 2 }               // reference number to numbering.xml
                    ),
                    new Indentation() { LeftChars = 0 }
                )
            );

            foreach (var run in runElements)
            {
                listItem.AppendChild(run);
            }

            return listItem;
        }

        public static Paragraph CreateImage(string imageRelationshipId, long iamgeWidthInEmus, long imageHeightInEmus, string fileName, string styleName = null)
        {
            if (styleName == null) styleName = string.Empty;

            // ID for image.
            var idNum = (uint)(DateTime.UtcNow.Ticks % Int32.MaxValue);

            return new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = styleName }
                ),
                new Run(
                    new Drawing(
                        new WP.Inline(
                            new WP.Extent() { Cx = iamgeWidthInEmus, Cy = imageHeightInEmus },
                            new WP.EffectExtent() { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                            new WP.DocProperties()
                            {
                                Id = idNum,
                                Name = fileName,
                            },
                            new WP.NonVisualGraphicFrameDrawingProperties(
                                new A.GraphicFrameLocks() { NoChangeAspect = true }
                            ),
                            new A.Graphic(
                                new A.GraphicData(
                                    new PIC.Picture(
                                        new PIC.NonVisualPictureProperties(
                                            new PIC.NonVisualDrawingProperties()
                                            {
                                                Id = idNum,
                                                Name = fileName,
                                            },
                                            new PIC.NonVisualPictureDrawingProperties()
                                        ),
                                        new PIC.BlipFill(
                                            new A.Blip() { Embed = imageRelationshipId },
                                            new A.Stretch(
                                                new A.FillRectangle()
                                            )
                                        ),
                                        new PIC.ShapeProperties(
                                            new A.Transform2D(
                                                new A.Offset() { X = 0, Y = 0 },
                                                new A.Extents() { Cx = iamgeWidthInEmus, Cy = imageHeightInEmus }
                                            ),
                                            new A.PresetGeometry(
                                                new A.AdjustValueList()
                                            )
                                            { Preset = A.ShapeTypeValues.Rectangle }
                                        )
                                    )
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                            )
                        )
                        {
                            DistanceFromTop = 0U,
                            DistanceFromBottom = 0U,
                            DistanceFromLeft = 0U,
                            DistanceFromRight = 0U,
                        }
                    )
                )
            );
        }

        public static Paragraph[] CreateCodeBlock(string codeBlockText, string styleName = null)
        {
            if (styleName == null) styleName = string.Empty;

            List<Paragraph> codeBlockParagraphs = new List<Paragraph>();

            var codeBlockLines = codeBlockText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            foreach (var codeBlockLine in codeBlockLines)
            {
                codeBlockParagraphs.Add(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId() { Val = styleName }
                        ),
                        new Run(
                            new Text(codeBlockLine) { Space = SpaceProcessingModeValues.Preserve }
                        )
                    )
                );
            }

            return codeBlockParagraphs.ToArray();
        }

        public static Paragraph CreateQuotation()
        {
            throw new NotImplementedException();
        }

        public static Paragraph CreateParagraph(Run[] runElements, string styleName = null)
        {
            if (styleName == null) styleName = string.Empty;

            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = styleName }
                )
            );

            foreach (var run in runElements)
            {
                paragraph.AppendChild(run);
            }

            return paragraph;
        }
    }
}
