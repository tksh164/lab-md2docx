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

namespace DocumentBuilder
{
    public class WordDocumentClient : IDisposable
    {
        private WordprocessingDocument wordDoc;

        public WordDocumentClient(string baseDocxFilePath, string outputDocxFilePath)
        {
            // Copy a base docx file as a putput docx file. 
            File.Copy(baseDocxFilePath, outputDocxFilePath, true);

            // Open a docx file.
            var openSettings = new OpenSettings()
            {
                AutoSave = false,
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.NoProcess, FileFormatVersions.Office2013),
                MaxCharactersInPart = 0,
            };
            wordDoc = WordprocessingDocument.Open(outputDocxFilePath, true, openSettings);

            // Remove all existing paragraphs.
            RemoveAllExistingParagraphs();
        }

        public void Dispose()
        {
            wordDoc.Dispose();
        }

        private void RemoveAllExistingParagraphs()
        {
            wordDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
        }

        public void Save()
        {
            wordDoc.MainDocumentPart.Document.Save();
            wordDoc.Save();
        }

        public void Append(Paragraph paragraph)
        {
            wordDoc.MainDocumentPart.Document.Body.AppendChild(paragraph);
        }

        public void Append(Paragraph[] paragraphs)
        {
            foreach (var paragraph in paragraphs)
            {
                wordDoc.MainDocumentPart.Document.Body.AppendChild(paragraph);
            }
        }

        public (string relationshipId, long widthInEmus, long heightInEmus) AddImagePart(string imageFilePath)
        {
            // Add a image part and load the image data from the file.
            MainDocumentPart mainDocPart = wordDoc.MainDocumentPart;
            var imagePart = mainDocPart.AddImagePart(DetectImageType(imageFilePath));
            using (var stream = new FileStream(imageFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                imagePart.FeedData(stream);
            }

            // Get image size in EMUs.
            var sectionProperties = mainDocPart.Document.Body.GetFirstChild<SectionProperties>();
            (var iamgeWidthInEmus, var imageHeightInEmus) = CalcuateImageSizeInEmus(imagePart, sectionProperties);

            // Relationship ID
            var imageRelationshipId = mainDocPart.GetIdOfPart(imagePart);

            return (imageRelationshipId, iamgeWidthInEmus, imageHeightInEmus);
        }

        private static ImagePartType DetectImageType(string imageFilePath)
        {
            var extension = Path.GetExtension(imageFilePath);
            if (extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
            {
                return ImagePartType.Png;
            }
            else if (extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase))
            {
                return ImagePartType.Jpeg;
            }

            throw new NotImplementedException();
        }

        private static (long width, long height) CalcuateImageSizeInEmus(ImagePart imagePart, SectionProperties sectionProperties)
        {
            // Get a BitmapImage from the image part.
            var bitmapImage = new BitmapImage();
            using (var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
            {
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = stream;
                bitmapImage.EndInit();
            }

            // Calculate width and height in EMUs.
            // https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size
            // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            const int EMUsPerInch = 914400;
            var widthInEmus = (long)(bitmapImage.PixelWidth / bitmapImage.DpiX * EMUsPerInch);
            var heightInEmus = (long)(bitmapImage.PixelHeight / bitmapImage.DpiY * EMUsPerInch);

            // Calculate content area width in EMUs.
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            var pageMargin = sectionProperties.GetFirstChild<PageMargin>();

            const double DXAsPerInch = 1440;  // 1 inch = 72 points = 72 x 20 DXAs
            var maxContentAreaWidthInDxa = pageSize.Width - pageMargin.Left - pageMargin.Right;
            var maxContentAreaWidthInEmus = (long)(maxContentAreaWidthInDxa / DXAsPerInch * EMUsPerInch);

            // Adjust the image size to content area width.
            if (widthInEmus > maxContentAreaWidthInEmus)
            {
                var ratio = (double)heightInEmus / widthInEmus;
                widthInEmus = maxContentAreaWidthInEmus;
                heightInEmus = (long)(widthInEmus * ratio);
            }

            return (width: widthInEmus, height: heightInEmus);
        }
    }
}
