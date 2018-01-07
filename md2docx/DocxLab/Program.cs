using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using DocumentBuilder;

namespace DocxLab
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("command <Markdown File Path> <Base docx Path>");
                return;
            }

            var markdownFilePath = args[0];
            var docxFilePath = args[1];

            // Create a output file by copying of base docx file.
            var outputFilePath = Path.GetDirectoryName(markdownFilePath) + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(markdownFilePath) + ".docx";

            using (var builder = new WordDocumentBuilder(docxFilePath, outputFilePath))
            {
                builder.AddHeader("Header 1", "Heading1");

                builder.AddParagraph(@"これはサンプル テキストです。これはサンプル テキストです。これはサンプル テキストです。これはサンプル テキストです。これはサンプル テキストです。これはサンプル テキストです。");

                builder.AddListItem("List Item 1", "ListParagraph");

                var imageFilePath = @"D:\Temp\md2docx\image1.png";
                builder.AddImage(imageFilePath, "Figure");


                builder.Save();
            }
        }
    }
}
