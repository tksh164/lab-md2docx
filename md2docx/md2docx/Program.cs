using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace md2docx
{
    class Program
    {
        private const int TabWidth = 4;

        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("command <Base docx Path> <Markdown File Path> <Output Directory Path>");
                return;
            }

            var baseDocxFilePath = args[0];
            var markdownFilePath = args[1];
            var outputFolderPath = args[2];

            var outputDocxFilePath = Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(markdownFilePath) + ".docx");
            var md2docx = new MD2Docx();
            md2docx.Convert(markdownFilePath, baseDocxFilePath, outputDocxFilePath);
        }
    }
}
