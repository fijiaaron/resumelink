using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace Resumelink.Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("using DocumentConverter");
            var converter = new DocumentConverter();

            var filename = @"C:\temp\resume\resume.doc";

            converter.SaveWordDocumentAs(filename, FileFormat.PDF);
            converter.SaveWordDocumentAs(filename, FileFormat.HTML);
            System.Environment.Exit(0);

            Console.WriteLine("converting DOCX to PDF");
            PdfConverter.ConvertDocxToPdf(@"C:\temp\resume.docx");

            Console.WriteLine("converting DOC to HTML");
            HtmlConverter.ConvertDocToHtml(@"C:\temp\docs", WdSaveFormat.wdFormatHTML);
        }
    }
}
