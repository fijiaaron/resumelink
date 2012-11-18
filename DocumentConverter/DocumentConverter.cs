using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Resumelink.Converter.MsWord.Parameters;
namespace Resumelink.Converter
{
    public enum FileFormat { DOC, DOCX, PDF, HTML, TXT, RTF };

    public class DocumentConverter
    {
        public Microsoft.Office.Interop.Word._Application MsWord {get; set;}
        public Microsoft.Office.Interop.Word._Document Document {get; set;}
        public Boolean Visible { get; set; }

        public ExportAsFixedFormatParams PdfExportParams { get; set; }
        public DocumentOpenParams DocOpenParams { get; set; }



        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="msword"></param>
        /// <param name="visible"></param>
        public DocumentConverter(Microsoft.Office.Interop.Word._Application msword = null, bool visible=false)
        {
            MsWord = msword;
            Visible = visible;

            PdfExportParams = new ExportAsFixedFormatParams();
            DocOpenParams = new DocumentOpenParams();
        }
        


        /// <summary>
        /// Start Microsoft Word
        /// </summary>
        /// <param name="visible"></param>
        public _Application StartMsWord(bool visible=false)
        {
            if (MsWord == null) {
                MsWord = new Microsoft.Office.Interop.Word.Application();
            }

            if (visible) {
                MsWord.Visible = true;
                MsWord.ScreenUpdating = true;
            }

            return MsWord;
        }



        /// <summary>
        /// Stop Microsoft Word
        /// </summary>
        public void StopMsWord() 
        {
            if (MsWord != null)
            {
                MsWord.Quit();
                MsWord = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }


        
        /// <summary>
        /// Open a Microsoft Word document
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public _Document OpenWordDocument(String filename)
        {
            Document = MsWord.Documents.Open(filename);
            return Document;
        }
        
        

        /// <summary>
        /// Close the Microsoft Word Document
        /// </summary>
        public void CloseWordDocument() 
        {
            if (Document != null)
            {
                Document.Close();
                Document = null;
            }
        }



        /// <summary>
        /// Save a Microsoft Word document as PDF
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="format"></param>
        public void SaveWordDocumentAs(String filename, FileFormat format)
        {

            switch (format) {
                case FileFormat.PDF:
                    SaveWordDocumentAsPdf(filename);
                    break;
                case FileFormat.HTML:
                    break;
                case FileFormat.TXT:
                    break;
                case FileFormat.DOC:
                    break;
                case FileFormat.DOCX:
                    break;
                case FileFormat.RTF:
                    break;
            }
        }

        
        
        /// <summary>
        /// Save a Microsoft Word document as PDF
        /// </summary>
        /// <param name="filename"></param>
        public void SaveWordDocumentAsPdf(String filename)
        {
            Console.WriteLine("Converting document to pdf: " + filename);

            if (!System.IO.File.Exists(filename))
            {
                throw new Exception("File not found: " + filename);
            }

            StartMsWord();
            Document = OpenWordDocument(filename);
            
            if (Document == null)
            {
                throw new Exception("Microsoft word document is null");
            }

            var outputFileName = ChangeFileExtension(filename, "pdf");

            try
            {
                SaveWordDocumentAsPdf(Document, outputFileName);
                Console.WriteLine("Document converted to pdf: " + outputFileName);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                CloseWordDocument();
                StopMsWord();
            }
        }



        /// <summary>
        /// Save a Microsoft Word Document as PDF
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="outputFileName"></param>
        public void SaveWordDocumentAsPdf(_Document doc, String outputFileName)
        {
            doc.ExportAsFixedFormat(outputFileName,
                PdfExportParams.ExportFormat, PdfExportParams.OpenAfterExport, PdfExportParams.OptimizeFor,
                PdfExportParams.ExportRange, PdfExportParams.StartPage, PdfExportParams.EndPage,
                PdfExportParams.ExportItem, PdfExportParams.IncludeDocProps, PdfExportParams.KeepIRM,
                PdfExportParams.CreateBookmarks, PdfExportParams.DocStructureTags, PdfExportParams.BitmapMissingFonts,
                PdfExportParams.UseISO19005_1, PdfExportParams.FixedFormatExtClassPtr
            );
        }



        /// <summary>
        /// Change a document file extension (e.g. file.doc -> file.pdf)
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="newExtension"></param>
        /// <returns></returns>
        public String ChangeFileExtension(String filename, String newExtension)
        {
            String oldExtension = System.IO.Path.GetExtension(filename);
            int extensionPosition = filename.LastIndexOf(oldExtension);
            String filenameWithoutPrefix = filename.Substring(0, extensionPosition);
            String newFilename = filenameWithoutPrefix + "." + newExtension;

            return newFilename;
        }

          
    }
}
