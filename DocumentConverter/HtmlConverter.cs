using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Resumelink.Converter
{
    public class HtmlConverter
    {

        public static void ConvertDocToHtml(string path, WdSaveFormat format)
        {

            DirectoryInfo dirInfo = new DirectoryInfo(path);
            FileInfo[] wordFiles = dirInfo.GetFiles("*.doc");
            if (wordFiles.Length == 0)
            {
                return;
            }

            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            try
            {
                word.Visible = false;
                word.ScreenUpdating = false;
                foreach (FileInfo wordFile in wordFiles)
                {
                    Object filename = (Object)wordFile.FullName;
                    Document doc = word.Documents.Open(ref filename, ref oMissing,
                                                       ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                       ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                       ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    try
                    {
                        doc.Activate();
                        object outputFileName = wordFile.FullName.Replace(".doc", ".html");
                        object fileFormat = format;
                        doc.SaveAs(ref outputFileName,
                                   ref fileFormat, ref oMissing, ref oMissing,
                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    }
                    finally
                    {
                        object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                        ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                        doc = null;
                    }
                }

            }
            finally
            {
                ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                word = null;
            }
        }
    }
}
