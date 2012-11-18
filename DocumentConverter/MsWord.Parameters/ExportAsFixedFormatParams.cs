using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Resumelink.Converter.MsWord.Parameters
{
    public class ExportAsFixedFormatParams
    {
        public string OutputFileName { get; set; }

        public Microsoft.Office.Interop.Word.WdExportFormat ExportFormat = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
        public bool OpenAfterExport = false;
        public Microsoft.Office.Interop.Word.WdExportOptimizeFor OptimizeFor = Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
        public Microsoft.Office.Interop.Word.WdExportRange ExportRange = Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument;
        public int StartPage = 0;
        public int EndPage = 0;
        public Microsoft.Office.Interop.Word.WdExportItem ExportItem = Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent;
        public bool IncludeDocProps = true;
        public bool KeepIRM = true;
        public Microsoft.Office.Interop.Word.WdExportCreateBookmarks CreateBookmarks = Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
        public bool DocStructureTags = true;
        public bool BitmapMissingFonts = true;
        public bool UseISO19005_1 = false;
        public object FixedFormatExtClassPtr = Type.Missing;
    }
}
