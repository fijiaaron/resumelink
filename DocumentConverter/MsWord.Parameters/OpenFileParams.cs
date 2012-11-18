using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Word;

namespace Resumelink.Converter.MsWord.Parameters
{
    public class DocumentOpenParams
    {
        private static Object Missing = Type.Missing;

        public String FileName { get; set; }

        public Boolean ConfirmConversions = false;
        public Boolean ReadOnly = true;
        public Boolean AddToRecentFiles = false;
        public Object PasswordDocument = Missing;
        public Object PasswordTemplate = Missing;
        public Boolean Revert = false;
        public Object WritePasswordDocument = Missing;
        public Object WritePasswordTemplate = Missing;
        public WdOpenFormat Format = WdOpenFormat.wdOpenFormatDocument;
        public Object Encoding = Missing;
        public Boolean Visible = false;
        public Boolean OpenAndRepair = true;
        public WdDocumentDirection DocumentDirection = WdDocumentDirection.wdLeftToRight;
        public Boolean NoEncodingDialog = true;
        public Object XMLTransform = Missing;
    }
}
