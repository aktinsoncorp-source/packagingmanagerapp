using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PackingListManager {
    // Own only the temporary workbook; never close another user workbook.
    internal sealed class NameExcelSession : IDisposable {
        internal string Folder {get;private set;}
        internal string FilePath {get;private set;}
        internal dynamic Application;
        internal dynamic Workbook;
        internal const string MissingMessage="Error: No file has been detected\nPlease do not exit excel file before clicking import";

        internal NameExcelSession() {
            Folder=Path.Combine(Path.GetTempPath(),"PackingNames-"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Folder);FilePath=Path.Combine(Folder,"Product Names.xlsx");
        }
        internal void Open(bool visible) {
            var type=Type.GetTypeFromProgID("Excel.Application");
            if(type==null)throw new Exception("Microsoft Excel is required for editing names. Install desktop Excel and export again.");
            Application=Activator.CreateInstance(type);
            Workbook=Application.Workbooks.Open(FilePath);
            Workbook.Worksheets["Names"].Rows.AutoFit();
            Workbook.Worksheets["Names"].Range["C2"].Select();
            Application.Visible=visible;
        }
        internal bool IsOpen() {
            try {
                if((object)Application==null||(object)Workbook==null)return false;
                int count=Application.Workbooks.Count;
                if(count==0)return false;
                return String.Equals(Convert.ToString(Workbook.FullName),FilePath,StringComparison.OrdinalIgnoreCase);
            } catch(COMException ex) {
                if(ex.ErrorCode==unchecked((int)0x80010001)||ex.ErrorCode==unchecked((int)0x8001010A))
                    throw new Exception("Excel is busy. Finish editing the cell or close its dialog, then click Import again.");
                return false;
            } catch(InvalidComObjectException){return false;}
        }
        internal void Save() {Workbook.Save();}
        internal void Close() {
            if((object)Workbook!=null) {
                try {Workbook.Close(false);}catch{}
                try {Marshal.FinalReleaseComObject((object)Workbook);}catch{}
                Workbook=null;
            }
            if((object)Application!=null) {
                try {if(Application.Workbooks.Count==0)Application.Quit();}catch{}
                try {Marshal.FinalReleaseComObject((object)Application);}catch{}
                Application=null;
            }
            try {File.Delete(FilePath);}catch{}
            try {Directory.Delete(Folder,false);}catch{}
        }
        public void Dispose(){Close();}
    }
}
