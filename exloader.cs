using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Exloader {
    public class ExcelLoader : IDisposableObject { 

        public Microsoft.Office.Interop.Excel.Application Application { get;set; }

        public ExcelLoader()
        {
            this.Application = new Microsoft.Office.Interop.Excel.Application();
        }

        public Workbook OpenWorkbook(string path) {
            return (new Workbooks(this)).Open(path);
        }

        public void Close() {
            if (this.Application != null) this.Application.Quit();
        }

        protected override void Dispose(bool disposing) {
            // if (disposing) 
            Console.WriteLine("Dispose: ExcelLoader");
            this.Close();
            Marshal.ReleaseComObject(this.Application);
        }

    }

    public class Workbooks : IDisposableObject {

        private ExcelLoader loader { get; set; }
        private Microsoft.Office.Interop.Excel.Workbooks workbooks { get; set; }

        public Workbooks(ExcelLoader loader) {
            this.loader = loader;
            this.workbooks = loader.Application.Workbooks;
        }

        public Workbook Open(string path) {
            return new Workbook(this, this.workbooks.Open(path));
        }

        protected override void Dispose(bool disposing) {
            // if (disposing) 
            Console.WriteLine("Dispose: Workbooks");
            Marshal.ReleaseComObject(this.workbooks);
        }
    }
    
    public class Workbook : IDisposableObject {
        private Workbooks workbooks { get; set; }
        private Microsoft.Office.Interop.Excel.Workbook workbook { get; set; }

        public Workbook(Workbooks workbooks, Microsoft.Office.Interop.Excel.Workbook workbook) {
            this.workbooks = workbooks;
            this.workbook = workbook;
        }

        public WorkSheet GetSheet(int index) {
            return null;
        }

        public WorkSheet[] GetWorkSheets() {            
            using (WorkSheets ws = new WorkSheets(this, workbook.Sheets)) {                
                return ws.GetWorkSheets();
            }
        }

        public void Close() {
            if (this.workbook != null) this.workbook.Close(false);
        }

        protected override void Dispose(bool disposing) {
            // if (disposing) 
            Console.WriteLine("Dispose: Workbook");
            this.Close();
            Marshal.ReleaseComObject(this.workbook);
        }
    }

    public class WorkSheets : IDisposableObject {
        private Workbook workbook { get; set; }
        private Microsoft.Office.Interop.Excel.Sheets sheets { get; set; }

        public WorkSheets(Workbook workbook, Microsoft.Office.Interop.Excel.Sheets sheets) {
            this.workbook = workbook;
            this.sheets = sheets;
        }

        public WorkSheet[] GetWorkSheets() {
            
            List<WorkSheet> sheets = new List<WorkSheet>();

            for (int i = 0; i < this.sheets.Count; i++) {
                sheets.Add(new WorkSheet(this, (Microsoft.Office.Interop.Excel.Worksheet)this.sheets[i + 1]));
            }

            return sheets.ToArray();            
        }

        protected override void Dispose(bool disposing) {
            Console.WriteLine("Dispose: Worksheets");
            Marshal.ReleaseComObject(this.sheets);
        }
    }
    
    public class WorkSheet : IDisposableObject {
        private WorkSheets sheets { get; set; }
        private Microsoft.Office.Interop.Excel.Worksheet sheet { get; set; }

        public string Name { get; private set; }

        public WorkSheet(WorkSheets sheets, Microsoft.Office.Interop.Excel.Worksheet sheet) {
            this.sheets = sheets;
            this.sheet = sheet;

            this.Name = this.sheet.Name;
        }

        public object[,] GetUsedRange() {
            using (Range range = new Range(this, this.sheet.UsedRange)) {

                object[,] r = (object[,])range.GetValue();
                object[,] values = new object[r.GetLength(0), r.GetLength(1)];
 
                for (int y = 0; y < values.GetLength(0); y++) {
                    for (int x = 0; x < values.GetLength(1); x++) {
                        values[y, x] = r[y + 1, x + 1];
                    }
                }

                return values;
            }
        }

        protected override void Dispose(bool disposing) {
            Console.WriteLine("Dispose: WorkSheet");
            Marshal.ReleaseComObject(this.sheet);
        }
    }
    
    public class Range : IDisposableObject {
        private WorkSheet sheet { get; set; }
        private Microsoft.Office.Interop.Excel.Range range { get; set; }

        public Range(WorkSheet sheet, Microsoft.Office.Interop.Excel.Range range) {
            this.range = range;
            this.sheet = sheet;
        }

        public Object GetValue() {
            object a = this.range.Value;
            return a;
        }

        protected override void Dispose(bool disposing) {
            Console.WriteLine("Dispose: Range");
            Marshal.ReleaseComObject(this.range);
        }
    }
    public class IDisposableObject : IDisposable {

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing) {

        }
    }
}