using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Chemtech.Autoproj.DocGenerator.Services
{
    class PrepareDocument
    {
        //Adopted arbitray naming convention for the name of the sheets in the documents related to this add-in
        private const int HeaderRow = 1;
        private const string TableNotesSheetName = "NOTES";
        private const string TableDataSheetName = "DATA";
        private readonly object _misValue = System.Reflection.Missing.Value;

        //Fills the header of a excel sheet with the marked items content in the template given 
        //so it can be used to automatically fill empty template documents
        public void PrepareDoc(Application xlApp, Range xlRangeTemplate)
        {
            xlApp.Visible = true;
            xlApp.SheetsInNewWorkbook = 2;
            //Maps template and gets all relevant info from it to fill the header
            var mapIt = new CommonServices();
            var mappingMatrix = mapIt.MapTemplate(xlRangeTemplate);
            //Generates the 'body' of the document
            var xlWorkBookData = xlApp.Workbooks.Add();
            _Worksheet xlWorkSheetData = xlWorkBookData.Sheets[1];
            xlWorkSheetData.Name = TableDataSheetName;
            _Worksheet xlWorkSheetNote = xlWorkBookData.Sheets[2];
            xlWorkSheetNote.Name = TableNotesSheetName;
            //Fills header line and transforms it into a list
            var xlRangeDataHeader = GetHeaderRange(mappingMatrix.GetLength(0) + 1, xlWorkSheetData);
            FillHeader(mappingMatrix, xlRangeDataHeader);
            xlWorkSheetData.ListObjects.Add(XlListObjectSourceType.xlSrcRange, xlRangeDataHeader, xlRangeDataHeader, XlYesNoGuess.xlYes, _misValue);
            MessageBox.Show("Planilha de dados pronta para ser preenchida.");
            xlWorkBookData.Activate();
        }
        //Fills the header of the table sheet header with the name of the items properties, including the default must have item property for every item: ID
        private void FillHeader(string[,] mappingMatrix2, Range xlRangeDataHeader)
        {
            xlRangeDataHeader.Cells[1, 1] = "ID";
            for (var i = 2; i != mappingMatrix2.GetLength(0) + 2; i++) xlRangeDataHeader.Cells[1, i] = mappingMatrix2[i - 2, 0];
        }
        //
        private Range GetHeaderRange(int numOfProps, _Worksheet xlWorkSheetData)
        {
            var xlRangeDataHeader = xlWorkSheetData.Range[xlWorkSheetData.Cells[HeaderRow, 1],
                                      xlWorkSheetData.Cells[HeaderRow, numOfProps]];
            return xlRangeDataHeader;
        }

    }
}
