using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Chemtech.Autoproj.DocGenerator.Services
{
    public class GenerateDocument
    {
        //Adopted convention for the template marks (arbitrary, renaming shouldn't interfere with the algorithm for as long the inputs are also changed accordingly)
        private const string PrefixMark = "##";
        private const string PrefixMarkCopy = "#$#";
        private const string PageMark = "%%PAGE";
        private const string NoteMark = "%%NOTES";

        //Adopted naming convention for the name of the sheets in the documents related to this add-in (same as before)
        private const string TableNotesSheetName = "NOTES";
        private const string TableDataSheetName = "DATA";

        private const int HeaderRow = 1;

        private readonly object _misValue = System.Reflection.Missing.Value;

        //[DllImport("user32.dll")]
        //private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public void GenerateDoc(Application xlApp, Range xlRangeTemplate, Workbook xlWorkbookData)
        {
            var xlWorkSheetData = xlWorkbookData.Sheets[TableDataSheetName];
            var xlRangeData = xlWorkSheetData.UsedRange;
            
            //Maps the relation between documents
            var mapIt = new CommonServices();
            var mappingMatrix = mapIt.MapTemplate(xlRangeTemplate);
            mapIt.MapDataTableHeader(mappingMatrix, xlRangeData);

            var numOfItemsPerSheet = mappingMatrix.GetLength(1) - 2;

            //Abstracts the data using the mapping matrix
            var table = AbstractsTable(xlRangeData, mappingMatrix);
            var sheets = AbstractsFilledSheets(numOfItemsPerSheet, table);

            var numOfSheets = GetNumOfSheets(numOfItemsPerSheet, table);

            //Creates the document
            var xlWorkBookDoc = xlApp.Workbooks.Add();
            _Worksheet xlWorkSheetDoc = xlWorkBookDoc.Sheets.Add();
            var xlRangeDoc = xlRangeTemplate;

            var noteAddr = new int[2];
            var notes = new Dictionary<string, string>();
            if (xlRangeTemplate.Find(NoteMark) != null)
            {
                noteAddr = NoteMarkerAddress(xlRangeTemplate);
                Range notesRange = xlWorkbookData.Sheets[TableNotesSheetName].UsedRange;
                notes = GetNotes(notesRange);
            }
            PrepareTemplateToFill(xlRangeDoc, numOfSheets, xlWorkSheetDoc, xlApp);
            //PrepareTemplateToFill(xlRangeTemplate, numOfSheets, xlWorkSheetTemplate, xlApp);
            PutPageMark(xlRangeDoc);

            //PutNotes(xlRangeTemplate);

            var rowCount = xlRangeTemplate.Rows.Count;
            //foreach (var sheet in sheets)
            //{
            //    var currentItemIndex = 0;
            //    var sheetIndex = sheets.IndexOf(sheet);
            //    sheet.Keys.ToList().ForEach(key =>
            //    {
            //        var enumdSheet = sheet[key].ToList();
            //        enumdSheet.ForEach(kvpItemProp =>
            //        {
            //            var propName = kvpItemProp.Key;
            //            var propValue = kvpItemProp.Value;
            //            //Searches fieldName in mapping matrix                             
            //            var row = 0;
            //            var col = 0;
            //            for (var j = 0; j != mappingMatrix.GetLength(0); j++)
            //            {
            //                if (mappingMatrix[j, 0] != propName) continue;
            //                row = Convert.ToInt32(mappingMatrix[j, 2 + 2 * currentItemIndex]);
            //                col = Convert.ToInt32(mappingMatrix[j, 2 + 2 * currentItemIndex + 1]);
            //            }
            //            //xlWorkSheetTemplate.Cells[row + sheetIndex * rowCount, col] = propValue;
            //            xlWorkSheetDOC.Cells[row + sheetIndex * rowCount, col] = propValue;
            //        });
            //        currentItemIndex++;
            //    });
            //    if (notes.Count != 0)
            //    {
            //        string note = "";
            //        var notesIndexes = GetNotesIndexes(sheet);
            //        notesIndexes.Sort();
            //        foreach (var noteIndex in notesIndexes) note = note + notes[noteIndex] + "\n";
            //        //xlWorkSheetTemplate.Cells[noteAddr[0] + sheetIndex * rowCount, noteAddr[1]] = note;
            //        xlWorkSheetDOC.Cells[noteAddr[0] + sheetIndex * rowCount, noteAddr[1]] = note;
            //    }
            //}
            //CleanAnyRemainingMarkedField(xlWorkSheetTemplate.UsedRange);
            CleanAnyRemainingMarkedField(xlWorkSheetDoc.UsedRange);
            MessageBox.Show("O documento foi preenchido!");
        }
        
        //Extracts the content of the table sheet to a Dictionary of Items using the mapping matrix
        private Dictionary<string, Dictionary<string, string>> AbstractsTable(Range xlRangeData, string[,] mappingMatrix)
        {
            var table = new Dictionary<string, Dictionary<string, string>>();

            for (var i = HeaderRow + 1; i != xlRangeData.Rows.Count + 1; i++)
            {
                //Gets the coordinates in the mapping matrix and adds the fieldname and fieldvalue to the dictionary
                var row = new Dictionary<string, string>();
                string rowID = "";
                for (var j = 0; j < mappingMatrix.GetLength(0); j++)
                {
                    var fieldName = mappingMatrix[j, 0];
                    int col = Convert.ToInt32(mappingMatrix[j, 1]);
                    string fieldValue = xlRangeData.Cells[i, col].Value2.ToString();
                    row.Add(fieldName, fieldValue);
                    if (fieldName == "ID") rowID = fieldValue;
                }
                var ID = rowID;
                table.Add(ID, row);
            }
            return table;
        }
        //Partitions the abstracted table into n sheets
        private List<Dictionary<string, Dictionary<string, string>>> AbstractsFilledSheets(int numOfItemsPerSheet, Dictionary<string, Dictionary<string, string>> table)
        {
            var sheets = new List<Dictionary<string, Dictionary<string, string>>>();
            var keys = table.Keys.ToArray();
            var numOfSheets = GetNumOfSheets(numOfItemsPerSheet, table);
            for (var sheetCounter = 0; sheetCounter != numOfSheets; sheetCounter++)
            {
                var sheet = new Dictionary<string, Dictionary<string, string>>();

                for (var itemNumber = sheetCounter * numOfItemsPerSheet;
                     itemNumber != (sheetCounter + 1) * numOfItemsPerSheet && itemNumber != keys.Count();
                     itemNumber++)
                {
                    var itemKey = keys[itemNumber];
                    var item = table[itemKey];
                    sheet.Add(itemKey, item);
                }
                sheets.Add(sheet);
            }
            return sheets;
        }

        private int GetNumOfSheets(int numOfItemsPerSheet, Dictionary<string, Dictionary<string, string>> table)
        {
            int numOfSheets;
            if (table.Count % numOfItemsPerSheet != 0) numOfSheets = (table.Count / numOfItemsPerSheet) + 1;
            else numOfSheets = table.Count / numOfItemsPerSheet;
            return numOfSheets;
        }
        //Extracts the notes in the note sheet to a dictionary
        private Dictionary<string, string> GetNotes(Range rangeNotes)
        {
            var notes = new Dictionary<string, string>();
            var firstRow = rangeNotes.Row;
            var rows = rangeNotes.Rows.Count;
            for (var currentRow = firstRow + 1; currentRow != firstRow + rows; currentRow++)
            {
                var noteKey = rangeNotes.Cells[currentRow, 1].Value2.ToString();
                var noteText = rangeNotes.Cells[currentRow, 2].Value2.ToString();
                if (noteKey != null && noteText != null) notes.Add(noteKey, noteText);
            }
            return notes;
        }

        private void PrepareTemplateToFill(Range xlRangeTemplate, int numOfSheets, _Worksheet xlWorkSheet, Application xlApp)
        {
            var rowCount = xlRangeTemplate.Rows.Count;
            var colCount = xlRangeTemplate.Columns.Count;
            var firstRow = xlRangeTemplate.Row;
            var firstCol = xlRangeTemplate.Column;

            //var firstDataPage = 1;
            //var pageMarker = PageMarkerAddress(xlRangeTemplate);
            //xlRangeTemplate.Cells[pageMarker[0], pageMarker[1]] = firstDataPage;

            for (var sheetCount = 2; sheetCount != numOfSheets + 1; sheetCount++)
            {
                //Make n duplicates the first page of the document
                var newSheetRange = xlApp.Range[xlWorkSheet.Cells[firstRow + (sheetCount - 1) * rowCount, firstCol],
                                                xlWorkSheet.Cells[firstRow + sheetCount * rowCount + 1, firstCol + colCount]];
                xlRangeTemplate.Copy(newSheetRange);
                //Adds horizontal page break 
                var breakCell = xlWorkSheet.Cells[firstRow + (sheetCount - 1) * rowCount, firstCol + colCount];
                xlWorkSheet.HPageBreaks.Add(breakCell);
                //Adds page counter
                //xlRangeTemplate.Cells[pageMarker[0] + (sheetCount - 1) * rowCount, pageMarker[1]] = firstDataPage + (sheetCount - 1);
            }
            xlWorkSheet.PageSetup.PrintArea = xlWorkSheet.UsedRange.Address[_misValue, _misValue, XlReferenceStyle.xlA1, _misValue, _misValue];
        }
        //Searches for every page mark in the document and replaces it for the correspondent page
        private void PutPageMark(Range xlRange)
        {
            var pageMarkCounter = GetFirstPageNumber(xlRange);
            while (xlRange.Find(PageMark, LookAt: XlLookAt.xlPart) != null)
            {
                pageMarkCounter++;
                xlRange.Find(PageMark, LookAt: XlLookAt.xlPart).Value2 = pageMarkCounter;
            }
        }
        //Gets the first page for the page counter
        private int GetFirstPageNumber(Range xlRange)
        {
            if (xlRange.Find(PageMark, LookAt: XlLookAt.xlPart) != null)
            {
                var foundCell = xlRange.Find(PageMark, LookAt: XlLookAt.xlPart);
                var pageNumber = foundCell.Value2.ToString().Replace(PageMark, "");
                foundCell.Value2 = pageNumber;
                return pageNumber;
            }
            else return 1;
        }

        private List<string> GetNotesIndexes(Dictionary<string, Dictionary<string, string>> Sheet)
        {
            var notesInSheet = new List<string>();
            foreach (var itemNote in Sheet
                    .Where(item => item.Value.ContainsKey("NOTES"))
                    .Select(item => item.Value["NOTES"]))
            {
                var notesArray = Regex.Split(itemNote, ", ");
                foreach (var note in notesArray) if (notesInSheet.Contains(note) == false) notesInSheet.Add(note);
            }
            return notesInSheet;
        }
        private void CleanAnyRemainingMarkedField(Range xlRange)
        {
            while (xlRange.Find(PrefixMark) != null) xlRange.Find(PrefixMark).Value2 = "";
            while (xlRange.Find(PrefixMarkCopy) != null) xlRange.Find(PrefixMarkCopy).Value2 = "";
        }

        private int[] NoteMarkerAddress(Range xlRangeTemplate)
        {
            var pageMarkerAddress = new int[2];
            pageMarkerAddress[0] = xlRangeTemplate.Find(NoteMark).Row;
            pageMarkerAddress[1] = xlRangeTemplate.Find(NoteMark).Column;
            return pageMarkerAddress;
        }
    }
}
