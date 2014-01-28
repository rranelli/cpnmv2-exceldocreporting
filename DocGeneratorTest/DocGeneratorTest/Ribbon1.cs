using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel; 
namespace DocGeneratorTest
{
    public partial class Ribbon1
    {
        private const int mappingColumns = 4; //[fieldname, sheet row, sheet col, table col]    
        private const string templateFilepath = @"C:\Users\augusto-ortiz\Desktop\TemplateTeste.xlsm";
        private const string emptyTableFilePath = @"C:\Users\augusto-ortiz\Desktop\DataTeste.xlsm";
        private const string filledTableFilePath = @"C:\Users\augusto-ortiz\Desktop\DataTestefilled.xlsm";
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void button1_Click(object sender, RibbonControlEventArgs e)//Preparar Documento
        {
            
            //string fileName = @"C:\Users\augusto-ortiz\Desktop\TemplateTeste.xlsm";
            //if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    // set the file name from the open file dialog
            //    fileName = openFileDialog1.FileName;
            //    //object fileName = openFileDialog1.FileName;
            //    //object readOnly = false;
            //    //object isVisible = true;
            //    // Here is the way to handle parameters you don't care about in .NET
            //    //object missing = System.Reflection.Missing.Value;
            //    //// Make word visible, so you can see what's happening
            //    //WordApp.Visible = true;
            //    //// Open the document that was chosen by the dialog
            //    //Word.Document aDoc = WordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible);
            //    //// Activate the document so it shows up in front
            //    //aDoc.Activate();
            //    //// Add the copyright text and a line break
            //    //WordApp.Selection.TypeText("Copyright C# Corner");
            //    //Selection.TypeParagraph();
            //}
            Excel.Application xlApp;
            Workbook xlWorkBookTemplate = null;
            Workbook xlWorkBookData = null;
            _Worksheet xlWorkSheetTemplate = null;
            _Worksheet xlWorkSheetData = null;
            Range xlRange;
            xlApp = new Excel.Application();
            xlApp.Visible = true;

            var form = new Form1();
            form.ShowDialog();
            string templateFilePath2 = form.templateFilePath;
            form.Close();
            object misValue = System.Reflection.Missing.Value;

            try
            {
                //1. Reads the template and searches for the marked fields
                xlWorkBookTemplate = xlApp.Workbooks.Open(templateFilepath);
                //xlWorkBookTemplate = xlApp.Workbooks.Open((string) fileName);
                xlWorkBookTemplate.Activate();
                xlWorkSheetTemplate = xlWorkBookTemplate.Sheets[1];
                xlRange = xlWorkSheetTemplate.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int firstRow = xlRange.Row;
                int firstCol = xlRange.Column;

                string cellText;
                int numOfProps = 0;
                int countAllProps = 0;
                //a. Counts the number of distinct items properties to size the mapping matrix accordingly (its considered that all the items in a particular template sheet shares the same properties and theres at least 1 item)
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRange.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("##Item") == true) countAllProps++;
                            if (cellText.StartsWith("##Item-01") == true) numOfProps++;
                        }
                    }
                }

                int numOfItems = countAllProps/numOfProps;//useless for this button, kill it?
                string[,] mappingMatrix = new string[numOfProps, mappingColumns];//[fieldname, sheet row, sheet col, table col]
                //b.Fills the mapping matrix
                int aux = 0;
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRange.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("##Item-01") == true)//change it for a function that searches a custom prefix and extracts the relevant content
                            {
                                //mappingMatrix[aux, 0] = cellText;//fieldname
                                mappingMatrix[aux, 0] = ExtractMarkContent(cellText)[1];//fieldname
                                mappingMatrix[aux, 1] = (i+firstRow-1).ToString();
                                mappingMatrix[aux, 2] = (j+firstCol-1).ToString();
                                mappingMatrix[aux, 3] = (aux + 1).ToString();
                                aux++;
                            }
                        }
                    }
                }
                //2.Builds the data sheet and fills its header line
                xlWorkBookData = xlApp.Workbooks.Open(emptyTableFilePath);
                xlWorkSheetData = (Excel.Worksheet)xlWorkBookData.Worksheets.get_Item(1);
                int headerRow = 1;
                int atualCol, atualRow;
                for (int i = 0; i < numOfProps ; i++)
                {
                    atualCol = Convert.ToInt32(mappingMatrix[i, 2]);
                    atualRow = Convert.ToInt32(mappingMatrix[i, 1]);
                    xlWorkSheetData.Cells[headerRow, i+1] = xlWorkSheetTemplate.Cells[atualRow, atualCol];
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to open file ");
                xlApp.Quit();
                releaseObject(xlApp);
            }
            finally
            {
                xlWorkBookTemplate.Close(true, misValue, misValue);
                xlWorkBookData.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlWorkSheetTemplate);
                releaseObject(xlWorkBookTemplate);
                releaseObject(xlApp);
            }
        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlApp;
            Workbook xlWorkBookTemplate = null;
            Workbook xlWorkBookData = null;
            _Worksheet xlWorkTemplate = null;
            _Worksheet xlWorkData = null;
            Range xlRangeTemplate;
            Range xlRangeData;
            xlApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            try
            {
                //1. Reads the marked template
                xlWorkBookTemplate = xlApp.Workbooks.Open(templateFilepath);//replace it with a user generated path through a form 
                xlWorkTemplate = xlWorkBookTemplate.Sheets[1];
                xlRangeTemplate = xlWorkTemplate.UsedRange;

                int rowCount = xlRangeTemplate.Rows.Count;
                int colCount = xlRangeTemplate.Columns.Count;
                int firstRow = xlRangeTemplate.Row;
                int firstCol = xlRangeTemplate.Column;

                string cellText;
                int numOfItems = 0;
                ////a.Counts the number of items to size the mapping matrix accordingly             
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRangeTemplate.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRangeTemplate.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("#") == true) numOfItems++;//tratar como não contar mesmos fields para itens distintos
                        }
                    }
                }
                string[,] mappingMatrix;
                int mappingColumns = 4;
                mappingMatrix = new string[numOfItems, mappingColumns];
                //b.Fills the mapping matrix whose format it's [fieldname, sheet row, sheet col, table col]
                int aux = 0;
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRangeTemplate.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRangeTemplate.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("#") == true)//change it for a function that searches a custom prefix and extracts the relevant content
                            {
                                mappingMatrix[aux, 0] = cellText;
                                mappingMatrix[aux, 1] = (i + firstRow - 1).ToString();
                                mappingMatrix[aux, 2] = (j + firstCol - 1).ToString();
                                mappingMatrix[aux, 3] = (aux + 1).ToString();
                                aux++;
                            }
                        }
                    }
                }
                //2.Checks/completes the mapping matrix with its data sheet relevant content
                xlWorkBookData = xlApp.Workbooks.Open(filledTableFilePath);//replace it with a user generated path through a form 
                xlWorkData = (Excel.Worksheet)xlWorkBookData.Worksheets.get_Item(1);//different sheet? 
                xlRangeData = xlWorkData.UsedRange;
                int headerRow = 1;//make it a constant?
                string colField;
                bool foundAllFields = true;
                int tableCols = xlRangeData.Columns.Count;
                long tableRows = xlRangeData.Rows.Count;
                //a. Iterates through each fieldname in the mapping mattrix and searches it in the data sheet in order to complete the mapping
                for (int i = 0; i < numOfItems; i++)
                {
                    cellText = mappingMatrix[i, 0];
                    for(int j=0; j<tableCols;j++)
                    {
                        colField = xlWorkData.Cells[headerRow, j + 1].Value2.ToString();
                        foundAllFields = false;
                        if(colField==cellText)
                        {
                            mappingMatrix[i, 3] = (j + 1).ToString();
                            foundAllFields = true;//work in the 'template not found in data sheet' exception
                        }
                    }
                }
                //3. Data abstraction using the mapping matrix
                //Abstracts the data sheet
                Dictionary<long, Dictionary<string, string>> Table = new Dictionary<long, Dictionary<string, string>>();
                string fieldName, fieldValue;
                int col;
                for (int i = headerRow+1; i < tableRows; i++)//selects row
                {
                    //Gets the coordinates in the mapping matrix and adds the fieldname and fieldvalue to the dictionary
                    Dictionary<string, string> Row = new Dictionary<string, string>();
                    for(int j=0; j<numOfItems; j++)
                    {
                        fieldName = mappingMatrix[j, 0];
                        col = Convert.ToInt32(mappingMatrix[j, 3]);
                        fieldValue = xlWorkData.Cells[i, col].Value2.ToString();
                        Row.Add(fieldName, fieldValue);
                    }
                    Table.Add(i, Row);
                }
                //Abstracts Filled Template (initially containing just 1 item, add more in the future)
                List<Dictionary<long, Dictionary<string, string>>> Sheets =
                    new List<Dictionary<long, Dictionary<string, string>>>();
                int sheetItems = 1;
                int numOfSheets = Table.Count/sheetItems;
                
                foreach (long key in Table.Keys)// selects a row in the table
                {
                    for(int sheetCounter=0; sheetCounter!=sheetItems; )
                    {
                        Dictionary<long, Dictionary<string, string>> Sheet = new Dictionary<long, Dictionary<string, string>>();
                        Dictionary<string, string> Item = new Dictionary<string, string>();
                        Item = Table[key];
                        Sheet.Add(key, Item);
                        sheetCounter++;
                        if(sheetCounter==sheetItems)
                        {
                            Sheets.Add(Sheet);        
                        }
                    }
                }
                //4.Gets the abstracted Filled Template and builds each page of the final document using the mapping matrix
                string propName, propValue;
                int row;
                int space = 0;
                foreach (var Sheet in Sheets)
                {
                    foreach (long key in Sheet.Keys)
                    {
                        xlWorkTemplate = (Excel.Worksheet)xlWorkBookTemplate.Worksheets.get_Item(1);
                        foreach (KeyValuePair<string, string>ItemProp in Sheet[key])
                        {
                            propName = ItemProp.Key;//Reads fieldName
                            //Searches fieldName in mapping matrix
                            row = 0;
                            col = 0;
                            for (int j = 0; j != numOfItems; j++)
                            {
                                if (mappingMatrix[j, 0] == propName)//Gets coordinates in matrix
                                {
                                    row = Convert.ToInt32(mappingMatrix[j, 1]);
                                    col = Convert.ToInt32(mappingMatrix[j, 2]);
                                }
                            }
                            propValue = ItemProp.Value;
                            xlWorkTemplate.Cells[row+space*rowCount, col] = propValue;//Fill template with the pertinent data
                        }
                    }
                    space++;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to open file ");
                xlApp.Quit();
                releaseObject(xlApp);
            }
                finally
            {
                xlWorkBookData.Close(true, misValue, misValue);
                xlWorkBookTemplate.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlWorkTemplate);
                releaseObject(xlWorkData);
                releaseObject(xlWorkBookTemplate);
                releaseObject(xlWorkBookData);
                releaseObject(xlApp);
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
        }
        private string[] ExtractMarkContent(string markedContent)// Mark format = #Item-i#propName
        {
            //string markContentTest = "##Item-01&TAG";
            string aux=null;
            if (markedContent.StartsWith("##") == true)
            {
                aux = markedContent.Remove(0, 2);
            }
            string[] Split = aux.Split(new Char[] { '&' });
            return Split;
        }
        private bool CheckMarkConsistency(string markedContent)
        {
            bool Ok = false;
            string markContentTest = "##Item-01&TAG";
            if (markContentTest.StartsWith("##Item-") == true)
            {
                markContentTest = markContentTest.Remove(0, 7);
                string itemNumber = markContentTest[0].ToString() + markContentTest[1].ToString();
                int Num;
                if (int.TryParse(itemNumber, out Num) == true)
                {
                    if (markContentTest[2].ToString() == "&")
                    {
                        Ok = true;
                    }
                }
            }
            else Ok = false;
            return Ok;
        }
    }
}
