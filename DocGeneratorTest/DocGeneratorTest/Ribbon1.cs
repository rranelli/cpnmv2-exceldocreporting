﻿using System;
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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e){
            Excel.Application xlApp;
            Workbook xlWorkBookTemplate = new Excel.Workbook();
            Workbook xlWorkBookData = new Excel.Workbook();
            _Worksheet xlWorkSheetTemplate=xlWorkBookTemplate.Sheets[1] ;
            _Worksheet xlWorkSheetData=xlWorkBookData.Sheets[1];
            Range xlRange;
            xlApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            try
            {
                //1. Reads the template and searches for the marked fields
                xlWorkBookTemplate = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\TemplateTeste.xlsm");
                xlWorkSheetTemplate = xlWorkBookTemplate.Sheets[1];
                xlRange = xlWorkSheetTemplate.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int firstRow = xlRange.Row;
                int firstCol = xlRange.Column;

                string cellText;
                int numOfItems = 0;
                ////a. Counts the number of items to size the mapping matrix accordingly
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRange.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("#") == true) numOfItems++;
                        }
                    }
                }
                string[,] mappingMatrix;
                int mappingColumns = 4;//[fieldname, sheet row, sheet col, table col]
                mappingMatrix = new string[numOfItems, mappingColumns];
                //b.Fills the mapping matrix
                int aux = 0;
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRange.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("#") == true)//change it for a function that searches a custom prefix and extracts the relevant content
                            {
                                mappingMatrix[aux, 0] = cellText;
                                mappingMatrix[aux, 1] = (i+firstRow-1).ToString();
                                mappingMatrix[aux, 2] = (j+firstCol-1).ToString();
                                mappingMatrix[aux, 3] = (aux + 1).ToString();
                                aux++;
                            }
                        }
                    }
                }
                //2.Builds the data sheet and fills its header line
                xlWorkBookData = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\DataTeste.xlsm");
                xlWorkSheetData = (Excel.Worksheet)xlWorkBookData.Worksheets.get_Item(1);
                int headerRow = 1;
                int atualCol, atualRow;
                for (int i = 0; i < numOfItems ; i++)
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
            Excel.Application xlApp = new Excel.Application();
            Workbook xlWorkBookTemplate = new Excel.Workbook();
            Workbook xlWorkBookData = new Excel.Workbook();
            _Worksheet xlWorkTemplate = xlWorkBookTemplate.Sheets[1];
            _Worksheet xlWorkData = xlWorkBookData.Sheets[1];;
            Range xlRangeTemplate;
            Range xlRangeData;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                //1. Reads the marked template
                xlWorkBookTemplate = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\TemplateTeste.xlsm");//replace it with a user generated path through a form 
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
                xlWorkBookData = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\DataTestefilled.xlsm");//replace it with a user generated path through a form 
                xlWorkData = (Excel.Worksheet)xlWorkBookData.Worksheets.get_Item(1);//different sheet? 
                xlRangeData = xlWorkData.UsedRange;
                int headerRow = 1;//make it a constant?
                string colField;
                bool foundAllFields = true;
                int tableItems = xlRangeData.Columns.Count;
                long tableRows = xlRangeData.Rows.Count;
                //a. Iterates through each fieldname in the mapping mattrix and searches it in the data sheet in order to complete the mapping
                for (int i = 0; i < numOfItems; i++)
                {
                    cellText = mappingMatrix[i, 0];
                    for(int j=0; j<tableItems;j++)
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
    }
}
