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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e){
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBook2;
            Excel._Worksheet xlWorkSheet;
            Excel._Worksheet xlWorkSheet2;
            Excel.Range xlRange;
            //Excel.Range xlRange2;
            xlApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            try
            {
                //1. Lê o template e procura pelos campos marcados
                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\TemplateTeste.xlsm");
                xlWorkSheet = xlWorkBook.Sheets[1];
                xlRange = xlWorkSheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int firstRow = xlRange.Row;
                int firstCol = xlRange.Column;

                string[,] contentMatrix;
                string cellText;
                int numOfItems = 0;
                contentMatrix = new string[rowCount + 1,colCount + 1];
                ////a.Calcula número de itens a serem preenchidos no template para dimensionar a matriz de mapeamento             
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
                mappingMatrix = new string[numOfItems, 4];
                //b.Preenche a matriz de mapeamento
                int aux = 0;
                //int[] aux2;
                //aux2 = new int[3];
                //Dictionary<string, int[]> fieldMap = new Dictionary<string, int[]>();
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRange.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("#") == true)//criar função que procura determinado prefixo e retira o conteúdo pertinente
                            {
                                //aux2[0] = i;
                                //aux2[1] = j;
                                //aux2[2] = aux+1;
                                //fieldMap.Add(cellText,aux2);
                                mappingMatrix[aux, 0] = cellText;
                                mappingMatrix[aux, 1] = (i+firstRow-1).ToString();
                                mappingMatrix[aux, 2] = (j+firstCol-1).ToString();
                                mappingMatrix[aux, 3] = (aux + 1).ToString();
                                aux++;
                            }
                        }
                    }
                }
                //2.Cria o esqueleto da planilha de dados e preenche a linha de cabeçalho
                xlWorkBook2 = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\DataTeste.xlsm");
                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);
                int headerRow = 1;
                int atualCol, atualRow;
                for (int i = 0; i < numOfItems ; i++)
                {
                    atualCol = Convert.ToInt32(mappingMatrix[i, 2]);
                    atualRow = Convert.ToInt32(mappingMatrix[i, 1]);
                    xlWorkSheet2.Cells[headerRow, i+1] = xlWorkSheet.Cells[atualRow, atualCol];
                }
                xlWorkBook.Close(true, misValue, misValue);
                xlWorkBook2.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to open file ");
                xlApp.Quit();
                releaseObject(xlApp);
            }
        }

        

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBook2;
            Excel._Worksheet xlWorkSheet;
            Excel._Worksheet xlWorkSheet2;
            Excel.Range xlRange;
            Excel.Range xlRange2;
            xlApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            try
            {
                //1. Lê o template marcado
                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\TemplateTeste.xlsm");//substituir com caminho fornecido pelo usuário
                xlWorkSheet = xlWorkBook.Sheets[1];
                xlRange = xlWorkSheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int firstRow = xlRange.Row;
                int firstCol = xlRange.Column;

                string[,] contentMatrix;
                string cellText;
                int numOfItems = 0;
                contentMatrix = new string[rowCount + 1, colCount + 1];
                ////a.Calcula número de itens a serem preenchidos no template para dimensionar a matriz de mapeamento             
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRange.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("#") == true) numOfItems++;//tratar como não contar mesmos fields para itens distintos
                        }
                    }
                }
                string[,] mappingMatrix;
                mappingMatrix = new string[numOfItems, 4];
                //b.Preenche a matriz de mapeamento cujo formato é: [fieldname, sheet row, sheet col, table col]
                int aux = 0;
                for (int i = 1; i < rowCount + 1; i++)
                {
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cellText = xlRange.Cells[i, j].Value2.ToString();
                            if (cellText.StartsWith("#") == true)//criar função que procura determinado prefixo e retira o conteúdo pertinente
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
                //2.Checa/completa o mapeamento com a parte da planilha de dados
                xlWorkBook2 = xlApp.Workbooks.Open(@"C:\Users\augusto-ortiz\Desktop\DataTestefilled.xlsm");//substituir com caminho fornecido pelo usuário
                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);//planilha diferente?
                xlRange2 = xlWorkSheet2.UsedRange;
                int headerRow = 1;//tornar constante?
                int atualCol, atualRow;
                string colField;
                bool foundAllFields = true;
                int tableItems = xlRange2.Columns.Count;
                long tableRows = xlRange2.Rows.Count;
                //a. Varre cada fieldname da matriz de mapeamento e procura na planilha de dados para completar o mapeamento
                for (int i = 0; i < numOfItems; i++)
                {
                    cellText = mappingMatrix[i, 0];
                    for(int j=0; j<tableItems;j++)
                    {
                        colField = xlWorkSheet2.Cells[headerRow, j + 1].Value2.ToString();
                        foundAllFields = false;
                        if(colField==cellText)
                        {
                            mappingMatrix[i, 3] = (j + 1).ToString();
                            foundAllFields = true;//tratar exceção em que fieldname do template não é encontrado na planilha de dados
                        }
                    }
                }
                //3. Abstração dos dados com base no mapeamento validado
                //Abstrai planilha de dados
                Dictionary<long, Dictionary<string, string>> Table = new Dictionary<long, Dictionary<string, string>>();
                string fieldName, fieldValue;
                int col;
                for (int i = headerRow+1; i < tableRows; i++)//seleciona linha
                {
                    //Pega as coordenadas da matriz de mapeamento e adiciona o fieldname e o fieldvalue ao dicionário
                    Dictionary<string, string> Row = new Dictionary<string, string>();
                    for(int j=0; j<numOfItems; j++)
                    {
                        fieldName = mappingMatrix[j, 0];
                        col = Convert.ToInt32(mappingMatrix[j, 3]);
                        fieldValue = xlWorkSheet2.Cells[i, col].Value2.ToString();
                        Row.Add(fieldName, fieldValue);
                    }
                    Table.Add(i, Row);
                }
                //Abstrai Sheet preenchida(inicialmente só possui 1 item, adicionar mais futuramente)
                Dictionary<long, Dictionary<string, string>> Sheet = new Dictionary<long, Dictionary<string, string>>();
                string propName, propValue;
                long items = 5;
                //Depois iterar para cada item da sheet
                for (long i = 1; i != items; i++)
                {
                    Dictionary<string, string> Item = new Dictionary<string, string>();
                    //copia campos salvos em Row no Item
                    foreach (long key in Table.Keys)
                    {
                        foreach (KeyValuePair<string, string> propPair in Table[key])
                        {
                            propName = propPair.Key;
                            propValue = propPair.Value;
                            Item.Add(propName, propValue);
                        }
                        //Table.Remove(key);
                        Sheet.Add(i, Item);
                    }
                }
                //4.Pega dados da Sheet abstraída e preenche cada folha do documento final usando o mapeamentoxlWorkBook.Close(true, misValue, misValue);// passar liberação dos recursos para o bloco de finally
                xlWorkBook2.Close(true, misValue, misValue);
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkSheet2);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkBook2);
                xlApp.Quit();
                releaseObject(xlApp);
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to open file ");
                xlApp.Quit();
                releaseObject(xlApp);
            }
            finally
            {
                //xlApp.Quit();
                //releaseObject(xlApp);
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
            finally
            {
                GC.Collect();
            }
        }
    }
}
