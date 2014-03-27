using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Chemtech.Autoproj.DocGenerator.Services
{
    class CommonServices
    {
        private const string PrefixMark = "##";
        private const string PrefixMarkCopy = "#$#";

        /*Maps the marked items in the template to a mattrix whose format is: 
        [fieldname , tablecol, sheetrowforItem1, sheetcolforItem1, sheetrowforItem2, sheetcolforItem2 ,...]*/
        public string[,] MapTemplate(Range xlRange)
        {
            List<string> props;
            props = new List<string> {"ID"};

            var firstMarkedItem = xlRange.Find(PrefixMark, LookAt: XlLookAt.xlPart);

            var addressRefs = GetPropsNameNLocation(xlRange, firstMarkedItem, props);

            var deltas = Deltas(xlRange, firstMarkedItem);

            var addressMatrix = GetAddressMatrix(deltas, props, addressRefs);

            var mappingMatrix = new string[props.Count, 2 + 2 * (1 + deltas.Count)];
            var propsArray = props.ToArray();

            for (var i = 0; i != mappingMatrix.GetLength(0); i++)
            {
                mappingMatrix[i, 0] = propsArray[i];
                mappingMatrix[i, 1] = "0";
            }

            for (var i = 0; i != addressMatrix.GetLength(0); i++)
            {
                for (var j = 0; j != addressMatrix.GetLength(1); j++)
                    mappingMatrix[i, j + 2] = addressMatrix[i, j].ToString(CultureInfo.InvariantCulture);
            }
            return mappingMatrix;
        }
        //Gets the columns in the header for each field in the mapping matrix in order to complete the matrix.The Header must have a height of a 1-cell and a width of n-cells.
        public void MapDataTableHeader(string[,] mappingMatrix2, Range xlRangeTableHeader)
        {
            for (var i = 0; i != mappingMatrix2.GetLength(0); i++)
            {
                var prop = mappingMatrix2[i, 0];
                if (xlRangeTableHeader.Find(prop) != null) mappingMatrix2[i, 1] = xlRangeTableHeader.Find(prop).Column.ToString(CultureInfo.InvariantCulture);
                else
                {
                    MessageBox.Show("Não foi encontrado a propriedade" + prop + "na tabela. Confira a tabela preenchida.");
                    throw new ArgumentException();
                }
            }
        }
        //Gets the location in the sheet for each property which must be filled using the info from the marked items and the copy marks
        private int[,] GetAddressMatrix(List<int[]> deltas, List<string> props, List<int[]> addressRefs)
        {
            var addressRefsArray = addressRefs.ToArray();

            var deltasArray = deltas.ToArray();

            var addressMatrix = new int[props.Count, 2 * (1 + deltas.Count)];
            for (var curPropCnt = 0; curPropCnt != props.Count(); curPropCnt++)
            {
                //Copies the addresses of the referrence marked items to the mattrix
                addressMatrix[curPropCnt, 0] = addressRefsArray[curPropCnt][0];
                addressMatrix[curPropCnt, 1] = addressRefsArray[curPropCnt][1];
                //Calculates the addresses of the other additional marked items and put them into the mattrix
                for (var additItemCnt = 0; additItemCnt != deltasArray.Count(); additItemCnt++)
                {
                    var delRow = deltasArray[additItemCnt][0];
                    var delCol = deltasArray[additItemCnt][1];
                    addressMatrix[curPropCnt, 2 + 2 * additItemCnt] = 
                        addressRefsArray[curPropCnt][0] + delRow;
                    
                    addressMatrix[curPropCnt, 2 + 2 * additItemCnt + 1] = 
                        addressRefsArray[curPropCnt][1] + delCol;
                }
            }
            return addressMatrix;
        }
        //Gets the number of rows and columns of difference between the first marked item and each mark for copy
        private List<int[]> Deltas(Range xlRange, Range firstMarkedItem)
        {
            var firstAdditionalMarkedItem = xlRange.Find(PrefixMarkCopy, LookAt: XlLookAt.xlPart);
            var deltas = new List<int[]>();
            if (firstAdditionalMarkedItem == null) return deltas;
            var currentAdditionalMarkedItem = firstAdditionalMarkedItem;
            var searchCompleted = false;

            while(searchCompleted==false)
            {
                deltas.Add(GetDel(firstMarkedItem, currentAdditionalMarkedItem));

                if (xlRange.FindNext(currentAdditionalMarkedItem) != null)
                {
                    if (xlRange.FindNext(currentAdditionalMarkedItem).Address == firstAdditionalMarkedItem.Address) searchCompleted = true;
                    else currentAdditionalMarkedItem = xlRange.FindNext(currentAdditionalMarkedItem);
                }
                else searchCompleted = true;
            }
            return deltas;
        }
        //Extracts the property name from each marked item and gets its location in the sheet, storing it in the addressref mattrix
        private List<int[]> GetPropsNameNLocation(Range xlRange, Range firstMarkedItem, List<string> props)
        {
            
            var addressRefs = new List<int[]>();
            var currentMarkedItem = firstMarkedItem;
            var searchCompleted = false;
            while(searchCompleted==false)
            {
                props.Add(currentMarkedItem.Value2.ToString().Replace("##", ""));
                addressRefs.Add(GetCellRangeAddress(currentMarkedItem));

                if (xlRange.FindNext(currentMarkedItem) != null)
                {
                    if (xlRange.FindNext(currentMarkedItem).Address == firstMarkedItem.Address) searchCompleted = true;
                    else currentMarkedItem = xlRange.FindNext(currentMarkedItem);
                }
                else searchCompleted = true;
            }
            return addressRefs;
        }
        private int[] GetCellRangeAddress(Range xlRange)
        {
            var address = new int[2];
            address[0] = xlRange.Row;
            address[1] = xlRange.Column;
            return address;
        }
        private static int[] GetDel(Range firstMarkedItem, Range currentAdditionalMarkedItem)
        {
            var del = new int[2];
            del[0] = currentAdditionalMarkedItem.Row - firstMarkedItem.Row;
            del[1] = currentAdditionalMarkedItem.Column - firstMarkedItem.Column;
            return del;
        }
    }
}
