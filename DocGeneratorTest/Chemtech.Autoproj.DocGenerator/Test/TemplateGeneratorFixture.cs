using Microsoft.Office.Interop.Excel;
using NUnit.Framework;

namespace Chemtech.Autoproj.DocGenerator.Test
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    public class TemplateGeneratorFixture
    {
        private Application _excelApp;
        private Worksheet _ws;
        private Workbook _wb;
        private Range _range;

        [SetUp]
        public void SetUp()
        {
            _excelApp = new Application { Visible = false, DisplayAlerts = false };
            _excelApp.Visible = true;
            _excelApp.Workbooks.Add();
            _wb = _excelApp.ActiveWorkbook;
            _ws = _wb.ActiveSheet;

            _ws.Cells[1, 1].Value = "##REV";
            _ws.Cells[1, 2].Value = "##TAG";
            _ws.Cells[1, 3].Value = "##STATUS";
            _ws.Cells[1, 5].Value = "##INSTTYPE";
            _ws.Cells[2, 5].Value = "##FUN";
            Range itemRef = _excelApp.Range[_ws.Cells[1, 1], _ws.Cells[1, 3]];

            _ws.Cells[3, 1].Value = "#$#";
            _ws.Cells[5, 1].Value = "#$#";
            _ws.Cells[7, 1].Value = "#$#";

            _range = _ws.UsedRange;

        }

        [TearDown]
        public void TearDown()
        {
            _wb.Close();
            _excelApp.Quit();
        }

        [Test]
        public void CanGetSingleCellNamedRangeAddress()
        {
            var namedRangeAddress = _ws.Names.Item("testNamedRange");

            Assert.AreEqual(namedRangeAddress.RefersTo, "=\"A1\"");
        }

        [Test]
        public void CanGetValueByAddress()
        {
            Assert.AreEqual("testValue", _ws.Range["A1"].Text);
        }
        [Test]
        public void Canfindmarkforcopies()
        {
            var copyRange = _range.Find("#$#");
            Assert.IsNotNull(copyRange);
        }
        [Test]
        public void canCopyrange()
        {
            var copyRange = _excelApp.Range[_ws.Cells[1, 1], _ws.Cells[1, 3]];

            var pasteRange = copyRange.Copy(_range.Find("#$#").Address);
            Assert.AreEqual(_ws.Cells[1, 1].Value, _ws.Cells[2, 1].Value);

        }
        [Test]
        public void Copyisok()
        {
        }
        [Test]
        public void MapisOk()
        {
            var addressMatrix = new int[5, 8];
            //REV
            addressMatrix[0, 0] = 1;
            addressMatrix[0, 1] = 1;
            addressMatrix[0, 2] = 2;
            addressMatrix[0, 3] = 1;
            addressMatrix[0, 4] = 3;
            addressMatrix[0, 5] = 1;
            addressMatrix[0, 6] = 4;
            addressMatrix[0, 7] = 1;
            //TAG
            addressMatrix[1, 0] = 1;
            addressMatrix[1, 1] = 2;
            addressMatrix[1, 2] = 2;
            addressMatrix[1, 3] = 2;
            addressMatrix[1, 4] = 3;
            addressMatrix[1, 5] = 2;
            addressMatrix[1, 6] = 4;
            addressMatrix[1, 7] = 2;
            //STATUS
            addressMatrix[2, 0] = 1;
            addressMatrix[2, 1] = 2;
            addressMatrix[2, 2] = 2;
            addressMatrix[2, 3] = 2;
            addressMatrix[2, 4] = 3;
            addressMatrix[2, 5] = 2;
            addressMatrix[2, 6] = 4;
            addressMatrix[2, 7] = 2;
            //INST TYPE
            addressMatrix[3, 0] = 1;
            addressMatrix[3, 1] = 1;
            addressMatrix[3, 2] = 2;
            addressMatrix[3, 3] = 1;
            addressMatrix[3, 4] = 3;
            addressMatrix[3, 5] = 1;
            addressMatrix[3, 6] = 4;
            addressMatrix[3, 7] = 1;
            //FUN
            addressMatrix[4, 0] = 1;
            addressMatrix[4, 1] = 1;
            addressMatrix[4, 2] = 2;
            addressMatrix[4, 3] = 1;
            addressMatrix[4, 4] = 3;
            addressMatrix[4, 5] = 1;
            addressMatrix[4, 6] = 4;
            addressMatrix[4, 7] = 1;
            var serv = new Services.CommonServices();
            //var actualAddressMatrix = serv.MapTemplate(_range);
            //Assert.AreEqual(actualAddressMatrix, addressMatrix);
        }
    }
}
