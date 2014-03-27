using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
namespace Chemtech.Autoproj.DocGenerator.Interface
{
    public partial class SelectTemplateFormForGen : Form
    {
        private const string DefaultTemplateDirectory = @"C:\Users\augusto-ortiz\Desktop";
        private const string TemplateModelSheetName = "MODELO";

        public SelectTemplateFormForGen()
        {
            InitializeComponent();
        }
        
        private void PopulateList(string path)
        {
            string[] files = Directory.GetFiles(path, "*.xlsx");

            foreach (string file in files)
            {
                Templates_Liberados.Items.Add(file.Replace(DefaultTemplateDirectory, ""));
            }
        }

        private void SelectTemplateFormForGenLoad(object sender, EventArgs e)
        {
            PopulateList(DefaultTemplateDirectory);
        }

        private void BtnConfirmClick(object sender, EventArgs e)
        {
            var file = Templates_Liberados.SelectedItem.ToString();
            var xlApp = (Application)Marshal.GetActiveObject("Excel.Application");
            var xlWorkbookData = xlApp.ActiveWorkbook;
            Range xlRangeTemplate = null;
            try
            {
                var xlWorkBookTemplate = xlApp.Workbooks.Open(DefaultTemplateDirectory + file);
                var xlWorkSheetTemplate = xlWorkBookTemplate.Sheets[TemplateModelSheetName];
                xlRangeTemplate = xlWorkSheetTemplate.UsedRange;
            }
            catch
            {
                MessageBox.Show("Não foi possível abrir o template. \n Favor tentar novamente.");
                Close();
                Show();
            }
            Close();
            if(xlRangeTemplate!=null)
            {
                var serviceCall = new Services.GenerateDocument();
                serviceCall.GenerateDoc(xlApp, xlRangeTemplate, xlWorkbookData);
                Marshal.ReleaseComObject(xlRangeTemplate);
            }
            else MessageBox.Show("Ocorreu um erro ao executar a macro.");
        }

        private void CancelClick(object sender, EventArgs e)
        {
            Close();
        }

        private void Templates_Liberados_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
