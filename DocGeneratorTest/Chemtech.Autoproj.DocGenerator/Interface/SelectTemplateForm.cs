using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Chemtech.Autoproj.DocGenerator.Interface
{
    public partial class SelectTemplateForm : Form
    {
        private const string DefaultTemplateDirectory = @"C:\Users\augusto-ortiz\Desktop";
        private const string TemplateModelSheetName = "MODELO";
        public SelectTemplateForm()
        {
            InitializeComponent();
        }

        private  void PopulateList(string path)
        {
            string[] files = Directory.GetFiles(path, "*.xlsx");
            
            foreach (string file in files) {
                Templates_Liberados.Items.Add(file.Replace(DefaultTemplateDirectory, ""));
            }
        }

        private void ListBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void BtnConfirmClick(object sender, EventArgs e)
        {
            var file = Templates_Liberados.SelectedItem.ToString();
            var xlApp = (Application)Marshal.GetActiveObject("Excel.Application");
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
                var serviceCall = new Services.PrepareDocument();
                serviceCall.PrepareDoc(xlApp, xlRangeTemplate);
                Marshal.ReleaseComObject(xlRangeTemplate);
            }
            else MessageBox.Show("Ocorreu um erro ao executar a macro.");
        }

        private void FormLoad(object sender, EventArgs e)
        {
            PopulateList(DefaultTemplateDirectory);
        }

        private void CancelClick(object sender, EventArgs e)
        {
            Close();
        }

        private void Label1Click(object sender, EventArgs e)
        {

        }

    }
}
