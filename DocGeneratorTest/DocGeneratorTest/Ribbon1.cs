using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace DocGeneratorTest
{
    public partial class Ribbon1
    {
        private void Ribbon1Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void PrepDoc(object sender, RibbonControlEventArgs e)
        {
            var prepCall = new Chemtech.Autoproj.DocGenerator.Interface.SelectTemplateForm();
            prepCall.Show();
        }

        private void GenDoc(object sender, RibbonControlEventArgs e)
        {
            var genCall = new Chemtech.Autoproj.DocGenerator.Interface.SelectTemplateFormForGen();
            genCall.Show();
        }
    }
}

 
