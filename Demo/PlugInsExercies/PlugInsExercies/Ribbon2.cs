using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using HS.ExcelExt;
using Hs.Tools;
using System;

namespace PlugInsExercies
{
    public partial class Ribbon2
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            sheet.CleanPassword();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string url = @"http://www.matrix67.com/blog/feed";
            string txt = MyGrab.GetContent(url);
            System.Windows.Forms.MessageBox.Show(txt);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            
        }
    }
}
