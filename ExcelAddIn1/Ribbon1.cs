using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void NamedRange_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);


            string Name = "MyNamedRange";

            if (((RibbonCheckBox)sender).Checked)
            {
                Microsoft.Office.Interop.Excel.Range selection = Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
                if (selection != null)
                {
                    worksheet.Controls.AddNamedRange(selection, Name);
                }
            }
            else
            {
                worksheet.Controls.Remove(Name);
            }
        }

        private void ListObject_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);


            string listObjectName = "MyListObject";

            if (((RibbonCheckBox)sender).Checked)
            {
                Microsoft.Office.Interop.Excel.Range selection = Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
                if (selection != null)
                {
                    worksheet.Controls.AddListObject(selection, listObjectName);
                }
            }
            else
            {
                worksheet.Controls.Remove(listObjectName);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("bulkrefresh");
        }
    }
}
