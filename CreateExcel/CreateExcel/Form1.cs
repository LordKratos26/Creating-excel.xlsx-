using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;

namespace CreateExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        object missing = Type.Missing;
        private void button1_Click(object sender, EventArgs e)
        {

            Excel.Application oXL = new Excel.Application();
            oXL.Visible = false;
            Excel.Workbook oWB = oXL.Workbooks.Add(missing);
            Excel.Worksheet oSheet = oWB.ActiveSheet as Excel.Worksheet;
            oSheet.Name = "defterMain";
            oSheet.Cells[1, 1] = "vkn";
            oSheet.Cells[1, 2] = "Period_start";
            oSheet.Cells[1, 3] = "Period_end";
            oSheet.Cells[1, 4] = "Sube_kodu";
            Excel.Worksheet oSheet2 = oWB.Sheets.Add(missing, missing, 1, missing)
                            as Excel.Worksheet;
            oSheet2.Name = "entryHeader";
            
            oSheet2.Cells[1, 1] = "enteredBy";
            oSheet2.Cells[1, 2] = "enteredDate";
            oSheet2.Cells[1, 3] = "entryNumber";
            oSheet2.Cells[1, 4] = "entryComment";
            oSheet2.Cells[1, 5] = "totalDebit";
            oSheet2.Cells[1, 6] = "totalCredit";
            oSheet2.Cells[1, 7] = "entryNumberCounter";
            Excel.Worksheet oSheet3 = oWB.Sheets.Add(missing, missing, 1, missing)
                            as Excel.Worksheet;
            oSheet3.Name = "entryDetail";
            
            oSheet3.Cells[1, 1] = "lineNumber";
            oSheet3.Cells[1, 2] = "lineNumberCounter";
            oSheet3.Cells[1, 3] = "accountMainID";
            oSheet3.Cells[1, 4] = "accountMainDescription";
            oSheet3.Cells[1, 5] = "accountSubDescription";
            oSheet3.Cells[1, 6] = "accountSubID";
            oSheet3.Cells[1, 7] = "amount";
            oSheet3.Cells[1, 8] = "debitCreditCode";
            oSheet3.Cells[1, 9] = "postingDate";
            oSheet3.Cells[1, 10] = "documentType";
            oSheet3.Cells[1, 11] = "documentTypeDescription";
            oSheet3.Cells[1, 12] = "documentNumber";
            oSheet3.Cells[1, 13] = "documentReference";
            oSheet3.Cells[1, 14] = "documentDate";
            oSheet3.Cells[1, 15] = "paymentMethot";
            oSheet3.Cells[1, 16] = "detailComment";

            var headcolorsh1 = oSheet.Range[
                oSheet.Cells[1, 1],
                oSheet.Cells[1, 4]];
            headcolorsh1.Interior.Color = Excel.XlRgbColor.rgbDarkGray;

            var headcolorsh2 = oSheet2.Range[
                 oSheet2.Cells[1, 1],
                 oSheet2.Cells[1, 7]];
            headcolorsh2.Interior.Color = Excel.XlRgbColor.rgbDarkGray;

            var headcolorsh3 = oSheet3.Range[
                oSheet3.Cells[1, 1],
                oSheet3.Cells[1, 16]];
            headcolorsh3.Interior.Color = Excel.XlRgbColor.rgbDarkGray;


            oSheet.Columns.AutoFit();
            oSheet2.Columns.AutoFit();
            oSheet3.Columns.AutoFit();
            string fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                                    + "\\turkexcel.xlsx";
            oWB.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook,
                missing, missing, missing, missing,
                Excel.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);
            oWB.Close(missing, missing, missing);
            oXL.UserControl = true;
            oXL.Quit();


            MessageBox.Show("Excel saved succesfully");



        }

    
    }
}

     








