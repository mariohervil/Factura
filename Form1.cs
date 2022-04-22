using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Diagnostics;

//3 puntos


namespace Facturas
{

    public partial class Form1 : Form
    {
       
        int ap = 0;
        Excel.Application appExcel = null;
        Excel.Workbook workBookExcel = null;
        Excel.Worksheet workSheetExcel = null;
        Word.Application apWord = null;
        Word.Document docWord = null;
        string currency = "€";
        double[] iva;
        double[] prices;
        string[] productName;
        ComboBox[] comboBoxProducts, comboBoxUnits;
        TextBox[] textBoxUnitPrices, textBoxAmounts;

        public Form1()
        {
            InitializeComponent();
            initConfig();
            timer1.Enabled = true;
            time();
        }
        private void time()
        {
            string fecha = DateTime.Now.ToString("dddd, dd/MM/yyyy, HH:mm:ss");
            label9.Text = fecha;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {


            time();

        }

        void initConfig()
        {
            comboBoxProducts = new ComboBox[] { comboBoxProduct1, comboBoxProduct2, comboBoxProduct3, comboBoxProduct4, comboBoxProduct5 };
            comboBoxUnits = new ComboBox[] { comboBoxUnit1, comboBoxUnit2, comboBoxUnit3, comboBoxUnit4, comboBoxUnit5 };
            textBoxUnitPrices = new TextBox[] { textBoxUnit1, textBoxUnit2, textBoxUnit3, textBoxUnit4, textBoxUnit5 };
            textBoxAmounts = new TextBox[] { textBoxAmount1, textBoxAmount2, textBoxAmount3, textBoxAmount4, textBoxAmount5 };
            productName = new string[] { "Chair", "Office Chair", "Office Table", "Sofa", "Coffee Maker", "Select" };
            prices = new double[] { 50, 120, 200, 500, 180, 0 };
            iva = new double[] { 0, 10, 14, 21 };
            for (int i = 0; i < productName.Length; i++)
            {
                for (int j = 0; j < comboBoxProducts.Length; j++)
                {
                    comboBoxProducts[j].Items.Add(productName[i]);
                }
            }

            for (int i = 0; i < iva.Length; i++)
            {
                comboBoxTax.Items.Add(iva[i]);
            }

            for (int i = 0; i <= 10; i++)
            {
                for (int j = 0; j < comboBoxUnits.Length; j++)
                {
                    comboBoxUnits[j].Items.Add(i);
                }


            }

            for (int i = 0; i < comboBoxUnits.Length; i++)
            {
                if (comboBoxProducts[i].SelectedIndex <= 0)
                {
                    comboBoxUnits[i].Enabled = false;

                }

            }
            for (int i = 0; i < textBoxAmounts.Length; i++)
            {
                textBoxAmounts[i].Text = "0";
            }



        }


        private void Process(object sender, EventArgs e)
        {

            double totalprice = 0;

            ComboBox cb = (ComboBox)sender;

            try
            {

                for (int row = 0; row < comboBoxProducts.Length; row++)
                {
                    if (comboBoxProducts[row].SelectedIndex >= 0)
                    {
                        textBoxUnitPrices[row].Text = Convert.ToString(prices[comboBoxProducts[row].SelectedIndex]);
                        comboBoxUnits[row].Enabled = true;
                    }
                }
                for (int row = 0; row < comboBoxUnits.Length; row++)
                {

                    if (comboBoxUnits[row].SelectedIndex >= 0)
                    {
                        textBoxAmounts[row].Text = Convert.ToString(prices[comboBoxProducts[row].SelectedIndex] * Convert.ToDouble(comboBoxUnits[row].SelectedItem));


                    }

                }
                for (int i = 0; i < textBoxAmounts.Length; i++)
                {
                    totalprice += Convert.ToDouble(textBoxAmounts[i].Text); //+ Convert.ToDouble(textBoxAmount2.Text) + Convert.ToDouble(textBoxAmount3.Text) + Convert.ToDouble(textBoxAmount4.Text) + Convert.ToDouble(textBoxAmount5.Text);
                    textBoxTotal.Text = totalprice.ToString();
                }


                double totalinvoice;

                if (comboBoxTax.SelectedIndex >= 0)
                {
                    totalinvoice = (Convert.ToDouble(textBoxTotal.Text) + (Convert.ToDouble(textBoxTotal.Text) * (Convert.ToDouble(comboBoxTax.SelectedItem) / 100)));
                    textBoxTotalInvoice.Text = totalinvoice.ToString() + currency;
                    double taxes = (Convert.ToDouble(textBoxTotal.Text) * (Convert.ToDouble(comboBoxTax.SelectedItem) / 100));
                    textBoxTaxes.Text = taxes.ToString() + currency;

                }



            }

            catch (FormatException)
            {


            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void createExcelFile()
        {
            try
            {
       appExcel = new Excel.Application();
                string path = Directory.GetCurrentDirectory() + "\\Invoice\\";
                string ExcelFileNameIn = "ExcelTemplate.xlsx";
                string ExcelFileNameOut = "Invoice.xlsx";
                appExcel.DisplayAlerts = false;
                workBookExcel = appExcel.Workbooks.Open(path + ExcelFileNameIn);
                workSheetExcel = (Excel.Worksheet)workBookExcel.Worksheets.get_Item(1);
                workBookExcel.SaveAs(path + ExcelFileNameOut);
            }
            catch (System.Runtime.InteropServices.COMException e)
            {

               
            }
         
        }
        private void ExcelWriter(Control control)
        {
            try
            {
                if (control.Tag != null && control.Text != "")
                {
                    string tag = control.Tag.ToString();
                    int row = Convert.ToInt32(tag) / 100;//0705 /100=     07 
                    int col = Convert.ToInt32(tag) % 100;//0705 %100=     05
                    //MessageBox.Show(row + ":" + col);
                    workSheetExcel.Cells[row, col] = control.Text;
                }
            }
            catch (Exception e) { }
        }
        private void closeExcel()
        {
            workBookExcel.Save();
            workBookExcel.Close(true);//guardar los cambios:true
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheetExcel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBookExcel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
        }
        private void ExcelKiller()
        {
            System.Diagnostics.Process[] procs =
            System.Diagnostics.Process.GetProcessesByName("EXCEL.EXE");
            if (procs.Length >= 1)
            {
                for (int i = 0; i < procs.Length; i++)
                {
                    try { procs[i].Kill(); }
                    catch (Exception e) { }
                }
            }
        }


        private void createWordFile()
        {
            try
            {
                apWord = new Word.Application();
                docWord = new Word.Document();
                string path = Directory.GetCurrentDirectory() + "\\Invoice\\";
                string wordFileNameIn = "WordTemplate.docx";
                string wordFileNameOut = "Invoice.docx";
                docWord = apWord.Documents.Open(path + wordFileNameIn);
                docWord.SaveAs(path + wordFileNameOut);
            }
            catch (System.Runtime.InteropServices.COMException)
            {

                
            }

        }

        private void wordWriter(Control control)
        {
            Object bookMarkName = control.Name;
            string text = control.Text;
            try
            {
                docWord.Bookmarks[ref bookMarkName].Select();
                apWord.Selection.TypeText(Text: text);

            }
            catch (Exception e)
            {
            }
        }

        private void viewObjects()
        {
            foreach (Control control in this.Controls)
            {
                if (control is TextBox || control is ComboBox)
                {
                    if (ap == 1)
                    {
                        wordWriter(control);
                    }
                    if (ap == 2)
                    {
                        ExcelWriter(control);
                    }
                }

                else if (control is GroupBox)
                {
                    foreach (Control c in control.Controls)
                    {
                        {
                            if (c is TextBox || c is ComboBox) 
                            {
                                if (ap == 1)
                                {
                                    wordWriter(c);
                                }
                                if (ap == 2)
                                {
                                    ExcelWriter(c);
                                }

                            }
                        }
                    }
                }


            }
        }

        // fin de viewObject


        private void SaveWordFile()
        {
            docWord.Save();
        }



        private void CloseWord()
        {
            apWord.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(apWord);
            apWord = null;
        }

        private void buttonPDF_Click(object sender, EventArgs e)
        {
            ap = 3;
            createPDFFile();
            System.Threading.Thread.Sleep(2000);
        }

        private void WordKiller()
        {
            System.Diagnostics.Process[] procs =
            System.Diagnostics.Process.GetProcessesByName("WINWORD");
            if (procs.Length >= 1)
            {
                for (int i = 0; i < procs.Length; i++)
                {
                    try { procs[i].Kill(); }
                    catch (Exception e) { }
                }
            }
        }
        private void createPDFFile()
        {
            try
            {
                apWord = new Word.Application();
                docWord = new Word.Document();
                string path = Directory.GetCurrentDirectory() + "\\Invoice\\";
                string wordFileNameIn = "Invoice.docx";
                string wordFileNameOut = "Invoice";
                docWord = apWord.Documents.Open(path + wordFileNameIn);
                docWord.ExportAsFixedFormat(path + wordFileNameOut + ".pdf", Word.WdExportFormat.wdExportFormatPDF);
            }
            catch (System.Runtime.InteropServices.COMException)
            {

                throw;
            }

        }

        private void buttonWord_Click(object sender, EventArgs e)
        {
            ap = 1;
            WordKiller();
            createWordFile();
            viewObjects();
            SaveWordFile();
            CloseWord();
            WordKiller();
            WordKiller();
            WordKiller();
            System.Threading.Thread.Sleep(2000);
        }
        private void buttonExcel_Click(object sender, EventArgs e)
        {
            ExcelKiller();
            ap = 2;
            createExcelFile();
            viewObjects();
            closeExcel();
            ExcelKiller();
            ExcelKiller();
            ExcelKiller();
            System.Threading.Thread.Sleep(2000);

        }
        private void MailMsg(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + "\\Invoice\\";
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
          
            mail.From = new MailAddress("mariomonlau@gmail.com", "Mario", Encoding.UTF8);
            mail.Subject = "Mario PDF";
            mail.Body = "Correo control";
            mail.To.Add("amiriamg@monlau.com");
            mail.Attachments.Add(new Attachment(path + "Invoice.pdf"));
            SmtpServer.Port = 25; 
            SmtpServer.Credentials = new System.Net.NetworkCredential("mariomonlau@gmail.com", "Mariohv007!");
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(mail);
            System.Threading.Thread.Sleep(2000);

        }
        private void SendToPrinter(object sender, EventArgs e)
        {
            string pathPdf = Directory.GetCurrentDirectory() + "\\Invoice\\Invoice.pdf";
            ProcessStartInfo infoPrintPdf = new ProcessStartInfo();
            infoPrintPdf.FileName = pathPdf;
           
            string printerName = "";
            string driverName = "";
            string portName = "";
            infoPrintPdf.FileName = @"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe";
            infoPrintPdf.Arguments = string.Format("/t {0} \"{1}\" \"{2}\" \"{3}\"",
                pathPdf, printerName, driverName, portName);
            infoPrintPdf.CreateNoWindow = true;
            infoPrintPdf.UseShellExecute = false;
            infoPrintPdf.WindowStyle = ProcessWindowStyle.Hidden;
            Process printPdf = new Process();
            printPdf.StartInfo = infoPrintPdf;
            printPdf.Start();

            
            System.Threading.Thread.Sleep(10000);

            if (!printPdf.CloseMainWindow())              
                printPdf.Kill(); printPdf.WaitForExit();  

            printPdf.Close();
            System.Threading.Thread.Sleep(2000);
        }







    }


}
