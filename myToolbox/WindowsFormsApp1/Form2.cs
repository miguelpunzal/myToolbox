using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CefSharp.WinForms;
using CefSharp;
using System.Xml;
using System.Net;
using Microsoft.VisualBasic.FileIO;
using System.Threading;
using System.IO;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Syncfusion.XlsIO;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public ChromiumWebBrowser browser;
        public String applicationFileName = "DPM.csv";
        public String queryString = "(((Iteration.Name%20%3D%20%22{ITERATIONANME}%22)%20AND%20(Iteration.StartDate%20%3D%20%22{ITERATIONSTARTDATE}T04%3A00%3A00.000Z%22))%20AND%20(Iteration.EndDate%20%3D%20%22{ITERATIONENDDATE}T03%3A59%3A59.000Z%22))";
        public String urlString = "https://rally1.rallydev.com/#/244788293444ud/iterationstatus";
        public Form2()
        {
            InitializeComponent();
            InitBrowser();
        }

        public void InitBrowser()
        {
            Cef.Initialize(new CefSettings());
            browser = new ChromiumWebBrowser(urlString);
            this.Controls.Add(browser);
            DownloadHandler downer = new DownloadHandler(this);
            browser.DownloadHandler = downer;
            downer.OnBeforeDownloadFired += OnBeforeDownloadFired;
            downer.OnDownloadUpdatedFired += OnDownloadUpdatedFired;
            //browser.Dock = DockStyle.Fill;

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            timer1.Start();
            timer2.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "https://rally1.rallydev.com/slm/webservice/v2.x/artifact.csv?workspace=%2Fworkspace%2F43050365031&project=%2Fproject%2F244788293444&projectScopeDown=true&projectScopeUp=true&fetch=FormattedID%2CFormattedID%2CName%2CScheduleState%2CDefects%2CBlocked%2CTags%2CPlanEstimate%2CTasks%2CTaskEstimateTotal%2CTaskRemainingTotal%2COwner%2CDiscussion%2CMilestones%2CLastUpdateDate%2CParent%2CIteration%2Cc_DefectType%2CEnvironment%2Cc_RejectedReason%2CDefectStatus%2CProject%2CDescription&order=DragAndDropRank%20ASC&types=hierarchicalrequirement%2Cdefect%2Cdefectsuite%2Ctestset&query=";
            applicationFileName = "DPM.csv";
            DownloadFile();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "https://rally1.rallydev.com/slm/webservice/v2.x/artifact.csv?workspace=%2Fworkspace%2F43050365031&project=%2Fproject%2F254510370492&projectScopeDown=true&projectScopeUp=true&fetch=FormattedID%2CFormattedID%2CName%2CScheduleState%2CDefects%2CBlocked%2CTags%2CPlanEstimate%2CTasks%2CTaskEstimateTotal%2CTaskRemainingTotal%2COwner%2CDiscussion%2CMilestones%2CLastUpdateDate%2CParent%2CIteration%2Cc_DefectType%2CEnvironment%2Cc_RejectedReason%2CDefectStatus%2CProject%2CDescription&order=DragAndDropRank%20ASC&types=hierarchicalrequirement%2Cdefect%2Cdefectsuite%2Ctestset&query=";
            applicationFileName = "OPSSUITE.csv";
            DownloadFile();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "https://rally1.rallydev.com/slm/webservice/v2.x/artifact.csv?workspace=%2Fworkspace%2F43050365031&project=%2Fproject%2F254511250364&projectScopeDown=true&projectScopeUp=true&fetch=FormattedID%2CFormattedID%2CName%2CScheduleState%2CDefects%2CBlocked%2CTags%2CPlanEstimate%2CTasks%2CTaskEstimateTotal%2CTaskRemainingTotal%2COwner%2CDiscussion%2CMilestones%2CLastUpdateDate%2CParent%2CIteration%2Cc_DefectType%2CEnvironment%2Cc_RejectedReason%2CDefectStatus%2CProject%2CDescription&order=DragAndDropRank%20ASC&types=hierarchicalrequirement%2Cdefect%2Cdefectsuite%2Ctestset&query=";
            applicationFileName = "DEALSTRACKING.csv";
            DownloadFile();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = "https://rally1.rallydev.com/slm/webservice/v2.x/artifact.csv?workspace=%2Fworkspace%2F43050365031&project=%2Fproject%2F254510364788&projectScopeDown=true&projectScopeUp=true&fetch=FormattedID%2CFormattedID%2CName%2CScheduleState%2CDefects%2CBlocked%2CTags%2CPlanEstimate%2CTasks%2CTaskEstimateTotal%2CTaskRemainingTotal%2COwner%2CDiscussion%2CMilestones%2CLastUpdateDate%2CParent%2CIteration%2Cc_DefectType%2CEnvironment%2Cc_RejectedReason%2CDefectStatus%2CProject%2CDescription&order=DragAndDropRank%20ASC&types=hierarchicalrequirement%2Cdefect%2Cdefectsuite%2Ctestset&query=";
            applicationFileName = "BDW.csv";
            DownloadFile();
        }

        //===================================================OPERATION METHODS================================================================
        String formattedDate1;
        String formattedDate2;
        public async void getSprintAndDateAsync()
        {
            try
            {
               
                String x = await browser.GetSourceAsync();

                string date = getBetween(x, textBox8.Text, textBox7.Text);
                string replacedDate = date.Replace(" - ", ";");
                string sprint = getBetween(x, textBox10.Text, textBox9.Text);
                string[] twoDates = replacedDate.Split(';');

                DateTime date1 = DateTime.Parse(twoDates[0]);
                DateTime date2 = DateTime.Parse(twoDates[1]).AddDays(1);

                formattedDate1 = date1.ToString("yyyy-MM-dd");
                formattedDate2 = date2.ToString("yyyy-MM-dd");

                queryString = queryString.Replace("{ITERATIONSTARTDATE}", formattedDate1);
                queryString = queryString.Replace("{ITERATIONENDDATE}", formattedDate2);
                queryString = queryString.Replace("{ITERATIONANME}", sprint);
                timer1.Stop();
                label1.Text = "Success";
            }
            catch (Exception e)
            {
                label1.Text = e.Message.ToString();
            }



        }

        public void DownloadFile()
        {
            if (label1.Text == "Success")
            {
                string downloadUrl = textBox1.Text + queryString;
                textBox1.Text = downloadUrl;
                browser.Load(downloadUrl);
                button5.PerformClick();
                button6.PerformClick();
            }
            else
            {
                MessageBox.Show("Not yet initialized");
            }

        }

        DataTable dtAll = new DataTable();
        public void FormatDataGridView()
        {
            DataTable dt = new DataTable();


            dt = NewDataTable(applicationFileName, ",", true);
            if (applicationFileName.Contains("OPS"))
            {
                dataGridView2.DataSource = dt.Copy();
                dtAll.Merge(dt);
            }
            else if (applicationFileName.Contains("DPM"))
            {
                dataGridView1.DataSource = dt.Copy();
                dtAll.Merge(dt);
            }
            else if (applicationFileName.Contains("BDW"))
            {
                dataGridView3.DataSource = dt.Copy();
                dtAll.Merge(dt);
            }
            else if (applicationFileName.Contains("DEALS"))
            {
                dataGridView4.DataSource = dt.Copy();
                dtAll.Merge(dt);
            }



        }

        public void formatDataGridView()
        {

            DataGridView dtg = new DataGridView();

            if (applicationFileName.Contains("OPS"))
            {
                dtg = dataGridView2;
            }
            else if (applicationFileName.Contains("DPM"))
            {
                dtg = dataGridView1;
            }
            else if (applicationFileName.Contains("BDW"))
            {
                dtg = dataGridView3;
            }
            else if (applicationFileName.Contains("DEALS"))
            {
                dtg = dataGridView4;
            }

            for (int x = 0; x < dtg.Rows.Count - 1;)
            {

                // MessageBox.Show(dataGridView1.Rows[x].Cells[0].Value.ToString());
                if (!dtg.Rows[x].Cells[0].Value.ToString().StartsWith("DE"))
                {
                    dtg.Rows.RemoveAt(x);
                }
                else
                {
                    x++;
                }
            }
            /*
            foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                MessageBox.Show(row.Cells[0].Value.ToString());
                    if (!row.Cells[0].Value.ToString().StartsWith("DE"))
                    {
                        dataGridView1.Rows.RemoveAt(row.Index);
                    }
                    //More code here
                }*/



        }

        public static DataTable NewDataTable(string fileName, string delimiters, bool firstRowContainsFieldNames = true)
        {
            DataTable result = new DataTable();

            using (TextFieldParser tfp = new TextFieldParser(fileName))
            {
                tfp.SetDelimiters(delimiters);

                // Get Some Column Names
                if (!tfp.EndOfData)
                {
                    string[] fields = tfp.ReadFields();

                    for (int i = 0; i < fields.Count(); i++)
                    {
                        if (firstRowContainsFieldNames)
                            result.Columns.Add(fields[i]);

                        else
                            result.Columns.Add("Col" + i);
                    }

                    // If first line is data then add it
                    if (!firstRowContainsFieldNames)
                        result.Rows.Add(fields);
                }

                // Get Remaining Rows
                while (!tfp.EndOfData)
                    result.Rows.Add(tfp.ReadFields());
            }

            return result;
        }

        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }

        //======================================================DOWNLOAD HANDLER============================================================
        public void OnBeforeDownloadFired(object sender, DownloadItem e)
        {
            this.UpdateDownloadAction("OnBeforeDownload", e);
        }

        private void OnDownloadUpdatedFired(object sender, DownloadItem e)
        {
            this.UpdateDownloadAction("OnDownloadUpdated", e);

        }

        private void UpdateDownloadAction(string downloadAction, DownloadItem downloadItem)
        {
            /*
            this.Dispatcher.Invoke(() =>
            {
                var viewModel = (BrowserTabViewModel)this.DataContext;
                viewModel.LastDownloadAction = downloadAction;
                viewModel.DownloadItem = downloadItem;
            });
            */
        }

        // ...

        public class DownloadHandler : IDownloadHandler
        {
            public event EventHandler<DownloadItem> OnBeforeDownloadFired;

            public event EventHandler<DownloadItem> OnDownloadUpdatedFired;

            Form2 mainForm;

            public DownloadHandler(Form2 form)
            {
                mainForm = form;
            }

            public void OnBeforeDownload(IBrowser browser, DownloadItem downloadItem, IBeforeDownloadCallback callback)
            {
                var handler = OnBeforeDownloadFired;
                if (handler != null)
                {
                    handler(this, downloadItem);
                }

                if (!callback.IsDisposed)
                {
                    using (callback)
                    {

                        callback.Continue(downloadItem.SuggestedFileName = mainForm.applicationFileName, showDialog: false);
                        mainForm.browser.Reload();

                    }
                }
            }

            public void OnDownloadUpdated(IBrowser browser, DownloadItem downloadItem, IDownloadItemCallback callback)
            {
                var handler = OnDownloadUpdatedFired;
                if (handler != null)
                {
                    handler(this, downloadItem);

                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            getSprintAndDateAsync();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FormatDataGridView();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            formatDataGridView();
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void button3_Click_1(object sender, EventArgs e)
        {

        }

        public DataGridViewRow CloneWithValues(DataGridViewRow row)
        {
            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
            for (Int32 index = 0; index < row.Cells.Count; index++)
            {
                clonedRow.Cells[index].Value = row.Cells[index].Value;
            }
            return clonedRow;
        }

        public void savetoExcel()
        {
            DataGridView dtg = new DataGridView();
            dtg = dataGridView6;

            string path = Directory.GetCurrentDirectory();

            for (int x = 0; x < 4; x++)
            {


                // creating Excel Application  
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                app.AlertBeforeOverwriting = false;
                // creating new WorkBook within Excel application  
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                
                // creating new Excelsheet in workbook  
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program  
                //  app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.  
                // store its reference to worksheet  
                
                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet  
                worksheet.Name = "Defect_Tracker";
                
                // storing header part in Excel  
                for (int i = 1; i < dtg.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dtg.Columns[i - 1].HeaderText;
                }
                // storing Each row and column value to excel sheet  
                
                    for (int i = 0; i < dtg.Rows.Count - 2; i++)
                    {
                        for (int j = 0; j < dtg.Columns.Count; j++)
                        {
                        try
                            {
                                worksheet.Cells[i + 2, j + 1] = dtg.Rows[i].Cells[j].Value.ToString();
                            }
                            catch (Exception ex) {

                            }
                        
                        }
                    }
                
                

                app.DisplayAlerts = false;
                //workbook.SaveAs(path + "\\MASTER-EXPORT.csv", Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows);
    
                workbook.SaveAs(path + "\\MASTER-EXPORT.csv", Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
       
                app.Quit();
            }
            // save the application  

            // Exit from the application  
            Form3 frm3 = new Form3();
            frm3.Show();
            this.Hide();
            frm3.Focus();
           // Process.Start(path + "\\MASTER-EXPORT.xlsx");
          //  this.Close();
           
        }

        private void button7_Click(object sender, EventArgs e)
        {
            /*
            
            List<String> fnames = new List<String>();
            string path = Directory.GetCurrentDirectory();
            fnames.Add(path+"\\M_DPM.xlsx");
            fnames.Add(path + "\\M_OPSSUITE.xlsx");
            fnames.Add(path + "\\M_BDW.xlsx");
            fnames.Add(path + "\\M_DEALSTRACKING.xlsx");

            MergeWorkbooks( path+"\\testMerge.xlsx", fnames);
            */
            combineFiles();

            //  savetoExcel();
        }

       public static void CombineFiles(List<string> FiletNames, string ExcleOutPutFilePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            try //try to open it. If its a proper excel file 
            {
                Microsoft.Office.Interop.Excel.Workbook TargetWorkbook = app.Workbooks.Open(FiletNames[0]);
                ((Microsoft.Office.Interop.Excel.Worksheet)TargetWorkbook.Sheets[1]).Name = Path.GetFileNameWithoutExtension(FiletNames[0]);

                for (int i = 1; i < FiletNames.Count(); i++)
                {
                    Microsoft.Office.Interop.Excel.Workbook curWorkBook = app.Workbooks.Open(FiletNames[i]);

                    Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)curWorkBook.Sheets[1];
                    workSheet.Name = Path.GetFileNameWithoutExtension(FiletNames[i]);

                    //Copy from source to destination 
                    workSheet.Copy(Type.Missing, TargetWorkbook.Sheets[TargetWorkbook.Sheets.Count]);

                    // Save both workboooks 
                    TargetWorkbook.Save();
                    curWorkBook.Save();
                    curWorkBook.Close();


                }
                TargetWorkbook.SaveAs(Path.ChangeExtension(ExcleOutPutFilePath, "xlsx"), Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                TargetWorkbook.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                app.Quit();
                
            }
        }

        public static void MergeSheets() {
            string path = Directory.GetCurrentDirectory();
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Initialize Application
                IApplication application = excelEngine.Excel;

                //Set default application version
                application.DefaultVersion = ExcelVersion.Excel2013;

                //Open existing workbooks with data entered
                Assembly assembly = typeof(Program).GetTypeInfo().Assembly;

                Stream fileStream1 = assembly.GetManifestResourceStream("MergeExcelFiles.M_DPM.xlsx");
                IWorkbook workbook1 = application.Workbooks.Open(fileStream1);
                IWorksheet worksheet = workbook1.Worksheets[0];

                Stream fileStream2 = assembly.GetManifestResourceStream("MergeExcelFiles.M_OPSSUITE.xlsx");
                IWorkbook workbook2 = application.Workbooks.Open(fileStream2);

                Stream fileStream3 = assembly.GetManifestResourceStream("MergeExcelFiles.M_BDW.xlsx");
                IWorkbook workbook3 = application.Workbooks.Open(fileStream3);

                Stream fileStream4 = assembly.GetManifestResourceStream("MergeExcelFiles.M_DEALSTRACKING.xlsx");
                IWorkbook workbook4 = application.Workbooks.Open(fileStream4);

                //The content in the range A1:C8 from workbook2 and workbook3 is copied to workbook1
                int lastRow = worksheet.UsedRange.LastRow;
                workbook2.Worksheets[0].UsedRange.CopyTo(worksheet.Range[lastRow + 1, 1]);

                lastRow = worksheet.UsedRange.LastRow;
                workbook3.Worksheets[0].UsedRange.CopyTo(worksheet.Range[lastRow + 1, 1]);

                lastRow = worksheet.UsedRange.LastRow;
                workbook4.Worksheets[0].UsedRange.CopyTo(worksheet.Range[lastRow + 1, 1]);

                worksheet.UsedRange.AutofitColumns();

                //Save the workbook
                workbook1.SaveAs(path + "\\OUTPUT_FINAL.csv");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            savetoExcel();
        }

        private static void MergeWorkbooks(string destinationFilePath, List<String> sourceFilePaths)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false; // No prompt when overriding

            // Create a new workbook (index=1) and open source workbooks (index=2,3,...)
            Workbook destinationWb = app.Workbooks.Add();
            foreach (var sourceFilePath in sourceFilePaths)
            {
                app.Workbooks.Add(sourceFilePath);
            }

            // Copy all worksheets
            Worksheet after = destinationWb.Worksheets[1];
            for (int wbIndex = app.Workbooks.Count; wbIndex >= 2; wbIndex--)
            {
                Workbook wb = app.Workbooks[wbIndex];
                for (int wsIndex = wb.Worksheets.Count; wsIndex >= 1; wsIndex--)
                {
                    Worksheet ws = wb.Worksheets[wsIndex];
                    ws.Copy(After: after);
                }
            }

            // Close source documents before saving destination. Otherwise, save will fail
            for (int wbIndex = 2; wbIndex <= app.Workbooks.Count; wbIndex++)
            {
                Workbook wb = app.Workbooks[wbIndex];
                wb.Close();
            }

            // Delete default worksheet
            after.Delete();

            // Save new workbook
            destinationWb.SaveAs(destinationFilePath);
            destinationWb.Close();

            app.Quit();
        }

        public void combineFiles() {
           

        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView5.DataSource = dtAll.Copy();
            for (int x = 0; x < dataGridView5.Rows.Count - 1;)
            {

                // MessageBox.Show(dataGridView1.Rows[x].Cells[0].Value.ToString());
                if (!dataGridView5.Rows[x].Cells[0].Value.ToString().StartsWith("DE"))
                {
                    dataGridView5.Rows.RemoveAt(x);
                }
                else
                {
                    x++;
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (label1.Text == "Success") {
                button1.PerformClick();
                Thread.Sleep(2000);
                button2.PerformClick();
                Thread.Sleep(2000);
                button3.PerformClick();
                Thread.Sleep(2000);
                button4.PerformClick();
                Thread.Sleep(3000);
                button9.PerformClick();
                Thread.Sleep(5000);
                button12.PerformClick();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            savetoExcel();
        }

        public void formatCells() {
            int x = 1;
            foreach (DataGridViewRow row in dataGridView5.Rows)
            {
                try
                {
                    // for (int x = 1; x < dataGridView5.Columns.Count-1; x++)
                    //  {
                    //MessageBox.Show(row.Cells[x].Value.ToString());
                    dataGridView6.Rows.Add();
                    dataGridView6.Rows[row.Index].Cells[0].Value = x.ToString();
                        dataGridView6.Rows[row.Index].Cells[1].Value = row.Cells[14].Value.ToString();
                        dataGridView6.Rows[row.Index].Cells[2].Value = row.Cells[13].Value.ToString();
                        dataGridView6.Rows[row.Index].Cells[3].Value = row.Cells[0].Value.ToString();
                        dataGridView6.Rows[row.Index].Cells[4].Value = " ";
                        dataGridView6.Rows[row.Index].Cells[5].Value = row.Cells[19].Value.ToString();
                        dataGridView6.Rows[row.Index].Cells[6].Value = " ";
                        dataGridView6.Rows[row.Index].Cells[7].Value = row.Cells[9].Value.ToString();
                        dataGridView6.Rows[row.Index].Cells[8].Value = " ";
                    
                    x++;
                  //  }
                }
                catch (Exception ex) {
                  
                }
                
              
                //More code here
            }

            button11.PerformClick();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            formatCells();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (label1.Text == "Success") {
                browser.SetBounds(0, 0, 0, 0);
                timer2.Stop();
                panel1.Visible = true;
                this.SetBounds(this.Location.X, this.Location.Y, 534, 127);
                
                button10.PerformClick();
            }
        }
    }
}
