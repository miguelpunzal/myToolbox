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

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Boolean isStart = false;
        public ChromiumWebBrowser browser;
        public ChromiumWebBrowser browser2;
        public string oldurl = "";
        public string newurl = "https://rally1.rallydev.com/#/244788293444ud/iterationstatus";
        public string originalurl = "https://rally1.rallydev.com/#/244788293444ud/iterationstatus";
        public string dlUrl = "";
        public string notifStatement = "";
        public Boolean dclicks = false;
        public Form1()
        {
            Cef.Initialize(new CefSettings());
            InitializeComponent();
            InitBrowser();
          
        }

        public void InitBrowser() {
           
            browser = new ChromiumWebBrowser(newurl);
            this.Controls.Add(browser);
            //browser.Dock = DockStyle.Fill;
         
        }

        public void DownloadCSV()
        {
            browser2 = new ChromiumWebBrowser(dlUrl);
            this.Controls.Add(browser2);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Visible = true;
            timer1.Start();
            timer3.Start();
            button6.PerformClick();
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                this.Hide();
                notifyIcon1.Visible = true;
                timer2.Start();
            }
        }

        private void IPXMLread()
        {
            string ioIP = "";
            string ioPort = "";
            string conIP = "";
            string conPort = "";

            using (XmlReader reader = XmlReader.Create("Config.xml"))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        switch (reader.Name)
                        {
                            case "io_ip":
                                if (reader.Read())
                                    ioIP = reader.Value;
                                Console.WriteLine(ioIP);
                                break;
                            case "io_port":
                                if (reader.Read())
                                    ioPort = reader.Value;
                                Console.WriteLine(ioPort);
                                break;
                            case "con_ip":
                                if (reader.Read())
                                    conIP = reader.Value;
                                Console.WriteLine(conIP);
                                break;
                            case "con_port":
                                if (reader.Read())
                                    conPort = reader.Value;
                                Console.WriteLine(conPort);
                                break;
                        }
                    }
                }
            }

            var script = string.Format("getIP('{0}', '{1}', '{2}', '{3}');", ioIP, ioPort, conIP, conPort);
            browser.ExecuteScriptAsync(script);
        }

        private void WebBrowserFrameLoadEnded(object sender, FrameLoadEndEventArgs e)
        {
            if (e.Frame.IsMain)
            {
                browser.ViewSource();
                browser.GetSourceAsync().ContinueWith(taskHtml =>
                {
                    var html = taskHtml.Result;
                    richTextBox1.Text = html; 
                });
            }
        }

        private async void button2_ClickAsync(object sender, EventArgs e)
        {
            //browser.ViewSource();
            String x = await browser.GetSourceAsync();
            
            string task = getBetween(x, textBox1.Text, textBox2.Text);
            string defect = getBetween(x, textBox3.Text, textBox4.Text);
            richTextBox1.Text = task+" "+defect;
            try
            {
                int newActive = Convert.ToInt32(task);
                int newDefect = Convert.ToInt32(defect);

                RALLY_NOTIFYER.textnotification = "";
                
                oldActive = Convert.ToInt32(task);
                oldDefect = Convert.ToInt32(defect);
                notifyIcon1.Text = "Active: " + task + " Defects:" + defect;
                //this.Hide();
                if (FormWindowState.Minimized == this.WindowState || dclicks==false)
                {
                    //this.Visible = false;
                    dclicks = false;
                    this.SetBounds(10000, 0, 449, 656);
                    this.ShowInTaskbar = false;
                }
                notifyIcon1.Visible = true;
                timer2.Start();
            }
            catch (Exception ex) {
                richTextBox1.Text = ex.Message;
            }
                
           
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
        public int oldActive = 0;
        public int oldDefect = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            browser.Refresh();
            button5.PerformClick();
           // browser.Reload();
        }

        private async void timer2_TickAsync(object sender, EventArgs e)
        {
            button2.PerformClick();
            //browser.ViewSource();
            String x = await browser.GetSourceAsync();

            string task = getBetween(x, textBox1.Text, textBox2.Text);
            string defect = getBetween(x, textBox3.Text, textBox4.Text);
            richTextBox1.Text = task + " " + defect;
            try
            {
                int newActive = Convert.ToInt32(task);
                int newDefect = Convert.ToInt32(defect);
                Boolean hasOne = false;
                RALLY_NOTIFYER.textnotification = "";


                if (newActive > oldActive)
                {

                    //NEGATIVE NEWS TO NOTIFY
                    RALLY_NOTIFYER.textnotification += Environment.NewLine + "NEW ACTIVE TASK ADDED!";
                    hasOne = true;
                }
                else if (newActive < oldActive)
                {
                    //POSITIVE NEWS TO NOTIFY
                    RALLY_NOTIFYER.textnotification += Environment.NewLine + "AN ACTIVE TASK HAS BEEN CLOSED!";
                    hasOne = true;
                }

                if (newDefect > oldDefect)
                {
                    //NEGATIVE NEWS TO NOTIFY
                    RALLY_NOTIFYER.textnotification += Environment.NewLine + "NEW DEFECT HAS BEEN RAISED!";
                    hasOne = true;
                }
                else if (newDefect < oldDefect)
                {
                    //POSITIVE NEWS TO NOTIFY
                    RALLY_NOTIFYER.textnotification += Environment.NewLine + "A DEFECT HAS BEEN CLOSED!";
                    hasOne = true;
                }
                if (hasOne) {
                    RALLY_NOTIFYER frm2 = new RALLY_NOTIFYER();
                    frm2.Show();
                }


                oldActive = Convert.ToInt32(task);
                oldDefect = Convert.ToInt32(defect);
                notifyIcon1.Text = "Active: " + task + " Defects:" + defect;

            }
            catch (Exception ex)
            {
                richTextBox1.Text = ex.Message;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            notifyIcon1.Visible = true;
            timer2.Start();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //this.Show();
            //this.Visible = true;
            dclicks = true;
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;
            int screenWidth = Screen.PrimaryScreen.Bounds.Width/2;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height/2;
            this.SetBounds(screenWidth, screenHeight, 449, 656);
          
            //notifyIcon1.Visible = false;

        }

        private async void timer3_TickAsync(object sender, EventArgs e)
        {
            //browser.ViewSource();
            String x = await browser.GetSourceAsync();

            string task = getBetween(x, textBox1.Text, textBox2.Text);
            string defect = getBetween(x, textBox3.Text, textBox4.Text);
            richTextBox1.Text = task + " " + defect;
            try
            {
                int newActive = Convert.ToInt32(task);
                int newDefect = Convert.ToInt32(defect);

                RALLY_NOTIFYER.textnotification = "";

                oldActive = Convert.ToInt32(task);
                oldDefect = Convert.ToInt32(defect);
                notifyIcon1.Text = "Active: " + task + " Defects: " + defect;
                //this.Hide();
                notifyIcon1.Visible = true;
                timer2.Start();
                timer3.Stop();
            }
            catch (Exception ex)
            {
                richTextBox1.Text = ex.Message;
            }

        }

        public static List<String> TextCompare(String oldString, String newString) {

            string s1 = oldString;
            string s2 = newString;

            List<string> diff;
            IEnumerable<string> set1 = s1.Split(' ').Distinct();
            IEnumerable<string> set2 = s2.Split(' ').Distinct();

            if (set2.Count() > set1.Count())
            {
                diff = set2.Except(set1).ToList();
            }
            else
            {
                diff = set1.Except(set2).ToList();
            }

            return diff;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<String> sList = TextCompare(richTextBox2.Text, richTextBox3.Text);
            String xx = "";
            for (var x = 0; x < sList.Count; x++) {
               // MessageBox.Show(sList[x]);
                xx += sList[x] + Environment.NewLine;
            }

            richTextBox4.Text = xx;           
        }
        string origvalue_1;
        string origvalue_2;
        string oldhtml="";
        private async void button5_ClickAsync(object sender, EventArgs e)
        {
            try
            {
                origvalue_1 = txt_csvurl.Text;
                origvalue_2 = queryUrl.Text;
               
                DownloadHandler downer = new DownloadHandler(this);
                browser.DownloadHandler = downer;
                downer.OnBeforeDownloadFired += OnBeforeDownloadFired;
                downer.OnDownloadUpdatedFired += OnDownloadUpdatedFired;


                //timer1.Enabled = false;
                //timer2.Enabled = false;
                //timer3.Enabled = false;

                string x = await browser.GetSourceAsync();
                
                string date = getBetween(x, textBox5.Text, textBox6.Text);
                string replacedDate = date.Replace(" - ", ";");
                string sprint = getBetween(x, textBox7.Text, textBox8.Text);
                string[] twoDates = replacedDate.Split(';');

                DateTime date1 = DateTime.Parse(twoDates[0]);
                DateTime date2 = DateTime.Parse(twoDates[1]).AddDays(1);

                String formattedDate1 = date1.ToString("yyyy-MM-dd");
                String formattedDate2 = date2.ToString("yyyy-MM-dd");

                queryUrl.Text = queryUrl.Text.Replace("^^^^^^", formattedDate1);
                queryUrl.Text = queryUrl.Text.Replace("######", formattedDate2);
                queryUrl.Text = queryUrl.Text.Replace("!!!!!!", sprint);

                txt_csvurl.Text += queryUrl.Text;
                //  MessageBox.Show(txt_csvurl.Text);
                oldurl = browser.Address;
                //  browser.Load(txt_csvurl.Text);
                //browser.Reload();

                // dlUrl = txt_csvurl.Text;
                browser.Load(txt_csvurl.Text);

                txt_csvurl.Text = origvalue_1;
                queryUrl.Text = origvalue_2;
                FillDataGrids();
            }
            catch (Exception ex) {
                
            }
            
            
            
        }

       

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

            Form1 mainForm;

            public DownloadHandler(Form1 form)
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
                        
                        callback.Continue(downloadItem.SuggestedFileName, showDialog: false);
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

        public void FillDataGrids() {
            string filename = "export.csv";
            if (filename.Trim() != string.Empty)
            {
                try
                {
                    DataTable dt = new DataTable();
                    DataTable dt2 = new DataTable();
                    if (dataGridView1.Rows.Count < 1)
                    {
                        dt = NewDataTable(filename, ",", false);
                        dataGridView1.DataSource = dt.DefaultView;
                    }
                    else
                    {

                        dt2 = NewDataTable(filename, ",", false);
                        dataGridView2.DataSource = dt2.DefaultView;
                        //COMPARE
                        
                        for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                        {
                            for (int j = 0; j < dataGridView1.ColumnCount - 1; j++)
                            {
                                var arg1 = dataGridView1.Rows[i].Cells[j].Value.ToString();
                                var arg2 = dataGridView2.Rows[i].Cells[j].Value.ToString();

                                if (!arg1.Equals(arg2, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (dataGridView2.Rows[0].Cells[j].Value.ToString() != "Last Update Date") {
                                       
                                        dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                        dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                        notifStatement += Environment.NewLine + dataGridView2.Rows[i].Cells[0].Value.ToString() + " " + dataGridView2.Rows[i].Cells[1].Value.ToString();
                                        notifStatement += Environment.NewLine + "has changed its " + dataGridView2.Rows[0].Cells[j].Value.ToString();
                                        notifStatement += Environment.NewLine + "from " + dataGridView1.Rows[i].Cells[j].Value.ToString().ToUpper();
                                        notifStatement += " to " + dataGridView2.Rows[i].Cells[j].Value.ToString().ToUpper() + ".";
                                        
                                        RALLY_NOTIFYER.textnotification = notifStatement;
                                        RALLY_NOTIFYER frm2 = new RALLY_NOTIFYER();
                                        frm2.Show();
                                    }
                                      
                                    


                                }
                            }
                        }
                        //THEN REPLACE  dataGridView1.DataSource = dt.DefaultView
                        dt = NewDataTable(filename, ",", false);
                        dataGridView1.DataSource = dt.DefaultView;
                        //THEN CLEAR  dataGridView2.DataSource
                        dataGridView2.DataSource = null;
                    }


                }
                catch (Exception ex)
                {
                    
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FillDataGrids();
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Cef.Shutdown();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void queryUrl_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_csvurl_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {


        
        }
        private void Form1_ResizeBegin(object sender, EventArgs e)
        {

        }
        private void Form1_StyleChanged(object sender, EventArgs e)
        {
           
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                notifyIcon1.Visible = true;
               // this.Visible = false;
                
                timer2.Start();
            }
        }
    }
}
