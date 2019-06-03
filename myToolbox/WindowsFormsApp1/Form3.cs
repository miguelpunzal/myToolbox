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
using System.Runtime.InteropServices;
using System.Data.OleDb;
using MaterialSkin.Controls;

namespace WindowsFormsApp1
{
    public partial class Form3 : MaterialForm
    {
        public ChromiumWebBrowser browser;
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        public Form3()
        {

            InitializeComponent();

          

            //InitBrowser();
        }

        public void InitBrowser()
        {
            Cef.Initialize(new CefSettings());
            browser = new ChromiumWebBrowser("https://myoffice.accenture.com/:x:/r/personal/michael_j_d_layda_accenture_com/_layouts/15/Doc.aspx?sourcedoc=%7Bf2c4636f-ab6b-4e95-89d3-af4d6a91c097%7D&action=default&uid=%7BF2C4636F-AB6B-4E95-89D3-AF4D6A91C097%7D&ListItemId=3639&ListId=%7B0F2C9253-2369-4DBE-A33B-933D1D79EE51%7D&odsp=1&env=prod");
            this.Controls.Add(browser);
            /*
            DownloadHandler downer = new DownloadHandler(this);
            browser.DownloadHandler = downer;
            downer.OnBeforeDownloadFired += OnBeforeDownloadFired;
            downer.OnDownloadUpdatedFired += OnDownloadUpdatedFired;*/
            //browser.Dock = DockStyle.Fill;

        }

        private async void Form3_LoadAsync(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + "\\MASTER-EXPORT.csv";
        
       
            DataTable dt = new DataTable();


            dt = NewDataTable(path, ",", false);

               dataGridView1.DataSource = dt.Copy();
            //await PutTaskDelay();
            //await PutTaskDelay();
            //button3.PerformClick();

        }

        async Task PutTaskDelay()
        {
            await Task.Delay(2000);
        }

        public DataTable NewDataTable(string fileName, string delimiters, bool firstRowContainsFieldNames = false)
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


        private async void button1_ClickAsync(object sender, EventArgs e)
        {
            string user = "miguel.g.punzal@accenture.com";
            browser.ExecuteScriptAsync("document.getElementById('i0116').value=" + '\'' + user + '\'');
            //idSIButton9
            await PutTaskDelay();
            browser.ExecuteScriptAsync("document.getElementById('idSIButton9').click()");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            browser.ExecuteScriptAsync("document.getElementsByClassName('ewr-hdrcornerwidget')[0].click()");

        }

        private async void button3_ClickAsync(object sender, EventArgs e)
        {
           

            Process[] processes = Process.GetProcessesByName("WindowsFormsApp1");
            Process game1 = processes[0];

            IntPtr p = game1.MainWindowHandle;

            SetForegroundWindow(p);

            webBrowser1.Focus();
            //textBox1.Focus();
            SendKeys.SendWait("^(a)");
            await PutTaskDelay();
            SendKeys.SendWait("{DELETE}");
            await PutTaskDelay();
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            Clipboard.SetDataObject(dataObj, true);
            await PutTaskDelay();
            webBrowser1.Focus();
            SendKeys.SendWait("^(v)");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            browser.ExecuteScriptAsync("document.getElementsByClassName('ewr-hdrcornerwidget')[1].click()");
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form2 frm = new Form2();
            frm.Close();

        }
    }
}
