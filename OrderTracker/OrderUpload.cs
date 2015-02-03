using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using OrderTracker.Entity;
using OrderTracker.OrderDBTableAdapters;
using ClosedXML.Excel;
using System.Globalization;
using System.Threading;
using System.Text.RegularExpressions;
//using SevenZip;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Configuration;

namespace OrderTracker
{
    public partial class OrderUpload : Form
    {
      
        #region [Global Variable]
        private bool isProcessRunning = false;
         OpenFileDialog oOpenFileDialog = new OpenFileDialog();
         BackgroundWorker startWorker = new BackgroundWorker();
       
        #endregion
        DataTable oDataTable = new DataTable();
        orderdetailsTableAdapter oorderdetailsTableAdapter = new orderdetailsTableAdapter();
        orderhosTableAdapter oorderhosTableAdapter = new orderhosTableAdapter();
        orderpackedTableAdapter oorderpackedTableAdapter = new orderpackedTableAdapter();
        ordertransectionTableAdapter oordertransectionTableAdapter = new ordertransectionTableAdapter();
        orderstatusTableAdapter oorderstatusTableAdapter = new orderstatusTableAdapter();
        orderallamountTableAdapter oorderallamountTableAdapter = new orderallamountTableAdapter();
        usermasterTableAdapter ousermasterTableAdapter = new usermasterTableAdapter();
        fileheaderTableAdapter ofileheaderTableAdapter = new fileheaderTableAdapter();
        public OrderUpload()
        {
            InitializeComponent();
        }

        #region [Mainfest]
        private void btnMainfest_Click(object sender, EventArgs e)
        {
            if (GetFileExtension(txtLocation.Text) == "csv")
            {
                BackgroundWorker CompressBack = new BackgroundWorker();
                CompressBack.DoWork += b_Compress;
                CompressBack.RunWorkerAsync();

                startWorker = new BackgroundWorker();
                startWorker.DoWork += btnMainfestClick;
                startWorker.WorkerReportsProgress = true;
                startWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
                startWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                startWorker.RunWorkerAsync();
            }
            else
                MessageBox.Show("Format not supported");
        }
        #region [Save Mainfest]
        private void saveMainfest(List<OrderMainfest> listOrderMainfest)
        {
            int i = 0;
            foreach (OrderMainfest or in listOrderMainfest)
            {
                try
                {
                    if (oorderpackedTableAdapter.GetSubOrderCount(or.Suborder_Id) > 0)
                    { }
                    else
                    {
                        //insert
                        oorderpackedTableAdapter.InsertQuery(or.Suborder_Id, or.Category, or.Courier, or.Product, or.Reference_Code, or.SKU_Code, or.AWB_Number, or.Order_Verified_Date, or.Order_Created_Date, or.Customer_Name, or.Shipping_Name, or.City, or.State, or.PINCode, or.Selling_Price.ToString(), or.IMEI_SERIAL, or.PromisedShipDate, or.MRP.ToString(), or.InvoiceCode, or.CreationDate);
                        OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
                        oOrderDetailsEntity.Amount = 0;
                        oOrderDetailsEntity.Status = "Packed";
                        oOrderDetailsEntity.SuborderId = or.Suborder_Id;
                        oOrderDetailsEntity.RefNo = or.Reference_Code;
                        
                        DataTable rdt = oorderdetailsTableAdapter.GetDataByRefSuborderID(oOrderDetailsEntity.RefNo, oOrderDetailsEntity.SuborderId);
                        if (rdt.Rows.Count > 0)
                            oOrderDetailsEntity.Status = rdt.Rows[0]["Status"].ToString();
                        OrderDetailsAU(oOrderDetailsEntity);
                        oordertransectionTableAdapter.InsertQuery(or.Suborder_Id, OrderConstant.Packed, DateTime.Now);
                    }
                }
                catch (Exception ex)
                {
                    Utility.ErrorLog(ex, "btnMainfestClick-Data Insert");
                }
                i++;
                startWorker.ReportProgress(((100 * i) / listOrderMainfest.Count()));
            }
        }
        #endregion

        private void btnMainfestClick(object sender, EventArgs e)
        {

            try
            {
                startWorker.ReportProgress(20);
                DisableAllUpload();
                oDataTable = new DataTable();
                //gridProduct.DataSource = oDataTable;
                ChangeGridViewdt(gridProduct, oDataTable);
                ChangeTextControlState(lblSaveType, SaveType.Mainfest.ToString());              

                oDataTable = ImportExport.CSVToDataTable(txtLocation.Text, true);
                List<OrderMainfest> listOrderMainfest = new List<OrderMainfest>();
                listOrderMainfest = ParseMainfest(oDataTable);
                if (listOrderMainfest.Count() > 0)
                {
                    int p = 100 / listOrderMainfest.Count();
                    int i = 0;                   
                    ChangeGridView(gridProduct, listOrderMainfest.ToList<dynamic>());
                    lblchange(Path.GetFileName(oOpenFileDialog.FileName), SaveType.Mainfest.ToString(), listOrderMainfest.Count());                   
                    startWorker.ReportProgress(100);
                    //gridProduct.DataSource = listOrderMainfest;
                    if (listOrderMainfest.Count() > 0)
                    {
                        MessageBox.Show("Total Record Found- " + listOrderMainfest.Count().ToString());
                    }
                }
                else
                {
                    MessageBox.Show("No Record Found");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Close your excel file or contact admin");
                Utility.ErrorLog(ex, "orderMainfest import"); 
            }
            finally
            {
                EnableAllUpload();
            }
            //EnableAllUpload();
        }
        #region[Parse HOS]
        private List<HOS> ParseHOS(DataTable dt)
        {
            List<HOS> listHOS = new List<HOS>();
            if (dt.Rows.Count > 0)
            {

                foreach (DataColumn dc in dt.Columns)
                {
                    dc.ColumnName = dc.ColumnName.Trim();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    HOS oHOS = new HOS();

                    oHOS.HosNo = DataTableValidation("HosNo", dt, i);
                   oHOS.HosDate = DataTableValidation("HosDate", dt, i);
                    oHOS.Ref = DataTableValidation("Ref", dt, i);
                    oHOS.Sku = DataTableValidation("Sku", dt, i);
                    oHOS.SubOrderID = DataTableValidation("SubOrderID", dt, i);
                    oHOS.Supc = DataTableValidation("Supc", dt, i);
                    oHOS.AWB = DataTableValidation("AWB", dt, i);
                    oHOS.Weight = DataTableValidation("Weight", dt, i);
                    oHOS.MobileNo = DataTableValidation("MobileNo", dt, i);
                    oHOS.RecDetails = DataTableValidation("RecDetails", dt, i);
                    oHOS.HSDate = DataTableValidationHOSDate("HosDate", dt, i);
                    oHOS.ManifestID = DataTableValidation("ManifestID", dt, i);
                    listHOS.Add(oHOS);
                }
            }
            return listHOS;
        }
        #endregion
        #region [Parse Mainfest]
        #endregion
        private List<OrderMainfest> ParseSaveMainfest(DataTable dt)
        {
            List<OrderMainfest> listOrderMainfest = new List<OrderMainfest>();
            if (dt.Rows.Count > 0)
            {

                foreach (DataColumn dc in dt.Columns)
                {
                    dc.ColumnName = dc.ColumnName.Trim();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    OrderMainfest oOrderMainfest = new OrderMainfest();

                    oOrderMainfest.Category = DataTableValidation("Category", dt, i);
                    oOrderMainfest.Courier = DataTableValidation("Courier", dt, i);
                    oOrderMainfest.Product = DataTableValidation("Product", dt, i);
                    oOrderMainfest.Reference_Code = DataTableValidation("Reference_Code", dt, i);
                    oOrderMainfest.Suborder_Id = dt.Rows[i]["Suborder_Id"].ToString();
                    oOrderMainfest.SKU_Code = DataTableValidation("SKU Code", dt, i);
                    oOrderMainfest.AWB_Number = DataTableValidation("AWB_Number", dt, i);
                    oOrderMainfest.Order_Verified_Date = DataTableValidationDate("Order_Verified_Date", dt, i);
                    oOrderMainfest.Order_Created_Date = DataTableValidationDate("Order_Created_Date", dt, i);
                    oOrderMainfest.Customer_Name = DataTableValidation("Shipping_Name", dt, i);
                    oOrderMainfest.Shipping_Name = DataTableValidation("Shipping_Name", dt, i);
                    oOrderMainfest.City = DataTableValidation("City", dt, i);
                    oOrderMainfest.State = DataTableValidation("State", dt, i);
                    oOrderMainfest.PINCode = DataTableValidation("PINCode", dt, i);
                    oOrderMainfest.Selling_Price = DataTableValidationDecimal("Selling_Price", dt, i);
                    oOrderMainfest.IMEI_SERIAL = DataTableValidation("IMEI_SERIAL", dt, i);
                    oOrderMainfest.PromisedShipDate = DataTableValidationDate("PromisedShipDate", dt, i);
                    oOrderMainfest.MRP = DataTableValidationDecimal("MRP", dt, i);
                    oOrderMainfest.InvoiceCode = DataTableValidation("InvoiceCode", dt, i);
                    oOrderMainfest.CreationDate = System.DateTime.Now;
                    listOrderMainfest.Add(oOrderMainfest);
                }
            }
            return listOrderMainfest;
        }
        private List<OrderMainfest> ParseMainfest(DataTable dt)
        {
            List<OrderMainfest> listOrderMainfest = new List<OrderMainfest>();           
            if (dt.Rows.Count > 0)
            {

                foreach (DataColumn dc in dt.Columns)
                {
                    dc.ColumnName = dc.ColumnName.Trim();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    OrderMainfest oOrderMainfest = new OrderMainfest();

                    oOrderMainfest.Category = DataTableValidation("Category", dt, i);
                    oOrderMainfest.Courier = DataTableValidation("Courier", dt, i);
                    oOrderMainfest.Product = DataTableValidation("Product", dt, i);
                    oOrderMainfest.Reference_Code = DataTableValidation("Reference Code", dt, i);
                    oOrderMainfest.Suborder_Id = dt.Rows[i]["Suborder Id"].ToString();
                    oOrderMainfest.SKU_Code = DataTableValidation("SKU Code", dt, i);
                    oOrderMainfest.AWB_Number = DataTableValidation("AWB Number", dt, i);
                    oOrderMainfest.Order_Verified_Date = DataTableValidationDate("Order Verified Date", dt, i);
                    oOrderMainfest.Order_Created_Date = DataTableValidationDate("Order Created Date", dt, i);
                    oOrderMainfest.Customer_Name = DataTableValidation("Customer Name", dt, i);
                    oOrderMainfest.Shipping_Name = DataTableValidation("Shipping Name", dt, i);
                    oOrderMainfest.City = DataTableValidation("City", dt, i);
                    oOrderMainfest.State = DataTableValidation("State", dt, i);
                    oOrderMainfest.PINCode = DataTableValidation("PIN Code", dt, i);
                    oOrderMainfest.Selling_Price = DataTableValidationDecimal("Selling Price Per Item", dt, i);
                    oOrderMainfest.IMEI_SERIAL = DataTableValidation("IMEI/SERIAL", dt, i);
                    oOrderMainfest.PromisedShipDate = DataTableValidationDate("PromisedShipDate", dt, i);
                    oOrderMainfest.MRP = DataTableValidationDecimal("MRP", dt, i);
                    oOrderMainfest.InvoiceCode = DataTableValidation("InvoiceCode", dt, i);
                    oOrderMainfest.CreationDate = System.DateTime.Now;
                    listOrderMainfest.Add(oOrderMainfest);
                }
            }
            return listOrderMainfest;
        }

        private string DataTableValidation(string col, DataTable dt, int i)
        {
            if (dt.Columns.Contains(col))
            {
                return dt.Rows[i][col].ToString();
            }
            else
            {
                return "NA";
            }
        }

        private DateTime DataTableValidationDate(string col, DataTable dt, int i)
        {
            DateTime validdate = DateTime.ParseExact("01-01-1900", "dd-MM-yyyy", CultureInfo.InvariantCulture); ;
            if (dt.Columns.Contains(col))
            {
                string format = "dd-MM-yyyy H:mm";
                string format1 = "MM-dd-yyyy H:mm";
                string format4 = "yyyy-MM-dd H:mm";
                string format2 = "dd/MM/yyyy H:mm";
                string format3 = "MM/dd/yyyy H:mm";
                string format5 = "yyyy/MM/dd H:mm";
                string format6 = "dd-MM-yyyy";
                string format7 = "MM-dd-yyyy";
                string format8 = "yyyy-MM-dd";
                string format9 = "dd/MM/yyyy";
                string format10 = "MM/dd/yyyy";
                string format11 = "yyyy/MM/dd";
                string format12 = "dd-MM-yyyy H:mm:ss";
                string format13 = "MM-dd-yyyy H:mm:ss";
                string format14 = "yyyy-MM-dd H:mm:ss";
                string format15 = "dd/MM/yyyy H:mm:ss";
                string format16 = "MM/dd/yyyy H:mm:ss";
                string format17= "yyyy/MM/dd H:mm:ss";
                DateTime dateTime;
                string[] allFormat = { format, format1, format2, format3, format4, format5, format6, format7, format8, format9, format10, format11, format12, format13, format14, format15, format16, format17 };
                foreach (string f in allFormat)
                {
                    if (DateTime.TryParseExact(dt.Rows[i][col].ToString(), f, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dateTime))
                    {
                        validdate = DateTime.ParseExact(dt.Rows[i][col].ToString(), f, CultureInfo.InvariantCulture);
                        break;
                    }
                }                

            }

            return validdate;
        }
        private DateTime DataTableValidationHOSDate(string col, DataTable dt, int i)
        {
            DateTime validdate = DateTime.ParseExact("01-01-1900", "dd-MM-yyyy", CultureInfo.CurrentCulture);

            if (dt.Columns.Contains(col) && dt.Rows[i][col].ToString().Length>5)
            {
                string dateString = dt.Rows[i][col].ToString().Substring(8, 2) + "-" + dt.Rows[i][col].ToString().Substring(4, 3) + "-" + dt.Rows[i][col].ToString().Substring(24, 4) + " " + dt.Rows[i][col].ToString().Substring(11, 8);
                string format = "dd-MMM-yyyy HH:mm:ss";
                DateTime dateTime;
                if (DateTime.TryParseExact(dateString, format, CultureInfo.CurrentCulture,
                    DateTimeStyles.None, out dateTime))
                {
                    if (dateString.Length > 10)
                        validdate = DateTime.ParseExact(dateString, format, CultureInfo.CurrentCulture);
                    else
                        validdate = DateTime.ParseExact(dateString, "dd-MMM-yyyy", CultureInfo.CurrentCulture);
                }

            }

            return validdate;
        }

        private DateTime DataTableValidationHOS1MainfestDate(string col)
        {
            DateTime validdate = DateTime.ParseExact("01-01-1900", "dd-MM-yyyy", CultureInfo.CurrentCulture);

            if (col.Length > 5)
            {
                string dateString = col.Substring(8, 2) + "-" + col.Substring(5, 2) + "-" + col.Substring(0, 4) + " " + col.Substring(11, 8);
                string format = "dd-MM-yyyy HH:mm:ss";
                DateTime dateTime;
                if (DateTime.TryParseExact(dateString, format, CultureInfo.CurrentCulture,
                    DateTimeStyles.None, out dateTime))
                {
                    if (dateString.Length > 10)
                        validdate = DateTime.ParseExact(dateString, format, CultureInfo.CurrentCulture);
                    else
                        validdate = DateTime.ParseExact(dateString, "dd-MMM-yyyy", CultureInfo.CurrentCulture);
                }
            }
           

            return validdate;
        }
        private decimal DataTableValidationDecimal(string col, DataTable dt, int i)
        {
            if (dt.Columns.Contains(col))
            {
                return decimal.Parse(dt.Rows[i][col].ToString());
            }
            else
            {
                return decimal.Parse("0.00");
            }
        }
        #endregion

        #region [Delegated & Events]
        private delegate void ChangeButtonTextDelegate(Button button, String Text);
        private void ChangeButtonText(Button button, String Text)
        {
            if (this.InvokeRequired)
            {
                ChangeButtonTextDelegate d = new ChangeButtonTextDelegate(ChangeButtonText);
                this.Invoke(d, button, Text);
            }
            else
                button.Text = Text;
        }

        private delegate void ChangeControlStateDelegate(Control control, Boolean Enabled);
        private void ChangeControlState(Control control, Boolean Enabled)
        {
            if (this.InvokeRequired)
            {
                ChangeControlStateDelegate d = new ChangeControlStateDelegate(ChangeControlState);
                this.Invoke(d, control, Enabled);
            }
            else
                control.Enabled = Enabled;
        }

        private delegate void ChangeTextControlStateDelegate(Control control, string s);
        private void ChangeTextControlState(Control control, string s)
        {
            if (this.InvokeRequired)
            {
                ChangeTextControlStateDelegate d = new ChangeTextControlStateDelegate(ChangeTextControlState);
                this.Invoke(d, control, s);
            }
            else
                control.Text = s;
        }

        private delegate void ChangeGridViewDelegate(DataGridView lvItem, List<dynamic> listHos);
        private void ChangeGridView(DataGridView lvItem, List<dynamic> listHos)
        {
            if (this.InvokeRequired)
            {
                ChangeGridViewDelegate d = new ChangeGridViewDelegate(ChangeGridView);
                this.Invoke(d, lvItem, listHos);
            }
            else
            {
                lvItem.DataSource = listHos;
                lvItem.PerformLayout();
            }
                
        }

        private delegate void ChangeGridViewDelegatedt(DataGridView lvItem, DataTable dt);
        private void ChangeGridViewdt(DataGridView lvItem, DataTable dt)
        {
            if (this.InvokeRequired)
            {
                ChangeGridViewDelegatedt d = new ChangeGridViewDelegatedt(ChangeGridViewdt);
                this.Invoke(d, lvItem, dt);
            }
            else
                lvItem.DataSource = dt;
        }

        private delegate void ChangeProgressDelegate(ProgressBar p, int i);
        private void ChangeProgress(ProgressBar p, int i)
        {
            if (this.InvokeRequired)
            {
                ChangeProgressDelegate d = new ChangeProgressDelegate(ChangeProgress);
                this.Invoke(d, p, i);
            }
            else
                //p.MarqueeAnimationSpeed=i;
            p.Style = ProgressBarStyle.Blocks;
            p.Value = i;
        }

        private delegate void ChangeProgressDelegatev(ProgressBar p, int i);
        private void ChangeProgressv(ProgressBar p, int i)
        {
            if (this.InvokeRequired)
            {
                ChangeProgressDelegatev d = new ChangeProgressDelegatev(ChangeProgressv);
                this.Invoke(d, p, i);
            }
            else
                p.Value = i;
        }
        #endregion
        

        #region [HOS]
        private void btnHOS_Click(object sender, EventArgs e)
        {
            if (GetFileExtension(txtLocation.Text) == "pdf")
            {
                BackgroundWorker CompressBack = new BackgroundWorker();
                CompressBack.DoWork += b_Compress;
                CompressBack.RunWorkerAsync();

                startWorker = new BackgroundWorker();
                startWorker.WorkerReportsProgress = true;
                startWorker.DoWork += btnHOSClick;
                startWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
                startWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                startWorker.RunWorkerAsync();
            }
            else
                MessageBox.Show("Format not supported");
        }
        void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Value = progressBar1.Minimum;
            lblper.Text = "";
            
            //do the code when bgv completes its work
        }
        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // The progress percentage is a property of e
            if (e.ProgressPercentage >= 100)
            {
                progressBar1.Value = 100;
                lblper.Text = "100%";
            }
            else
            {
                progressBar1.Value = e.ProgressPercentage;
                lblper.Text = e.ProgressPercentage+"%";
            }
           
        }

        private void startWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                //btnStart.Text = "Please wait...";
                //btnStart.Enabled = false;
                //rbtnCrawl.Enabled = rbtnUpdate.Enabled = false;
                //treeCatagory.Enabled = false;
                //ChangeButtonText(btnHOS, "Please wait...");
                //ChangeControlState(btnHOS, false);
            }
            catch { }
        }
        private void btnHOSClick(object sender, EventArgs e)
        {
            DisableAllUpload();
            
            if (!File.Exists(oOpenFileDialog.FileName))
            {
                MessageBox.Show("Input file missing", "Alert");
                EnableAllUpload();
                return;
            }


            //String destDir, tempFile = string.Empty;

            FileInfo srcFinfo = new FileInfo(oOpenFileDialog.FileName);
            //int numberOfPages = new PdfReader(srcFinfo.FullName).NumberOfPages;
            //try
            //{
            //    destDir = String.Format("{0}\\Output", Application.StartupPath);
            //    tempFile = String.Format("{0}\\_temp.pdf", destDir);
            //    if (!Directory.Exists(destDir))
            //        Directory.CreateDirectory(destDir);

            //    foreach (FileInfo destFinfo in new DirectoryInfo(destDir).GetFiles())
            //        destFinfo.Delete();
            //}
            //catch { }

            int index;
            int p=0;
            string hosCode = "", hosdate = "";
            List<HOS> listHOS = new List<HOS>();
            ChangeGridView(gridProduct, listHOS.ToList<dynamic>());
            ChangeTextControlState(lblSaveType, SaveType.HOS.ToString());
           
            bool status = false;
            p = 10;
            //for (int page = 1; page <= numberOfPages; page++)
            //{
                try
                {
                   
                    startWorker.ReportProgress(p);
                    #region[File read-20]
                    #endregion
                    String fileText = ImportExport.PDFText(srcFinfo.FullName);
                    p=30;
                    int t = 1;
                    List<HOSDetails> oHOSDetailsList = new List<HOSDetails>();
                    
                    fileText=fileText.Replace("FORMS NEEDED", "");                   
                    #region [Hos Date Extract]

                             string MatchEmailPattern = "HOS\\w{7}";
                            Regex rx = new Regex(MatchEmailPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            MatchCollection matches = rx.Matches(fileText);
                            // Report the number of matches found.
                            int noOfMatches = matches.Count;
                            if (matches.Count == 0)
                            {
                                MessageBox.Show("Not Valid File");
                                EnableAllUpload();
                                return;
                            }
                            startWorker.ReportProgress(p);
                            // Report on each match.
                            foreach (Match match in matches)
                            {
                                try
                                {
                                    HOSDetails oHOSDetails = new HOSDetails();
                                    //hosCode = match.Captures[0].Value;
                                    oHOSDetails.index = match.Index;
                                    oHOSDetails.HosNo = match.Value;
                                    string s = fileText.Substring(oHOSDetails.index, 100);
                                    index = s.IndexOf("Handover Sheet Date:");
                                    if (index > -1)
                                    {
                                        string hosdatePat = "19[789]\\d|20[01]\\d";
                                        Regex rxhosdate = new Regex(hosdatePat, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                                        MatchCollection matcheshosdate = rxhosdate.Matches(s);
                                        if (matcheshosdate.Count > 0)
                                            hosdate = s.Substring(index + 21, matcheshosdate[0].Index + 4 - index - 21);
                                    }
                                    oHOSDetails.HosDate = hosdate.Replace("*", "");
                                    oHOSDetailsList.Add(oHOSDetails);
                                }
                                catch (Exception ex)
                                {
                                    Utility.ErrorLog(ex, "Hos Date Extract"); 
                                }
                            }
                            oHOSDetailsList = oHOSDetailsList.GroupBy(x => x.HosNo).Select(x => x.First()).ToList();
                           
                    #endregion
                    #region[Reference Code-20]
                            List<RefDetails> oRefDetailsList = new List<RefDetails>();
                            string MatchEmailPatternR = "Reference Code";
                            Regex rxR = new Regex(MatchEmailPatternR, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            MatchCollection matchesR = rxR.Matches(fileText);
                            // Report the number of matches found.
                            int noOfMatchesR = matchesR.Count;
                            // Report on each match.
                            int k = 0;
                            for (k = 0; k < matchesR.Count; )
                            {
                                RefDetails oRefDetails = new RefDetails();
                                oRefDetails.index = matchesR[k].Index;
                                oRefDetailsList.Add(oRefDetails);
                                k = k + 3;
                            }
                            #endregion
                    p = 50;
                    startWorker.ReportProgress(p);
                    int j = 1;
                    #region[Retrive Date-30]
                           
                            foreach (RefDetails d in oRefDetailsList)
                            {
                                var l = oHOSDetailsList.Where(n =>n.index<d.index).ToList();
                                index = d.index;
                                int index1 = fileText.Length - 1;
                                string totalline = fileText.Substring(index + 15, index1 - index - 15);
                                string[] line = totalline.Split('\n');

                                foreach (string str in line)
                                {
                                    if (str.Contains("Handover Sheet Date"))
                                        break;
                                    try
                                    {
                                        if (str.Length > 0)
                                        {
                                            string[] orderdetails = str.Split(' ');
                                            int result;
                                            if (orderdetails.Count() == 6 && int.TryParse(orderdetails[0], out result))
                                            {
                                                HOS oHOS = new HOS();
                                                oHOS.SubOrderID = orderdetails[1];
                                                oHOS.Sku = orderdetails[2];
                                                oHOS.Supc = orderdetails[3];
                                                oHOS.AWB = orderdetails[4];
                                                oHOS.Ref = orderdetails[5].Replace("\r", "");
                                                oHOS.CreationDate = System.DateTime.Now;

                                                oHOS.HosNo = l[l.Count-1].HosNo;
                                                oHOS.HosDate = l[l.Count - 1].HosDate;

                                                listHOS.Add(oHOS);
                                            }
                                            else if (orderdetails.Count() == 9 && int.TryParse(orderdetails[0], out result))
                                            {
                                                //for Multiple Suborder ID
                                                HOS oHOS = new HOS();
                                                HOS oHOS1 = new HOS();

                                                oHOS.SubOrderID = orderdetails[1].Replace(",", "");
                                                oHOS.Sku = orderdetails[4].Replace(",", "");
                                                oHOS.Supc = orderdetails[6].Replace(",", "");
                                                oHOS.AWB = orderdetails[7].Replace(",", "");
                                                oHOS.Ref = orderdetails[8].Replace("\r", "");
                                                oHOS.CreationDate = System.DateTime.Now;
                                                oHOS.HosNo = l[l.Count - 1].HosNo;
                                                oHOS.HosDate = l[l.Count - 1].HosDate;

                                                listHOS.Add(oHOS);

                                                oHOS1.SubOrderID = orderdetails[2].Replace(",", "");
                                                oHOS1.Sku = orderdetails[4].Replace(",", "");
                                                oHOS1.Supc = orderdetails[6].Replace(",", "");
                                                oHOS1.AWB = orderdetails[7].Replace(",", "");
                                                oHOS1.Ref = orderdetails[8].Replace("\r", "");
                                                oHOS1.CreationDate = System.DateTime.Now;
                                                oHOS1.HosNo = l[l.Count - 1].HosNo;
                                                oHOS1.HosDate = l[l.Count - 1].HosDate;
                                                listHOS.Add(oHOS1);
                                            }
                                        }
                                    }
                                    catch(Exception ex)
                                    {
                                        Utility.ErrorLog(ex, "HOS Import- HOS Date"); 
                                    }
                                   
                                    
                                   
                                }
                                j++;
                                startWorker.ReportProgress(p + ((j * 30) / oRefDetailsList.Count()));
                            }
                        #endregion
                            p = 80;
                }
                catch(Exception ex)
                {
                    Utility.ErrorLog(ex, "orderMainfest Click"); 
                    MessageBox.Show(ex.ToString(), "Error");
                }
                //startWorker.ReportProgress(((page *100)/ numberOfPages)-10);
           // }
            ////
             int c= 1;
             #region[Data Insert-20]
             
             listHOS = listHOS.GroupBy(x => x.SubOrderID).Select(x => x.First()).ToList();
            //foreach (HOS oHOS in listHOS)
            //{
            //    if (oorderhosTableAdapter.GetSuborderCount(oHOS.SubOrderID) > 0)
            //    {
            //        //update
            //        // oorderhosTableAdapter.UpdateQuery(oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.SubOrderID);
            //    }
            //    else
            //    {
            //        //Insert                               
            //        oorderhosTableAdapter.InsertQuery(oHOS.SubOrderID, oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.CreationDate, oHOS.HosNo, oHOS.HosDate);
            //        oordertransectionTableAdapter.InsertQuery(oHOS.SubOrderID, OrderConstant.Shipped, DateTime.Now);
            //        OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
            //        oOrderDetailsEntity.Amount = 0;
            //        oOrderDetailsEntity.Status = "Shipped";
            //        oOrderDetailsEntity.SuborderId = oHOS.SubOrderID;
            //        OrderDetailsAU(oOrderDetailsEntity);
                   
            //    }
            //    c++;
            //    startWorker.ReportProgress(p + ((c * 20) / listHOS.Count()));
            //}
             #endregion
             var selectedFields = from s in listHOS
                                  select new
                                  {
                                      s.SubOrderID,
                                      s.Sku,
                                      s.Supc,
                                      s.AWB,
                                      s.Ref,
                                      s.HosNo,
                                      s.HosDate

                                  };
             ChangeGridView(gridProduct, selectedFields.ToList<dynamic>());
             lblchange(Path.GetFileName(oOpenFileDialog.FileName), SaveType.HOS.ToString(), listHOS.Count());   
            startWorker.ReportProgress(100);
            if (listHOS.Count() > 0)
            {
                MessageBox.Show("Total Record Found- " + listHOS.Count().ToString());

            }
            else
            {
                MessageBox.Show("No record found", "Alert");
            }
            EnableAllUpload();
        }
        #region[HOS Insert]
        private void HOSInsert(List<HOS> listHOS)
        {
            int c = 0;
            foreach (HOS oHOS in listHOS)
            {
                if (oorderhosTableAdapter.GetSuborderCount(oHOS.SubOrderID) > 0)
                {
                    //update
                    // oorderhosTableAdapter.UpdateQuery(oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.SubOrderID);
                }
                else
                {
                    //Insert                               
                    oorderhosTableAdapter.InsertQuery(oHOS.SubOrderID, oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.CreationDate, oHOS.HosNo, oHOS.HosDate, oHOS.HSDate, oHOS.Weight, oHOS.MobileNo, oHOS.RecDetails,oHOS.ManifestID);
                    oordertransectionTableAdapter.InsertQuery(oHOS.SubOrderID, OrderConstant.Shipped, DateTime.Now);
                    OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
                    oOrderDetailsEntity.Amount = 0;
                    oOrderDetailsEntity.Status = "Shipped";
                    oOrderDetailsEntity.SuborderId = oHOS.SubOrderID;
                    oOrderDetailsEntity.RefNo = oHOS.Ref;
                    OrderDetailsAU(oOrderDetailsEntity);

                }
                c++;
                startWorker.ReportProgress(((c * 100) / listHOS.Count()));
            }
        }
        #endregion

        #region [Read Text]
        public String ExtractPageText(string sourcePdfPath)
        {
            FileInfo fInfo = new FileInfo(sourcePdfPath);
            if (fInfo.Exists && fInfo.Extension == ".pdf")
            {
                StringBuilder text = new StringBuilder();
                PdfReader pdfReader = new PdfReader(fInfo.FullName);
                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    text.Append(System.Environment.NewLine);
                    text.Append("\n Page Number:" + page);
                    text.Append(System.Environment.NewLine);
                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                    pdfReader.Close();
                }
                //pdftext.Text += text.ToString();
                //File.WriteAllText(outputPdfPath, text.ToString());
                return text.ToString();
            }
            return null;
        }
        #endregion

        #region [Extract Pages]
        public void ExtractPage(string sourcePdfPath, string outputPdfPath, int pageNumber, string password = "")
        {
            PdfReader reader = null;
            Document document = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // Capture the correct size and orientation for the page:
                document = new Document(reader.GetPageSizeWithRotation(pageNumber));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(document,
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                document.Open();

                // Extract the desired page number:
                importedPage = pdfCopyProvider.GetImportedPage(reader, pageNumber);
                pdfCopyProvider.AddPage(importedPage);
                document.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                Utility.ErrorLog(ex, "extract Page"); 
                throw ex;
            }
        }

        public void ExtractPages(string sourcePdfPath, string outputPdfPath, int startPage, int endPage)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new Document(reader.GetPageSizeWithRotation(startPage));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument,
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the specified range and add the page copies to the output file:
                for (int i = startPage; i <= endPage; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void ExtractPages(string sourcePdfPath, string outputPdfPath, int[] extractThesePages)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the 
                // contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new Document(reader.GetPageSizeWithRotation(extractThesePages[0]));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument,
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the array and add the page copies to the output file:
                foreach (int pageNumber in extractThesePages)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, pageNumber);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #endregion

        private void btnbrowse_Click(object sender, EventArgs e)
        {
            oOpenFileDialog.Title = "Open File Dialog";
           // oOpenFileDialog.InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            oOpenFileDialog.Filter = "Excel/PDF Files|*.xls;*.xlsx;*.pdf;*.csv";
            oOpenFileDialog.FilterIndex = 2;
            oOpenFileDialog.RestoreDirectory = true;

            if (oOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtLocation.Text = oOpenFileDialog.FileName;
            }
        }

        #region [Import Payment]
        private void btnPayment_Click(object sender, EventArgs e)
        {
            if (GetFileExtension(txtLocation.Text) == "xls" || GetFileExtension(txtLocation.Text) == "xlsx")
            {
                BackgroundWorker CompressBack = new BackgroundWorker();
                CompressBack.DoWork += b_Compress;
                CompressBack.RunWorkerAsync();

                startWorker = new BackgroundWorker();
                startWorker.DoWork += btnPaymentClick;
                startWorker.WorkerReportsProgress = true;
                startWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                startWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
                startWorker.RunWorkerAsync();
            }
            else
                MessageBox.Show("Format not supported");
        }
        private void btnPaymentClick(object sender, EventArgs e)
        {
            DisableAllUpload();
            try
            {
                startWorker.ReportProgress(20);
                DataSet oDataSet = new DataSet();
                oDataSet = ImportExport.ExcelToDataSet(txtLocation.Text);
                List<Payment> listOrderPayment = new List<Payment>();
                ChangeGridView(gridProduct, listOrderPayment.ToList<dynamic>());
                ChangeTextControlState(lblSaveType, SaveType.Payment.ToString());
                if (oDataSet != null)
                {
                    listOrderPayment = ParsePayment(oDataSet.Tables[0]);
                    //var distinctTypeIDs = listOrderPayment.Select(x => x.SuborderID).Distinct();                    
                    //int i = 0;
                    //foreach (string p in distinctTypeIDs)
                    //{
                    //    var paymentlist = (from o in listOrderPayment where o.SuborderID == p select o).ToList();
                    //    PaymentStatus(paymentlist);
                    //    i++;
                    //    startWorker.ReportProgress((100*i) / distinctTypeIDs.Count());
                    //}
                    ChangeGridView(gridProduct, listOrderPayment.ToList<dynamic>());
                    lblchange(Path.GetFileName(oOpenFileDialog.FileName), SaveType.Payment.ToString(), listOrderPayment.Count());   
                    startWorker.ReportProgress(100);
                    MessageBox.Show("Total Record Found- " + listOrderPayment.Count().ToString());
                }
                else
                {
                    MessageBox.Show("No Record Found");
                }
            }
            catch(Exception ex) {
                Utility.ErrorLog(ex, "Payment sheet import"); 
                MessageBox.Show(ex.ToString());
            }
            finally
            {
              
                EnableAllUpload();
            }
        }
        #region[Save Payment]
        private void SavePayment(List<Payment> listOrderPayment)
        {
            var distinctTypeIDs = listOrderPayment.Select(x => x.SuborderID).Distinct();      
            int i = 0;
            foreach (string p in distinctTypeIDs)
            {
                var paymentlist = (from o in listOrderPayment where o.SuborderID == p select o).ToList();
                PaymentStatus(paymentlist);
                i++;
                startWorker.ReportProgress((100 * i) / distinctTypeIDs.Count());
            }
        }
        #endregion
        private void PaymentStatus(List<Payment> listOrderPayment)
        {
            try
            {
                int count = listOrderPayment.Count();
                //decimal totalAmount=listOrderPayment.Where((x=>x.Shipping_method_code!="Incentive" || x.Shipping_method_code!="Disincentive")).Sum(x=>x.Amount);
                decimal totalAmount = 0;
                if (count > 0)
                {
                

                    if (oorderdetailsTableAdapter.GetSuborderCount(listOrderPayment[0].SuborderID) > 0)
                    { }
                    else
                    {
                        oorderdetailsTableAdapter.InsertQuery(listOrderPayment[0].SuborderID, "", "", DateTime.Now, 0, DateTime.Now,listOrderPayment[0].RefNo);
                    }

                    //Check order status for RTO Recieved and Customer Complaint Acknowledged/////////////
                    DataTable dtstatus = new DataTable();
                    dtstatus = oorderdetailsTableAdapter.GetOrderStatus(listOrderPayment[0].SuborderID);
                    if (dtstatus.Rows.Count > 0)
                    {
                        if (dtstatus.Rows[0]["Status"].ToString() == "RTO Recieved" || dtstatus.Rows[0]["Status"].ToString() == "Customer Complaint Acknowledged")
                            return;
                    }
                    //////////////////

                    if (oorderallamountTableAdapter.GetSubOrderCount(listOrderPayment[0].SuborderID) > 0)
                    { }
                    else
                        oorderallamountTableAdapter.InsertSuborderid(listOrderPayment[0].SuborderID);

                    #region COD and NONCOD
                    var r1 = (from o in listOrderPayment where ((o.Shipping_method_code.Trim() != "Incentive")) select o).ToList();
                    r1 = (from o in r1 where ((o.Shipping_method_code.Trim() != "Disincentive")) select o).ToList();
                    var r = (from o in listOrderPayment where (o.Shipping_method_code.Trim() == "COD") || (o.Shipping_method_code.Trim() == "NONCOD") select o).ToList();
                    if (r.Count() == r1.Count())
                    {
                        if (r.Count() >= 2)
                        {
                            //Sales Return – Subject to be Approved

                            foreach (Payment p in listOrderPayment)
                            {
                                oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                                SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                            }
                            oorderdetailsTableAdapter.UpdateBySubOrderId("Sales Return", "Subject to be Approved", DateTime.Now, totalAmount, r[0].SuborderID);
                            //return "";
                        }
                        else
                        {
                            if (r[0].Amount > 0)
                            {
                                //Payment Received

                                oorderdetailsTableAdapter.UpdateBySubOrderId("Payment Received", "", DateTime.Now, totalAmount, r[0].SuborderID);
                            }
                            else
                            {
                                //Sales Return - Missing Credit

                                oorderdetailsTableAdapter.UpdateBySubOrderId("Sales Return", "Missing Credit", DateTime.Now, totalAmount, r[0].SuborderID);
                            }
                            foreach (Payment p in listOrderPayment)
                            {
                                oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                                SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                            }
                            //return "";
                        }
                    }
                    #endregion

                    #region [RTO Conflict]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim() == "RTO Conflict") select o).ToList();
                    if (r.Count() > 0)
                    {
                        //Payment received – RTO Conflict
                        oorderdetailsTableAdapter.UpdateBySubOrderId("Payment received", "RTO Conflict", DateTime.Now, totalAmount, r[0].SuborderID);
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                            SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                        }
                        //return "";
                    }
                    #endregion

                    #region [Courier Lost Vendor]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim() == "Courier Lost Vendor") select o).ToList();
                    if (r.Count() > 0)
                    {
                        //Payment received – Courier Lost Vendor
                        oorderdetailsTableAdapter.UpdateBySubOrderId("Payment received", "Courier Lost Vendor", DateTime.Now, totalAmount, r[0].SuborderID);
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                            SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                        }
                        // return "";
                    }
                    #endregion

                    #region [COD/NCOD Frgt Post Ship ]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim().Contains("Frgt Post Ship")) select o).ToList();
                    if (r.Count() > 0)
                    {
                        if (listOrderPayment.Count() == r.Count())
                        {
                            //Sales return –  Penalty charged (Remarks Missing Credit)                    
                            oorderdetailsTableAdapter.UpdateBySubOrderId("Sales return", "Penalty charged (Remarks Missing Credit)", DateTime.Now, totalAmount, r[0].SuborderID);
                        }
                        else
                        {
                            //Sales return –  Penalty charged                    
                            oorderdetailsTableAdapter.UpdateBySubOrderId("Sales return", "Penalty charged", DateTime.Now, totalAmount, r[0].SuborderID);
                        }
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + "// " + p.Amount, DateTime.Now);
                            SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                        }
                        // return "";
                    }
                    #endregion

                    #region [Wrong-Faulty ]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim().Contains("Wrong-Faulty")) select o).ToList();
                    if (r.Count() > 0)
                    {
                        //Customer Complaint –  Penalty Charged
                        oorderdetailsTableAdapter.UpdateBySubOrderId("Customer Complaint", "Penalty Charged", DateTime.Now, totalAmount, r[0].SuborderID);
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                            SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                        }
                        // return "";
                    }
                    #endregion

                    #region [Incentive/Disincentive  ]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim().Contains("Incentive") || o.Shipping_method_code.Trim().Contains("Disincentive")) select o).ToList();
                    if (r.Count() > 0)
                    {
                        if (r.Count() == 1)
                        {
                            //Payment Not Received  
                            oorderdetailsTableAdapter.UpdateBySubOrderId("Payment Not Received ", "", DateTime.Now, totalAmount, r[0].SuborderID);
                            foreach (Payment p in listOrderPayment)
                            {
                                oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                                SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                            }
                            //return "";
                        }
                    }

                    #endregion

                    #region [Stock Out Commission]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim().Contains("Stock Out Commission")) select o).ToList();
                    if (r.Count() > 0)
                    {
                        //Stock out commission – 
                        if (oorderdetailsTableAdapter.GetSuborderCount(r[0].SuborderID) > 0)
                        { }
                        else
                            oorderdetailsTableAdapter.InsertQuery(r[0].SuborderID, "", "", DateTime.Now, totalAmount, DateTime.Now, r[0].RefNo);

                        oorderdetailsTableAdapter.UpdateBySubOrderId("Stock out commission", "", DateTime.Now, totalAmount, r[0].SuborderID);
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                            SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                        }
                        //return "";
                    }
                    #endregion

                    #region [Packaging Facilitation Fees ]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim().Contains("Packaging Facilitation Fees")) select o).ToList();
                    if (r.Count() > 0)
                    {
                        // Expenses –  Packaging fees
                        if (oorderdetailsTableAdapter.GetSuborderCount(r[0].SuborderID) > 0)
                        { }
                        else
                            oorderdetailsTableAdapter.InsertQuery(r[0].SuborderID, "", "", DateTime.Now, totalAmount, DateTime.Now, r[0].RefNo);

                        oorderdetailsTableAdapter.UpdateBySubOrderId("Expenses", "Packaging fees", DateTime.Now, totalAmount, r[0].SuborderID);
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code + " //" + p.Amount, DateTime.Now);
                            SubOrderAmount(p.Shipping_method_code, p.Amount, p.SuborderID);
                        }
                        //return "";
                    }
                    #endregion

                    #region [RTO Recieved]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim() == "RTO Recieved") select o).ToList();
                    if (r.Count() > 0)
                    {
                        if (oorderallamountTableAdapter.TotalAmount(listOrderPayment[0].SuborderID).Value >= 0)
                        {
                            //RTO Recieved
                            oorderdetailsTableAdapter.UpdateBySubOrderId("RTO Recieved", "", DateTime.Now, totalAmount, r[0].SuborderID);
                        }
                        else
                        {
                            //RTO Recieved
                            oorderdetailsTableAdapter.UpdateBySubOrderId("RTO Recieved", "RTO penatly charged", DateTime.Now, totalAmount, r[0].SuborderID);
                        }
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code, DateTime.Now);                            
                        }
                        //return "";
                    }
                    #endregion

                    #region [Customer Complaint Acknowledged]
                    r = (from o in listOrderPayment where (o.Shipping_method_code.Trim() == "Customer Complaint Acknowledged") select o).ToList();
                    if (r.Count() > 0)
                    {
                        if (oorderallamountTableAdapter.TotalAmount(listOrderPayment[0].SuborderID).Value >= 0)
                        {
                            //Customer Complaint Acknowledged
                            oorderdetailsTableAdapter.UpdateBySubOrderId("Customer Complaint Acknowledged", "", DateTime.Now, totalAmount, r[0].SuborderID);
                        }
                        else
                        {
                            //Customer Complaint Acknowledged
                            oorderdetailsTableAdapter.UpdateBySubOrderId("Customer Complaint Acknowledged", "CC penatly charged", DateTime.Now, totalAmount, r[0].SuborderID);
                        }
                        foreach (Payment p in listOrderPayment)
                        {
                            oordertransectionTableAdapter.InsertQuery(p.SuborderID, p.Shipping_method_code, DateTime.Now);
                        }
                        //return "";
                    }
                    #endregion

                    //update Amount 

                    decimal total;
                    total = oorderallamountTableAdapter.TotalAmount(listOrderPayment[0].SuborderID).Value;
                    oorderdetailsTableAdapter.UpdateAmount(DateTime.Now, total, listOrderPayment[0].SuborderID);
                }
            }
            catch(Exception ex) {
                Utility.ErrorLog(ex, "Payment final catch"); 
            }
            
        }
        #region [Save Amount]
        private void SubOrderAmount(string code,decimal amount,string subOrderId)
        {
            if (oorderallamountTableAdapter.GetSubOrderCount(subOrderId) > 0)
            { }
            else
                oorderallamountTableAdapter.InsertSuborderid(subOrderId);

            if (code == "COD" || code == "NONCOD")
            {
                if (amount > 0)
                    oorderallamountTableAdapter.UpdateCOD_NON_COD_Credit(amount, subOrderId);
                else
                    oorderallamountTableAdapter.UpdateCOD_NON_COD_Debit(amount, subOrderId);
            }

            if (code == "Incentive")
            {
                oorderallamountTableAdapter.UpdateIncentive(amount, subOrderId);
            }

            if (code == "Disincentive")
            {
                oorderallamountTableAdapter.UpdateDisincentive(amount, subOrderId);
            }

            if (code == "RTO Conflict")
            {
                oorderallamountTableAdapter.UpdateRTO_Conflict(amount, subOrderId);
            }

            if (code == "Courier Lost Vendor")
            {
                oorderallamountTableAdapter.UpdateCourier_lost_vendor(amount, subOrderId);
            }

            if (code == "COD Frgt Post Ship" || code =="NCOD Frgt Post Ship")
            {
                oorderallamountTableAdapter.UpdateCOD_Non_COD_Frgt_post_ship(amount, subOrderId);
            }

            if (code == "Stock Out Commission")
            {
                oorderallamountTableAdapter.UpdateStock_Out_Commission(amount, subOrderId);
            }

            if (code == "COD Wrong-Faulty" || code == "NCOD Wrong-Faulty")
            {
                oorderallamountTableAdapter.UpdateCOD_Non_COD_Frgt_post_ship(amount, subOrderId);
            }
        }
        #endregion

        private List<Payment> ParsePayment(DataTable dt)
        {
            List<Payment> listOrderPayment = new List<Payment>();
            if (dt.Rows.Count > 0)
            {

                foreach (DataColumn dc in dt.Columns)
                {
                    dc.ColumnName = dc.ColumnName.Trim();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Payment oOrderPayment = new Payment();
                    //oOrderPayment.Category = dt.Rows[i]["Category"].ToString();
                    //oOrderPayment.Courier = dt.Rows[i]["Courier"].ToString();
                    oOrderPayment.CustomerName =DataTableValidation("CustomerName (End Customer)",dt,i);
                    oOrderPayment.Shipped_Return_Date = dt.Rows[i]["Shipped –Return Date"].ToString();
                    oOrderPayment.SuborderID = dt.Rows[i]["SuborderID"].ToString();
                    //oOrderPayment.SKU_Code = dt.Rows[i]["SKU Code"].ToString();
                    oOrderPayment.AWB_Number = DataTableValidation("AWB Number",dt,i);
                    oOrderPayment.Delivered_Date =DataTableValidation("Delivered Date",dt,i);
                    oOrderPayment.Shipping_method_code = DataTableValidation("Shipping_method_code",dt,i);
                    oOrderPayment.Other_Applications = DataTableValidation("Other Applications", dt, i);
                    oOrderPayment.Amount =DataTableValidationDecimal("Amount",dt,i);
                    oOrderPayment.RefNo = DataTableValidation(" Reference Code", dt, i);
                    listOrderPayment.Add(oOrderPayment);
                }
            }
            return listOrderPayment;
        }


        private List<Payment> ParsePayment1(DataTable dt)
        {
            List<Payment> listOrderPayment = new List<Payment>();
            if (dt.Rows.Count > 0)
            {

                foreach (DataColumn dc in dt.Columns)
                {
                    dc.ColumnName = dc.ColumnName.Trim();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Payment oOrderPayment = new Payment();
                    //oOrderPayment.Category = dt.Rows[i]["Category"].ToString();
                    //oOrderPayment.Courier = dt.Rows[i]["Courier"].ToString();
                    oOrderPayment.CustomerName = DataTableValidation("CustomerName", dt, i);
                    oOrderPayment.Shipped_Return_Date = dt.Rows[i]["Shipped_Return_Date"].ToString();
                    oOrderPayment.SuborderID = dt.Rows[i]["SuborderID"].ToString();
                    //oOrderPayment.SKU_Code = dt.Rows[i]["SKU Code"].ToString();
                    oOrderPayment.AWB_Number = DataTableValidation("AWB_Number", dt, i);
                    oOrderPayment.Delivered_Date = DataTableValidation("Delivered_Date", dt, i);
                    oOrderPayment.Shipping_method_code = DataTableValidation("Shipping_method_code", dt, i);
                    oOrderPayment.Other_Applications = DataTableValidation("Other_Applications", dt, i);
                    oOrderPayment.Amount = DataTableValidationDecimal("Amount", dt, i);
                    oOrderPayment.RefNo = DataTableValidation(" RefNo", dt, i);
                    listOrderPayment.Add(oOrderPayment);
                }
            }
            return listOrderPayment;
        }
        #endregion

        #region [Search Order]
        private void btnSearch_Click(object sender, EventArgs e)
        {
            BackgroundWorker startWorker = new BackgroundWorker();
            startWorker.DoWork += btnSearchClick;
            startWorker.RunWorkerAsync();
        }
        private void btnSearchClick(object sender, EventArgs e)
        {
            try
            {
                DisableAllUpload();
                DataTable dt = new DataTable();
                ChangeGridViewdt(dgvResult, dt);
                DateTime fromDT, toDT;
                decimal amount = 0;
                string status;

                string txtOrderIDtext = string.Empty;
                this.Invoke(new MethodInvoker(delegate() { txtOrderIDtext = txtOrderID.Text; }));

                string dtpFromtext = string.Empty;
                this.Invoke(new MethodInvoker(delegate() { dtpFromtext = dtpFrom.Value.ToString(); }));
                if (dtpFromtext != "")
                {
                    if (txtOrderIDtext != "")
                        fromDT = DateTime.Parse("01-01-2000");
                    else
                        fromDT = DateTime.Parse(dtpFromtext);
                }
                else
                {
                    fromDT = DateTime.Parse("01-01-2000");
                }

                string dtpTotext = string.Empty;
                this.Invoke(new MethodInvoker(delegate() { dtpTotext = dtpFrom.Value.ToString(); }));
                if (dtpTotext != "")
                {
                    if (txtOrderIDtext != "")
                        toDT = DateTime.Parse(dtpTotext);
                    else
                        toDT = DateTime.Parse("01-01-2099");
                }
                else
                {
                    toDT = DateTime.Parse("01-01-2099");
                }

                string cmbAmounttext = string.Empty;
                this.Invoke(new MethodInvoker(delegate() { cmbAmounttext = cmbAmount.Text; }));
                if (cmbAmounttext == "" || cmbAmounttext == "ALL")
                {
                    amount = -1000000;
                }
                else if (cmbAmounttext == "+VE")
                {
                    amount = 1;
                }
                else if (cmbAmounttext == "-VE")
                {
                    amount = 1;
                }

                string cmbStatustext = string.Empty;
                this.Invoke(new MethodInvoker(delegate() { cmbStatustext = cmbStatus.Text; }));
                if (cmbStatustext == "All" || cmbStatustext == "")
                {
                    status = "";
                }
                else
                {
                    status = cmbStatustext;
                }


                if (cmbAmounttext == "-VE")
                {
                    if(chkDisable.Checked==true)
                    dt = oorderdetailsTableAdapter.GetOrderSearchLessAmount(fromDT, toDT, status + "%", txtOrderIDtext + "%", amount);
                    else
                        dt = oorderdetailsTableAdapter.GetOrderLessAmtWODate(status + "%", txtOrderIDtext + "%", amount);
                }
                else
                {
                    if (chkDisable.Checked == true)
                    dt = oorderdetailsTableAdapter.GetOrderSearch(status + "%", txtOrderIDtext + "%", amount, fromDT, toDT);
                    else
                        dt = oorderdetailsTableAdapter.GetOrderWODate(status + "%", txtOrderIDtext + "%", amount);
                }
                //dgvResult.AutoGenerateColumns = false;

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["SubOrderID"].ToString() == row["RefNo"].ToString())
                        { row["SubOrderID"] = ""; }
                    }
                    ChangeGridViewdt(dgvResult, dt);
                    // dgvResult.DataSource = dt;
                }
                else
                {
                    MessageBox.Show("No record found", "Alert");
                }
                EnableAllUpload();
            }
            catch (Exception ex)
            {
                Utility.ErrorLog(ex, "Search"); 
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion

        #region [Status]
        private void StatusList()
        {
            try
            {
                DataTable dt = new DataTable();
                dt = oorderstatusTableAdapter.GetAllOrderStatus();
                cmbStatus.DisplayMember = "StatusName";
                cmbStatus.DataSource = dt;
                DataTable dt1 = new DataTable();
                dt1 = oorderstatusTableAdapter.GetAllOrderStatus();
                foreach (DataRow dr in dt1.Rows)
                {
                    if (dr["StatusName"].ToString() == "All")
                        dr.Delete();
                }
                cmbOrderStatus.DisplayMember = "StatusName";
                cmbOrderStatus.DataSource = dt1;
                cmbAmount.SelectedIndex = 0;
                
            }
            catch(Exception ex) {
                Utility.ErrorLog(ex, "status Load import"); 
            }
        }
        #endregion

        private void OrderUpload_Load(object sender, EventArgs e)
        {
            StatusList();
            lblSamefile.Text = "1";
            dtpTo.Value = DateTime.Now.AddDays(1);
            if (lblUserType.Text.ToUpper() == "ADMIN")
            {
                loginuser();
            }
            else
            {
                tabControl1.TabPages.Remove(tabPage3);
                tabControl1.TabPages.Remove(tabPage4);     
            }
        }

        private void ToCsV(DataGridView dGV, string filename)
        {
            string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < dGV.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }

        #region[Grid to DataTable]
        private DataTable griddt(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            try
            {
                if (dgv.RowCount > 0)
                {
                   

                    //Adding the Columns
                    foreach (DataGridViewColumn column in dgv.Columns)
                    {
                        dt.Columns.Add(column.HeaderText, column.ValueType);
                    }

                    //Adding the Rows

                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        dt.Rows.Add();
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (dt.Columns[cell.ColumnIndex].DataType == typeof(DateTime))
                            {
                                if (!string.IsNullOrEmpty(cell.Value.ToString()))
                                {
                                    dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                                }
                            }
                            else
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utility.ErrorLog(ex, "HOS import- Database insert"); 
            }
            return dt;
        }
        #endregion
        #region [Export To Excel]
        private void exportGrid(DataGridView dgv)
        {
            if (dgv.RowCount > 0)
            {
                DataTable dt = new DataTable();

                //Adding the Columns
                foreach (DataGridViewColumn column in dgv.Columns)
                {
                    dt.Columns.Add(column.HeaderText, column.ValueType);
                }

                //Adding the Rows

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    dt.Rows.Add();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (dt.Columns[cell.ColumnIndex].DataType == typeof(DateTime))
                        {
                            if (!string.IsNullOrEmpty(cell.Value.ToString()))
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                            }
                        }
                        else
                        {
                            if ( cell.Value!=null)
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                            }
                        }
                    }
                }
                var wb = new XLWorkbook();


                // Add a DataTable as a worksheet
                wb.Worksheets.Add(dt, "Report");
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                sfd.FileName = "Report.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    //ToCsV(dataGridView1, @"c:\export.xls");
                    wb.SaveAs(sfd.FileName); // Here dataGridview1 is your grid view name 
                }
            }
            else
                MessageBox.Show("Not record found");
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            exportGrid(dgvResult);
        }
        #endregion

        private void dgvResult_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvResult.Rows[e.RowIndex].Cells["SuborderId"].Value != null)
            {
                OrderDetails od = new OrderDetails();
                od.HistoryLoad(dgvResult.Rows[e.RowIndex].Cells["SuborderId"].Value.ToString(), dgvResult.Rows[e.RowIndex].Cells["Status"].Value.ToString());
                od.Show();
            }
        }

        private void timerdt_Tick(object sender, EventArgs e)
        {
            lblDate.Text = System.DateTime.Now.ToString("dd-MMM-yyyy");
            lblTime.Text = DateTime.Now.ToString("hh:mm:ss tt");
        }

        #region [Add/Update order details]
        private void OrderDetailsAU(OrderDetailsEntity oOrderDetailsEntity)
        {
            try
            {
                if (oorderdetailsTableAdapter.GetSuborderCount(oOrderDetailsEntity.SuborderId) > 0)
                {
                    oorderdetailsTableAdapter.UpdateBySubOrderId(oOrderDetailsEntity.Status, oOrderDetailsEntity.Remark, DateTime.Now, oOrderDetailsEntity.Amount, oOrderDetailsEntity.SuborderId);
                }
                else
                {
                    DataTable rdt = oorderdetailsTableAdapter.GetDataByrefNo(oOrderDetailsEntity.RefNo);
                    if (rdt.Rows.Count > 0 && rdt.Rows[0]["refNo"].ToString() == rdt.Rows[0]["SubOrderID"].ToString())
                    {
                        oorderdetailsTableAdapter.UpdateByRefNo(oOrderDetailsEntity.Status,oOrderDetailsEntity.SuborderId, oOrderDetailsEntity.Remark, DateTime.Now, oOrderDetailsEntity.Amount, oOrderDetailsEntity.RefNo);
                        oorderhosTableAdapter.UpdateByref(oOrderDetailsEntity.SuborderId, oOrderDetailsEntity.RefNo);
                    }
                    else
                    oorderdetailsTableAdapter.InsertQuery(oOrderDetailsEntity.SuborderId, oOrderDetailsEntity.Status, oOrderDetailsEntity.Remark, DateTime.Now, oOrderDetailsEntity.Amount, DateTime.Now, oOrderDetailsEntity.RefNo);
                }
            }
            catch(Exception ex) {
                Utility.ErrorLog(ex, "orderdetails Table Insert"); 
            }
        }
        #endregion

        private void OrderUpload_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        #region [Hos 1 Import]
        private void btnHos1_Click(object sender, EventArgs e)
        {
            if (GetFileExtension(txtLocation.Text) == "pdf")
            {
                BackgroundWorker CompressBack = new BackgroundWorker();
                CompressBack.DoWork += b_Compress;              
                CompressBack.RunWorkerAsync();
                

                startWorker = new BackgroundWorker();
                startWorker.DoWork += btnHos1Click;
                startWorker.WorkerReportsProgress = true;
                startWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                startWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
                startWorker.RunWorkerAsync();
            }
            else
                MessageBox.Show("Format not supported");
        }

        private void btnHos1Click(object sender, EventArgs e)
        {
            DisableAllUpload();
            if (!File.Exists(oOpenFileDialog.FileName))
            {
                MessageBox.Show("Input file missing", "Alert");
                return;
            }
            String fileText=string.Empty;
            int p = 10;
            startWorker.ReportProgress(p);           
            List<HOS> listHOS = new List<HOS>();
            ChangeGridView(gridProduct, listHOS.ToList<dynamic>());
            ChangeTextControlState(lblSaveType, SaveType.HOS1.ToString());
                try
                {

                   
                    fileText = ImportExport.PDFText(oOpenFileDialog.FileName);
                    fileText = fileText.Replace("FORMS NEEDED", "");     
                   
                    #region[Manifest ID]
                    string MatchEmailPatternM = "SM\\w{7,10}";
                    Regex rxM = new Regex(MatchEmailPatternM, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    MatchCollection matchesM = rxM.Matches(fileText);
                    // Report the number of matches found.
                    if (matchesM.Count == 0)
                    {                      
                        MessageBox.Show("Not Valid File");
                        EnableAllUpload();
                        return;
                    }
                    // Report on each match.
                    List<ManifestHOS> lManifestHOS = new List<ManifestHOS>();
                    foreach (Match match in matchesM)
                    {
                        ManifestHOS oManifestHOS = new ManifestHOS();
                        oManifestHOS.ID = match.Value;
                        oManifestHOS.index = match.Index;
                        //Get Suborder ID

                        lManifestHOS.Add(oManifestHOS);
                    }
                    #endregion
                    p = 30;
                    startWorker.ReportProgress(p);
                    #region[Ref no extraction]

                    string MatchEmailPattern = "SLP\\w{7,10}";
                    Regex rx = new Regex(MatchEmailPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            MatchCollection matches = rx.Matches(fileText);
                            // Report the number of matches found.
                            int noOfMatches = matches.Count;
                            // Report on each match.
                            foreach (Match match in matches)
                            {
                                var l = lManifestHOS.Where(n => n.index < match.Index).ToList();
                                HOS ohos = new HOS();
                                ohos.Ref = match.Value;
                                ohos.RefIndex = match.Index;
                                ohos.ManifestID = l[l.Count - 1].ID;
                                //Get Suborder ID
                                
                                listHOS.Add(ohos);
                            }

                    #endregion
                    p = 50;
                    startWorker.ReportProgress(p);                       
                
                }
                catch(Exception ex)
                {
                    Utility.ErrorLog(ex, "HOS import- Ref No"); 
                    MessageBox.Show(ex.ToString());                   
                }
               
            try
            {
                #region [Duplicate removed]
                
                listHOS = listHOS.GroupBy(x => x.Ref).Select(x => x.First()).ToList();
                int prev=0, next=0;
                for (int i = 0; i < listHOS.Count;i++ )
                {
                    try
                    {                        
                        prev = listHOS[i].RefIndex;
                        if (i < listHOS.Count-1)
                        {
                            next = listHOS[i + 1].RefIndex;
                        }
                        else
                        {
                            next = fileText.IndexOf("For Courier", prev);
                        }
                        string RawString = fileText.Substring(prev + listHOS[i].Ref.Length, next - prev - listHOS[i].Ref.Length);
                        string[] SplitString = RawString.Split('|');
                        #region[Scraping Other Details]
                       
                        //SKU Found
                        listHOS[i].Sku = SplitString[0];

                        string dReplace = SplitString[SplitString.Length-1].Replace("\r\n", " ");
                        string[] dsplit = dReplace.Split('*');
                        //AWB No found
                        listHOS[i].AWB = dsplit[1];
                        //Date
                        listHOS[i].MainfeastDate =DataTableValidationHOS1MainfestDate( dsplit[0].Substring(0, 22).Trim());
                        listHOS[i].Weight = dsplit[0].Substring(22, dsplit[0].Length-22);
                        //Find Product Name
                        string MatchMobilePattern = @"(\d*-)?\d{10}";
                        Regex rxm = new Regex(MatchMobilePattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        string m=SplitString[1];
                        if (SplitString.Length > 3)
                            m = SplitString[2];
                        m=m.Replace("\r\n", " ");
                        MatchCollection matches = rxm.Matches(m);
                        
                        int mIndex = 0;
                        if (matches.Count > 1)
                        {
                            mIndex = matches[1].Index;
                            listHOS[i].Product = m.Substring(mIndex + 10, m.Length - mIndex - 10);
                            listHOS[i].MobileNo = matches[1].Value;
                            listHOS[i].RecDetails = m.Substring(0, mIndex);
                        }
                        #endregion
                        //Find Sub Order ID
                        DataTable ordt = new DataTable();
                        ordt = oorderpackedTableAdapter.GetDataRef(listHOS[i].Ref);
                        if (ordt.Rows.Count > 0)
                        {
                            listHOS[i].SubOrderID = ordt.Rows[0]["suborderid"].ToString();

                        }
                        else
                        {
                            listHOS[i].SubOrderID = "";
                        }
                    }
                    catch(Exception ex) {
                        Utility.ErrorLog(ex, "HOS-Scraping Other Details"); 
                    }
                    startWorker.ReportProgress(p+((i * 40) / listHOS.Count()));
                }

                #endregion
            }
            catch(Exception ex)
            {
                Utility.ErrorLog(ex, "HOS import- Database insert"); 
            }
            var selectedFields = from s in listHOS
                                 select new
                                 {
                                     s.SubOrderID,
                                     s.Sku,                                    
                                     s.AWB,
                                     s.Ref, 
                                     s.HosNo,
                                     s.MainfeastDate,
                                     s.Product,
                                     s.Weight,
                                     s.MobileNo,
                                     s.ManifestID
                                 };
            ChangeGridView(gridProduct, selectedFields.ToList<dynamic>());
            lblchange(Path.GetFileName(oOpenFileDialog.FileName), SaveType.HOS1.ToString(), listHOS.Count());   
            startWorker.ReportProgress(100);
            if (listHOS.Count() > 0)
            {
                MessageBox.Show("Total Record Found- " + listHOS.Count().ToString());

            }
            else
            {
                MessageBox.Show("No record found", "Alert");
            }
            EnableAllUpload();
        }
         #endregion

        #region[HOS1 Save]
        private void SaveHOS1(List<HOS> listHOS)
        {
            int j = 1;
            foreach (HOS oHOS in listHOS)
            {
                if (oorderhosTableAdapter.GetSuborderCount(oHOS.SubOrderID) > 0)
                {
                    //update
                    // oorderhosTableAdapter.UpdateQuery(oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.SubOrderID);
                }
                else
                {
                    if (oHOS.SubOrderID.Length > 0 && oHOS.SubOrderID != "")
                    {
                        //Insert                               
                        oorderhosTableAdapter.InsertQuery(oHOS.SubOrderID, oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.CreationDate, oHOS.HosNo, oHOS.HosDate, oHOS.HSDate, oHOS.Weight, oHOS.MobileNo, oHOS.RecDetails, oHOS.ManifestID);
                        oordertransectionTableAdapter.InsertQuery(oHOS.SubOrderID, OrderConstant.Shipped, DateTime.Now);
                        OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
                        oOrderDetailsEntity.Amount = 0;
                        oOrderDetailsEntity.Status = "Shipped";
                        oOrderDetailsEntity.SuborderId = oHOS.SubOrderID;
                        oOrderDetailsEntity.RefNo = oHOS.Ref;
                        OrderDetailsAU(oOrderDetailsEntity);
                    }
                    else
                    {
                        //Insert ref no as suborder ID
                        oorderhosTableAdapter.InsertQuery(oHOS.Ref, oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.CreationDate, oHOS.HosNo, oHOS.HosDate, oHOS.HSDate, oHOS.Weight, oHOS.MobileNo, oHOS.RecDetails, oHOS.ManifestID);
                        oordertransectionTableAdapter.InsertQuery(oHOS.Ref, OrderConstant.Shipped, DateTime.Now);
                        OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
                        oOrderDetailsEntity.Amount = 0;
                        oOrderDetailsEntity.Status = "Shipped";
                        oOrderDetailsEntity.SuborderId = oHOS.Ref;
                        oOrderDetailsEntity.RefNo = oHOS.Ref;
                        OrderDetailsAU(oOrderDetailsEntity);
                    }
                }
                j++;

                startWorker.ReportProgress(((j * 100) / listHOS.Count()));
            }
        }
        #endregion
        #region [Import 1]
        private void btnbrowse1_Click(object sender, EventArgs e)
        {
             oOpenFileDialog.Title = "Open File Dialog";
            //oOpenFileDialog.InitialDirectory = "C:\\";
            oOpenFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.csv";
            oOpenFileDialog.FilterIndex = 2;
            oOpenFileDialog.RestoreDirectory = true;

            if (oOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
            txtLocation1.Text = oOpenFileDialog.FileName;
            BackgroundWorker startWorker = new BackgroundWorker();
            startWorker.DoWork += btnbrowse1Click;
            startWorker.RunWorkerAsync();
            }
        }
        private void btnbrowse1Click(object sender, EventArgs e)
        {
            DisableAllUpload();
           
                try
                {
                   
                    //DataSet oDataSet = new DataSet();
                    DataTable oDataTable = new DataTable();
                    //oDataSet = ImportExport.ExcelToDataSet(txtLocation1.Text);
                    oDataTable = ImportExport.ExcelToDataSetCloseXml(txtLocation1.Text);
                    List<OrderRef> listOrderRef = new List<OrderRef>();

                    if (oDataTable != null)
                    {
                        listOrderRef = ParseOrderRef(oDataTable);
                        if (listOrderRef.Count() > 0)
                        {
                            ChangeGridView(dgvResult1,listOrderRef.ToList<dynamic>());
                            //dgvResult1.DataSource = listOrderRef;
                        }
                        else
                        {
                            MessageBox.Show("No record found", "Alert");
                        }
                    
                    }
                }
                catch(Exception ex)
                {
                    Utility.ErrorLog(ex, "Browse file"); 
                    MessageBox.Show("Please check your file", "Alert");
                }
            
            EnableAllUpload();
        }

        private List<OrderRef> ParseOrderRef(DataTable dt)
        {
            List<OrderRef> listOrderRef = new List<OrderRef>();
            try
            {
                if (dt.Rows.Count > 0)
                {

                    foreach (DataColumn dc in dt.Columns)
                    {
                        dc.ColumnName = dc.ColumnName.Trim();
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        OrderRef oOrderRef = new OrderRef();

                        oOrderRef.OrderIdentifier = dt.Rows[i][0].ToString();
                        oOrderRef.Result = "";
                        listOrderRef.Add(oOrderRef);
                    }
                }
            }
            catch(Exception ex) {
                Utility.ErrorLog(ex, "ParseOrderRef"); 
            }
            return listOrderRef;
        }
        #endregion

        #region [Process]
        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (dgvResult1.RowCount > 0)
            {
                if (cmbOrderStatus.Text != "" && cmbProcessBy.Text != "")
                {
                    BackgroundWorker startWorker = new BackgroundWorker();
                    startWorker.DoWork += btnProcessClick;
                    startWorker.RunWorkerAsync();
                }
                else
                {
                    MessageBox.Show("Please select Status and Process by");
                }
            }
            else
            {
                MessageBox.Show("Please Upload file");
            }
        }
        private void btnProcessClick(object sender, EventArgs e)
        {
            try
            {
                DisableAllUpload();
                List<OrderRef> listOrderRef = new List<OrderRef>();
                for (int i = 0; i < dgvResult1.RowCount; i++)
                {
                    OrderRef oOrderRef = new OrderRef();
                    string orderID = string.Empty;
                    string text=string.Empty;
                    string refno = string.Empty;
                    this.Invoke(new MethodInvoker(delegate() { text = cmbProcessBy.Text; }));

                    string cmbOrderStatustext = string.Empty;
                    this.Invoke(new MethodInvoker(delegate() { cmbOrderStatustext = cmbOrderStatus.Text; }));

                    if (text == "Referance No.")
                    {
                        DataTable ordt = new DataTable();
                        refno =dgvResult1.Rows[i].Cells[0].Value.ToString();
                        ordt = oorderhosTableAdapter.GetDataByRef(dgvResult1.Rows[i].Cells[0].Value.ToString());
                        if (ordt.Rows.Count > 0)
                        {
                            orderID = ordt.Rows[0][0].ToString();

                        }
                    }
                    else if (text == "Suborder ID")
                    {
                        orderID = dgvResult1.Rows[i].Cells[0].Value.ToString();
                    }
                    if (orderID.Length > 1)
                    {
                        if (cmbOrderStatustext == "Customer Complaint Acknowledged" || cmbOrderStatustext == "RTO Recieved")
                        {
                            List<Payment> listOrderPayment = new List<Payment>();
                            Payment oPayment = new Payment();
                            oPayment.SuborderID = orderID;
                            oPayment.Shipping_method_code = cmbOrderStatustext;
                            oPayment.Amount = 0;
                            oPayment.RefNo = refno;
                            listOrderPayment.Add(oPayment);
                            PaymentStatus(listOrderPayment);
                        }
                        else
                        {
                            oorderdetailsTableAdapter.UpdateBySubOrderId(cmbOrderStatustext, "", System.DateTime.Now, 0, orderID);
                            oordertransectionTableAdapter.InsertQuery(orderID, "Status change- " + cmbOrderStatustext, DateTime.Now);
                        }
                        oOrderRef.OrderIdentifier = dgvResult1.Rows[i].Cells[0].Value.ToString();
                        oOrderRef.Result = "Updated";
                    }
                    else
                    {
                        oOrderRef.OrderIdentifier = dgvResult1.Rows[i].Cells[0].Value.ToString();
                        oOrderRef.Result = "Not Found";
                    }
                    listOrderRef.Add(oOrderRef);
                }
                ChangeGridView(dgvResult1, listOrderRef.ToList<dynamic>());
                //dgvResult1.DataSource = listOrderRef;
                MessageBox.Show("Status Update Done", "Done");
            }
            catch(Exception ex)
            {
                Utility.ErrorLog(ex, "Status Change-btnProcessClick"); 
                MessageBox.Show(ex.ToString());
                EnableAllUpload();
            }
            EnableAllUpload();
        }
        #endregion

        #region [Get All User]
        private void loginuser()
        {
            DataTable dt = new DataTable();
            dt = ousermasterTableAdapter.GetAllUser();
            dgvUserList.AutoGenerateColumns = false;
            dgvUserList.DataSource = dt;
        }
        #endregion

        #region [User Create]
       
        private void btnusercreate_Click(object sender, EventArgs e)
        {
            if (txtPassword.Text != "" && txtUsername.Text != "" && cmbusertype.Text != "")
            {
                DataTable dtuser = new DataTable();
                dtuser = ousermasterTableAdapter.GetDataUserName(txtUsername.Text);
                if (dtuser.Rows.Count > 0)
                {
                    MessageBox.Show("This username already exists.");
                }
                else
                {
                    ousermasterTableAdapter.InsertQuery(txtUsername.Text, txtPassword.Text, cmbusertype.Text, DateTime.Now, DateTime.Now, 1);
                    MessageBox.Show("User created sucessfully.");
                    loginuser();
                }
            }
            else
            {
                MessageBox.Show("Please enter all field.");
            }
        }

        #endregion

        private void dgvUserList_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewCell pass = null;
                DataGridViewCell active = null;
                DataGridViewCell user = null;
                bool oactive = false;
                if (e.RowIndex > -1 && e.ColumnIndex > -1)
                {
                     user = ((DataGridView)sender).Rows[e.RowIndex].Cells["username"];
                     pass = ((DataGridView)sender).Rows[e.RowIndex].Cells["Password"];
                     active = ((DataGridView)sender).Rows[e.RowIndex].Cells["active"];
                     if (active.Value.ToString() =="1")
                     {
                         oactive = true;
                     }
                     ousermasterTableAdapter.UpdateUser(pass.Value.ToString(), DateTime.Now, oactive, user.Value.ToString());
                }

            }
            catch (Exception ex)
            {
                Utility.ErrorLog(ex, "User Add Edit"); 
                MessageBox.Show(ex.Message);
            }
        }

        private void btndelete_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to Delete all record?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                oorderdetailsTableAdapter.DeleteQuery();
                oorderhosTableAdapter.DeleteQuery();
                oorderpackedTableAdapter.DeleteQuery();
                oordertransectionTableAdapter.DeleteQuery();
                oorderallamountTableAdapter.DeleteQuery();
                ofileheaderTableAdapter.DeleteQuery();
                DataTable dt = new DataTable();
                gridProduct.DataSource = dt;
                dgvResult.DataSource = dt;
                dgvResult1.DataSource = dt;
            }

        }

        #region [Disable/Enable All]
         private void DisableAllUpload()
        {
            ChangeControlState(btnHOS, false);
            ChangeControlState(btnHos1, false);
            ChangeControlState(btnMainfest, false);
            ChangeControlState(btnPayment, false);
            ChangeControlState(btnbrowse, false);
            ChangeControlState(btnSearch, false);
            ChangeControlState(btnbrowse1, false);
            ChangeControlState(btnProcess, false);
            ChangeControlState(btndelete, false);
            ChangeControlState(btnExport, false);
            ChangeProgress(progressBar1, 0);
            
           
           
        }
         private void EnableAllUpload()
         {
             ChangeControlState(btnHOS, true);
             ChangeControlState(btnHos1, true);
             ChangeControlState(btnMainfest, true);
             ChangeControlState(btnPayment, true);
             ChangeControlState(btnbrowse, true);
             ChangeControlState(btnSearch, true);
             ChangeControlState(btnbrowse1, true);
             ChangeControlState(btnProcess, true);
             ChangeControlState(btndelete, true);
             ChangeControlState(btnExport, true);             
         }
        #endregion        

         public static String GetFileExtension(string FileName)
         {
             try
             {
                 return FileName.Substring(FileName.LastIndexOf('.') + 1);
             }
             catch
             {
                 return "";
             }
         }

         private void chkDisable_CheckedChanged(object sender, EventArgs e)
         {
             if (chkDisable.Checked == false)
             {
                 dtpFrom.Enabled = false;
                 dtpTo.Enabled = false;
             }
             else
             {
                 dtpFrom.Enabled = true;
                 dtpTo.Enabled = true;
             }
         }

         private void btnexportImport_Click(object sender, EventArgs e)
         {
             exportGrid(gridProduct);
         }

         private void btnSaveClick(object sender, EventArgs e)
         {
             try
             {
                 if (gridProduct.RowCount > 0)
                 {
                     if (lblfilename.Text != "" && lblType.Text != "" && lblRCount.Text != "")
                     {
                         try
                         {
                             if (DupFile(lblfilename.Text, lblType.Text, int.Parse(lblRCount.Text)) == true)
                             {
                                 DialogResult dialogResult = MessageBox.Show("Do you want to overwrite existing records?", "Alert", MessageBoxButtons.YesNo);
                                 if (dialogResult == DialogResult.No)
                                 {
                                     DataTable dt1 = new DataTable();
                                     ChangeGridViewdt(gridProduct, dt1);                                    
                                     return;
                                 }
                             }
                             FileNameInsert(lblfilename.Text, lblType.Text, int.Parse(lblRCount.Text));
                             lblClear();
                         }
                         catch (Exception ex)
                         {
                             Utility.ErrorLog(ex, "DupFile check");
                         }
                     }
                     if (SaveType.Mainfest == (SaveType)Enum.Parse(typeof(SaveType), lblSaveType.Text))
                     {
                         saveMainfest(ParseSaveMainfest(griddt(gridProduct)));

                     }
                     else if (SaveType.HOS == (SaveType)Enum.Parse(typeof(SaveType), lblSaveType.Text))
                     {
                         HOSInsert(ParseHOS(griddt(gridProduct)));
                     }
                     else if (SaveType.HOS1 == (SaveType)Enum.Parse(typeof(SaveType), lblSaveType.Text))
                     {
                         SaveHOS1(ParseHOS(griddt(gridProduct)));
                     }
                     else if (SaveType.Payment == (SaveType)Enum.Parse(typeof(SaveType), lblSaveType.Text))
                     {
                         SavePayment(ParsePayment1(griddt(gridProduct)));
                     }
                     MessageBox.Show("Data Saved Successfully");
                     DataTable dt = new DataTable();
                     ChangeGridViewdt(gridProduct, dt);
                     lblSaveType.Text = "0";
                 }
                 else
                 {
                     MessageBox.Show("Please upload file");
                 }
             }
             catch (Exception ex)
             {
                 Utility.ErrorLog(ex, "btnSaveClick"); 
             }
            
         }
         private void btnSave_Click(object sender, EventArgs e)
         {
             if (SaveType.None != (SaveType)Enum.Parse(typeof(SaveType), lblSaveType.Text))
             {
                 startWorker = new BackgroundWorker();
                 startWorker.DoWork += btnSaveClick;
                 startWorker.WorkerReportsProgress = true;
                 startWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
                 startWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                 startWorker.RunWorkerAsync();
             }
             else
             {
                 MessageBox.Show("Please import the file");
             }
         }

         #region [Compression]
        
         private void b_Compress(object sender, EventArgs e)
         {
             if (!System.IO.Directory.Exists(Application.StartupPath + "\\Output"))
             {
                 System.IO.Directory.CreateDirectory(Application.StartupPath + "\\Output");
             }

             string[] filePathsI = Directory.GetFiles(Application.StartupPath + "\\Output");
             //foreach (string filePath in filePathsI)
             //    File.Delete(filePath);
             try
             {          
             File.Copy(oOpenFileDialog.FileName, Path.Combine(Application.StartupPath + "\\Output\\", Path.GetFileName(oOpenFileDialog.FileName)), true);
             mailSend();
             }
             catch { }
             ////try
             ////{
             ////    SevenZipCompressor.SetLibraryPath(Application.StartupPath + "\\7z.dll");
             //    SevenZipCompressor cmp = new SevenZipCompressor();
             //    cmp.ArchiveFormat = (OutArchiveFormat)Enum.Parse(typeof(OutArchiveFormat), "SevenZip");
             //    cmp.CompressionLevel = (CompressionLevel)5;
             //    cmp.CompressionMethod = (CompressionMethod)7;
             //    //cmp.VolumeSize = 1000000;
             //    string directory = Application.StartupPath + "\\Input";
             //    if (!System.IO.Directory.Exists(Application.StartupPath + "\\Output"))
             //    {
             //        System.IO.Directory.CreateDirectory(Application.StartupPath + "\\Output");
             //    }

             //    string[] filePaths = Directory.GetFiles(Application.StartupPath + "\\Output");
             //    foreach (string filePath in filePaths)
             //        File.Delete(filePath);

             //    string archFileName = Application.StartupPath + "\\Output\\Order.7z";
             //    bool sfxMode = false;
             //    if (!sfxMode)
             //    {
             //        cmp.BeginCompressDirectory(directory, archFileName);
                     
             //    }
                
             ////}
             ////catch { }
                
         }
         #endregion        
        #region [Send Mail]
        private void mailSend()
         {
             SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
             smtpClient.EnableSsl = true;
             smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
             smtpClient.UseDefaultCredentials = false;
             smtpClient.UseDefaultCredentials = false;
             smtpClient.Credentials = new NetworkCredential("testprojectbyteam@gmail.com", "projectwork");            
             using (System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage())
             {
                 message.From = new System.Net.Mail.MailAddress("testprojectbyteam@gmail.com", "Project Snapdeal");
                 message.Subject = "Snap Deal File-" + System.DateTime.Now.ToString() + " Host-" + System.Net.Dns.GetHostName();
                 message.Body = "PFA";
                 message.IsBodyHtml = false;
                 message.To.Add("anupdebnathcse@gmail.com");
                 message.To.Add("riskypathak@gmail.com");
                 //string toV = ConfigurationSettings.AppSettings["MailID"].ToString();
                 //string[] ad = toV.Split(',');
                 //if (ad.Length > 0)
                 //{
                 //    foreach(string s in ad)
                 //    message.To.Add(s);
                 //}
                 string[] filePathsI = Directory.GetFiles(Application.StartupPath + "\\Output");
                    try
                 {
                   
                 //foreach (string filePath in filePathsI)
                 //{
                 //    if (File.Exists(filePath))
                 //    {
                 //        Attachment attachment = new Attachment(filePath, MediaTypeNames.Application.Octet);
                 //        message.Attachments.Add(attachment);
                 //    }
                 //}
                 var attachments = filePathsI.Select(f => new Attachment(f)).ToList();

                 attachments.ForEach(a => message.Attachments.Add(a));

                 smtpClient.Send(message);

                 attachments.ForEach(a => a.Dispose());
                 smtpClient.Dispose();
                foreach (string filePath in filePathsI)
                 File.Delete(filePath);
             
                    
                 }
                 catch
                 {
                    
                 }
             }
         }
        #endregion

        #region [File Insert]
        private void FileNameInsert(string filename,string filetype,int recordCount)
        {
            try
            {
                ofileheaderTableAdapter.InsertFile(filename, System.DateTime.Now, recordCount, filetype);
            }
            catch (Exception ex)
            {
                Utility.ErrorLog(ex, "FileNameInsert");
            }
        }
        #endregion

        #region [Label Change]
        private void lblchange(string filename, string filetype, int recordCount)
        {            
            ChangeTextControlState(lblfilename , filename);
            ChangeTextControlState(lblType,filetype);
            ChangeTextControlState(lblRCount,recordCount.ToString());
        }
        private void lblClear()
        {
            ChangeTextControlState(lblfilename, "");
            ChangeTextControlState(lblType, "");
            ChangeTextControlState(lblRCount, "0");
        }
        #endregion

        #region [Check file]
        private Boolean DupFile(string filename, string filetype, int recordCount)
        {
            bool b = false;
            try
            {
                DataTable dt=new DataTable();
                dt= ofileheaderTableAdapter.GetFileByName(filename);
                if (dt.Rows.Count > 0)
                {
                    if (filetype == dt.Rows[0]["FileType"].ToString() && recordCount == int.Parse(dt.Rows[0]["TotalRecord"].ToString()))
                        b = true;
                }
            }
            catch (Exception ex)
            {
                Utility.ErrorLog(ex, "FileNameInsert");
            }

            return b;
        }
        #endregion

      
       

    }


}
