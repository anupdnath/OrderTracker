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

namespace OrderTracker
{
    public partial class OrderUpload : Form
    {
      
        #region [Global Variable]
         OpenFileDialog oOpenFileDialog = new OpenFileDialog();
        #endregion
        DataTable oDataTable = new DataTable();
        orderdetailsTableAdapter oorderdetailsTableAdapter = new orderdetailsTableAdapter();
        orderhosTableAdapter oorderhosTableAdapter = new orderhosTableAdapter();
        orderpackedTableAdapter oorderpackedTableAdapter = new orderpackedTableAdapter();
        ordertransectionTableAdapter oordertransectionTableAdapter = new ordertransectionTableAdapter();
        orderstatusTableAdapter oorderstatusTableAdapter = new orderstatusTableAdapter();
        orderallamountTableAdapter oorderallamountTableAdapter = new orderallamountTableAdapter();
        usermasterTableAdapter ousermasterTableAdapter = new usermasterTableAdapter();
        public OrderUpload()
        {
            InitializeComponent();
        }

        #region [Mainfest]
        private void btnMainfest_Click(object sender, EventArgs e)
        {
            BackgroundWorker startWorker = new BackgroundWorker();
            startWorker.DoWork += btnMainfestClick;
            startWorker.RunWorkerAsync();
        }       
        private void btnMainfestClick(object sender, EventArgs e)
        {

            try
            {
                DisableAllUpload();
                oDataTable = new DataTable();
                //gridProduct.DataSource = oDataTable;
                ChangeGridViewdt(gridProduct, oDataTable);
                oDataTable = ImportExport.CSVToDataTable(txtLocation.Text, true);
                List<OrderMainfest> listOrderMainfest = new List<OrderMainfest>();
                listOrderMainfest = ParseMainfest(oDataTable);
                if (listOrderMainfest.Count() > 0)
                {
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
                                OrderDetailsAU(oOrderDetailsEntity);
                                oordertransectionTableAdapter.InsertQuery(or.Suborder_Id, OrderConstant.Packed, DateTime.Now);
                            }
                        }
                        catch
                        {
                        }
                    }
                    ChangeGridView(gridProduct, listOrderMainfest.ToList<dynamic>());
                    //gridProduct.DataSource = listOrderMainfest;
                    if (listOrderMainfest.Count() > 0)
                    {
                        MessageBox.Show("Total Record Updated- " + listOrderMainfest.Count().ToString());
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
            }
            finally
            {
                EnableAllUpload();
            }
            //EnableAllUpload();
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
                string format = "dd-MM-yyyy HH:mm";
                DateTime dateTime;
                if (DateTime.TryParseExact(dt.Rows[i][col].ToString(), format, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dateTime))
                {
                    if (dt.Rows[i][col].ToString().Length > 10)
                        validdate = DateTime.ParseExact(dt.Rows[i][col].ToString(), "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture);
                    else
                        validdate = DateTime.ParseExact(dt.Rows[i][col].ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture);
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

        private delegate void ChangeGridViewDelegate(DataGridView lvItem, List<dynamic> listHos);
        private void ChangeGridView(DataGridView lvItem, List<dynamic> listHos)
        {
            if (this.InvokeRequired)
            {
                ChangeGridViewDelegate d = new ChangeGridViewDelegate(ChangeGridView);
                this.Invoke(d, lvItem, listHos);
            }
            else
                lvItem.DataSource = listHos;
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
        #endregion
        

        #region [HOS]
        private void btnHOS_Click(object sender, EventArgs e)
        {
            BackgroundWorker startWorker = new BackgroundWorker();           
            startWorker.DoWork += btnHOSClick;           
            startWorker.RunWorkerAsync();
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


            String destDir, tempFile = string.Empty;

            FileInfo srcFinfo = new FileInfo(oOpenFileDialog.FileName);
            int numberOfPages = new PdfReader(srcFinfo.FullName).NumberOfPages;
            try
            {
                destDir = String.Format("{0}\\Output", Application.StartupPath);
                tempFile = String.Format("{0}\\_temp.pdf", destDir);
                if (!Directory.Exists(destDir))
                    Directory.CreateDirectory(destDir);

                foreach (FileInfo destFinfo in new DirectoryInfo(destDir).GetFiles())
                    destFinfo.Delete();
            }
            catch { }

            int index;
            string hosCode = "", hosdate = "";
            List<HOS> listHOS = new List<HOS>();
            ChangeGridView(gridProduct, listHOS.ToList<dynamic>());
            bool status = false;
            for (int page = 1; page <= numberOfPages; page++)
            {
                try
                {

                    ExtractPage(srcFinfo.FullName, tempFile, page);
                    // Read Text from temp file
                    // String fileText = ExtractPageText(tempFile);
                    String fileText = ImportExport.PDFText(tempFile);

                    fileText=fileText.Replace("FORMS NEEDED", "");
                    string[] sentences = fileText.Split('\n');
                    foreach (string s in sentences)
                    {
                        if (page == 1)
                        {
                            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(s, "HOS\\w{7}");
                            if (match.Success)
                            {
                                hosCode = match.Captures[0].Value;
                            }
                            index = s.IndexOf("Handover Sheet Date:");
                            if (index > -1)
                            {
                                hosdate = s.Substring(index + 21, s.Length - (index + 21));
                            }
                        }
                        //else
                        //{
                        //    //rider copy started
                        //    System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(s, "HOS\\w{7}");
                        //    if (match.Success)
                        //    {
                        //        status = true;
                        //        break;
                        //    }

                        //}
                    }

                    //if (status == true)
                    //    break;

                    index = fileText.IndexOf("Reference Code");
                    int index1 = fileText.Length - 1;
                    string totalline = fileText.Substring(index + 15, index1 - index - 15);
                    string[] line = totalline.Split('\n');

                    foreach (string str in line)
                    {
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
                                    oHOS.Ref = orderdetails[5].Replace("\r","");
                                    oHOS.CreationDate = System.DateTime.Now;
                                    oHOS.HosNo = hosCode;
                                    oHOS.HosDate = hosdate;
                                    listHOS.Add(oHOS);
                                }
                            }
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Something going wrong", "Error");
                }
            }
            ////
            listHOS = listHOS.GroupBy(x => x.SubOrderID).Select(x => x.First()).ToList();
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
                    oorderhosTableAdapter.InsertQuery(oHOS.SubOrderID, oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.CreationDate, oHOS.HosNo, oHOS.HosDate);
                    oordertransectionTableAdapter.InsertQuery(oHOS.SubOrderID, OrderConstant.Shipped, DateTime.Now);
                    OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
                    oOrderDetailsEntity.Amount = 0;
                    oOrderDetailsEntity.Status = "Shipped";
                    oOrderDetailsEntity.SuborderId = oHOS.SubOrderID;
                    OrderDetailsAU(oOrderDetailsEntity);
                   
                }
                               
            }
            ChangeGridView(gridProduct, listHOS.ToList<dynamic>());
            if (listHOS.Count() > 0)
            {
                MessageBox.Show("Total Record Updated- " + listHOS.Count().ToString());

            }
            else
            {
                MessageBox.Show("No record found", "Alert");
            }
            EnableAllUpload();
        }

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
            BackgroundWorker startWorker = new BackgroundWorker();
            startWorker.DoWork += btnPaymentClick;
            startWorker.RunWorkerAsync();
        }
        private void btnPaymentClick(object sender, EventArgs e)
        {
            DisableAllUpload();
            try
            {
                DataSet oDataSet = new DataSet();
                oDataSet = ImportExport.ExcelToDataSet(txtLocation.Text);
                List<Payment> listOrderPayment = new List<Payment>();
                ChangeGridView(gridProduct, listOrderPayment.ToList<dynamic>());
                if (oDataSet != null)
                {
                    listOrderPayment = ParsePayment(oDataSet.Tables[0]);
                    var distinctTypeIDs = listOrderPayment.Select(x => x.SuborderID).Distinct();
                    foreach (string p in distinctTypeIDs)
                    {
                        var paymentlist = (from o in listOrderPayment where o.SuborderID == p select o).ToList();
                        PaymentStatus(paymentlist);
                    }
                    ChangeGridView(gridProduct, listOrderPayment.ToList<dynamic>());
                    MessageBox.Show("Total Record Updated- " + listOrderPayment.Count().ToString());
                }
                else
                {
                    MessageBox.Show("No Record Found");
                }
            }
            catch {
                MessageBox.Show("Please contact Admin");
            }
            finally
            {
              
                EnableAllUpload();
            }
        }

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
                        oorderdetailsTableAdapter.InsertQuery(listOrderPayment[0].SuborderID, "", "", DateTime.Now, 0, DateTime.Now);
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
                            oorderdetailsTableAdapter.InsertQuery(r[0].SuborderID, "", "", DateTime.Now, totalAmount, DateTime.Now);

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
                            oorderdetailsTableAdapter.InsertQuery(r[0].SuborderID, "", "", DateTime.Now, totalAmount, DateTime.Now);

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
            catch { }
            
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
            DisableAllUpload();
            DataTable dt = new DataTable();
           ChangeGridViewdt(dgvResult,dt);
            DateTime fromDT,toDT;
            decimal amount=0;
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
               fromDT=DateTime.Parse("01-01-2000"); 
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
               toDT=DateTime.Parse("01-01-2099"); 
            }

            string cmbAmounttext = string.Empty;
            this.Invoke(new MethodInvoker(delegate() { cmbAmounttext = cmbAmount.Text; }));
            if (cmbAmounttext == "" || cmbAmounttext == "ALL")
            {
                amount=-1000000;
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
                dt = oorderdetailsTableAdapter.GetOrderSearchLessAmount(fromDT, toDT, status + "%", txtOrderIDtext + "%", amount);
            }
            else
            {
                dt = oorderdetailsTableAdapter.GetOrderSearch(status + "%", txtOrderIDtext + "%", amount, fromDT, toDT);
            }
             //dgvResult.AutoGenerateColumns = false;
             
             if (dt.Rows.Count > 0)
             {
                 ChangeGridViewdt(dgvResult,dt);
                // dgvResult.DataSource = dt;
             }
             else
             {
                 MessageBox.Show("No record found", "Alert");
             }
             EnableAllUpload();
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
                
                
            }
            catch { }
        }
        #endregion

        private void OrderUpload_Load(object sender, EventArgs e)
        {
            StatusList();
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

        #region [Export To Excel]        
       
        private void btnExport_Click(object sender, EventArgs e)
        {
            ////SaveFileDialog sfd = new SaveFileDialog();
            ////sfd.Filter = "Excel Documents (*.xls)|*.xls";
            ////sfd.FileName = "export.xls";
            ////if (sfd.ShowDialog() == DialogResult.OK)
            ////{
            ////    //ToCsV(dataGridView1, @"c:\export.xls");
            ////    ToCsV(dgvResult, sfd.FileName); // Here dataGridview1 is your grid view name 
            ////}  
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgvResult.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgvResult.Rows)
            {
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
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
                    oorderdetailsTableAdapter.UpdateBySubOrderId(oOrderDetailsEntity.Status,oOrderDetailsEntity.Remark,DateTime.Now,oOrderDetailsEntity.Amount,oOrderDetailsEntity.SuborderId);
                }
                else
                {
                    oorderdetailsTableAdapter.InsertQuery(oOrderDetailsEntity.SuborderId, oOrderDetailsEntity.Status, oOrderDetailsEntity.Remark, DateTime.Now, oOrderDetailsEntity.Amount,DateTime.Now);
                }
            }
            catch { }
        }
        #endregion

        private void OrderUpload_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        #region [Hos 1 Import]
        private void btnHos1_Click(object sender, EventArgs e)
        {
            BackgroundWorker startWorker = new BackgroundWorker();
            startWorker.DoWork += btnHos1Click;
            startWorker.RunWorkerAsync();
        }

        private void btnHos1Click(object sender, EventArgs e)
        {
            DisableAllUpload();
            if (!File.Exists(oOpenFileDialog.FileName))
            {
                MessageBox.Show("Input file missing", "Alert");
                return;
            }


            String destDir, tempFile = string.Empty;

            FileInfo srcFinfo = new FileInfo(oOpenFileDialog.FileName);
            int numberOfPages = new PdfReader(srcFinfo.FullName).NumberOfPages;
            try
            {
                destDir = String.Format("{0}\\Output", Application.StartupPath);
                tempFile = String.Format("{0}\\_temp.pdf", destDir);

                foreach (FileInfo destFinfo in new DirectoryInfo(destDir).GetFiles())
                    destFinfo.Delete();
            }
            catch { }
            List<HOS> listHOS = new List<HOS>();
            ChangeGridView(gridProduct, listHOS.ToList<dynamic>());

            for (int page = 1; page <= numberOfPages; page++)
            {
                try
                {

                    ExtractPage(srcFinfo.FullName, tempFile, page);
                    // Read Text from temp file
                    String fileText = ExtractPageText(tempFile);
                    String fileText1 = ImportExport.PDFText(tempFile);

                    // Read Employee Code from temp file                   
                    string[] sentences = fileText.Split('\n');
                  
                    foreach (string s in sentences)
                    {
                        System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(s, "SLP\\w{9}");
                        if (match.Success)
                        {
                            HOS ohos = new HOS();
                            ohos.Ref = match.Captures[0].Value;
                            //Get Suborder ID
                            DataTable ordt = new DataTable();
                            ordt = oorderpackedTableAdapter.GetDataRef(ohos.Ref);
                            if (ordt.Rows.Count > 0)
                            {
                                ohos.SubOrderID = ordt.Rows[0]["suborderid"].ToString();

                            }
                            listHOS.Add(ohos);
                        }
                    }


                
                }
                catch
                {
                    MessageBox.Show("Something going wrong","Error");                   
                }
            }
            try
            {
                listHOS = listHOS.GroupBy(x => x.Ref).Select(x => x.First()).ToList();
               // var v = listHOS.SelectMany(mo => mo.SubOrderID).Distinct();

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
                        oorderhosTableAdapter.InsertQuery(oHOS.SubOrderID, oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.CreationDate, oHOS.HosNo, oHOS.HosDate);
                        oordertransectionTableAdapter.InsertQuery(oHOS.SubOrderID, OrderConstant.Shipped, DateTime.Now);
                        OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
                        oOrderDetailsEntity.Amount = 0;
                        oOrderDetailsEntity.Status = "Shipped";
                        oOrderDetailsEntity.SuborderId = oHOS.SubOrderID;
                        OrderDetailsAU(oOrderDetailsEntity);

                    }

                }

            }
            catch
            {
                
            }
            ChangeGridView(gridProduct, listHOS.ToList<dynamic>());
            if (listHOS.Count() > 0)
            {
                MessageBox.Show("Total Record Updated- " + listHOS.Count().ToString());

            }
            else
            {
                MessageBox.Show("No record found", "Alert");
            }
            EnableAllUpload();
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
                catch
                {
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
            catch { }
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
                    this.Invoke(new MethodInvoker(delegate() { text = cmbProcessBy.Text; }));

                    string cmbOrderStatustext = string.Empty;
                    this.Invoke(new MethodInvoker(delegate() { cmbOrderStatustext = cmbOrderStatus.Text; }));

                    if (text == "Referance No.")
                    {
                        DataTable ordt = new DataTable();
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
            catch
            {
                MessageBox.Show("Please contact Admin");
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

    }


}
