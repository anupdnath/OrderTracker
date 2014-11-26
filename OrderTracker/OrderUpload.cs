﻿using System;
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
        public OrderUpload()
        {
            InitializeComponent();
        }

        #region [Mainfest]
        
       
        private void btnMainfest_Click(object sender, EventArgs e)
        {
            try
            {
                oDataTable = new DataTable();
                oDataTable = ImportExport.CSVToDataTable(txtLocation.Text, true);
                List<OrderMainfest> listOrderMainfest = new List<OrderMainfest>();
                listOrderMainfest=ParseMainfest(oDataTable);
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
                    gridProduct.DataSource = listOrderMainfest;
                    if(listOrderMainfest.Count()>0)
                    {
                        MessageBox.Show("Total Record Updated- " + listOrderMainfest.Count().ToString());
                    }
                }
                else
                {
                    MessageBox.Show("No Record Found");
                }
                             
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
            if (dt.Columns.Contains(col))
            {
                if(dt.Rows[i][col].ToString().Length>10)
                return DateTime.ParseExact(dt.Rows[i][col].ToString(), "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture);
                else
                    return DateTime.ParseExact(dt.Rows[i][col].ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture);
            }
            else
            {
                return DateTime.ParseExact("01-01-1900", "dd-MM-yyyy", CultureInfo.InvariantCulture);
            }
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

        #region [HOS]
        private void btnHOS_Click(object sender, EventArgs e)
        {
            if (!File.Exists(oOpenFileDialog.FileName))
            {
                MessageBox.Show("Input file missing","Alert");                
                return;
            }
           
           
            String destDir, tempFile=string.Empty;

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

            int index;           
             for (int page = 1; page <= numberOfPages; page++)
            {
                try
                {
                    string hosCode="", hosdate="";
                    ExtractPage(srcFinfo.FullName, tempFile, page);
                       // Read Text from temp file
                    String fileText = ExtractPageText(tempFile);
                    String fileText1 = ImportExport.PDFText(tempFile);

                    // Read Employee Code from temp file                   
                    string[] sentences = fileText.Split('\n');
                    foreach (string s in sentences)
                    {
                        System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(s, "HOS\\w{7}");
                        if (match.Success)
                        {
                            hosCode=match.Captures[0].Value;                            
                        }
                        index = s.IndexOf("Handover Sheet Date:");
                        if (index >-1)
                        {
                            hosdate = s.Substring(index + 21, s.Length - (index + 21));
                        }
                    }

                    index = fileText.IndexOf("Reference Code");
                    int index1 = fileText.IndexOf("ABC PVT Ltd");
                    string totalline = fileText.Substring(index + 15, index1 - index - 15);
                    string[] line=totalline.Split('\n');
                    List<HOS> listHOS = new List<HOS>();
                    foreach (string str in line)
                    {
                        try
                        {
                            if (str.Length > 0)
                            {
                                string[] orderdetails = str.Split(' ');
                                HOS oHOS = new HOS();
                                oHOS.SubOrderID = orderdetails[1];
                                oHOS.Sku = orderdetails[2];
                                oHOS.Supc = orderdetails[3];
                                oHOS.AWB = orderdetails[4];
                                oHOS.Ref = orderdetails[5];
                                oHOS.CreationDate = System.DateTime.Now;
                                oHOS.HosNo=hosCode;
                                oHOS.HosDate=hosdate;
                                if (oorderhosTableAdapter.GetSuborderCount(orderdetails[1]) > 0)
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
                                listHOS.Add(oHOS);
                            }
                        }
                        catch
                        {
                        }
                    }
                    gridProduct.DataSource = listHOS;
                    if (listHOS.Count() > 0)
                    {
                        MessageBox.Show("Total Record Updated- " + listHOS.Count().ToString());
                     
                    }
                    else
                    {
                        MessageBox.Show("No record found", "Alert");
                    }
                }
                 catch{
                     //MessageBox.Show("Something going wrong","Error");
                 }
             }
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
            oOpenFileDialog.InitialDirectory = "C:\\";
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
            DataSet oDataSet = new DataSet();
            oDataSet = ImportExport.ExcelToDataSet(txtLocation.Text);
            List<Payment> listOrderPayment = new List<Payment>();
            if (oDataSet != null)
            {
                listOrderPayment = ParsePayment(oDataSet.Tables[0]);
                var distinctTypeIDs = listOrderPayment.Select(x => x.SuborderID).Distinct();
                foreach (string p in distinctTypeIDs)
                {
                    var paymentlist = (from o in listOrderPayment where o.SuborderID == p select o).ToList();
                    PaymentStatus(paymentlist);
                }
                gridProduct.DataSource = listOrderPayment;
                MessageBox.Show("Total Record Updated- " + listOrderPayment.Count().ToString());
            }
            else
            {
                MessageBox.Show("No Record Found");
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
            DataTable dt = new DataTable();
            DateTime fromDT,toDT;
            decimal amount=0;
            string status;
            if(dtpFrom.Value.ToString()!="")
            {
                fromDT=dtpFrom.Value.Date;
            }
            else
            {
               fromDT=DateTime.Parse("01-01-2000"); 
            }

             if(dtpTo.Value.ToString()!="")
            {
                toDT=dtpTo.Value.Date;
            }
            else
            {
               toDT=DateTime.Parse("01-01-2099"); 
            }
            if(cmbAmount.Text=="" || cmbAmount.Text=="ALL")
            {
                amount=-1000000;
            }
            else if(cmbAmount.Text=="+VE" )
            {
                amount = 1;
            }
            else if (cmbAmount.Text == "-VE")
            {
                amount = 1;
            }
            if (cmbStatus.SelectedText == "All" || cmbStatus.SelectedText == "")
            {
                status = "";
            }
            else
            {
                status = cmbStatus.SelectedText;
            }

            if (cmbAmount.Text == "-VE")
            {
                dt = oorderdetailsTableAdapter.GetOrderSearchLessAmount(fromDT, toDT, status + "%", txtOrderID.Text + "%", amount);
            }
            else
            {
                dt = oorderdetailsTableAdapter.GetOrderSearch(fromDT, toDT, status + "%", txtOrderID.Text + "%", amount);
            }
             //dgvResult.AutoGenerateColumns = false;
             dgvResult.DataSource = dt;
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
            }
            catch { }
        }
        #endregion

        private void OrderUpload_Load(object sender, EventArgs e)
        {
            StatusList();
            dtpTo.Value = DateTime.Now.AddDays(1);
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
            
            if (!File.Exists(oOpenFileDialog.FileName))
            {
                MessageBox.Show("Input file missing","Alert");                
                return;
            }
           
           
            String destDir, tempFile=string.Empty;

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
                    List<HOS> listHOS = new List<HOS>();
                    foreach (string s in sentences)
                    {
                        System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(s, "SLP\\w{9}");
                        if (match.Success)
                        {
                            HOS ohos = new HOS();
                            ohos.Ref = match.Captures[0].Value;
                            listHOS.Add(ohos);
                        }                        
                    }


                    try
                    {

                        //if (oorderhosTableAdapter.GetSuborderCount(orderdetails[1]) > 0)
                        //{
                        //    //update
                        //    // oorderhosTableAdapter.UpdateQuery(oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.SubOrderID);
                        //}
                        //else
                        //{
                        //    //Insert                               
                        //    oorderhosTableAdapter.InsertQuery(oHOS.SubOrderID, oHOS.Sku, oHOS.Supc, oHOS.AWB, oHOS.Ref, oHOS.CreationDate, oHOS.HosNo, oHOS.HosDate);
                        //    oordertransectionTableAdapter.InsertQuery(oHOS.SubOrderID, OrderConstant.Shipped, DateTime.Now);
                        //    OrderDetailsEntity oOrderDetailsEntity = new OrderDetailsEntity();
                        //    oOrderDetailsEntity.Amount = 0;
                        //    oOrderDetailsEntity.Status = "Shipped";
                        //    oOrderDetailsEntity.SuborderId = oHOS.SubOrderID;
                        //    OrderDetailsAU(oOrderDetailsEntity);
                        //}
                        

                    }
                    catch
                    {
                    }
                    
                    gridProduct.DataSource = listHOS;
                    if (listHOS.Count() > 0)
                    {
                        MessageBox.Show("Total Record Updated- " + listHOS.Count().ToString());

                    }
                    else
                    {
                        MessageBox.Show("No record found", "Alert");
                    }
                }
                catch
                {
                    //MessageBox.Show("Something going wrong","Error");
                }
            }
        }
         #endregion
    }
}
