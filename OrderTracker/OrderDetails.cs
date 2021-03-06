﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OrderTracker.OrderDBTableAdapters;

namespace OrderTracker
{
    public partial class OrderDetails : Form
    {
        ordertransectionTableAdapter oordertransectionTableAdapter = new ordertransectionTableAdapter();
        orderstatusTableAdapter oorderstatusTableAdapter = new orderstatusTableAdapter();
        orderdetailsTableAdapter oorderdetailsTableAdapter = new orderdetailsTableAdapter();
        public OrderDetails()
        {
            InitializeComponent();
        }
        public void HistoryLoad(string orderId, string status)
        {
            DataTable dt = new DataTable();
            dt = oordertransectionTableAdapter.GetTransBySuborderId(orderId);
            dgvResult.AutoGenerateColumns = false;
            dgvResult.DataSource = dt;
            lblOrderNo.Text = orderId;
            lblStatus.Text = status;
            StatusList(status);
        }
        #region [Status]
        private void StatusList(string status)
        {
            try
            {
                DataTable dt = new DataTable();
                dt =oorderstatusTableAdapter.GetAllOrderStatus();
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["StatusName"].ToString() == "All")
                        dr.Delete();
                }
                cmbStatus.DisplayMember = "StatusName";
                cmbStatus.DataSource = dt;
                
            }
            catch { }
        }
        #endregion
 
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            oorderdetailsTableAdapter.UpdateBySubOrderId(cmbStatus.Text, "", System.DateTime.Now, 0, lblOrderNo.Text);
            oordertransectionTableAdapter.InsertQuery(lblOrderNo.Text, txtRemark.Text, DateTime.Now);
            //new OrdersTableAdapter().UpdateOrderStatus(cmbStatus.Text, lblOrderNo.Text);
            //new OrderTransectionTableAdapter().InsertQuery(lblOrderNo.Text,"Updated To "+cmbStatus.Text+" - "+ txtRemark.Text, DateTime.Now);
            //HistoryLoad(lblOrderNo.Text, cmbStatus.Text);
            MessageBox.Show("Order Details Updated");
            //txtRemark.Text = "";
        }
    }
}
