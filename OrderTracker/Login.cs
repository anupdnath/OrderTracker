using System;
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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void UserLogin()
        {
            try
            {
                if (txtUsername.Text.Trim() != "" && txtPassword.Text.Trim() != "")
                {
                    DataTable dt = new DataTable();
                    dt = new usermasterTableAdapter().GetUserLogin(txtUsername.Text.Trim(),txtPassword.Text.Trim(), true);
                    if (dt.Rows.Count > 0)
                    {
                        txtUsername.Text = "";
                        txtPassword.Text = "";
                        lblMessage.Text = "";
                        Hide();
                        OrderUpload o = new OrderUpload();
                        o.Show();
                    }
                    else
                    {
                        lblMessage.Text = "Enter Valid Username And Password";
                    }
                }
                else
                {
                    lblMessage.Text = "Enter UserName and Password";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            UserLogin();
        }

        private void txtUsername_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                UserLogin();
            }
        }

        private void txtPassword_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                UserLogin();
            }
        }
    }
}
