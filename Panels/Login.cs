using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Panels
{
    public partial class Login : Form
    {
        SqlConnection con = new SqlConnection("Data Source=DESKTOP-GT6UN2V\\SQLEXPRESS;Initial Catalog=registration;Integrated Security=True;");
        public Login()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtUName.Text=="" && txtPassword.Text=="")
                {
                    MessageBox.Show("Please Enter User name and Password");
                }
                else
                {
                    SqlCommand cmd = new SqlCommand("select * from LoginUsers where U_Name=@Name and U_Pass=@Pass", con);
                    cmd.Parameters.Add("@Name", txtUName.Text);
                    cmd.Parameters.Add("@Pass", txtPassword.Text);
                    SqlDataAdapter adpt = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    adpt.Fill(ds);

                    int count = ds.Tables[0].Rows.Count;
                    if(count == 1)
                    {
                        MessageBox.Show("You have Succefully Login");
                        Form1 ob = new Form1(); // conect with form 1
                        this.Hide();
                        ob.Show();
                    }
                    else
                    {
                        MessageBox.Show("Please Check Username and Password ");
                    }
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //SqlConnection conn = new SqlConnection
        }
    }
}
