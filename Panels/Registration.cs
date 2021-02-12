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
using Microsoft.Office.Interop.Excel;

namespace Panels
{
    public partial class Registration : Form
    {
        // set path in sqlserver
        string path = @"Data Source=DESKTOP-GT6UN2V\SQLEXPRESS;Initial Catalog=registration;Integrated Security=True";
        SqlConnection con;     
        SqlCommand cmd;      //Run sql server quires
        SqlDataAdapter adpt;
        System.Data.DataTable dt;          // DataTable object create // Reson for Add in "System.Data. by using Excel library Other wise DataTable word is enough"
        int ID;

        public Registration()
        {
            InitializeComponent();
            con = new SqlConnection(path); // connection in sql server
            display();
            button2.Enabled = false;  // If user open app first time they cannot click on update button
            button3.Enabled = false;  // If user open app first time they cannot click on delete button
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtName.Text == "" || txtFName.Text == "" || txtDesignation.Text == "" || txtID.Text == "" || txtEmail.Text == "" || txtID.Text == "" || txtAddress.Text == "")
            {
                MessageBox.Show("Please Fill in the Blanks");
            }
            else
            {
                try
                {
                    string gender;
                    if (rbtnMale.Checked)
                    {
                        gender = "Male";
                    }
                    else
                    {
                        gender = "Female";
                    }
                    cmd = new SqlCommand("insert into Employee (Employee_Name,Employee_FName,Employee_Designation,Employee_Email,Emp_ID,Gender,Addrss) values ('" + txtName.Text + "','" + txtFName.Text + "','" + txtDesignation.Text + "','" + txtEmail.Text + "','" + txtID.Text + "','" + gender + "','" + txtAddress.Text + "')", con);   // Insert query
                    con.Open(); // sql connection on
                    cmd.ExecuteNonQuery(); // execute the query
                    con.Close(); // close the sql connection
                    MessageBox.Show("You Data has Been Saved in the Database");
                    clear(); // after given insert value clean text fields
                    display();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        public void clear()
        {
            txtName.Text = "";
            txtFName.Text = "";
            txtDesignation.Text = "";
            txtEmail.Text = "";
            txtID.Text = "";
            txtAddress.Text = "";
        }
        public void display()
        {
            try
            {
                dt = new System.Data.DataTable();
                con.Open();
                adpt = new SqlDataAdapter("select * from Employee", con);
                adpt.Fill(dt);
                dataGridView1.DataSource = dt;
                con.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)  // Get data gridview for text boxese
        {
            ID = int.Parse(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString()); //assign the gridview id for the ID variable
            txtName.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString(); //assign gridview Name value for text box Name field
            txtFName.Text =dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString(); //assign gridview FName value for text box FName field
            txtDesignation.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString(); //assign gridview Designation value for text box Designation field
            txtEmail.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString(); //assign gridview txtEmail value for text box Email field
            txtID.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString(); //assign gridview txtID value for text box ID field

            rbtnMale.Checked = true;    //check the radio button values in gridview and assign them into the correct radio button
            rbtnFemale.Checked = false;

            if(dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString()=="Female")
            {
                rbtnMale.Checked = false;
                rbtnFemale.Checked = true;
            }
            txtAddress.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString(); //assign gridview txtAddress value for text box Address field
            
            button2.Enabled = true; //If user click on grid view update button is enable
            button3.Enabled = true; //If user click on grid view delete button is enable
        }

        private void button2_Click(object sender, EventArgs e)  //update button
        {
            try
            {
                string gender;

                if(rbtnMale.Checked)
                {
                    gender = "Male";
                }
                {
                    gender = "Female";
                }
                con.Open();
                cmd = new SqlCommand("update employee set Employee_Name = '" + txtName.Text + "',Employee_FName = '" + txtFName.Text + "',Employee_Designation='" + txtDesignation.Text + "',Employee_Email='" + txtEmail.Text + "',Emp_ID='" + txtID.Text + "', Gender='" + gender + "',Addrss='" + txtAddress.Text + "'where Employee_Id='"+ID+"'",con); // update sql query
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Your Data has been updated");
                display();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)  //Delete Button
        {
            try
            {
                con.Open();
                cmd = new SqlCommand("delete From Employee where Employee_Id='" + ID + "'", con);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Your Record has been Deleted");
                display();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            con.Open();
            adpt = new SqlDataAdapter("select * from employee where Employee_Name like '%" + txtSearch.Text + "%'", con); //search by using employee name
            //adpt = new SqlDataAdapter("select * from employee where Employee_Id like '%" + txtSearch.Text + "%'", con); // search by using employee id
            //adpt = new SqlDataAdapter("select * from employee where Employee_FName like '%" + txtSearch.Text + "%'", con);  // search by using employee fname
            //adpt = new SqlDataAdapter("select * from employee where Employee_Designation like '%" + txtSearch.Text + "%'", con);  // search by using employee designation
            //adpt = new SqlDataAdapter("select * from employee where Employee_Email '%" + txtSearch.Text + "%'", con);  // search by using employee Email
            //adpt = new SqlDataAdapter("select * from employee where Emp_ID like '%" + txtSearch.Text + "%'", con);  // search by using employee Emp ID
            //adpt = new SqlDataAdapter("select * from employee where Addrss like '%" + txtSearch.Text + "%'", con);  // search by using employee address

            dt = new System.Data.DataTable();
            adpt.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        private void button4_Click(object sender, EventArgs e)    //Export excel document
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excell = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = Excell.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)Excell.ActiveSheet;
                Excell.Visible = true;

                for (int j = 2; j <= dataGridView1.Rows.Count; j++)               // Count grid view Rows 
                {
                    for (int i = 1; i <= 1; i++)
                    {
                        ws.Cells[j, i] = dataGridView1.Rows[j - 2].Cells[i - 1].Value;
                    }
                }

                for (int i = 1; i <= dataGridView1.Columns.Count + 1; i++)         //Count grid view Colums
                {
                    ws.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Columns.Count - 1; i++)       //Get all data from grid view
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        ws.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
