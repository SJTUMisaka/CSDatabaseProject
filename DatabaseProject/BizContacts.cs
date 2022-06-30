using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace DatabaseProject
{
    public partial class BizContacts : Form
    {
        string connString = @"Data Source=DN-LAPTOP-233;Initial Catalog=AddressBook;Integrated Security=True;
            Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        SqlDataAdapter dataAdapter;
        System.Data.DataTable table;
        SqlConnection conn;
        string selectionCommand = "Select * from BizContacts";

        public delegate void RefreshForm();
        public RefreshForm refreshDelegate;
        public BizContacts()
        {
            InitializeComponent();
            refreshDelegate = new RefreshForm(RefreshFunc);
        }

        private void RefreshFunc()
        {
            GetData(selectionCommand);
            dataGridView1.Update();
        }

        private void RefreshForms()
        {
            foreach (Form frm in System.Windows.Forms.Application.OpenForms)
            {
                if (frm is BizContacts)
                    frm.Invoke((frm as BizContacts).refreshDelegate);
            }
        }

        private void BizContacts_Load(object sender, EventArgs e)
        {
            cboSearch.SelectedIndex = 0;
            dataGridView1.DataSource = bindingSource1;

            GetData(selectionCommand);
        }

        private void GetData(string selectCommand)
        {
            try
            {
                dataAdapter = new SqlDataAdapter(selectCommand, connString);
                table = new System.Data.DataTable();
                //table.Locale = System.Globalization.CultureInfo.InvariantCulture;
                dataAdapter.Fill(table);
                bindingSource1.DataSource = table;
                dataGridView1.Columns[0].ReadOnly = true;
            }
            catch(SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            SqlCommand command;
            string insertCommand = @"insert into BizContacts(Date_Added, Company, Website, Title, First_Name, Last_Name, Address, 
                                                            City, State, Postal_Code, Mobile, Notes, Image)
                                    values(@Date_Added, @Company, @Website, @Title, @First_Name, @Last_Name, @Address, 
                                                            @City, @State, @Postal_Code, @Mobile, @Notes, @Image)";
            using (conn = new SqlConnection(connString))
            {
                try 
                {
                    conn.Open();
                    command = new SqlCommand(insertCommand, conn);
                    command.Parameters.AddWithValue(@"Date_Added", dateTimePicker1.Value.Date);
                    command.Parameters.AddWithValue(@"Company", txtCompany.Text);
                    command.Parameters.AddWithValue(@"Website", txtWebsite.Text);
                    command.Parameters.AddWithValue(@"Title", txtTitle.Text);
                    command.Parameters.AddWithValue(@"First_Name", txtFName.Text);
                    command.Parameters.AddWithValue(@"Last_Name", txtLName.Text);
                    command.Parameters.AddWithValue(@"Address", txtAddress.Text);
                    command.Parameters.AddWithValue(@"City", txtCity.Text);
                    command.Parameters.AddWithValue(@"State", txtState.Text);
                    command.Parameters.AddWithValue(@"Postal_Code", txtZip.Text);
                    command.Parameters.AddWithValue(@"Mobile", txtMobile.Text);
                    command.Parameters.AddWithValue(@"Notes", txtNotes.Text);
                    if (dlgOpenImage.FileName != "")
                        command.Parameters.AddWithValue(@"Image", File.ReadAllBytes(dlgOpenImage.FileName));
                    else
                        command.Parameters.Add("@Image", SqlDbType.VarBinary).Value = DBNull.Value;
                    command.ExecuteNonQuery();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            RefreshForms();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
            dataAdapter.UpdateCommand = commandBuilder.GetUpdateCommand();
            try
            {
                bindingSource1.EndEdit();
                dataAdapter.Update(table);
                MessageBox.Show("Update Successful");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            string deleteCommand = @"Delete from BizContacts where ID = @ID";
            DataGridViewRow row = dataGridView1.CurrentCell.OwningRow;
            string value = row.Cells["ID"].Value.ToString();
            string fname = row.Cells["First_Name"].Value.ToString();
            string lname = row.Cells["Last_Name"].Value.ToString();
            
            if (MessageBox.Show($"Do you really want to delete {fname} {lname}, record {value}?", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                using(conn = new SqlConnection(connString))
                {
                    try
                    {
                        conn.Open();
                        var command = new SqlCommand(deleteCommand, conn);
                        command.Parameters.AddWithValue(@"ID", value);
                        command.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            RefreshForms();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            switch(cboSearch.SelectedItem.ToString())
            {
                case "First Name":
                    GetData("Select * from BizContacts where lower(First_Name) like '%" + txtSearch.Text.ToLower() + "%'");
                    break;
                case "Last Name":
                    GetData("Select * from BizContacts where lower(Last_Name) like '%" + txtSearch.Text.ToLower() + "%'");
                    break;
                case "Company":
                    GetData("Select * from BizContacts where lower(Company) like '%" + txtSearch.Text.ToLower() + "%'");
                    break;
            }
        }

        private void btnGetImage_Click(object sender, EventArgs e)
        {
            if (dlgOpenImage.ShowDialog() == DialogResult.OK)
                pictureBox1.Load(dlgOpenImage.FileName);
        }

        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            var picForm = new Form();
            picForm.BackgroundImage = pictureBox1.Image;
            picForm.Size = pictureBox1.Image.Size;
            picForm.Show();
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        private void btnExportOpen_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbooks bk = app.Workbooks;
            _Workbook workbook = bk.Add(Type.Missing);
            _Worksheet worksheet = null;
            try
            {
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Business Contacts";
                for(int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                for(int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for(int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        else
                            worksheet.Cells[i + 2, j + 1] = "";
                    }
                }
                if (saveFileDialog1.ShowDialog() == DialogResult.OK) 
                {
                    app.ActiveWorkbook.SaveAs(saveFileDialog1.FileName);
                    Process.Start(saveFileDialog1.FileName);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                app.Quit();
                GetWindowThreadProcessId(app.Hwnd, out int id);
                Process.GetProcessById(id).Kill();
            }
        }

        private void btnExportTxt_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Text Files (*.txt)|*.txt";
            try
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    using (var sw = new StreamWriter(saveFileDialog1.FileName))
                    {
                        foreach(DataGridViewRow row in dataGridView1.Rows)
                        {
                            foreach(DataGridViewCell cell in row.Cells)
                            {
                                sw.Write(cell.Value);
                            }
                            sw.WriteLine();
                        }
                    }
                    Process.Start("notepad.exe", saveFileDialog1.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
