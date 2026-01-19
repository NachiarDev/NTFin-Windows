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

namespace Estimate
{
    public partial class CustomerList : Form
    {
        public CustomerList()
        {
            InitializeComponent();
        }

        

        private void CustomerList_Load(object sender, EventArgs e)
        {
          MakeColumnsReadOnlyExceptButtons();
            LoadCustomergrid();
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["StateID"].Visible = false;
            dataGridView1.Columns["CountryID"].Visible = false;
            dataGridView1.Columns["GSTTypeID"].Visible = false;
            dataGridView1.Columns["GSTAreaCode"].Visible = false;
            dataGridView1.Columns["ContactNo"].HeaderText = "Primary Mobile Number";
            dataGridView1.Columns["MobileNo"].HeaderText = "Alternate Mobile Number";
            AddButtonColumns();
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#11345e");
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // Optional styling

        }

        DataTable originalDataTable = new DataTable(); // Global level

        private void LoadCustomergrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("Usp_CustomerGetGrid", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(originalDataTable); // Fill global DataTable
                        dataGridView1.DataSource = originalDataTable;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string filterText = textBox1.Text.Trim();

            if (string.IsNullOrEmpty(filterText))
            {
                dataGridView1.DataSource = originalDataTable; // If textbox is empty, show all rows
            }
            else
            {
                DataView dv = new DataView(originalDataTable);
                string filterExpression = "";

                // Build filter expression
                foreach (DataColumn column in originalDataTable.Columns)
                {
                    if (column.DataType == typeof(string))
                    {
                        filterExpression += $"{column.ColumnName} LIKE '%{filterText}%' OR ";
                    }
                }

                // Remove last "OR "
                if (filterExpression.EndsWith(" OR "))
                {
                    filterExpression = filterExpression.Substring(0, filterExpression.Length - 4);
                }

                if (!string.IsNullOrEmpty(filterExpression))
                {
                    dv.RowFilter = filterExpression;
                    dataGridView1.DataSource = dv;
                }
                else
                {
                    dataGridView1.DataSource = originalDataTable;
                }
            }
            foreach (DataGridViewColumn col in dataGridView1.Columns.Cast<DataGridViewColumn>().ToList())
            {
                if (col.Name == "Button1" || col.Name == "Button2")
                {
                    dataGridView1.Columns.Remove(col);
                }
            }
        }

        private void AddButtonColumns()
        {
            // Button 1 - Open Dialog
            if (!dataGridView1.Columns.Contains("Button1"))
            {
                DataGridViewButtonColumn btnCol1 = new DataGridViewButtonColumn();
                btnCol1.Name = "Button1";
                btnCol1.HeaderText = "Edit";
                btnCol1.Text = "\u270F";
                btnCol1.UseColumnTextForButtonValue = true;
                dataGridView1.Columns.Add(btnCol1);
            }

            // Button 2 - Delete
            if (!dataGridView1.Columns.Contains("Button2"))
            {
                DataGridViewButtonColumn btnCol2 = new DataGridViewButtonColumn();
                btnCol2.Name = "Button2";
                btnCol2.HeaderText = "Delete";
                btnCol2.Text = "🗑️";
                btnCol2.UseColumnTextForButtonValue = true;
                dataGridView1.Columns.Add(btnCol2);
            }
        }
        private void DeleteRowData(int rowIndex)
        {
            DataGridViewRow row = dataGridView1.Rows[rowIndex];

            // Assuming you have an "ID" column to identify rows
            var id = row.Cells["ID"].Value?.ToString();

            if (string.IsNullOrEmpty(id))
            {
                MessageBox.Show("Cannot find ID for this row.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Confirm delete
            var result = MessageBox.Show("Are you sure you want to delete this row?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                // 1. Delete from database
                DeleteFromDatabase(id);

                // 2. Delete from grid
                dataGridView1.Rows.RemoveAt(rowIndex);
            }
        }
        private void DeleteFromDatabase(string id)
        {
            try
            {
                // Adjust your connection string
                string connectionString = "Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = "DELETE FROM Customer WHERE ID = @ID"; // Change table name properly

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", id);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting from database: " + ex.Message);
            }
        }
        private void MakeColumnsReadOnlyExceptButtons()
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                if (column.Name != "Button1" && column.Name != "Button2")
                {
                    column.ReadOnly = true;
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Customer newForm = new Customer(); // Create an instance of Form2
            newForm.ShowDialog();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string columnName = dataGridView1.Columns[e.ColumnIndex].Name;

                if (columnName == "Button1") // View button clicked
                {
                    // Get data from the selected row
                    DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];

                    Customer newForm = new Customer();

                    // Set values into Organisation form controls
                    newForm.Controls["textBox3"].Text = selectedRow.Cells["ID"].Value?.ToString();
                    newForm.Controls["textBox1"].Text = selectedRow.Cells["Code"].Value?.ToString();
                    newForm.Controls["textBox2"].Text = selectedRow.Cells["CustomerName"].Value?.ToString();
                    newForm.Controls["maskedTextBox1"].Text = selectedRow.Cells["ContactNo"].Value?.ToString();
                    newForm.Controls["maskedTextBox2"].Text = selectedRow.Cells["MobileNo"].Value?.ToString();
                    newForm.Controls["textBox5"].Text = selectedRow.Cells["Email"].Value?.ToString();
                    
                    newForm.Controls["textBox6"].Text = selectedRow.Cells["GSTNumber"].Value?.ToString();
                    newForm.Shown += (s, args) =>
                    {
                        int countryId = Convert.ToInt32(selectedRow.Cells["CountryID"].Value ?? 0);
                        int stateId = Convert.ToInt32(selectedRow.Cells["StateID"].Value ?? 0);
                        int gsttypeid = Convert.ToInt32(selectedRow.Cells["GSTTypeID"].Value ?? 0);
                        int gstareacodeid = Convert.ToInt32(selectedRow.Cells["GSTAreaCode"].Value ?? 0);
                        if (newForm.Controls["comboBox1"] is ComboBox countryCombo &&
                            newForm.Controls["comboBox2"] is ComboBox stateCombo)
                        {
                            // Load countries
                            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                            {
                                con.Open();

                                SqlDataAdapter da = new SqlDataAdapter("SELECT ID, CountryName FROM Tbl_Country WHERE Active = 1", con);
                                DataTable dtCountries = new DataTable();
                                da.Fill(dtCountries);

                                countryCombo.DisplayMember = "CountryName";
                                countryCombo.ValueMember = "ID";
                                countryCombo.DataSource = dtCountries;
                                countryCombo.SelectedValue = countryId;

                                // Load states only after country set
                                SqlDataAdapter daStates = new SqlDataAdapter("SELECT ID, StateName FROM Tbl_State WHERE Active = 1 AND CountryID = @CountryID", con);
                                daStates.SelectCommand.Parameters.AddWithValue("@CountryID", countryId);
                                DataTable dtStates = new DataTable();
                                daStates.Fill(dtStates);

                                stateCombo.DisplayMember = "StateName";
                                stateCombo.ValueMember = "ID";
                                stateCombo.DataSource = dtStates;
                                stateCombo.SelectedValue = stateId;
                            }
                        }
                        

                        if (newForm.Controls["comboBox3"] is ComboBox gsttypeCombo )
                        {
                            // Load countries
                            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                            {
                                con.Open();

                                SqlDataAdapter da = new SqlDataAdapter("SELECT ID, GSTType FROM Tbl_GSTType", con);
                                DataTable dtGSTTypes = new DataTable();
                                da.Fill(dtGSTTypes);

                                gsttypeCombo.DisplayMember = "GSTType";
                                gsttypeCombo.ValueMember = "ID";
                                gsttypeCombo.DataSource = dtGSTTypes;
                                gsttypeCombo.SelectedValue = gsttypeid;

                                
                            }
                        }
                        
                        if (newForm.Controls["comboBox4"] is ComboBox gstareacodecombo)
                        {
                            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                            {
                                con.Open();

                                // Updated SQL to return a formatted string
                                SqlDataAdapter da = new SqlDataAdapter("SELECT ID, CAST(GSTTIN AS VARCHAR) + '-' + StateName AS DisplayText FROM Tbl_State", con);
                                DataTable dtGSTareacodes = new DataTable();
                                da.Fill(dtGSTareacodes);

                                gstareacodecombo.DisplayMember = "DisplayText";
                                gstareacodecombo.ValueMember = "ID";
                                gstareacodecombo.DataSource = dtGSTareacodes;
                                gstareacodecombo.SelectedValue = gstareacodeid;
                            }
                        }
                        if (newForm.Controls["pictureBox1"] is PictureBox picBox)
                        {
                            string imagePath = selectedRow.Cells["image"].Value?.ToString();
                            if (!string.IsNullOrEmpty(imagePath) && System.IO.File.Exists(imagePath))
                            {
                                picBox.Image = Image.FromFile(imagePath);
                                picBox.SizeMode = PictureBoxSizeMode.StretchImage;
                            }
                            else
                            {
                                picBox.Image = null;
                            }
                        }
                    };

                    newForm.ShowDialog();
                }
                else if (columnName == "Button2") // Delete button clicked
                {
                    DeleteRowData(e.RowIndex);
                }
            }

        }
    }
}
