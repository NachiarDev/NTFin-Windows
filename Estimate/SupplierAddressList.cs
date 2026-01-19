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
    public partial class SupplierAddressList : Form
    {
        public SupplierAddressList()
        {
            InitializeComponent();
        }

        private void SupplierAddressList_Load(object sender, EventArgs e)
        {
            MakeColumnsReadOnlyExceptButtons();
            LoadSupplierAddressgrid();
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["StateID"].Visible = false;
            dataGridView1.Columns["CountryID"].Visible = false;
            dataGridView1.Columns["SupplierID"].Visible = false;
            dataGridView1.Columns["AddressTypeID"].Visible = false;
            dataGridView1.Columns["Address"].HeaderText = "CustomerAddress";
            AddButtonColumns();
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#11345e");
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // Optional styling

        }
        DataTable originalDataTable = new DataTable(); // Global level

        private void LoadSupplierAddressgrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("Usp_SupplierAddressGetGrid", conn))
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

            // ⚡ NOW — Remove existing button columns (if any)
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

                    string query = "DELETE FROM Supplier_Address WHERE ID = @ID"; // Change table name properly

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
            SupplierAddress newForm = new SupplierAddress(); // Create an instance of Form2
            newForm.ShowDialog();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string columnName = dataGridView1.Columns[e.ColumnIndex].Name;

                if (columnName == "Button1") // View button clicked
                {
                    DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];

                    SupplierAddress newForm = new SupplierAddress();

                    newForm.Controls["textBox2"].Text = selectedRow.Cells["ID"].Value?.ToString();
                    newForm.Controls["richTextBox1"].Text = selectedRow.Cells["Address"].Value?.ToString();
                    newForm.Controls["textBox1"].Text = selectedRow.Cells["ZipCode"].Value?.ToString();



                    // Set Active Status (radio buttons)
                    if (newForm.Controls["radioButton1"] is RadioButton rbActive && newForm.Controls["radioButton2"] is RadioButton rbInactive)
                    {
                        int activeValue = Convert.ToInt32(selectedRow.Cells["Active"].Value ?? 0);
                        rbActive.Checked = (activeValue == 1);
                        rbInactive.Checked = (activeValue == 0);
                    }

                    // Show the form FIRST so ComboBoxes are created and loaded
                    newForm.Shown += (s, args) =>
                    {

                        int addresstypeId = Convert.ToInt32(selectedRow.Cells["AddressTypeID"].Value ?? 0);
                        if (newForm.Controls["comboBox3"] is ComboBox addresstypeCombo)
                        {
                            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                            {
                                con.Open();

                                SqlDataAdapter addresstypeDa = new SqlDataAdapter("SELECT ID, AddressType FROM Tbl_AddressType", con);
                                DataTable dtOrgs = new DataTable();
                                addresstypeDa.Fill(dtOrgs);
                                addresstypeCombo.DisplayMember = "AddressType";
                                addresstypeCombo.ValueMember = "ID";
                                addresstypeCombo.DataSource = dtOrgs;
                                addresstypeCombo.SelectedValue = addresstypeId;


                            }
                        }
                        int countryId = Convert.ToInt32(selectedRow.Cells["CountryID"].Value ?? 0);
                        int stateId = Convert.ToInt32(selectedRow.Cells["StateID"].Value ?? 0);

                        if (newForm.Controls["comboBox1"] is ComboBox countryCombo &&
                            newForm.Controls["comboBox2"] is ComboBox stateCombo)
                        {
                            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                            {
                                con.Open();

                                // Load countries
                                SqlDataAdapter countryDa = new SqlDataAdapter("SELECT ID, CountryName FROM Tbl_Country WHERE Active = 1", con);
                                DataTable dtCountries = new DataTable();
                                countryDa.Fill(dtCountries);
                                countryCombo.DisplayMember = "CountryName";
                                countryCombo.ValueMember = "ID";
                                countryCombo.DataSource = dtCountries;
                                countryCombo.SelectedValue = countryId;

                                // Load states based on selected country
                                SqlDataAdapter stateDa = new SqlDataAdapter("SELECT ID, StateName FROM Tbl_State WHERE Active = 1 AND CountryID = @CountryID", con);
                                stateDa.SelectCommand.Parameters.AddWithValue("@CountryID", countryId);
                                DataTable dtStates = new DataTable();
                                stateDa.Fill(dtStates);
                                stateCombo.DisplayMember = "StateName";
                                stateCombo.ValueMember = "ID";
                                stateCombo.DataSource = dtStates;
                                stateCombo.SelectedValue = stateId;
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
