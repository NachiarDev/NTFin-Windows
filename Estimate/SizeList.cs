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
    public partial class SizeList : Form
    {
        public SizeList()
        {
            InitializeComponent();
        }

        private void SizeList_Load(object sender, EventArgs e)
        {
            LoadSizegrid();

            MakeColumnsReadOnlyExceptButtons();

            dataGridView1.Columns["ID"].Visible = false;
            AddButtonColumns();
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#11345e");
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // Optional styling


        }
        DataTable originalDataTable = new DataTable(); // Global level

        private void LoadSizegrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("Usp_SizeGetGrid", conn))
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

        private void Loadaddresstypegrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("Usp_AddressTypeGetGrid", conn))
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

                    string query = "DELETE FROM Size WHERE ID = @ID"; // Change table name properly

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

        private void button1_Click(object sender, EventArgs e)
        {
            SizeForm newForm = new SizeForm(); // Create an instance of Form2
            newForm.ShowDialog();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string columnName = dataGridView1.Columns[e.ColumnIndex].Name;

                if (columnName == "Button1")
                {
                    DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];

                    SizeForm newForm = new SizeForm();

                    newForm.Controls["textBox2"].Text = selectedRow.Cells["ID"].Value?.ToString();
                    newForm.Controls["textBox1"].Text = selectedRow.Cells["SizeName"].Value?.ToString();
                    // Set Active Status (radio buttons)
                    if (newForm.Controls["radioButton1"] is RadioButton rbActive && newForm.Controls["radioButton2"] is RadioButton rbInactive)
                    {
                        int activeValue = Convert.ToInt32(selectedRow.Cells["Active"].Value ?? 0);
                        rbActive.Checked = (activeValue == 1);
                        rbInactive.Checked = (activeValue == 0);
                    }

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
