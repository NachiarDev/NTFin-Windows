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
    public partial class ProductList : Form
    {
        Dictionary<string, int> categoryDict = new Dictionary<string, int>();
        Dictionary<string, int> subCategoryDict = new Dictionary<string, int>();
        public ProductList()
        {
            InitializeComponent();
        }

        private void ProductList_Load(object sender, EventArgs e)
        {
            MakeColumnsReadOnlyExceptButtons();
            LoadProductgrid();
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["Unit"].Visible = false;
            dataGridView1.Columns["TaxType"].Visible = false;
            dataGridView1.Columns["SelectedCategory"].Visible = false;
            dataGridView1.Columns["selectedSubCategory"].Visible = false;
            dataGridView1.Columns["Name"].HeaderText = "Category Name";
            dataGridView1.Columns["gsttype"].HeaderText = "GST Type";
            dataGridView1.Columns["Types"].HeaderText = "Unit Type";
            dataGridView1.Columns["Qty"].HeaderText = "Quantity";

            AddButtonColumns();
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#11345e");
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // Optional styling


        }
        DataTable originalDataTable = new DataTable(); // Global level

        private void LoadProductgrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("Usp_ProductGetGrid", conn))
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

                    string query = "DELETE FROM Product WHERE ID = @ID"; // Change table name properly

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
            Product newForm = new Product(); // Create an instance of Form2
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

                    Product newForm = new Product();

                    // Set values into Organisation form controls
                    newForm.Controls["textBox10"].Text = selectedRow.Cells["ID"].Value?.ToString();
                    newForm.Controls["textBox1"].Text = selectedRow.Cells["Code"].Value?.ToString();
                    newForm.Controls["textBox9"].Text = selectedRow.Cells["ProductName"].Value?.ToString();
                    newForm.Controls["textBox2"].Text = selectedRow.Cells["HSNCode"].Value?.ToString();
                    newForm.Controls["textBox3"].Text = selectedRow.Cells["Qty"].Value?.ToString();
                    newForm.Controls["textBox4"].Text = selectedRow.Cells["Rate"].Value?.ToString();
                    newForm.Controls["textBox5"].Text = selectedRow.Cells["Discount"].Value?.ToString();
                    newForm.Controls["textBox6"].Text = selectedRow.Cells["Tax"].Value?.ToString();
                    newForm.Controls["textBox7"].Text = selectedRow.Cells["MRP"].Value?.ToString();
                    newForm.Controls["textBox8"].Text = selectedRow.Cells["TotalQty"].Value?.ToString();

                    //newForm.Shown += (s, args) =>
                    //{
                    //    int unitId = Convert.ToInt32(selectedRow.Cells["Unit"].Value ?? 0);
                    //    int taxId = Convert.ToInt32(selectedRow.Cells["TaxType"].Value ?? 0);

                    //    if (newForm.Controls["comboBox2"] is ComboBox countryCombo)
                    //    {
                    //        // Load countries
                    //        using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                    //        {
                    //            con.Open();

                    //            SqlDataAdapter da = new SqlDataAdapter("SELECT ID, Types FROM Unit", con);
                    //            DataTable dtunits = new DataTable();
                    //            da.Fill(dtunits);

                    //            countryCombo.DisplayMember = "Types";
                    //            countryCombo.ValueMember = "ID";
                    //            countryCombo.DataSource = dtunits;
                    //            countryCombo.SelectedValue = unitId;
                    //        }
                    //    }


                    //    if (newForm.Controls["comboBox3"] is ComboBox gsttypeCombo)
                    //    {
                    //        // Load countries
                    //        using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                    //        {
                    //            con.Open();

                    //            SqlDataAdapter da = new SqlDataAdapter("SELECT ID, gsttype FROM Tax", con);
                    //            DataTable dtGSTTypes = new DataTable();
                    //            da.Fill(dtGSTTypes);

                    //            gsttypeCombo.DisplayMember = "gsttype";
                    //            gsttypeCombo.ValueMember = "ID";
                    //            gsttypeCombo.DataSource = dtGSTTypes;
                    //            gsttypeCombo.SelectedValue = taxId;


                    //        }
                    //    }
                    //    if (newForm.Controls["pictureBox1"] is PictureBox picBox)
                    //    {
                    //        string imagePath = selectedRow.Cells["Image"].Value?.ToString();
                    //        if (!string.IsNullOrEmpty(imagePath) && System.IO.File.Exists(imagePath))
                    //        {
                    //            picBox.Image = Image.FromFile(imagePath);
                    //            picBox.SizeMode = PictureBoxSizeMode.StretchImage;
                    //        }
                    //        else
                    //        {
                    //            picBox.Image = null;
                    //        }
                    //    }
                    //};
                    newForm.Shown += (s, args) =>
                    {
                        int unitId = Convert.ToInt32(selectedRow.Cells["Unit"].Value ?? 0);
                        int taxId = Convert.ToInt32(selectedRow.Cells["TaxType"].Value ?? 0);

                        // Load Unit ComboBox
                        if (newForm.Controls["comboBox2"] is ComboBox countryCombo)
                        {
                            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True"))
                            {
                                con.Open();
                                SqlDataAdapter da = new SqlDataAdapter("SELECT ID, Types FROM Unit", con);
                                DataTable dtunits = new DataTable();
                                da.Fill(dtunits);

                                countryCombo.DisplayMember = "Types";
                                countryCombo.ValueMember = "ID";
                                countryCombo.DataSource = dtunits;
                                countryCombo.SelectedValue = unitId;
                            }
                        }

                        // Load Tax ComboBox
                        if (newForm.Controls["comboBox3"] is ComboBox gsttypeCombo)
                        {
                            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True"))
                            {
                                con.Open();
                                SqlDataAdapter da = new SqlDataAdapter("SELECT ID, gsttype FROM Tax", con);
                                DataTable dtGSTTypes = new DataTable();
                                da.Fill(dtGSTTypes);

                                gsttypeCombo.DisplayMember = "gsttype";
                                gsttypeCombo.ValueMember = "ID";
                                gsttypeCombo.DataSource = dtGSTTypes;
                                gsttypeCombo.SelectedValue = taxId;
                            }
                        }

                        // Load Category CheckedListBox
                        if (newForm.Controls["clbCategory"] is CheckedListBox clbCategory)
                        {
                            clbCategory.Items.Clear();  // Clear existing
                            Dictionary<int, string> categoryDict = new Dictionary<int, string>(); // id -> name

                            using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True"))
                            {
                                conn.Open();
                                SqlCommand cmd = new SqlCommand("SELECT ID, Name FROM Category", conn);
                                SqlDataReader reader = cmd.ExecuteReader();

                                while (reader.Read())
                                {
                                    int id = Convert.ToInt32(reader["ID"]);
                                    string name = reader["Name"].ToString();
                                    categoryDict[id] = name;
                                    clbCategory.Items.Add(name); // Add by name
                                }
                                reader.Close();
                            }

                            // Read selected IDs from DataGridView
                            string selectedCategoryIds = selectedRow.Cells["SelectedCategory"].Value?.ToString();

                            if (!string.IsNullOrEmpty(selectedCategoryIds))
                            {
                                List<int> selectedIds = selectedCategoryIds
                                    .Split(',')
                                    .Select(id => Convert.ToInt32(id.Trim()))
                                    .ToList();

                                foreach (int selectedId in selectedIds)
                                {
                                    if (categoryDict.TryGetValue(selectedId, out string categoryName))
                                    {
                                        int index = clbCategory.Items.IndexOf(categoryName);
                                        if (index >= 0)
                                        {
                                            clbCategory.SetItemChecked(index, true);
                                        }
                                    }
                                }
                            }
                        }

                        // Load product image
                        if (newForm.Controls["pictureBox1"] is PictureBox picBox)
                        {
                            string imagePaths = selectedRow.Cells["Image"].Value?.ToString();
                            if (!string.IsNullOrEmpty(imagePaths))
                            {
                                string[] images = imagePaths.Split('|');
                                string firstImage = images[0];

                                if (System.IO.File.Exists(firstImage))
                                {
                                    picBox.Image = Image.FromFile(firstImage);
                                    picBox.SizeMode = PictureBoxSizeMode.StretchImage;
                                }
                                else
                                {
                                    picBox.Image = null;
                                }
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
