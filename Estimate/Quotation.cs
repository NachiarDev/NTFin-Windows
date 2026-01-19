using QRCoder;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace Estimate
{
    public partial class Quotation : Form
    {
        public Quotation()
        {
            InitializeComponent();
        }

        private void Quotation_Load(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            LoadStates();
            LoadProductsIntoGridComboBox();

        }
        #region Combobox
        private void LoadStates()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT ID, CAST(GSTTIN AS VARCHAR) + ' - ' + StateName AS DisplayText FROM Tbl_State WHERE Active = 1";
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                DataTable dt = new DataTable();
                da.Fill(dt);

                comboBox1.DisplayMember = "DisplayText"; // what is shown
                comboBox1.ValueMember = "ID";            // actual value
                comboBox1.DataSource = dt;
                comboBox1.SelectedIndex = -1;
            }
        }
        #endregion

        #region Grid
        private void LoadProductsIntoGridComboBox()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, ProductName,HSNCode,Unit,Rate,Tax,Discount,TaxType FROM Product";
                    SqlDataAdapter da = new SqlDataAdapter(query, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dataGridView1.Columns["Column1"] is DataGridViewComboBoxColumn comboBoxColumn)
                    {
                        comboBoxColumn.DataSource = dt;
                        comboBoxColumn.DisplayMember = "ProductName"; // Show Product Name
                        comboBoxColumn.ValueMember = "ProductName"; // Store Product ID
                        dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                        //dataGridView1.DataError += dataGridView1_DataError;
                    }
                    else
                    {
                        MessageBox.Show("Product Name' not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                    if (e.ColumnIndex == dataGridView1.Columns["Column1"].Index)
                    {
                        if (row.Cells["Column1"].Value == null)
                            return;

                        string selectedProductName = row.Cells["Column1"].Value.ToString();

                        using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                        {
                            conn.Open();
                            string query = "SELECT HSNCode, TaxType, Unit, Rate, Discount, Tax FROM Product WHERE ProductName = @ProductName";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@ProductName", selectedProductName);
                                SqlDataReader reader = cmd.ExecuteReader();

                                if (reader.Read())
                                {
                                    row.Cells["Column5"].Value = reader["TaxType"].ToString();
                                    row.Cells["Column6"].Value = reader["HSNCode"].ToString();
                                    row.Cells["Column7"].Value = reader["Unit"].ToString();
                                    row.Cells["Column8"].Value = reader["Rate"].ToString();
                                    row.Cells["Column9"].Value = reader["Discount"].ToString();
                                    row.Cells["Column11"].Value = reader["Tax"].ToString();
                                }
                            }
                        }


                        LoadBatchForProduct(selectedProductName, row);
                    }


                    if (e.ColumnIndex == dataGridView1.Columns["Column2"].Index)
                    {
                        if (row.Cells["Column2"].Value == null)
                            return;

                        string selectedBatch = row.Cells["Column2"].Value.ToString();
                        string selectedProductName = row.Cells["Column1"].Value.ToString();
                        AutoFillSerialNo(selectedProductName, selectedBatch, row);
                    }

                    if (e.ColumnIndex == dataGridView1.Columns["Column4"].Index || e.ColumnIndex == dataGridView1.Columns["Column8"].Index || e.ColumnIndex == dataGridView1.Columns["Column9"].Index || e.ColumnIndex == dataGridView1.Columns["Column11"].Index)
                    {
                        if (row.Cells["Column4"].Value != null && row.Cells["Column8"].Value != null && row.Cells["Column9"].Value != null)
                        {
                            int quantity = Convert.ToInt32(row.Cells["Column4"].Value);
                           
                            decimal rate = Convert.ToDecimal(row.Cells["Column8"].Value);
                            decimal discountPercent = Convert.ToDecimal(row.Cells["Column9"].Value);


                            decimal total = quantity * rate;
                            decimal discountAmount = (total * discountPercent) / 100;
                            row.Cells["Column10"].Value = discountAmount;


                            decimal netAmount = total - discountAmount;


                            decimal additionalDiscountAmount = 0;
                            if (row.Cells["Column10"].Value != null)
                            {
                                decimal additionalDiscountPercent = Convert.ToDecimal(row.Cells["Column11"].Value);
                                additionalDiscountAmount = (netAmount * additionalDiscountPercent) / 100;
                                row.Cells["Column12"].Value = additionalDiscountAmount;
                            }
                            else
                            {
                                row.Cells["Column14"].Value = 0;
                            }


                            decimal finalAmount = netAmount + additionalDiscountAmount;
                            row.Cells["Column13"].Value = finalAmount;
                            CalculateTotal();

                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void CalculateTotal()
        {
            decimal totalQuantity = 0;
            decimal totalAmount = 0;
            decimal totalDiscount = 0;
            decimal totalTax = 0;
            decimal subtotal = 0;
            decimal totalSGST = 0;
            decimal totalCGST = 0;
            decimal totalIGST = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                decimal taxAmount = 0; // Initialize taxAmount to prevent unassigned variable error

                if (row.Cells["Column4"].Value != null && decimal.TryParse(row.Cells["Column4"].Value.ToString(), out decimal quantity))
                {
                    totalQuantity += quantity;
                }
                if (row.Cells["Column13"].Value != null && decimal.TryParse(row.Cells["Column13"].Value.ToString(), out decimal amount))
                {
                    totalAmount += amount;
                }
                if (row.Cells["Column10"].Value != null && decimal.TryParse(row.Cells["Column10"].Value.ToString(), out decimal discountAmount))
                {
                    totalDiscount += discountAmount;
                }
                if (row.Cells["Column12"].Value != null && decimal.TryParse(row.Cells["Column12"].Value.ToString(), out taxAmount))
                {
                    totalTax += taxAmount;
                }
                if (row.Cells["Column4"].Value != null && row.Cells["Column8"].Value != null &&
                    decimal.TryParse(row.Cells["Column4"].Value.ToString(), out decimal qty) &&
                    decimal.TryParse(row.Cells["Column8"].Value.ToString(), out decimal price))
                {
                    subtotal += qty * price; // Multiply Quantity (Column4) with Price (Column8) and sum up
                }

                // Handle SGST, CGST, and IGST based on Column5 value
                if (row.Cells["Column5"].Value != null && int.TryParse(row.Cells["Column5"].Value.ToString(), out int taxType))
                {
                    if (taxType == 2 || taxType == 3)
                    {
                        // Split Column14 value into SGST and CGST
                        decimal halfTax = taxAmount / 2;
                        totalSGST += halfTax;
                        totalCGST += halfTax;
                    }
                    else if (taxType == 1)
                    {
                        // Assign full Column14 value to IGST
                        totalIGST += taxAmount;
                    }
                }
            }

            label19.Text = totalQuantity.ToString("0");
            textBox4.Text = totalAmount.ToString("0.00");
            label12.Text = subtotal.ToString("0.00"); // Updated to show the correct subtotal
            label13.Text = totalDiscount.ToString("0.00");
            label14.Text = totalTax.ToString("0.00");
            label15.Text = totalIGST.ToString("0.00"); // SGST Total
            label16.Text = totalSGST.ToString("0.00"); // CGST Total
            label17.Text = totalCGST.ToString("0.00"); // IGST Total
           
        }

        private void LoadBatchForProduct(string productName, DataGridViewRow row)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT batch FROM StockIN WHERE ProductName = @ProductName";
                    SqlDataAdapter da = new SqlDataAdapter(query, conn);
                    da.SelectCommand.Parameters.AddWithValue("@ProductName", productName);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    // Create a new DataGridViewComboBoxCell for the specific row
                    DataGridViewComboBoxCell batchComboBoxCell = new DataGridViewComboBoxCell();
                    batchComboBoxCell.DataSource = dt;
                    batchComboBoxCell.DisplayMember = "batch";
                    batchComboBoxCell.ValueMember = "batch";

                    // Assign this cell to the specific row
                    row.Cells["Column2"] = batchComboBoxCell;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AutoFillSerialNo(string productName, string batch, DataGridViewRow row)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT serialno FROM StockIN WHERE ProductName = @ProductName AND batch = @Batch";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ProductName", productName);
                        cmd.Parameters.AddWithValue("@Batch", batch);
                        SqlDataReader reader = cmd.ExecuteReader();

                        if (reader.Read())
                        {
                            row.Cells["Column3"].Value = reader["serialno"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == dataGridView1.Columns["Column14"].Index)
            {

                if (dataGridView1.IsCurrentRowDirty)
                {
                    dataGridView1.EndEdit();


                    DialogResult result = MessageBox.Show("Are you sure you want to delete this row?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {

                        if (!dataGridView1.Rows[e.RowIndex].IsNewRow)
                        {
                            dataGridView1.Rows.RemoveAt(e.RowIndex);
                        }
                    }
                }
            }
        }
        #endregion

        #region Insert
        private void button1_Click(object sender, EventArgs e)
        {
            InsertData();
        }

        private async void InsertData()
        {
            try
            {
                if (!ValidateFields()) 
                    return;
                decimal estimateTotal = 0;
                decimal.TryParse(textBox4.Text.Trim(), out estimateTotal);
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();


                   
                    string insertQuotationQuery = @"
            INSERT INTO Quotation (RefNo,CustomerName, RefDate, CustomerCode, StateSupply, EstimateTotal, 
                                  CreatedBy, CreatedDate, BnchID, LocID) 
            VALUES (0,@CustomerName, @RefDate, @CustomerCode, @StateSupply, @EstimateTotal, 
                    @CreatedBy, GETDATE(), @BnchID, @LocID);
            SELECT SCOPE_IDENTITY();";

                    int refNo;
                    using (SqlCommand cmd = new SqlCommand(insertQuotationQuery, conn))
                    {
                       
                        cmd.Parameters.AddWithValue("@CustomerName", textBox1.Text);
                        cmd.Parameters.AddWithValue("@RefDate", textBox3.Text);
                        cmd.Parameters.AddWithValue("@CustomerCode", textBox2.Text);
                        cmd.Parameters.AddWithValue("@StateSupply", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@EstimateTotal", estimateTotal);
                        cmd.Parameters.AddWithValue("@CreatedBy", 1);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        object newIDObj = cmd.ExecuteScalar();
                        refNo = Convert.ToInt32(newIDObj);
                    }


                      


                    await Task.Delay(10000);



                    string updateQuotationQuery = "UPDATE Quotation SET RefNo = @RefNo WHERE ID = @ID";
                    using (SqlCommand updateCmd = new SqlCommand(updateQuotationQuery, conn))
                    {
                        updateCmd.Parameters.AddWithValue("@RefNo", refNo);
                        updateCmd.Parameters.AddWithValue("@ID", refNo);
                        updateCmd.ExecuteNonQuery();
                    }


                    if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows.Cast<DataGridViewRow>().All(r => r.IsNewRow))
                    {
                        MessageBox.Show("No data available in the grid to insert into Quotation Details!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }


                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null)
                            continue;
                        Random rnd = new Random();
                        int tempID = rnd.Next(100000, 999999);

                        string insertQuotationDetailsQuery = @"
                INSERT INTO Quotation_Details (temID,InvoNo,Items, batch, serialno, HSNCode, Qty, UnitType, 
                                              Price, Discount, DiscountAmount, Taxtype, TaxPercentage, TaxAmount, 
                                              TotalAmount, CreatedBy, CreatedDate, BnchID, LocID) 
                VALUES (@tempID,@RefNo,@Items, @Batch, @SerialNo, @HSNCode, @Qty, @UnitType, 
                        @Price, @Discount, @DiscountAmount, @Taxtype, @TaxPercentage, @TaxAmount, 
                        @TotalAmount, @CreatedBy, GETDATE(), @BnchID, @LocID);";

                        using (SqlCommand cmd = new SqlCommand(insertQuotationDetailsQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@tempID", tempID);
                            cmd.Parameters.AddWithValue("@RefNo", refNo);
                            cmd.Parameters.AddWithValue("@Items", row.Cells[0].Value?.ToString() ?? DBNull.Value.ToString());
                            cmd.Parameters.AddWithValue("@Batch", row.Cells[1].Value?.ToString() ?? DBNull.Value.ToString());
                            cmd.Parameters.AddWithValue("@SerialNo", row.Cells[2].Value?.ToString() ?? DBNull.Value.ToString());
                            cmd.Parameters.AddWithValue("@Qty", Convert.ToDecimal(row.Cells[3].Value ?? 0));
                            cmd.Parameters.AddWithValue("@Taxtype", row.Cells[4].Value?.ToString() ?? DBNull.Value.ToString());
                            cmd.Parameters.AddWithValue("@HSNCode", row.Cells[5].Value?.ToString() ?? DBNull.Value.ToString());
                            cmd.Parameters.AddWithValue("@UnitType", row.Cells[6].Value?.ToString() ?? DBNull.Value.ToString());
                            cmd.Parameters.AddWithValue("@Price", Convert.ToDecimal(row.Cells[7].Value ?? 0));
                            cmd.Parameters.AddWithValue("@Discount", Convert.ToDecimal(row.Cells[8].Value ?? 0));
                            cmd.Parameters.AddWithValue("@DiscountAmount", Convert.ToDecimal(row.Cells[9].Value ?? 0));
                            cmd.Parameters.AddWithValue("@TaxPercentage", Convert.ToDecimal(row.Cells[10].Value ?? 0));
                            cmd.Parameters.AddWithValue("@TaxAmount", Convert.ToDecimal(row.Cells[11].Value ?? 0));
                            cmd.Parameters.AddWithValue("@TotalAmount", Convert.ToDecimal(row.Cells[12].Value ?? 0));
                            
                          
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            cmd.ExecuteNonQuery();
                        }

                        
                    }

                    MessageBox.Show("Inserted Successfully");
                }
                ResetFormFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

        }

        #endregion

        #region Reset
        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        private void ResetFormFields()
        {

            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textBox && textBox != textBox3)
                {
                    textBox.Text = "";
                }
               

            }


            label12.Text = "";
            label13.Text = "";
            label14.Text = "";
            label15.Text = "";
            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
            label19.Text = "";
           

            comboBox1.SelectedIndex = -1;


           

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (i != 13)
                        {
                            row.Cells[i].Value = null;
                        }
                    }
                }
            }


            textBox1.Focus();
        }

        #endregion

        #region validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Customer Name is required" },
        { textBox2, "Customer Code is required" },
    };

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Customer State is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }

           




            return true;
           
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        #endregion

       

       
    }
}
