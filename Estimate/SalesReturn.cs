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
    public partial class SalesReturn : Form
    {
        public SalesReturn()
        {
            InitializeComponent();
        }

        private void SalesReturn_Load(object sender, EventArgs e)
        {
            textBox5.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            LoadModeofreturn();
            LoadRefNO();
            LoadProductsIntoGridComboBox();
           
        }
        #region ComboBox
        private void LoadRefNO()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT ID, RefNo FROM Estimate";
                SqlCommand cmd = new SqlCommand(query, con);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                comboBox1.DisplayMember = "RefNo";
                comboBox1.ValueMember = "ID";
                comboBox1.DataSource = dt;
                comboBox1.SelectedIndex = -1;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue == null)
                return;

            int selectedID = Convert.ToInt32(comboBox1.SelectedValue);

            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT CustomerCode, StateSupply, RefDate, CustomerName FROM Estimate WHERE ID = @ID";
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@ID", selectedID);

                con.Open();
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    textBox3.Text = reader["CustomerCode"].ToString();
                    textBox2.Text = reader["StateSupply"].ToString();
                    textBox1.Text = Convert.ToDateTime(reader["RefDate"]).ToString("yyyy-MM-dd");
                    textBox4.Text = reader["CustomerName"].ToString();
                }
                con.Close();
            }
        }
        private void LoadModeofreturn()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT ID, ReturnType FROM Mode_of_Return";
                SqlCommand cmd = new SqlCommand(query, con);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                comboBox2.DisplayMember = "ReturnType";
                comboBox2.ValueMember = "ID";
                comboBox2.DataSource = dt;
                comboBox2.SelectedIndex = -1;

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
                        MessageBox.Show("ComboBox column 'Column1' not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            label17.Text = totalQuantity.ToString("0");
            textBox6.Text = totalAmount.ToString("0.00");
            label18.Text = subtotal.ToString("0.00"); // Updated to show the correct subtotal
            label19.Text = totalDiscount.ToString("0.00");
            label20.Text = totalTax.ToString("0.00");
            label23.Text = totalIGST.ToString("0.00"); // SGST Total
            label22.Text = totalSGST.ToString("0.00"); // CGST Total
            label21.Text = totalCGST.ToString("0.00"); // IGST Total

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
                decimal.TryParse(textBox6.Text.Trim(), out estimateTotal);
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();



                    string insertQuotationQuery = @"
            INSERT INTO SalesReturn (ReturnID,ReturnDate,RefNo,CustomerName, RefDate, CustomerCode, StateSupply, EstimateTotal, 
                                  CreatedBy, CreatedDate, BnchID, LocID,ModeofReturn,Description) 
            VALUES (0,@ReturnDate,@RefNo,@CustomerName, @RefDate, @CustomerCode, @StateSupply, @EstimateTotal, 
                    @CreatedBy, GETDATE(), @BnchID, @LocID,@ModeofReturn,@Description);
            SELECT SCOPE_IDENTITY();";

                    int refNo;
                    using (SqlCommand cmd = new SqlCommand(insertQuotationQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@ReturnDate", textBox5.Text);
                        cmd.Parameters.AddWithValue("@RefNo", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@RefDate", textBox1.Text);
                        cmd.Parameters.AddWithValue("@CustomerCode", textBox3.Text);
                        cmd.Parameters.AddWithValue("@ModeofReturn", comboBox2.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Description", richTextBox1.Text);
                        cmd.Parameters.AddWithValue("@StateSupply", textBox2.Text);
                        cmd.Parameters.AddWithValue("@CustomerName", textBox4.Text);
                        //cmd.Parameters.AddWithValue("@EstimateTotal", Convert.ToDecimal(textBox6.Text));
                        cmd.Parameters.AddWithValue("@EstimateTotal", estimateTotal);
                        cmd.Parameters.AddWithValue("@CreatedBy", 1);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        object newIDObj = cmd.ExecuteScalar();
                        refNo = Convert.ToInt32(newIDObj);
                    }





                    await Task.Delay(10);



                    string updateSalesreturnQuery = "UPDATE SalesReturn SET ReturnID = @ReturnID WHERE ID = @ID";
                    using (SqlCommand updateCmd = new SqlCommand(updateSalesreturnQuery, conn))
                    {
                        updateCmd.Parameters.AddWithValue("@ReturnID", refNo);
                        updateCmd.Parameters.AddWithValue("@ID", refNo);
                        updateCmd.ExecuteNonQuery();
                    }


                    if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows.Cast<DataGridViewRow>().All(r => r.IsNewRow))
                    {
                        MessageBox.Show("No data available in the grid to insert into SalesReturn Details!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null)
                            continue;

                        string productName = row.Cells[0].Value?.ToString() ?? "";
                        string batchNo = row.Cells[1].Value?.ToString() ?? "";
                        decimal qty = Convert.ToDecimal(row.Cells[3].Value ?? 0);

                        // 1. Insert into SalesReturn_Details
                        string insertSalesDetailsQuery = @"
        INSERT INTO SalesReturn_Details (SRReturnID, Items, batch, serialno, HSNCode, Qty, UnitType, 
                                         Price, Discount, DiscountAmount, Taxtype, TaxPercentage, TaxAmount, 
                                         TotalAmount, CreatedBy, CreatedDate, BnchID, LocID) 
        VALUES (@SRReturnID, @Items, @Batch, @SerialNo, @HSNCode, @Qty, @UnitType, 
                @Price, @Discount, @DiscountAmount, @Taxtype, @TaxPercentage, @TaxAmount, 
                @TotalAmount, @CreatedBy, GETDATE(), @BnchID, @LocID);";

                        using (SqlCommand cmd = new SqlCommand(insertSalesDetailsQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@SRReturnID", refNo);
                            cmd.Parameters.AddWithValue("@Items", productName);
                            cmd.Parameters.AddWithValue("@Batch", batchNo);
                            cmd.Parameters.AddWithValue("@SerialNo", row.Cells[2].Value?.ToString() ?? DBNull.Value.ToString());
                            cmd.Parameters.AddWithValue("@Qty", qty);
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

                        // 2. Update TotalQty in Product table
                        string updateProductQtyQuery = @"
        UPDATE Product 
        SET TotalQty = ISNULL(TotalQty, 0) + @Qty 
        WHERE ProductName = @ProductName;";

                        using (SqlCommand updateProductCmd = new SqlCommand(updateProductQtyQuery, conn))
                        {
                            updateProductCmd.Parameters.AddWithValue("@Qty", qty);
                            updateProductCmd.Parameters.AddWithValue("@ProductName", productName);
                            updateProductCmd.ExecuteNonQuery();
                        }

                        // 3. Update TotalQty in StockIn table
                        string updateStockInQtyQuery = @"
        UPDATE StockIn 
        SET TotalQty = ISNULL(TotalQty, 0) + @Qty 
        WHERE ProductName = @ProductName AND batch = @BatchNo;";

                        using (SqlCommand updateStockCmd = new SqlCommand(updateStockInQtyQuery, conn))
                        {
                            updateStockCmd.Parameters.AddWithValue("@Qty", qty);
                            updateStockCmd.Parameters.AddWithValue("@ProductName", productName);
                            updateStockCmd.Parameters.AddWithValue("@BatchNo", batchNo);
                            updateStockCmd.ExecuteNonQuery();
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
                if (ctrl is System.Windows.Forms.TextBox textBox && textBox != textBox5)
                {
                    textBox.Text = "";
                }


            }


            label18.Text = "";
            label19.Text = "";
            label20.Text = "";
            label21.Text = "";
            label22.Text = "";
            label23.Text = "";
            label17.Text = "";


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


            comboBox1.Focus();
        }

        #endregion

        #region validation
        private bool ValidateFields()
        {
            
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Customer State is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }


            return true;

        }

      
        #endregion

        
    }
}
