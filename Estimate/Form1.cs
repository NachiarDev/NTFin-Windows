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
using QRCoder;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Estimate
{
    public partial class Form1 : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
         DataTable productTable;
        DataSet ds;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            this.estimateTableAdapter.Fill(this.estimateDataSet.Estimate);

            conn = new SqlConnection();
            conn.ConnectionString = "Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;";

            MessageBox.Show("ConnectionSuccessful");
            textBox3.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            
            LoadProductsIntoGridComboBox();

            dataGridView2.CellClick += dataGridView2_CellClick;


            radioButton1.TabIndex = 0;  // Initial Focus
            comboBox1.TabIndex = 1;      // Next
            checkBox1.TabIndex = 2;      // Next
            dataGridView2.TabIndex = 3;  // Move into DataGridView
            button2.TabIndex = 4;        // Final focus

            this.BeginInvoke((MethodInvoker)delegate {
                radioButton1.Focus();
            });



        }

        #region  the grid
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

                    if (dataGridView2.Columns["Column1"] is DataGridViewComboBoxColumn comboBoxColumn)
                    {
                        comboBoxColumn.DataSource = dt;
                        comboBoxColumn.DisplayMember = "ProductName"; // Show Product Name
                        comboBoxColumn.ValueMember = "ProductName"; // Store Product ID
                        dataGridView2.CellValueChanged += dataGridView2_CellValueChanged;
                        dataGridView2.DataError += dataGridView2_DataError;
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


        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
               
                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[e.RowIndex];

                    if (e.ColumnIndex == dataGridView2.Columns["Column1"].Index)
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
                                    row.Cells["Column10"].Value = reader["Tax"].ToString();
                                }
                            }
                        }

                       
                        LoadBatchForProduct(selectedProductName, row);
                    }

                   
                    if (e.ColumnIndex == dataGridView2.Columns["Column2"].Index)
                    {
                        if (row.Cells["Column2"].Value == null)
                            return;

                        string selectedBatch = row.Cells["Column2"].Value.ToString();
                        string selectedProductName = row.Cells["Column1"].Value.ToString();
                        AutoFillSerialNo(selectedProductName, selectedBatch, row);
                    }

                    if (e.ColumnIndex == dataGridView2.Columns["Column4"].Index || e.ColumnIndex == dataGridView2.Columns["Column8"].Index || e.ColumnIndex == dataGridView2.Columns["Column9"].Index || e.ColumnIndex == dataGridView2.Columns["Column10"].Index)
                    {
                        if (row.Cells["Column4"].Value != null && row.Cells["Column8"].Value != null && row.Cells["Column9"].Value != null)
                        {
                            decimal quantity = Convert.ToDecimal(row.Cells["Column4"].Value);
                            decimal rate = Convert.ToDecimal(row.Cells["Column8"].Value);
                            decimal discountPercent = Convert.ToDecimal(row.Cells["Column9"].Value);

                          
                            decimal total = quantity * rate;
                            decimal discountAmount = (total * discountPercent) / 100;
                            row.Cells["Column13"].Value = discountAmount; 

                           
                            decimal netAmount = total - discountAmount;

                            
                            decimal additionalDiscountAmount = 0;
                            if (row.Cells["Column10"].Value != null)
                            {
                                decimal additionalDiscountPercent = Convert.ToDecimal(row.Cells["Column10"].Value);
                                additionalDiscountAmount = (netAmount * additionalDiscountPercent) / 100;
                                row.Cells["Column14"].Value = additionalDiscountAmount; 
                            }
                            else
                            {
                                row.Cells["Column14"].Value = 0;
                            }

                           
                            decimal finalAmount = netAmount + additionalDiscountAmount;
                            row.Cells["Column11"].Value = finalAmount;
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

        #endregion

        #region Insert and Reset and calculate


        private async void InsertData()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    if (!ValidateFields())
                        return;

                    int quotationName = radioButton1.Checked ? 1 : 0;
                    int customerName = radioButton2.Checked ? 1 : 0;

                    decimal estimateTotal = 0;
                    decimal.TryParse(textBox17.Text.Trim(), out estimateTotal);

                    string insertEstimateQuery = @"
            INSERT INTO Estimate (RefNo, QuotationName, CustomerName, RefDate, CustomerCode, StateSupply, EstimateTotal, 
                                  CreatedBy, CreatedDate, BnchID, LocID) 
            VALUES (0, @QuotationName, @CustomerName, @RefDate, @CustomerCode, @StateSupply, @EstimateTotal, 
                    @CreatedBy, GETDATE(), @BnchID, @LocID);
            SELECT SCOPE_IDENTITY();";

                    int refNo;
                    using (SqlCommand cmd = new SqlCommand(insertEstimateQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@QuotationName", quotationName);
                        cmd.Parameters.AddWithValue("@CustomerName", customerName);
                        cmd.Parameters.AddWithValue("@RefDate", textBox3.Text);
                        cmd.Parameters.AddWithValue("@CustomerCode", textBox2.Text);
                        cmd.Parameters.AddWithValue("@StateSupply", textBox1.Text);
                        cmd.Parameters.AddWithValue("@EstimateTotal", estimateTotal);
                        cmd.Parameters.AddWithValue("@CreatedBy", 1);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        object newIDObj = cmd.ExecuteScalar();
                        refNo = Convert.ToInt32(newIDObj);
                    }

                   
                    string qrImagePath = GenerateUPIQRCode(refNo, textBox17.Text);

                    if (!string.IsNullOrEmpty(qrImagePath))
                    {
                        
                        string updateQRPathQuery = "UPDATE Estimate SET QRImagePath = @QRImagePath WHERE ID = @ID";

                        using (SqlCommand updateCmd = new SqlCommand(updateQRPathQuery, conn))
                        {
                            updateCmd.Parameters.AddWithValue("@QRImagePath", qrImagePath);
                            updateCmd.Parameters.AddWithValue("@ID", refNo);
                            updateCmd.ExecuteNonQuery();
                        }
                    }

                    await Task.Delay(10);

                 
                    pictureBox1.Image = null;

                    string updateEstimateQuery = "UPDATE Estimate SET RefNo = @RefNo WHERE ID = @ID";
                    using (SqlCommand updateCmd = new SqlCommand(updateEstimateQuery, conn))
                    {
                        updateCmd.Parameters.AddWithValue("@RefNo", refNo);
                        updateCmd.Parameters.AddWithValue("@ID", refNo);
                        updateCmd.ExecuteNonQuery();
                    }


                    if (dataGridView2.Rows.Count == 0 || dataGridView2.Rows.Cast<DataGridViewRow>().All(r => r.IsNewRow))
                    {
                        MessageBox.Show("No data available in the grid to insert into Estimate_Details!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }


                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null)
                            continue;
                        Random rnd = new Random();
                        int tempID = rnd.Next(100000, 999999);
                        MessageBox.Show($"Inserting: Items={row.Cells[0].Value}, Qty={row.Cells[3].Value}, Price={row.Cells[7].Value}");

                        string insertEstimateDetailsQuery = @"
                INSERT INTO Estimate_Details (temID,InvoNo,Items, batch, serialno, HSNCode, Qty, UnitType, 
                                              Price, Discount, DiscountAmount, Taxtype, TaxPercentage, TaxAmount, 
                                              TotalAmount, CreatedBy, CreatedDate, BnchID, LocID) 
                VALUES (@tempID,@RefNo,@Items, @Batch, @SerialNo, @HSNCode, @Qty, @UnitType, 
                        @Price, @Discount, @DiscountAmount, @Taxtype, @TaxPercentage, @TaxAmount, 
                        @TotalAmount, @CreatedBy, GETDATE(), @BnchID, @LocID);";

                        using (SqlCommand cmd = new SqlCommand(insertEstimateDetailsQuery, conn))
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

                            cmd.Parameters.AddWithValue("@TaxPercentage", Convert.ToDecimal(row.Cells[10].Value ?? 0));
                            cmd.Parameters.AddWithValue("@TotalAmount", Convert.ToDecimal(row.Cells[12].Value ?? 0));
                            cmd.Parameters.AddWithValue("@DiscountAmount", Convert.ToDecimal(row.Cells[9].Value ?? 0));
                            cmd.Parameters.AddWithValue("@TaxAmount", Convert.ToDecimal(row.Cells[11].Value ?? 0));
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            cmd.ExecuteNonQuery();
                        }

                        string updateProductQtyQuery = "UPDATE Product SET TotalQty = TotalQty - @Qty WHERE ProductName = @ProductName";
                        using (SqlCommand cmdUpdate = new SqlCommand(updateProductQtyQuery, conn))
                        {
                            cmdUpdate.Parameters.AddWithValue("@Qty", Convert.ToDecimal(row.Cells[3].Value ?? 0));
                            cmdUpdate.Parameters.AddWithValue("@ProductName", row.Cells[0].Value?.ToString());
                            cmdUpdate.ExecuteNonQuery();
                        }
                        string updateStockinQtyQuery = "UPDATE StockIN SET TotalQty = TotalQty - @Qty WHERE batch = @batch";
                        using (SqlCommand cmdUpdate = new SqlCommand(updateStockinQtyQuery, conn))
                        {
                            cmdUpdate.Parameters.AddWithValue("@Qty", Convert.ToDecimal(row.Cells[3].Value ?? 0));
                            cmdUpdate.Parameters.AddWithValue("@batch", row.Cells[1].Value?.ToString());
                            cmdUpdate.ExecuteNonQuery();
                        }
                        string insertStockOutQuery = @"
INSERT INTO StockOut (ProductName, batch, serialno, Price, TotalQty, BnchID, LocID)
VALUES (@ProductName, @Batch, @SerialNo, @Price, @Qty, @BnchID, @LocID);";

                        using (SqlCommand cmdStockOut = new SqlCommand(insertStockOutQuery, conn))
                        {
                            cmdStockOut.Parameters.AddWithValue("@ProductName", row.Cells[0].Value?.ToString() ?? DBNull.Value.ToString());
                            cmdStockOut.Parameters.AddWithValue("@Batch", row.Cells[1].Value?.ToString() ?? DBNull.Value.ToString());
                            cmdStockOut.Parameters.AddWithValue("@SerialNo", row.Cells[2].Value?.ToString() ?? DBNull.Value.ToString());
                            cmdStockOut.Parameters.AddWithValue("@Price", Convert.ToDecimal(row.Cells[7].Value ?? 0));
                            cmdStockOut.Parameters.AddWithValue("@Qty", Convert.ToDecimal(row.Cells[3].Value ?? 0));
                            cmdStockOut.Parameters.AddWithValue("@BnchID", 1);
                            cmdStockOut.Parameters.AddWithValue("@LocID", 1);

                            cmdStockOut.ExecuteNonQuery();
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
        private void button2_Click(object sender, EventArgs e)
        {
            InsertData();

        }
        private void button3_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ResetFormFields();

        }
        private string GenerateUPIQRCode(int refNo, string amount)
        {
            try
            {
               
                string upiId = "7358590955@okbizaxis";
                string payeeName = "Nachiar Systems";

               
                if (!decimal.TryParse(amount, out decimal amt) || amt <= 0)
                {
                    MessageBox.Show("Invalid amount in EstimateTotal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }

                
                string upiUri = $"upi://pay?pa={upiId}&pn={Uri.EscapeDataString(payeeName)}&am={amt}&cu=INR";

              
                using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                using (QRCodeData qrCodeData = qrGenerator.CreateQrCode(upiUri, QRCodeGenerator.ECCLevel.Q))
                using (QRCode qrCode = new QRCode(qrCodeData))
                {
                    Bitmap qrImage = qrCode.GetGraphic(5);
                    pictureBox1.Image = qrImage;
                    pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;

                    
                    string folderPath = @"C:\Malathi\Estimate\QRCode";

                   
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    string filePath = Path.Combine(folderPath, $"QR_{refNo}.png");

                  
                    qrImage.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

                    MessageBox.Show($"QR Code saved at: {filePath}", "QR Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    return filePath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error generating QR Code: " + ex.Message);
                return null;
            }
        }


        private void ResetFormFields()
        {
           
            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textBox && textBox != textBox3)
                {
                    textBox.Text = "";
                }
                if (ctrl is System.Windows.Forms.RadioButton radioButton)
                {
                    radioButton.Checked = false;
                }
                
            }


            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
            label18.Text = "";
            label19.Text = "";
            label20.Text = "";
            label27.Text = "";

            comboBox1.SelectedIndex = -1; 

           
            richTextBox1.Clear();
            richTextBox2.Clear();

          
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (!row.IsNewRow)
                {
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (i != 13)
                        {
                            row.Cells[i].Value = null;
                        }
                    }
                }
            }

           
            textBox3.Focus();
        }
      
        #region Calculation Details
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

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                decimal taxAmount = 0; // Initialize taxAmount to prevent unassigned variable error

                if (row.Cells["Column4"].Value != null && decimal.TryParse(row.Cells["Column4"].Value.ToString(), out decimal quantity))
                {
                    totalQuantity += quantity;
                }
                if (row.Cells["Column11"].Value != null && decimal.TryParse(row.Cells["Column11"].Value.ToString(), out decimal amount))
                {
                    totalAmount += amount;
                }
                if (row.Cells["Column13"].Value != null && decimal.TryParse(row.Cells["Column13"].Value.ToString(), out decimal discountAmount))
                {
                    totalDiscount += discountAmount;
                }
                if (row.Cells["Column14"].Value != null && decimal.TryParse(row.Cells["Column14"].Value.ToString(), out taxAmount))
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

            label20.Text = totalQuantity.ToString("0");
            textBox17.Text = totalAmount.ToString("0.00");
               label14.Text = subtotal.ToString("0.00"); // Updated to show the correct subtotal
            label15.Text = totalDiscount.ToString("0.00");
            label16.Text = totalTax.ToString("0.00");
            label17.Text = totalSGST.ToString("0.00"); // SGST Total
            label18.Text = totalCGST.ToString("0.00"); // CGST Total
            label19.Text = totalIGST.ToString("0.00"); // IGST Total
            label27.Text = textBox17.Text;
        }
        #endregion


        #region Delete Button
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            
            
                if (e.RowIndex >= 0 && e.ColumnIndex == dataGridView2.Columns["Column12"].Index)
                {
                    
                    if (dataGridView2.IsCurrentRowDirty)
                    {
                        dataGridView2.EndEdit();

                       
                        DialogResult result = MessageBox.Show("Are you sure you want to delete this row?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            
                            if (!dataGridView2.Rows[e.RowIndex].IsNewRow)
                            {
                                dataGridView2.Rows.RemoveAt(e.RowIndex);
                            }
                        }
                    }
                }
           

        }
        #endregion

        #endregion

        #region Billing Address Prefill
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                LoadBillingAddress();
            }
            else
            {
                richTextBox2.Clear(); // Clear the textbox when unchecked
            }

        }
        private void LoadBillingAddress()
        {
            try
            {
                // Define the SQL query to get the Address from BillAddress table
                string query = "SELECT TOP 1 Address FROM BillAddress WHERE ID = 1"; // Adjust WHERE clause if needed

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                // Prefill the richtextbox2 with the address
                                richTextBox2.Text = dr["Address"].ToString();
                            }
                            else
                            {
                                richTextBox2.Text = "No Address Found!";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading billing address: " + ex.Message);
            }
        }
        #endregion     

        #region Party select
       
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                comboBox1.SelectedIndex = -1; // Reset selection
                comboBox1.Items.Clear(); // Clear combobox items
                textBox1.Clear(); // Clear related text fields
                textBox2.Clear();
                LoadQuotationCustomerNames();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                comboBox1.SelectedIndex = -1; // Reset selection
                comboBox1.Items.Clear(); // Clear combobox items
                textBox1.Clear(); // Clear related text fields
                textBox2.Clear();
                LoadCustomerCustomerNames();
            }
        }

        private void LoadQuotationCustomerNames()
        {
            LoadCustomerNames("SELECT DISTINCT CustomerName FROM Quotation");
        }

        private void LoadCustomerCustomerNames()
        {
            LoadCustomerNames("SELECT DISTINCT CustomerName FROM Customer");
        }

        private void LoadCustomerNames(string query)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            List<string> newItems = new List<string>();

                            while (dr.Read())
                            {
                                newItems.Add(dr["CustomerName"].ToString());
                            }

                            comboBox1.Items.Clear(); // Ensure comboBox is cleared before adding new items
                            comboBox1.Items.AddRange(newItems.ToArray());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading customer names: " + ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                string selectedCustomer = comboBox1.SelectedItem.ToString();
                LoadCustomerDetails(selectedCustomer);
            }
        }

        private void LoadCustomerDetails(string customerName)
        {
            try
            {
                string query = @"
SELECT c.StateID AS StateSupply, c.Code AS CustomerCode 
FROM Customer c WHERE c.CustomerName = @CustomerName
UNION
SELECT q.StateSupply, q.CustomerCode 
FROM Quotation q WHERE q.CustomerName = @CustomerName";

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@CustomerName", customerName);

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                textBox1.Text = dr["StateSupply"].ToString();
                                textBox2.Text = dr["CustomerCode"].ToString();
                            }
                            else
                            {
                                textBox1.Clear();
                                textBox2.Clear();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading customer details: " + ex.Message);
            }
        }
        #endregion

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }

        

        private void dataGridView2_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            CalculateTotal();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2(); // Create an instance of Form2
            newForm.ShowDialog();
        }
        #region validation
        private bool ValidateFields()
        {
           

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("party is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }






            return true;

        }

       

        #endregion


    }
}

