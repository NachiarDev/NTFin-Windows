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
    public partial class Purchase : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
        DataTable productTable;
        DataSet ds;
        public Purchase()
        {
            InitializeComponent();
        }
        private void Purchase_Load(object sender, EventArgs e)
        {
            LoadSuppliers();
            LoadProductsIntoGridComboBox();

            textBox4.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }
        #region combo box
        private void LoadSuppliers()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT ID, SupplierName FROM Supplier";
                SqlDataAdapter da = new SqlDataAdapter(query, con); // ✅ FIXED
                DataTable dt = new DataTable();
                da.Fill(dt);

                comboBox1.DataSource = dt;
                comboBox1.DisplayMember = "SupplierName";
                comboBox1.ValueMember = "ID";
                comboBox1.SelectedIndex = -1;
                comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex < 0)
                return;

            // Safely get supplierId
            if (comboBox1.SelectedValue != null && int.TryParse(comboBox1.SelectedValue.ToString(), out int supplierId))
            {
                using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
                {
                    con.Open();
                    string query = "SELECT StateID, Code FROM Supplier WHERE ID = @ID";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@ID", supplierId);
                        SqlDataReader reader = cmd.ExecuteReader();

                        if (reader.Read())
                        {
                            textBox1.Text = reader["StateID"].ToString();
                            textBox2.Text = reader["Code"].ToString();
                        }
                    }
                }
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


                       
                    }




                    if (e.ColumnIndex == dataGridView1.Columns["Column4"].Index ||
      e.ColumnIndex == dataGridView1.Columns["Column8"].Index ||
      e.ColumnIndex == dataGridView1.Columns["Column9"].Index ||
      e.ColumnIndex == dataGridView1.Columns["Column11"].Index)
                    {
                        if (row.Cells["Column4"].Value != null &&
                            row.Cells["Column8"].Value != null &&
                            row.Cells["Column9"].Value != null)
                        {
                            int quantity = Convert.ToInt32(row.Cells["Column4"].Value);
                            decimal rate = Convert.ToDecimal(row.Cells["Column8"].Value);
                            decimal discountPercent = Convert.ToDecimal(row.Cells["Column9"].Value);

                            decimal total = quantity * rate;
                            decimal discountAmount = (total * discountPercent) / 100;
                            row.Cells["Column10"].Value = Math.Round(discountAmount, 2); // ⭐ Rounded

                            decimal netAmount = total - discountAmount;

                            decimal additionalDiscountAmount = 0;
                            if (row.Cells["Column11"].Value != null)
                            {
                                decimal additionalDiscountPercent = Convert.ToDecimal(row.Cells["Column11"].Value);
                                additionalDiscountAmount = (netAmount * additionalDiscountPercent) / 100;
                                row.Cells["Column12"].Value = Math.Round(additionalDiscountAmount, 2); // ⭐ Rounded
                            }
                            else
                            {
                                row.Cells["Column12"].Value = 0;
                            }

                            decimal finalAmount = netAmount - additionalDiscountAmount;
                            row.Cells["Column13"].Value = Math.Round(finalAmount, 2); // ⭐ Rounded

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

            label18.Text = totalQuantity.ToString("0");
            textBox3.Text = totalAmount.ToString("0.00");
            label12.Text = subtotal.ToString("0.00"); // Updated to show the correct subtotal
            label13.Text = totalDiscount.ToString("0.00");
            label14.Text = totalTax.ToString("0.00");
            label15.Text = totalIGST.ToString("0.00"); // SGST Total
            label16.Text = totalSGST.ToString("0.00"); // CGST Total
            label17.Text = totalCGST.ToString("0.00"); // IGST Total

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
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();



                    string insertPurchaseQuery = @"
            INSERT INTO Purchase (RefNo,Supplier, RefDate, SupplierCode, StatePurchase, PurchaseTotal, 
                                  CreatedBy, CreatedDate, BnchID, LocID) 
            VALUES (0,@Supplier, @RefDate, @SupplierCode, @StatePurchase, @PurchaseTotal, 
                    @CreatedBy, GETDATE(), @BnchID, @LocID);
            SELECT SCOPE_IDENTITY();";

                    int refNo;
                    using (SqlCommand cmd = new SqlCommand(insertPurchaseQuery, conn))
                    {

                        cmd.Parameters.AddWithValue("@Supplier", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@RefDate", textBox4.Text);
                        cmd.Parameters.AddWithValue("@SupplierCode", textBox2.Text);
                        cmd.Parameters.AddWithValue("@StatePurchase", textBox1.Text);
                        cmd.Parameters.AddWithValue("@PurchaseTotal", textBox3.Text);
                        cmd.Parameters.AddWithValue("@CreatedBy", 1);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        object newIDObj = cmd.ExecuteScalar();
                        refNo = Convert.ToInt32(newIDObj);
                    }





                    await Task.Delay(10);



                    string updatePurchaseQuery = "UPDATE Purchase SET RefNo = @RefNo WHERE ID = @ID";
                    using (SqlCommand updateCmd = new SqlCommand(updatePurchaseQuery, conn))
                    {
                        updateCmd.Parameters.AddWithValue("@RefNo", refNo);
                        updateCmd.Parameters.AddWithValue("@ID", refNo);
                        updateCmd.ExecuteNonQuery();
                    }


                    if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows.Cast<DataGridViewRow>().All(r => r.IsNewRow))
                    {
                        MessageBox.Show("No data available in the grid to insert into Purchase_Details!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }



                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null)
                            continue;

                        // Generate a random ID for the Purchase Detail
                        Random rnd = new Random();
                        int tempID = rnd.Next(100000, 999999);

                        string batch = row.Cells[1].Value?.ToString() ?? DBNull.Value.ToString();
                        string serialNo = row.Cells[2].Value?.ToString() ?? DBNull.Value.ToString();
                        string productName = row.Cells[0].Value?.ToString() ?? DBNull.Value.ToString();
                        decimal qty = Convert.ToDecimal(row.Cells[3].Value ?? 0);
                        decimal price = Convert.ToDecimal(row.Cells[7].Value ?? 0);
                        decimal discount = Convert.ToDecimal(row.Cells[8].Value ?? 0);
                        decimal discountAmount = Convert.ToDecimal(row.Cells[9].Value ?? 0);
                        decimal taxPercentage = Convert.ToDecimal(row.Cells[10].Value ?? 0);
                        decimal taxAmount = Convert.ToDecimal(row.Cells[11].Value ?? 0);
                        decimal totalAmount = Convert.ToDecimal(row.Cells[12].Value ?? 0);
                        string hsnCode = row.Cells[5].Value?.ToString() ?? DBNull.Value.ToString();
                        string unitType = row.Cells[6].Value?.ToString() ?? DBNull.Value.ToString();
                        string taxtype = row.Cells[4].Value?.ToString() ?? DBNull.Value.ToString();

                        // Insert Purchase Detail into Purchase_Detail table
                        string insertPurchaseDetailQuery = @"
        INSERT INTO Purchase_Detail (temID, InvoNo, Items, batch, serialno, HSNCode, Qty, UnitType, 
                                      Price, Discount, DiscountAmount, Taxtype, TaxPercentage, TaxAmount, 
                                      TotalAmount, CreatedBy, CreatedDate, BnchID, LocID) 
        VALUES (@tempID, @RefNo, @Items, @Batch, @SerialNo, @HSNCode, @Qty, @UnitType, 
                @Price, @Discount, @DiscountAmount, @Taxtype, @TaxPercentage, @TaxAmount, 
                @TotalAmount, @CreatedBy, GETDATE(), @BnchID, @LocID);";

                        using (SqlCommand cmd = new SqlCommand(insertPurchaseDetailQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@tempID", tempID);
                            cmd.Parameters.AddWithValue("@RefNo", refNo);
                            cmd.Parameters.AddWithValue("@Items", productName);
                            cmd.Parameters.AddWithValue("@Batch", batch);
                            cmd.Parameters.AddWithValue("@SerialNo", serialNo);
                            cmd.Parameters.AddWithValue("@HSNCode", hsnCode);
                            cmd.Parameters.AddWithValue("@Qty", qty);
                            cmd.Parameters.AddWithValue("@Taxtype", taxtype);
                            cmd.Parameters.AddWithValue("@UnitType", unitType);
                            cmd.Parameters.AddWithValue("@Price", price);
                            cmd.Parameters.AddWithValue("@Discount", discount);
                            cmd.Parameters.AddWithValue("@DiscountAmount", discountAmount);
                            cmd.Parameters.AddWithValue("@TaxPercentage", taxPercentage);
                            cmd.Parameters.AddWithValue("@TaxAmount", taxAmount);
                            cmd.Parameters.AddWithValue("@TotalAmount", totalAmount);
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            cmd.ExecuteNonQuery();
                        }

                        // Check if the batch and serial number already exist in the StockIn table
                        string checkStockQuery = "SELECT COUNT(*) FROM StockIn WHERE Batch = @Batch AND SerialNo = @SerialNo";

                        using (SqlCommand cmd = new SqlCommand(checkStockQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@Batch", batch);
                            cmd.Parameters.AddWithValue("@SerialNo", serialNo);

                            int count = (int)cmd.ExecuteScalar();

                            // If the batch and serial number don't exist, insert a new record
                            if (count == 0)
                            {
                                string insertStockQuery = @"
                INSERT INTO StockIn (ProductName, Batch, SerialNo, Price, TotalQty, BnchID, LocID)
                VALUES (@ProductName, @Batch, @SerialNo, @Price, @TotalQty, @BnchID, @LocID);";

                                using (SqlCommand cmdInsertStock = new SqlCommand(insertStockQuery, conn))
                                {
                                    cmdInsertStock.Parameters.AddWithValue("@ProductName", productName);
                                    cmdInsertStock.Parameters.AddWithValue("@Batch", batch);
                                    cmdInsertStock.Parameters.AddWithValue("@SerialNo", serialNo);
                                    cmdInsertStock.Parameters.AddWithValue("@Price", price);
                                    cmdInsertStock.Parameters.AddWithValue("@TotalQty", qty);
                                    cmdInsertStock.Parameters.AddWithValue("@BnchID", 1);
                                    cmdInsertStock.Parameters.AddWithValue("@LocID", 1);

                                    cmdInsertStock.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                // If the batch and serial number exist, update the TotalQty by adding the new Qty
                                string updateStockQuery = @"
                UPDATE StockIn
                SET TotalQty = TotalQty + @Qty
                WHERE ProductName = @ProductName AND Batch = @Batch AND SerialNo = @SerialNo;";

                                using (SqlCommand cmdUpdateStock = new SqlCommand(updateStockQuery, conn))
                                {
                                    cmdUpdateStock.Parameters.AddWithValue("@Qty", qty);
                                    cmdUpdateStock.Parameters.AddWithValue("@ProductName", productName);
                                    cmdUpdateStock.Parameters.AddWithValue("@Batch", batch);
                                    cmdUpdateStock.Parameters.AddWithValue("@SerialNo", serialNo);

                                    cmdUpdateStock.ExecuteNonQuery();
                                }
                            }
                        }

                        // Now, also update the TotalQty for the product in the Product table
                        string updateProductQuery = @"
        UPDATE Product
        SET TotalQty = TotalQty + @Qty
        WHERE ProductName = @ProductName;";

                        using (SqlCommand cmdUpdateProduct = new SqlCommand(updateProductQuery, conn))
                        {
                            cmdUpdateProduct.Parameters.AddWithValue("@Qty", qty);
                            cmdUpdateProduct.Parameters.AddWithValue("@ProductName", productName);

                            cmdUpdateProduct.ExecuteNonQuery();
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
                if (ctrl is System.Windows.Forms.TextBox textBox && textBox != textBox4)
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
            label18.Text = "";


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
    

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Party Name is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }
            return true;

        }


        #endregion

        
    }
}
