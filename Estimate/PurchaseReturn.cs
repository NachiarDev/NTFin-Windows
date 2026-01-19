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
    public partial class PurchaseReturn : Form
    {
        public PurchaseReturn()
        {
            InitializeComponent();
        }

        private void PurchaseReturn_Load(object sender, EventArgs e)
        {
            LoadInvoice();
            LoadModeofreturn();
            LoadProductsIntoGridComboBox();
            textBox5.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        }
        #region ComboBox
        private void LoadInvoice()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT ID, RefNo FROM Purchase";
                SqlDataAdapter da = new SqlDataAdapter(query, con); 
                DataTable dt = new DataTable();
                da.Fill(dt);

                comboBox1.DataSource = dt;
                comboBox1.DisplayMember = "RefNO";
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
                    string query = "SELECT RefDate,SupplierCode,Supplier,StatePurchase FROM Purchase WHERE ID = @ID";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@ID", supplierId);
                        SqlDataReader reader = cmd.ExecuteReader();

                        if (reader.Read())
                        {
                            textBox1.Text = reader["RefDate"].ToString();
                            textBox2.Text = reader["StatePurchase"].ToString();
                            textBox3.Text = reader["SupplierCode"].ToString();
                            textBox4.Text = reader["Supplier"].ToString();
                        }
                    }
                }
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
                    string query = "SELECT ID, ProductName,Rate FROM Product";
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
                            string query = "SELECT Rate FROM Product WHERE ProductName = @ProductName";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@ProductName", selectedProductName);
                                SqlDataReader reader = cmd.ExecuteReader();

                                if (reader.Read())
                                {
                                    row.Cells["Column3"].Value = reader["Rate"].ToString();
                                }
                            }
                        }
                    }

                    if (e.ColumnIndex == dataGridView1.Columns["Column2"].Index || e.ColumnIndex == dataGridView1.Columns["Column3"].Index)
                    {
                        decimal qty = 0, rate = 0;

                        decimal.TryParse(Convert.ToString(row.Cells["Column2"].Value), out qty);
                        decimal.TryParse(Convert.ToString(row.Cells["Column3"].Value), out rate);

                        decimal total = qty * rate;
                        row.Cells["Column4"].Value = total.ToString("0.00");
                        UpdateTotalQuantity();
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
            if (e.RowIndex >= 0 && e.ColumnIndex == dataGridView1.Columns["Column5"].Index)
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
                            UpdateTotalQuantity();
                        }
                    }
                }
            }
        }

        private void UpdateTotalQuantity()
        {
            int totalQty = 0;
            decimal estimatetotal = 0.00m;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                if (int.TryParse(Convert.ToString(row.Cells["Column2"].Value), out int qty))
                {
                    totalQty += qty;
                }

                if (decimal.TryParse(Convert.ToString(row.Cells["Column4"].Value), out decimal totalamount))
                {
                    estimatetotal += totalamount;
                }
            }

            label17.Text = totalQty.ToString();
            textBox6.Text = estimatetotal.ToString("0.00");
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



                    string insertPurchasereturnQuery = @"
            INSERT INTO PurchaseReturn (ReturnID,ReturnDate,RefNo,SupplierName, RefDate, SupplierCode, StatePurchase, EstimateTotal, 
                                  CreatedBy, CreatedDate, BnchID, LocID,ModeofReturn,Description) 
            VALUES (0,@ReturnDate,@RefNo,@SupplierName, @RefDate, @SupplierCode, @StatePurchase, @EstimateTotal, 
                    @CreatedBy, GETDATE(), @BnchID, @LocID,@ModeofReturn,@Description);
            SELECT SCOPE_IDENTITY();";

                    int refNo;
                    using (SqlCommand cmd = new SqlCommand(insertPurchasereturnQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@ReturnDate", textBox5.Text);
                        cmd.Parameters.AddWithValue("@RefNo", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@RefDate", textBox1.Text);
                        cmd.Parameters.AddWithValue("@SupplierCode", textBox3.Text);
                        cmd.Parameters.AddWithValue("@ModeofReturn", comboBox2.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Description", richTextBox1.Text);
                        cmd.Parameters.AddWithValue("@StatePurchase", textBox2.Text);
                        cmd.Parameters.AddWithValue("@SupplierName", textBox4.Text);
                        cmd.Parameters.AddWithValue("@EstimateTotal", estimateTotal);
                        cmd.Parameters.AddWithValue("@CreatedBy", 1);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        object newIDObj = cmd.ExecuteScalar();
                        refNo = Convert.ToInt32(newIDObj);
                    }





                    await Task.Delay(10);



                    string updateSalesreturnQuery = "UPDATE PurchaseReturn SET ReturnID = @ReturnID WHERE ID = @ID";
                    using (SqlCommand updateCmd = new SqlCommand(updateSalesreturnQuery, conn))
                    {
                        updateCmd.Parameters.AddWithValue("@ReturnID", refNo);
                        updateCmd.Parameters.AddWithValue("@ID", refNo);
                        updateCmd.ExecuteNonQuery();
                    }


                    if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows.Cast<DataGridViewRow>().All(r => r.IsNewRow))
                    {
                        MessageBox.Show("No data available in the grid to insert into PurchaseReturn Details!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null)
                            continue;
                        string productName = row.Cells[0].Value?.ToString() ?? "";
                       
                        decimal qty = Convert.ToDecimal(row.Cells[1].Value ?? 0);

                        // 1. Insert into SalesReturn_Details
                        string insertPurchaseQuery = @"
        INSERT INTO Purchasereturn_Details (PRReturnID,ProductName, price,Qty,TotalAmount,BnchID, LocID)                              
        VALUES (@PRReturnID, @ProductName, @price, @Qty, @TotalAmount, @BnchID, @LocID);";

                        using (SqlCommand cmd = new SqlCommand(insertPurchaseQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@PRReturnID", refNo);
                            cmd.Parameters.AddWithValue("@ProductName", productName);
                            
                            cmd.Parameters.AddWithValue("@Qty", qty);
                            
                            cmd.Parameters.AddWithValue("@price", Convert.ToDecimal(row.Cells[2].Value ?? 0));
                            cmd.Parameters.AddWithValue("@TotalAmount", Convert.ToDecimal(row.Cells[3].Value ?? 0));
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            cmd.ExecuteNonQuery();
                        }

                        // 2. Update TotalQty in Product table
                        string updateProductQtyQuery = @"
        UPDATE Product 
        SET TotalQty = ISNULL(TotalQty, 0) - @Qty 
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
        SET TotalQty = ISNULL(TotalQty, 0) - @Qty 
        WHERE ProductName = @ProductName";

                        using (SqlCommand updateStockCmd = new SqlCommand(updateStockInQtyQuery, conn))
                        {
                            updateStockCmd.Parameters.AddWithValue("@Qty", qty);
                            updateStockCmd.Parameters.AddWithValue("@ProductName", productName);
                           
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
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;

            richTextBox1.Text = "";


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
                MessageBox.Show("Invoice Number is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }

            return true;

        }






        #endregion

    }
}
