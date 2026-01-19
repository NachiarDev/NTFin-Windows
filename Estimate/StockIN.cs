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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Estimate
{
    public partial class StockIN : Form
    {
        public StockIN()
        {
            InitializeComponent();
        }

        private void StockIN_Load(object sender, EventArgs e)
        {
            textBox5.Visible = false;
            LoadProductName();
            textBox3.KeyPress += OnlyAllowDigits;
            textBox4.KeyPress += OnlyAllowDigits;
        }

        #region Combo box
        private void LoadProductName()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, ProductName FROM Product";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "ProductName";  // Show country name
                        comboBox1.ValueMember = "ProductName";            // Store country ID
                        comboBox1.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Product Name: " + ex.Message);
            }
        }
        #endregion

        #region Validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Batch  is required" },
       { textBox2, "Serial Number is required" },
        { textBox3, "Price is required" },
        { textBox4, "Total Qty is required" },
    
    };

            foreach (var field in fields)
            {
                if (string.IsNullOrWhiteSpace(field.Key.Text))
                {
                    MessageBox.Show(field.Value, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    field.Key.Focus();
                    return false;
                }
            }

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Product Name is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }

           


            return true;
        }
        private void OnlyAllowDigits(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        #endregion


        #region Reset
        private void ResetFormFields()
        {

            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textBox)
                {
                    textBox.Text = "";
                }
            }
           
            comboBox1.SelectedIndex = -1;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();

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
                if (!ValidateFields()) return;

                // Get ID from textBox5 (if it's empty, we will treat it as a new entry)
                int id = string.IsNullOrEmpty(textBox5.Text) ? 0 : Convert.ToInt32(textBox5.Text);

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    if (id == 0) // If ID is 0, it's an insert operation
                    {
                        // Check for duplicate entry before inserting
                        string checkQuery = "SELECT COUNT(*) FROM StockIN WHERE ProductName = @ProductName AND serialno=@serialno";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue ?? DBNull.Value);
                            checkCmd.Parameters.AddWithValue("@serialno", textBox2.Text);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This ProductName already exists with this Serial Number. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        // Insert new data into StockIN
                        string insertStockinQuery = @"INSERT INTO StockIN (ProductName, serialno, batch, price, TotalQty, BnchID, LocID)
                                              VALUES (@ProductName, @serialno, @batch, @price, @TotalQty, @BnchID, @LocID);
                                              SELECT SCOPE_IDENTITY();";
                        using (SqlCommand cmd = new SqlCommand(insertStockinQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@batch", textBox1.Text);
                            cmd.Parameters.AddWithValue("@serialno", textBox2.Text);
                            cmd.Parameters.AddWithValue("@price", textBox3.Text);
                            cmd.Parameters.AddWithValue("@TotalQty", textBox4.Text);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            object newIDObj = await cmd.ExecuteScalarAsync(); // Get the new inserted ID

                            // Update Product TotalQty
                            string updateProductQtyQuery = @"UPDATE Product SET TotalQty = TotalQty + @Qty WHERE ProductName = @ProductName";
                            using (SqlCommand cmdUpdate = new SqlCommand(updateProductQtyQuery, conn))
                            {
                                cmdUpdate.Parameters.AddWithValue("@Qty", Convert.ToDecimal(textBox4.Text));
                                cmdUpdate.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue?.ToString());
                                await cmdUpdate.ExecuteNonQueryAsync();
                            }
                        }

                        MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else // If ID exists, it's an update operation
                    {
                        // Check if the record exists to update
                        string checkUpdateQuery = "SELECT COUNT(*) FROM StockIN WHERE ID = @ID";
                        using (SqlCommand checkCmd = new SqlCommand(checkUpdateQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@ID", id);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count == 0)
                            {
                                MessageBox.Show("No record found with the specified ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        // Update the StockIN data
                        string updateStockinQuery = @"UPDATE StockIN SET ProductName = @ProductName, serialno = @serialno, batch = @batch,
                                              price = @price, TotalQty = @TotalQty, BnchID = @BnchID, LocID = @LocID
                                              WHERE ID = @ID";
                        using (SqlCommand cmd = new SqlCommand(updateStockinQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", id);
                            cmd.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@batch", textBox1.Text);
                            cmd.Parameters.AddWithValue("@serialno", textBox2.Text);
                            cmd.Parameters.AddWithValue("@price", textBox3.Text);
                            cmd.Parameters.AddWithValue("@TotalQty", textBox4.Text);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            await cmd.ExecuteNonQueryAsync();
                        }

                        // Update Product TotalQty
                        string updateProductQtyQuery = @"UPDATE Product SET TotalQty = TotalQty + @Qty WHERE ProductName = @ProductName";
                        using (SqlCommand cmdUpdate = new SqlCommand(updateProductQtyQuery, conn))
                        {
                            cmdUpdate.Parameters.AddWithValue("@Qty", Convert.ToDecimal(textBox4.Text));
                            cmdUpdate.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue?.ToString());
                            await cmdUpdate.ExecuteNonQueryAsync();
                        }

                        MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    ResetFormFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //        private async void InsertData()
        //        {
        //            try
        //            {

        //                if (!ValidateFields()) return;

        //                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
        //                {
        //                    await conn.OpenAsync();
        //                    string checkQuery = "SELECT COUNT(*) FROM StockIN WHERE ProductName = @ProductName AND serialno=@serialno";
        //                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
        //                    {
        //                        checkCmd.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue ?? DBNull.Value);
        //                        checkCmd.Parameters.AddWithValue("@serialno", textBox2.Text);
        //                        int count = (int)await checkCmd.ExecuteScalarAsync();

        //                        if (count > 0)
        //                        {
        //                            MessageBox.Show("This ProductName already exists with this Serial Number. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                            return;
        //                        }
        //                    }


        //                    //        string insertStockinQuery = @"
        //                    //INSERT INTO StockIN ( ProductName, serialno, batch, price, TotalQty,BnchID,LocID) 
        //                    //VALUES              (@ProductName, @serialno, @batch, @price,  @TotalQty,@BnchID,@LocID);

        //                    //SELECT SCOPE_IDENTITY();";

        //                    //        using (SqlCommand cmd = new SqlCommand(insertStockinQuery, conn))
        //                    //        {

        //                    //            cmd.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue ?? DBNull.Value);
        //                    //            cmd.Parameters.AddWithValue("@batch", textBox1.Text);
        //                    //            cmd.Parameters.AddWithValue("@serialno", textBox2.Text);
        //                    //            cmd.Parameters.AddWithValue("@price", textBox3.Text);
        //                    //            cmd.Parameters.AddWithValue("@TotalQty", textBox4.Text);
        //                    //            cmd.Parameters.AddWithValue("@BnchID", 1);
        //                    //            cmd.Parameters.AddWithValue("@LocID", 1);
        //                    //            object newIDObj = await cmd.ExecuteScalarAsync();
        //                    //        }

        //                    //        MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    //        ResetFormFields();
        //                    string insertStockinQuery = @"
        //INSERT INTO StockIN (ProductName, serialno, batch, price, TotalQty, BnchID, LocID) 
        //VALUES (@ProductName, @serialno, @batch, @price, @TotalQty, @BnchID, @LocID);
        //SELECT SCOPE_IDENTITY();";

        //                    using (SqlCommand cmd = new SqlCommand(insertStockinQuery, conn))
        //                    {
        //                        cmd.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue ?? DBNull.Value);
        //                        cmd.Parameters.AddWithValue("@batch", textBox1.Text);
        //                        cmd.Parameters.AddWithValue("@serialno", textBox2.Text);
        //                        cmd.Parameters.AddWithValue("@price", textBox3.Text);
        //                        cmd.Parameters.AddWithValue("@TotalQty", textBox4.Text);
        //                        cmd.Parameters.AddWithValue("@BnchID", 1);
        //                        cmd.Parameters.AddWithValue("@LocID", 1);

        //                        object newIDObj = await cmd.ExecuteScalarAsync();
        //                    }

        //                    // ✅ Now increase the Product.TotalQty for the same product
        //                    string updateProductQtyQuery = @"
        //UPDATE Product 
        //SET TotalQty = TotalQty + @Qty 
        //WHERE ProductName = @ProductName";

        //                    using (SqlCommand cmdUpdate = new SqlCommand(updateProductQtyQuery, conn))
        //                    {
        //                        cmdUpdate.Parameters.AddWithValue("@Qty", Convert.ToDecimal(textBox4.Text));
        //                        cmdUpdate.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue?.ToString());

        //                        await cmdUpdate.ExecuteNonQueryAsync();
        //                    }

        //                    MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    ResetFormFields();

        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            }
        //        }

        #endregion


    }
}
