using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;

namespace Estimate
{
    public partial class Form2 : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
        DataTable productTable;
        DataSet ds;
        public Form2()
        {
            InitializeComponent();
          
        }
   

        private void button1_Click(object sender, EventArgs e)
        {
            LoadEstimateDetails();
        }

        #region data fetch in grid
        private void LoadEstimateDetails()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    // Ensure textBox1 contains a valid number
                    if (!int.TryParse(textBox1.Text.Trim(), out int invoiceNo))
                    {
                        MessageBox.Show("Please enter a valid Invoice Number.");
                        return;
                    }

                    string query = @"
                SELECT 
                    Items,temID, batch, serialno, Qty, TaxType, HSNCode, 
                    UnitType, Price, Discount, TaxPercentage, TaxAmount, TotalAmount, 
                    DiscountAmount 
                FROM estimate_Details 
                WHERE InvoNo = @InvoNo";  // Filtering by InvoNo

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@InvoNo", invoiceNo);  // Prevents SQL injection

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            dataGridView1.Rows.Clear(); // Clear previous data before adding new rows

                            if (dataGridView1.Columns.Count > 0)
                            {
                                foreach (DataRow row in dt.Rows)
                                {
                                    int rowIndex = dataGridView1.Rows.Add();
                                    dataGridView1.Rows[rowIndex].Cells["Column1"].Value = row["Items"];
                                    dataGridView1.Rows[rowIndex].Cells["Column15"].Value = row["temID"];
                                    dataGridView1.Rows[rowIndex].Cells["Column2"].Value = row["batch"];
                                    dataGridView1.Rows[rowIndex].Cells["Column3"].Value = row["serialno"];
                                    dataGridView1.Rows[rowIndex].Cells["Column4"].Value = row["Qty"];
                                    dataGridView1.Rows[rowIndex].Cells["Column5"].Value = row["TaxType"];
                                    dataGridView1.Rows[rowIndex].Cells["Column6"].Value = row["HSNCode"];
                                    dataGridView1.Rows[rowIndex].Cells["Column7"].Value = row["UnitType"];
                                    dataGridView1.Rows[rowIndex].Cells["Column8"].Value = row["Price"];
                                    dataGridView1.Rows[rowIndex].Cells["Column9"].Value = row["Discount"];
                                    dataGridView1.Rows[rowIndex].Cells["Column10"].Value = row["TaxPercentage"];
                                    dataGridView1.Rows[rowIndex].Cells["Column14"].Value = row["TaxAmount"];
                                    dataGridView1.Rows[rowIndex].Cells["Column11"].Value = row["TotalAmount"];
                                    dataGridView1.Rows[rowIndex].Cells["Column13"].Value = row["DiscountAmount"];

                                    // Call calculation method for each row
                                    CalculateRowValues(dataGridView1.Rows[rowIndex]);
                                }

                                // Recalculate total after all rows are added
                                CalculateTotal();
                            }
                            else
                            {
                                MessageBox.Show("Please define columns in DataGridView first.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data: " + ex.Message);
            }
        }
        #endregion
      
        
        private void CalculateRowValues(DataGridViewRow row)
        {
            try
            {
                // Ensure required values are not null
                if (row.Cells["Column4"].Value != null && row.Cells["Column8"].Value != null && row.Cells["Column9"].Value != null)
                {
                    decimal quantity = Convert.ToDecimal(row.Cells["Column4"].Value);
                    decimal rate = Convert.ToDecimal(row.Cells["Column8"].Value);
                    decimal discountPercent = Convert.ToDecimal(row.Cells["Column9"].Value);
                    decimal taxPercent = Convert.ToDecimal(row.Cells["Column10"].Value ?? 0);

                    // Calculate total before discount
                    decimal total = quantity * rate;

                    // Calculate discount amount
                    decimal discountAmount = (total * discountPercent) / 100;
                    row.Cells["Column13"].Value = discountAmount;

                    // Calculate net amount after discount
                    decimal netAmount = total - discountAmount;

                    // Calculate tax amount
                    decimal taxAmount = (netAmount * taxPercent) / 100;
                    row.Cells["Column14"].Value = taxAmount;

                    // Final Amount = Net Amount + Tax
                    decimal finalAmount = netAmount + taxAmount;
                    row.Cells["Column11"].Value = finalAmount;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in row calculation: " + ex.Message);
            }
        }
        private void CalculateTotal()
        {
            try
            {
                decimal grandTotal = 0;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells["Column11"].Value != null)
                    {
                        grandTotal += Convert.ToDecimal(row.Cells["Column11"].Value);
                    }
                }

                // Show total in a label or textbox
                label3.Text =grandTotal.ToString("0.00");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error calculating total: " + ex.Message);
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // Get the edited row
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                // Check if the edited column is Column4 (Qty)
                if (e.ColumnIndex == dataGridView1.Columns["Column4"].Index)
                {
                    CalculateRowValues(row);
                    CalculateTotal();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        #region Insert 
        private void button2_Click(object sender, EventArgs e)
        {
            UpdateData();
        }
        private void UpdateData()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    if (!int.TryParse(textBox1.Text.Trim(), out int invoiceNo))
                    {
                        MessageBox.Show("Please enter a valid Invoice Number.");
                        return;
                    }

                    // **Update Estimate Table**
                    string updateEstimateQuery = @"
                UPDATE Estimate 
                SET 
                   
                    EstimateTotal = @EstimateTotal
                    
                WHERE RefNo = @InvoNo";

                    using (SqlCommand cmd = new SqlCommand(updateEstimateQuery, conn))
                    {
                        
                        cmd.Parameters.AddWithValue("@EstimateTotal", Convert.ToDecimal(label3.Text));
                        
                        cmd.Parameters.AddWithValue("@InvoNo", invoiceNo);

                        cmd.ExecuteNonQuery();
                    }

                    // **Update Estimate_Details Table**
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null) continue;

                        string updateEstimateDetailsQuery = @"
                    UPDATE Estimate_Details 
                    SET 
                        Items = @Items,
                        batch = @Batch,
                        serialno = @SerialNo,
                        HSNCode = @HSNCode,
                        Qty = @Qty,
                        UnitType = @UnitType,
                        Price = @Price,
                        Discount = @Discount,
                        DiscountAmount = @DiscountAmount,
                        Taxtype = @Taxtype,
                        TaxPercentage = @TaxPercentage,
                        TaxAmount = @TaxAmount,
                        TotalAmount = @TotalAmount,
                        ModifiedBy = @UpdatedBy,
                    ModifiedDate = GETDATE()
                    WHERE InvoNo = @InvoNo AND Items = @Items";

                        using (SqlCommand cmd = new SqlCommand(updateEstimateDetailsQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@InvoNo", invoiceNo);
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
                            cmd.Parameters.AddWithValue("@UpdatedBy", 1);

                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Updated Successfully");
                    ResetFormFields();
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        #endregion
        private void button3_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        private void ResetFormFields()
        {
            
            textBox1.Clear();
            label3.Text ="0.00";

            dataGridView1.Rows.Clear();

           
            textBox1.Focus();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex >= 0)
            {
                // Get the unique identifier for the row
                int tempID;
                if (dataGridView1.Rows[e.RowIndex].Cells["Column15"].Value != null &&
                    int.TryParse(dataGridView1.Rows[e.RowIndex].Cells["Column15"].Value.ToString(), out tempID))
                {
                    DialogResult result = MessageBox.Show("Are you sure you want to delete this row from the database?",
                                                          "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                            {
                                conn.Open();

                                // Delete the row from the database using tempID (unique row identifier)
                                string deleteQuery = "DELETE FROM estimate_Details WHERE temID = @tempID";
                                using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                                {
                                    cmd.Parameters.AddWithValue("@tempID", tempID);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                           
                            dataGridView1.Rows.RemoveAt(e.RowIndex);

                          
                            CalculateTotal();
                            UpdateData();

                            MessageBox.Show("Row deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error deleting row: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Invalid Row Identifier. Cannot delete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
    }
}
