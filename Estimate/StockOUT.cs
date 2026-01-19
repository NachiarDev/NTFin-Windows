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
    public partial class StockOUT : Form
    {
        public StockOUT()
        {
            InitializeComponent();
        }

        private void StockOUT_Load(object sender, EventArgs e)
        {
            LoadProductName();
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        { // Clear batch and text fields when product is changed
            comboBox2.DataSource = null;
            textBox1.Clear(); // serialno
            textBox3.Clear(); // price
            textBox2.Clear(); // totalQty

            if (comboBox1.SelectedIndex != -1)
            {
                string productName = comboBox1.SelectedValue.ToString();

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT DISTINCT batch FROM StockIN WHERE ProductName = @ProductName";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ProductName", productName);

                        DataTable dt = new DataTable();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);

                        comboBox2.DisplayMember = "batch";
                        comboBox2.ValueMember = "batch";
                        comboBox2.DataSource = dt;
                        comboBox2.SelectedIndex = -1; // Don't auto-select anything


                    }
                }

            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1)
            {
                string productName = comboBox1.SelectedValue.ToString();
                string batch = comboBox2.SelectedValue.ToString();

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = @"SELECT TOP 1 serialno, price, TotalQty 
                             FROM StockIN 
                             WHERE ProductName = @ProductName AND batch = @Batch";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ProductName", productName);
                        cmd.Parameters.AddWithValue("@Batch", batch);

                        SqlDataReader reader = cmd.ExecuteReader();
                        if (reader.Read())
                        {
                            textBox1.Text = reader["serialno"]?.ToString();
                            textBox2.Text = reader["price"]?.ToString();
                            textBox3.Text = reader["TotalQty"]?.ToString();
                        }
                    }
                }
            }
        }
        #endregion


        #region Validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        
        { textBox3, "Total Quantity is required" },
        

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
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Batch is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Focus();
                return false;
            }



            return true;
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
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
            comboBox2.SelectedIndex = -1;
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

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();
                   


                  
                    string insertStockinQuery = @"
INSERT INTO StockOUT (ProductName, serialno, batch, price, TotalQty, BnchID, LocID) 
VALUES (@ProductName, @serialno, @batch, @price, @TotalQty, @BnchID, @LocID);
SELECT SCOPE_IDENTITY();";

                    using (SqlCommand cmd = new SqlCommand(insertStockinQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@ProductName", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@batch", comboBox2.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@serialno", textBox1.Text);
                        cmd.Parameters.AddWithValue("@price", textBox2.Text);
                        cmd.Parameters.AddWithValue("@TotalQty", textBox3.Text);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        object newIDObj = await cmd.ExecuteScalarAsync();
                    }

                   

                    MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ResetFormFields();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        #endregion

       
    }
}
