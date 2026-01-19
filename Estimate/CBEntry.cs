using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Estimate
{ 
    public partial class CBEntry : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
        DataTable productTable;
        DataSet ds;
        public CBEntry()
        {
            InitializeComponent();
        }


        #region Combo box & data fetch
        private void CBEntry_Load(object sender, EventArgs e)
        {
            LoadInvoice();
        }
        private void LoadInvoice()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = @"
            SELECT ID, 
                   CAST(RefNo AS VARCHAR) + ' - ' + CustomerCode AS DisplayText 
            FROM Estimate";

                SqlDataAdapter da = new SqlDataAdapter(query, con);
                DataTable dt = new DataTable();
                da.Fill(dt);

                comboBox1.DisplayMember = "DisplayText";
                comboBox1.ValueMember = "ID";
                comboBox1.DataSource = dt;
            }

            // Event hookup (ensure not duplicated)
            comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;

            textBox2.TextChanged -= textBox2_TextChanged;
            textBox2.TextChanged += textBox2_TextChanged;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue is int selectedId)
            {
                using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
                {
                    string query = "SELECT EstimateTotal FROM Estimate WHERE ID = @ID";
                    SqlCommand cmd = new SqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@ID", selectedId);
                    con.Open();
                    object result = cmd.ExecuteScalar();

                    if (result != null && decimal.TryParse(result.ToString(), out decimal total))
                    {
                        textBox1.Text = total.ToString("0.00");
                        UpdateBalance(); // recalculate balance
                    }
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            UpdateBalance();
        }
        private void UpdateBalance()
        {
            decimal total = 0, paid = 0;

            decimal.TryParse(textBox1.Text, out total);
            decimal.TryParse(textBox2.Text, out paid);

            decimal balance = total - paid;

            // Round to nearest whole number
            int roundedBalance = (int)Math.Round(balance, MidpointRounding.AwayFromZero);
            textBox3.Text = roundedBalance.ToString();
        }

        #endregion

        #region Reset
        private void ResetFormFields()
        {
            comboBox1.SelectedIndex = -1;
            textBox1.Text = "";
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            textBox2.Text = "";
            textBox3.Text = "";

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

                    bool isUpdate = !string.IsNullOrWhiteSpace(textBox4.Text); // ID textbox check
                    string query;

                    if (isUpdate)
                    {
                        query = @"
                UPDATE CBEntry
                SET 
                    InvoiceNo = @InvoiceNo,
                    TotalAmount = @TotalAmount,
                    Amount = @Amount,
                    Balance = @Balance,
                    TransactionDate = GETDATE(),
                    TransactionType = @TransactionType,
                    BnchID = @BnchID,
                    LocID = @LocID
                WHERE ID = @ID;";
                    }
                    else
                    {
                        query = @"
                INSERT INTO CBEntry 
                (InvoiceNo, TotalAmount, Amount, Balance, TransactionDate, TransactionType, BnchID, LocID) 
                VALUES 
                (@InvoiceNo, @TotalAmount, @Amount, @Balance, GETDATE(), @TransactionType, @BnchID, @LocID);";
                    }

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        int transactionType = radioButton1.Checked ? 1 : radioButton2.Checked ? 2 : 0;

                        if (isUpdate)
                        {
                            cmd.Parameters.AddWithValue("@ID", Convert.ToInt32(textBox4.Text)); // Use ID for update
                        }

                        cmd.Parameters.AddWithValue("@InvoiceNo", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TotalAmount", textBox1.Text);
                        cmd.Parameters.AddWithValue("@Amount", textBox2.Text);
                        cmd.Parameters.AddWithValue("@Balance", textBox3.Text);
                        cmd.Parameters.AddWithValue("@TransactionType", transactionType);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        await cmd.ExecuteNonQueryAsync();
                    }

                    MessageBox.Show(isUpdate ? "Updated Successfully" : "Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ResetFormFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        //private async void InsertData()
        //{
        //    try
        //    {
        //        if (!ValidateFields()) return;
        //        using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
        //        {
        //            await conn.OpenAsync();

        //            string insertCBEntryQuery = @"
        //INSERT INTO CBEntry (InvoiceNo, TotalAmount, Amount, Balance, TransactionDate, TransactionType, BnchID, LocID) 
        //VALUES (@InvoiceNo, @TotalAmount, @Amount, @Balance, GETDATE(), @TransactionType, @BnchID, @LocID);

        //SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertCBEntryQuery, conn))
        //            {
        //                int transactionType = radioButton1.Checked ? 1 : radioButton2.Checked ? 2 : 0;

        //                cmd.Parameters.AddWithValue("@InvoiceNo", comboBox1.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@TotalAmount", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@Amount", textBox2.Text);
        //                cmd.Parameters.AddWithValue("@Balance", textBox3.Text);
        //                cmd.Parameters.AddWithValue("@TransactionType", transactionType);
        //                cmd.Parameters.AddWithValue("@BnchID", 1);
        //                cmd.Parameters.AddWithValue("@LocID", 1);

        //                object newIDObj = await cmd.ExecuteScalarAsync();
        //            }

        //            MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            ResetFormFields();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }

        //}

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; 
            }
        }

        #endregion

        #region validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox2, "Amount is required" },
    };

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Invoice Number is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }

            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                MessageBox.Show("Please select a Transaction Type (Cash or Bank)", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                radioButton1.Focus(); // or radioButton2.Focus()
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Amount is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox2.Focus();
                return false;
            }

            return true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }


        #endregion


    }
}
