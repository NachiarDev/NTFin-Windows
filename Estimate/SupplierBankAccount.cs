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
    public partial class SupplierBankAccount : Form
    {
        public SupplierBankAccount()
        {
            InitializeComponent();
        }
        private void SupplierBankAccount_Load(object sender, EventArgs e)
        {
            textBox7.Visible = false;
        }
        #region validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Bank Name is required" },
        { textBox2, "Bank Branch Name is required" },
        { textBox3, "Account Number is required" },
        { textBox4, "IFSC Code is required" },

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


            return true;
        }
        private void textBox4_Leave(object sender, EventArgs e)
        {
            string email = textBox4.Text.Trim();
            string emailPattern = @"^^[A-Z]{4}0[A-Z0-9a-z]{6}$";

            if (!Regex.IsMatch(email, emailPattern))
            {
                MessageBox.Show("Invalid IFSC Code! Please enter a valid IFSC.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox4.Focus();
            }
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }
        private void textBox3_Leave(object sender, EventArgs e)
        {
            string input = textBox3.Text.Trim();
            string digitPattern = @"^\d{9,18}$"; // Only digits, length 9 to 18

            if (!Regex.IsMatch(input, digitPattern))
            {
                MessageBox.Show("Invalid input! Please enter only digits, between 9 to 18 characters.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox3.Focus();
            }
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
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
                if (!ValidateFields()) return;

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    int.TryParse(textBox7.Text, out int id); // ID for update

                    if (id > 0)
                    {
                        // UPDATE MODE
                        string updateQuery = @"
                UPDATE Supplier_BankAccount
                SET 
                    SupliBankAccNo = @SupliBankAccNo,
                    SupliBankName = @SupliBankName,
                    SupliBankBranch = @SupliBankBranch,
                    SupliBankIFSC = @SupliBankIFSC,
                    SwiftCode = @SwiftCode,
                    SwiftBranch = @SwiftBranch,
                    SwiftDetails = @SwiftDetails,
                    BnchID = @BnchID,
                    LocID = @LocID
                WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", id);
                            cmd.Parameters.AddWithValue("@SupliBankAccNo", textBox3.Text);
                            cmd.Parameters.AddWithValue("@SupliBankName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@SupliBankBranch", textBox2.Text);
                            cmd.Parameters.AddWithValue("@SupliBankIFSC", textBox4.Text);
                            cmd.Parameters.AddWithValue("@SwiftCode", textBox5.Text);
                            cmd.Parameters.AddWithValue("@SwiftBranch", textBox6.Text);
                            cmd.Parameters.AddWithValue("@SwiftDetails", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            int rowsAffected = await cmd.ExecuteNonQueryAsync();
                            if (rowsAffected > 0)
                                MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("Update failed. Record not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        // INSERT MODE: Check for duplicate
                        string checkQuery = "SELECT COUNT(*) FROM Supplier_BankAccount WHERE SupplierID = @SupplierID AND SupliBankAccNo = @SupliBankAccNo";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@SupplierID", 8);
                            checkCmd.Parameters.AddWithValue("@SupliBankAccNo", textBox3.Text);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("A Supplier with this Account Number already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                INSERT INTO Supplier_BankAccount (SupplierID, SupliBankAccNo, SupliBankName, SupliBankBranch, SupliBankIFSC, SwiftCode, SwiftBranch, SwiftDetails, BnchID, LocID, CreatedBy, CreatedDate) 
                VALUES (@SupplierID, @SupliBankAccNo, @SupliBankName, @SupliBankBranch, @SupliBankIFSC, @SwiftCode, @SwiftBranch, @SwiftDetails, @BnchID, @LocID, @CreatedBy, GETDATE());
                SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@SupplierID", 8);
                            cmd.Parameters.AddWithValue("@SupliBankAccNo", textBox3.Text);
                            cmd.Parameters.AddWithValue("@SupliBankName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@SupliBankBranch", textBox2.Text);
                            cmd.Parameters.AddWithValue("@SupliBankIFSC", textBox4.Text);
                            cmd.Parameters.AddWithValue("@SwiftCode", textBox5.Text);
                            cmd.Parameters.AddWithValue("@SwiftBranch", textBox6.Text);
                            cmd.Parameters.AddWithValue("@SwiftDetails", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);

                            object newIDObj = await cmd.ExecuteScalarAsync();
                            MessageBox.Show("Inserted Successfully. New ID: " + newIDObj, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

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

        //            // Check if Organisation Name and City already exist together
        //            //string checkQuery = "SELECT COUNT(*) FROM Organisation WHERE Name = @Name AND City = @City";
        //            string checkQuery = "SELECT COUNT(*) FROM Supplier_BankAccount WHERE SupplierID = @SupplierID AND SupliBankAccNo=@SupliBankAccNo";

        //            using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
        //            {
        //                checkCmd.Parameters.AddWithValue("@SupplierID", 8);
        //                checkCmd.Parameters.AddWithValue("@SupliBankAccNo", textBox3.Text);
        //                int count = (int)await checkCmd.ExecuteScalarAsync();

        //                if (count > 0)
        //                {
        //                    MessageBox.Show("An Supplier with this Account Number already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                    return;
        //                }
        //            }




        //            string insertSupplierBankAccQuery = @"
        //    INSERT INTO Supplier_BankAccount (SupplierID, SupliBankAccNo, SupliBankName, SupliBankBranch, SupliBankIFSC,SwiftCode,SwiftBranch,SwiftDetails,BnchID,LocID,CreatedBy,CreatedDate) 
        //    VALUES              (@SupplierID, @SupliBankAccNo, @SupliBankName, @SupliBankBranch, @SupliBankIFSC,@SwiftCode,@SwiftBranch,@SwiftDetails,@BnchID,@LocID,@CreatedBy,GETDATE());

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertSupplierBankAccQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@SupplierID", 8);
        //                cmd.Parameters.AddWithValue("@SupliBankAccNo", textBox3.Text);
        //                cmd.Parameters.AddWithValue("@SupliBankName", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@SupliBankBranch", textBox2.Text);
        //                cmd.Parameters.AddWithValue("@SupliBankIFSC", textBox4.Text);
        //                cmd.Parameters.AddWithValue("@SwiftCode", textBox5.Text);
        //                cmd.Parameters.AddWithValue("@SwiftBranch", textBox6.Text);
        //                cmd.Parameters.AddWithValue("@SwiftDetails", richTextBox1.Text); // ✅ FIXED HERE
        //                cmd.Parameters.AddWithValue("@BnchID", 1);
        //                cmd.Parameters.AddWithValue("@LocID", 1);
        //                cmd.Parameters.AddWithValue("@CreatedBy", 1);
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
        #endregion



        #region Reset
        private void ResetFormFields()
        {

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is TextBox textBox)
                {
                    textBox.Text = "";
                }
            }

            richTextBox1.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }






        #endregion

       
    }
}
