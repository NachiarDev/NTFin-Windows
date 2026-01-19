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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Estimate
{
    public partial class CustomerBanckAccount : Form
    {
        public CustomerBanckAccount()
        {
            InitializeComponent();
        }
        private void CustomerBanckAccount_Load(object sender, EventArgs e)
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

                    int.TryParse(textBox7.Text, out int id); // this is your primary key (e.g., CustomerBankAccountID)

                    if (id > 0)
                    {
                        // Perform Update
                        string updateQuery = @"
                UPDATE Customer_BankAccount 
                SET 
                    CustBankAccNo = @CusBankAccNo,
                    CustBankName = @CusBankName,
                    CustBankBranch = @CusBankBranch,
                    CustBankIFSC = @CusBankIFSC,
                    SwiftCode = @SwiftCode,
                    SwiftBranch = @SwiftBranch,
                    SwiftDetails = @SwiftDetails,
                    BnchID = @BnchID,
                    LocID = @LocID
                WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", id);
                            cmd.Parameters.AddWithValue("@CusBankAccNo", textBox3.Text);
                            cmd.Parameters.AddWithValue("@CusBankName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CusBankBranch", textBox2.Text);
                            cmd.Parameters.AddWithValue("@CusBankIFSC", textBox4.Text);
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
                        // Check duplicate before insert
                        string checkQuery = "SELECT COUNT(*) FROM Customer_BankAccount WHERE CustomerID = @CustomerID AND CustBankAccNo = @CustBankAccNo";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@CustomerID", 1);
                            checkCmd.Parameters.AddWithValue("@CustBankAccNo", textBox3.Text);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("A customer with this account number already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        // Insert new record
                        string insertQuery = @"
                INSERT INTO Customer_BankAccount (CustomerID, CustBankAccNo, CustBankName, CustBankBranch, CustBankIFSC, SwiftCode, SwiftBranch, SwiftDetails, BnchID, LocID, CreatedBy, CreatedDate)
                VALUES (@CustomerID, @CusBankAccNo, @CusBankName, @CusBankBranch, @CusBankIFSC, @SwiftCode, @SwiftBranch, @SwiftDetails, @BnchID, @LocID, @CreatedBy, GETDATE());
                SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CustomerID", 1);
                            cmd.Parameters.AddWithValue("@CusBankAccNo", textBox3.Text);
                            cmd.Parameters.AddWithValue("@CusBankName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CusBankBranch", textBox2.Text);
                            cmd.Parameters.AddWithValue("@CusBankIFSC", textBox4.Text);
                            cmd.Parameters.AddWithValue("@SwiftCode", textBox5.Text);
                            cmd.Parameters.AddWithValue("@SwiftBranch", textBox6.Text);
                            cmd.Parameters.AddWithValue("@SwiftDetails", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);

                            object newID = await cmd.ExecuteScalarAsync();
                            MessageBox.Show("Inserted Successfully. New ID: " + newID.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        //            string checkQuery = "SELECT COUNT(*) FROM Customer_BankAccount WHERE CustomerID = @CustomerID AND CustBankAccNo=@CustBankAccNo";

        //            using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
        //            {
        //                checkCmd.Parameters.AddWithValue("@CustomerID", 1);
        //                checkCmd.Parameters.AddWithValue("@CustBankAccNo", textBox3.Text);
        //                int count = (int)await checkCmd.ExecuteScalarAsync();

        //                if (count > 0)
        //                {
        //                    MessageBox.Show("An Customer with this Account Number already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                    return;
        //                }
        //            }




        //            string insertCustomerBankAccQuery = @"
        //    INSERT INTO Customer_BankAccount (CustomerID, CustBankAccNo, CustBankName, CustBankBranch, CustBankIFSC,SwiftCode,SwiftBranch,SwiftDetails,BnchID,LocID,CreatedBy,CreatedDate) 
        //    VALUES              (@CustomerID, @CusBankAccNo, @CusBankName, @CusBankBranch, @CusBankIFSC,@SwiftCode,@SwiftBranch,@SwiftDetails,@BnchID,@LocID,@CreatedBy,GETDATE());

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertCustomerBankAccQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@CustomerID", 1);
        //                cmd.Parameters.AddWithValue("@CusBankAccNo", textBox3.Text);
        //                cmd.Parameters.AddWithValue("@CusBankName", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@CusBankBranch", textBox2.Text);
        //                cmd.Parameters.AddWithValue("@CusBankIFSC", textBox4.Text);
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

            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textBox)
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
