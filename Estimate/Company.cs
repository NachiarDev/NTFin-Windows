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
    public partial class Company : Form
    {
        public Company()
        {
            InitializeComponent();
        }

        private void Company_Load(object sender, EventArgs e)
        {
            textBox4.Visible = false;
        }
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
                // Validation block: All fields must be filled except textBox4
                if (string.IsNullOrWhiteSpace(textBox1.Text) || // CompanyName
                    string.IsNullOrWhiteSpace(maskedTextBox1.Text.Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "")) || // Contact
                    string.IsNullOrWhiteSpace(richTextBox1.Text) || // Address
                    string.IsNullOrWhiteSpace(textBox3.Text)) // GSTNumber
                {
                    MessageBox.Show("Please fill in all required fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    string CompanyName = textBox1.Text;
                    string formattedPrimaryContact = $"({maskedTextBox1.Text.Substring(0, 3)}) {maskedTextBox1.Text.Substring(3, 3)}-{maskedTextBox1.Text.Substring(6, 4)}";

                    if (!string.IsNullOrWhiteSpace(textBox4.Text)) // UPDATE logic
                    {
                        int locationId = int.Parse(textBox4.Text);

                        string updateQuery = @"
                    UPDATE Company 
                    SET Name = @Name, Address = @Address, Contact = @Contact, GSTNumber = @GSTNumber
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@Name", CompanyName);
                            cmd.Parameters.AddWithValue("@Address", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@Contact", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@GSTNumber", textBox3.Text);
                            cmd.Parameters.AddWithValue("@ID", locationId);

                            int rowsAffected = await cmd.ExecuteNonQueryAsync();

                            if (rowsAffected > 0)
                                MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("No matching record found to update.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else // INSERT logic
                    {
                        string checkQuery = "SELECT COUNT(*) FROM Company WHERE Name = @Name AND Contact = @Contact";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@Name", CompanyName);
                            checkCmd.Parameters.AddWithValue("@Contact", formattedPrimaryContact);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This Company already Existing. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                    INSERT INTO Company (Name, Address, Contact, GSTNumber) 
                    VALUES (@Name, @Address, @Contact, @GSTNumber);

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@Name", CompanyName);
                            cmd.Parameters.AddWithValue("@Address", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@Contact", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@GSTNumber", textBox3.Text);

                            object newID = await cmd.ExecuteScalarAsync();
                        }

                        MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    ResetFormFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        #endregion


        #region Validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Company Name Code is required" },
      
        { textBox3, "GST Number is required" },
         { maskedTextBox1, "Mobile Number is required" },
         { richTextBox1, "Address is required" },

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

        private void textBox3_Leave(object sender, EventArgs e)
        {
            string gst = textBox3.Text.Trim();
            string gstpattern = @"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[0-9]{1}[A-Z]{1}[0-9]{1}$";
            if (!Regex.IsMatch(gst, gstpattern))
            {
                MessageBox.Show("Invaid GST Format!Please enter a valid GST Number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox3.Focus();
            }
        }
        #endregion
    }
}
