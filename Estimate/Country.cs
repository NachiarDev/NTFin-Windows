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
using System.Windows.Controls;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Estimate
{
    public partial class Country : Form
    {
        public Country()
        {
            InitializeComponent();
        }

        private void Country_Load(object sender, EventArgs e)
        {
            textBox4.Visible = false;
            
        }

        #region Insert

        private async void button1_Click(object sender, EventArgs e)
        {
          await  InsertData();
        }
        private async Task InsertData()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Country Code is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox1.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBox3.Text))
                {
                    MessageBox.Show("Country Name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox3.Focus();
                    return;
                }

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    // Determine if this is an update or an insert based on Branch ID in textBox4.
                    // If textBox4 contains a value (e.g. "5"), then update the branch with ID 5.
                    int.TryParse(textBox4.Text, out int countryID);
                    int activeStatus = radioButton1.Checked ? 1 : 0;

                    if (countryID > 0)
                    {
                        // Perform UPDATE for an existing Branch.
                        string updateQuery = @"
                    UPDATE tbl_Country 
                    SET 
                        CountryCode = @CountryCode,
                        CountryNumber = @CountryNumber,
                        CountryName = @CountryName,
                        Active = @Active
                       
                    WHERE ID = @ID"; // Assuming your primary key column is 'ID'

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", countryID);
                            cmd.Parameters.AddWithValue("@CountryCode", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CountryNumber", textBox2.Text);
                            cmd.Parameters.AddWithValue("@CountryName", textBox3.Text);

                            cmd.Parameters.AddWithValue("@Active", activeStatus);

                            await cmd.ExecuteNonQueryAsync();
                        }

                        MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Before inserting, check for a duplicate branch with the same OrgID, Location, and Name
                        string checkQuery = "SELECT COUNT(*) FROM Tbl_Country WHERE CountryCode = @CountryCode AND CountryNumber = @CountryNumber AND CountryName = @CountryName";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@CountryCode", textBox1.Text);
                            checkCmd.Parameters.AddWithValue("@CountryNumber", textBox2.Text);
                            checkCmd.Parameters.AddWithValue("@CountryName", textBox3.Text.Trim());

                            int count = (int)await checkCmd.ExecuteScalarAsync();
                            if (count > 0)
                            {
                                MessageBox.Show("A Country Nmae for this Country Code with the same Country Number already exists. Please enter different details.",
                                    "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        // Perform INSERT for a new Branch.
                        string insertQuery = @"
                    INSERT INTO tbl_Country (CountryCode, CountryNumber, CountryName, Active) 
                    VALUES (@CountryCode, @CountryNumber, @CountryName, @Active);
                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CountryCode", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CountryNumber", textBox2.Text);
                            cmd.Parameters.AddWithValue("@CountryName", textBox3.Text);

                            cmd.Parameters.AddWithValue("@Active", activeStatus);

                            object result = await cmd.ExecuteScalarAsync();
                            // Save the new Branch ID into textBox4
                            textBox4.Text = result.ToString();
                        }

                        MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    ResetFormFields(); // Clear or reset the form fields as needed
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
            radioButton1.Checked = false;
            radioButton2.Checked = false;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }



        #endregion

        #region Validation
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar))
            {
                // Suppress the key press
                e.Handled = true;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                // Convert lowercase to uppercase
                e.KeyChar = char.ToUpper(e.KeyChar);
            }
            else if (!char.IsControl(e.KeyChar))
            {
                // Block non-letter and non-control keys (like numbers, symbols)
                e.Handled = true;
            }
        }



        #endregion

    }
}

