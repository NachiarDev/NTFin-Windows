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
using System.Text.RegularExpressions;

namespace Estimate
{
    public partial class Organisation : Form
    {
        public Organisation()
        {
            InitializeComponent();
        }

        private void Organisation_Load(object sender, EventArgs e)
        {
            LoadCountries();
            textBox6.Visible = false;
        }

        #region Combobox load
        private void LoadCountries()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, CountryName FROM tbl_Country WHERE Active = 1 ORDER BY CountryName";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "CountryName";  // Show country name
                        comboBox1.ValueMember = "ID";            // Store country ID
                        comboBox1.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading countries: " + ex.Message);
            }
        }
        private void LoadStates(int countryID)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, StateName FROM Tbl_State WHERE CountryID = @CountryID AND Active = 1 ORDER BY StateName";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@CountryID", countryID);

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            comboBox2.DataSource = dt;
                            comboBox2.DisplayMember = "StateName";  // Show state name
                            comboBox2.ValueMember = "ID";           // Store state ID
                            comboBox2.SelectedIndex = -1;          // No pre-selection
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading states: " + ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue != null && comboBox1.SelectedIndex != -1)
            {
                int selectedCountryID;
                if (int.TryParse(comboBox1.SelectedValue.ToString(), out selectedCountryID))
                {
                    LoadStates(selectedCountryID);
                }
            }
            else
            {
                comboBox2.DataSource = null;  // Clear states if no country is selected
            }
        }

        #endregion


        #region Insert 
        private async void InsertData()
        {
            try
            {
                if (!ValidateFields()) return;

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();
                   
                    int.TryParse(textBox6.Text, out int organisationId); // Use as OrganisationID (0 or empty = insert)

                    int activeStatus = radioButton1.Checked ? 1 : 0;
                    string formattedPrimaryContact = $"({maskedTextBox1.Text.Substring(0, 3)}) {maskedTextBox1.Text.Substring(3, 3)}-{maskedTextBox1.Text.Substring(6, 4)}";
                    string formattedAlternateContact = $"({maskedTextBox2.Text.Substring(0, 3)}) {maskedTextBox2.Text.Substring(3, 3)}-{maskedTextBox2.Text.Substring(6, 4)}";

                    if (organisationId > 0)
                    {
                        // Perform UPDATE
                        string updateQuery = @"
                    UPDATE Organisation 
                    SET Name = @Name, ContactNo = @ContactNo, Mobile = @Mobile, Email = @Email,
                        Address1 = @Address1, Address2 = @Address2, City = @City, Zip = @Zip,
                        CountryID = @CountryID, StateID = @StateID, Active = @Active
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", organisationId);
                            cmd.Parameters.AddWithValue("@Name", textBox2.Text);
                            cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@Mobile", formattedAlternateContact);
                            cmd.Parameters.AddWithValue("@Email", textBox3.Text);
                            cmd.Parameters.AddWithValue("@Address1", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@Address2", richTextBox2.Text);
                            cmd.Parameters.AddWithValue("@City", textBox4.Text);
                            cmd.Parameters.AddWithValue("@Zip", textBox5.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@Active", activeStatus);

                            await cmd.ExecuteNonQueryAsync();
                        }

                        MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string checkQuery = "SELECT COUNT(*) FROM Organisation WHERE Name = @Name";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@Name", textBox2.Text);
                            checkCmd.Parameters.AddWithValue("@City", textBox4.Text);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("An organisation with this Name already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        // Perform INSERT
                        string insertQuery = @"
                    INSERT INTO Organisation (Code, Name, ContactNo, Mobile, Email, Address1, Address2, City, Zip, CountryID, StateID, Active) 
                    VALUES (@Code, @Name, @ContactNo, @Mobile, @Email, @Address1, @Address2, @City, @Zip, @CountryID, @StateID, @Active);

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@Code", textBox1.Text);
                            cmd.Parameters.AddWithValue("@Name", textBox2.Text);
                            cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@Mobile", formattedAlternateContact);
                            cmd.Parameters.AddWithValue("@Email", textBox3.Text);
                            cmd.Parameters.AddWithValue("@Address1", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@Address2", richTextBox2.Text);
                            cmd.Parameters.AddWithValue("@City", textBox4.Text);
                            cmd.Parameters.AddWithValue("@Zip", textBox5.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@Active", activeStatus);

                            object result = await cmd.ExecuteScalarAsync();
                            textBox6.Text = result.ToString(); // Save the new ID
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
        //            string checkQuery = "SELECT COUNT(*) FROM Organisation WHERE Name = @Name";

        //            using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
        //            {
        //                checkCmd.Parameters.AddWithValue("@Name", textBox2.Text);
        //                checkCmd.Parameters.AddWithValue("@City", textBox4.Text);
        //                int count = (int)await checkCmd.ExecuteScalarAsync();

        //                if (count > 0)
        //                {
        //                    MessageBox.Show("An organisation with this Name already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                    return;
        //                }
        //            }


        //            int activeStatus = radioButton1.Checked ? 1 : 0;
        //            string formattedPrimaryContact = $"({maskedTextBox1.Text.Substring(0, 3)}) {maskedTextBox1.Text.Substring(3, 3)}-{maskedTextBox1.Text.Substring(6, 4)}";
        //            string formattedAlternateContact = $"({maskedTextBox2.Text.Substring(0, 3)}) {maskedTextBox2.Text.Substring(3, 3)}-{maskedTextBox2.Text.Substring(6, 4)}";

        //            string insertOrganisationQuery = @"
        //    INSERT INTO Organisation (Code, Name, ContactNo, Mobile, Email, Address1, Address2, City, Zip, CountryID, StateID, Active) 
        //    VALUES (@Code, @Name, @ContactNo, @Mobile, @Email, @Address1, @Address2, @City, @Zip, @CountryID, @StateID, @Active);

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertOrganisationQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@Code", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@Name", textBox2.Text);
        //                cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
        //                cmd.Parameters.AddWithValue("@Mobile", formattedAlternateContact);
        //                cmd.Parameters.AddWithValue("@Email", textBox3.Text);
        //                cmd.Parameters.AddWithValue("@Address1", richTextBox1.Text);
        //                cmd.Parameters.AddWithValue("@Address2", richTextBox2.Text);
        //                cmd.Parameters.AddWithValue("@City", textBox4.Text);
        //                cmd.Parameters.AddWithValue("@Zip", textBox5.Text);
        //                cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@Active", activeStatus);

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

        private void button1_Click(object sender, EventArgs e)
        {
            InsertData();
        }
        #endregion

        #region validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<Control, string>
    {
        { textBox1, "Organisation Code is required" },
        { textBox2, "Organisation Name is required" },
        { maskedTextBox1, "primary contact Number is required" },
        { maskedTextBox2, "Alternate contact Number is required" },
        { textBox3, "Email is required" },
        { richTextBox1, "Primary Address is required" },
       
        { textBox4, "City is required" },
        { textBox5, "ZIP Code is required" }
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
            string primaryContact = maskedTextBox1.Text.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "");
            string alternateContact = maskedTextBox2.Text.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "");

            // Validate phone number length (10 digits)
            if (primaryContact.Length != 10)
            {
                MessageBox.Show("Primary Contact Number must be in the format (000) 000-0000", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                maskedTextBox1.Focus();
                return false;
            }

            if (alternateContact.Length != 10)
            {
                MessageBox.Show("Alternate Contact Number must be in the format (000) 000-0000", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                maskedTextBox2.Focus();
                return false;
            }
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Country is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }

            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("State is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Focus();
                return false;
            }

            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                MessageBox.Show("Active status must be selected", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            string email = textBox3.Text.Trim();
            string emailPattern = @"^[a-zA-Z0-9._%+-]{3,}@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";

            if (!Regex.IsMatch(email, emailPattern))
            {
                MessageBox.Show("Invalid email format! Please enter a valid email.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox3.Focus();  // Set focus back to the textbox
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        #endregion

        #region Reset
        private void ResetFormFields()
        {

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textBox)
                {
                    textBox.Text = "";
                }
            }


            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is RadioButton radioButton)
                {
                    radioButton.Checked = false;
                }
            }


            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            maskedTextBox2.Text = "";
            maskedTextBox1.Text = "";

            richTextBox1.Clear();
            richTextBox2.Clear();


           


            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }

        #endregion
    }
}
