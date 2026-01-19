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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Estimate
{
    public partial class Branch : Form
    {
        public Branch()
        {
            InitializeComponent();
        }
        private void Branch_Load(object sender, EventArgs e)
        {   
            textBox4.Visible= false;
            LoadOrganisation();
            LoadCountries();
        }
        #region LoadOrganisation
        private void LoadOrganisation()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, Name FROM Organisation WHERE Active = 1";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "Name";  // Show Organisation Name
                        comboBox1.ValueMember = "ID";      // Store Organisation ID
                        comboBox1.SelectedIndex = -1;      // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading organisations: " + ex.Message);
            }
        }

        private void LoadLocation(int orgID)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, LocationName FROM Tbl_Location WHERE OrgID = @OrgID AND Active = 1"; // Fixed SQL syntax

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@OrgID", orgID);

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                comboBox2.DataSource = dt;
                                comboBox2.DisplayMember = "LocationName";
                                comboBox2.ValueMember = "ID";
                                comboBox2.SelectedIndex = -1;
                            }
                            else
                            {
                                comboBox2.DataSource = null; // No locations, clear comboBox2
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading locations: " + ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue != null && comboBox1.SelectedIndex != -1)
            {
                int selectedOrgID;
                if (int.TryParse(comboBox1.SelectedValue.ToString(), out selectedOrgID))
                {
                    LoadLocation(selectedOrgID);
                }
            }
            else
            {
                comboBox2.DataSource = null;  // Clear locations if no organisation is selected
            }
        }

        #endregion
       
        
        #region LoadCountries 
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

                        comboBox3.DataSource = dt;
                        comboBox3.DisplayMember = "CountryName";  // Show country name
                        comboBox3.ValueMember = "ID";            // Store country ID
                        comboBox3.SelectedIndex = -1;           // No pre-selection
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

                            comboBox4.DataSource = dt;
                            comboBox4.DisplayMember = "StateName";  // Show state name
                            comboBox4.ValueMember = "ID";           // Store state ID
                            comboBox4.SelectedIndex = -1;          // No pre-selection
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading states: " + ex.Message);
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedValue != null && comboBox1.SelectedIndex != -1)
            {
                int selectedCountryID;
                if (int.TryParse(comboBox3.SelectedValue.ToString(), out selectedCountryID))
                {
                    LoadStates(selectedCountryID);
                }
            }
            else
            {
                comboBox4.DataSource = null;  // Clear states if no country is selected
            }
        }

        #endregion
         
        #region Insert
        private async void button1_Click(object sender, EventArgs e)
        {
            await InsertData();
        }

        private async Task InsertData()
        {
            try
            {
               

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    // Determine if this is an update or an insert based on Branch ID in textBox4.
                    // If textBox4 contains a value (e.g. "5"), then update the branch with ID 5.
                    int.TryParse(textBox4.Text, out int branchID);

                    // For Branch, OrgID usually comes from comboBox1.
                    int orgID = Convert.ToInt32(comboBox1.SelectedValue ?? 0);

                    int activeStatus = radioButton1.Checked ? 1 : 0;

                    // Format the phone numbers (using inline substring formatting as in your Organisation code).
                    string formattedPrimaryContact = $"({maskedTextBox1.Text.Substring(0, 3)}) {maskedTextBox1.Text.Substring(3, 3)}-{maskedTextBox1.Text.Substring(6, 4)}";
                    string formattedAlternateContact = $"({maskedTextBox2.Text.Substring(0, 3)}) {maskedTextBox2.Text.Substring(3, 3)}-{maskedTextBox2.Text.Substring(6, 4)}";

                    if (branchID > 0)
                    {
                        // Perform UPDATE for an existing Branch.
                        string updateQuery = @"
                    UPDATE Branch 
                    SET 
                        OrgID = @OrgID,
                        Code = @Code,
                        Name = @Name,
                        Location = @Location,
                        GSTNumber = @GSTNumber,
                        ContactNo = @ContactNo,
                        Mobile = @Mobile,
                        Email = @Email,
                        BranchAddress = @BranchAddress,
                        CountryID = @CountryID,
                        StateID = @StateID,
                        City = @City,
                        Zip = @Zip,
                        Active = @Active
                    WHERE ID = @ID"; // Assuming your primary key column is 'ID'

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", branchID);
                            cmd.Parameters.AddWithValue("@OrgID", orgID);
                            cmd.Parameters.AddWithValue("@Code", textBox1.Text);
                            cmd.Parameters.AddWithValue("@Name", textBox2.Text);
                            cmd.Parameters.AddWithValue("@Location", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTNumber", textBox3.Text);
                            cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@Mobile", formattedAlternateContact);
                            cmd.Parameters.AddWithValue("@Email", textBox5.Text);
                            cmd.Parameters.AddWithValue("@BranchAddress", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox3.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox4.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@City", textBox6.Text);
                            cmd.Parameters.AddWithValue("@Zip", textBox7.Text);
                            cmd.Parameters.AddWithValue("@Active", activeStatus);

                            await cmd.ExecuteNonQueryAsync();
                        }

                        MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Before inserting, check for a duplicate branch with the same OrgID, Location, and Name
                        string checkQuery = "SELECT COUNT(*) FROM Branch WHERE OrgID = @OrgID AND Location = @Location AND Name = @Name";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@OrgID", orgID);
                            checkCmd.Parameters.AddWithValue("@Location", comboBox2.SelectedValue ?? DBNull.Value);
                            checkCmd.Parameters.AddWithValue("@Name", textBox2.Text.Trim());

                            int count = (int)await checkCmd.ExecuteScalarAsync();
                            if (count > 0)
                            {
                                MessageBox.Show("A branch for this organisation with the same Name and Location already exists. Please enter different details.",
                                    "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        // Perform INSERT for a new Branch.
                        string insertQuery = @"
                    INSERT INTO Branch (OrgID, Code, Name, Location, GSTNumber, ContactNo, Mobile, Email, BranchAddress, CountryID, StateID, City, Zip, Active) 
                    VALUES (@OrgID, @Code, @Name, @Location, @GSTNumber, @ContactNo, @Mobile, @Email, @BranchAddress, @CountryID, @StateID, @City, @Zip, @Active);
                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@OrgID", orgID);
                            cmd.Parameters.AddWithValue("@Code", textBox1.Text);
                            cmd.Parameters.AddWithValue("@Name", textBox2.Text);
                            cmd.Parameters.AddWithValue("@Location", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTNumber", textBox3.Text);
                            cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@Mobile", formattedAlternateContact);
                            cmd.Parameters.AddWithValue("@Email", textBox5.Text);
                            cmd.Parameters.AddWithValue("@BranchAddress", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox3.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox4.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@City", textBox6.Text);
                            cmd.Parameters.AddWithValue("@Zip", textBox7.Text);
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
        private bool ValidateZip(string zip, int zipMaxLength, int zipFormat)
        {
            int maxLength = (zipMaxLength == 1) ? 6 : 5;
            
            if (zip.Length > maxLength)
            {
                MessageBox.Show($"Invalid Zip Code. Maximum length allowed is {maxLength}.",
                    "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (zipFormat == 1 && !zip.All(char.IsLetter))
            {
                MessageBox.Show("Invalid Zip Code. Only characters are allowed.",
                    "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (zipFormat == 2 && !zip.All(char.IsDigit))
            {
                MessageBox.Show("Invalid Zip Code. Only digits are allowed.",
                    "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
       
        
        
        
        
        
        
        
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        private void textBox5_Leave(object sender, EventArgs e)
        {
            string email = textBox5.Text.Trim();
            string emailPattern = @"^[a-zA-Z0-9._%+-]{3,}@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";

            if (!Regex.IsMatch(email, emailPattern))
            {
                MessageBox.Show("Invalid email format! Please enter a valid email.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox5.Focus();
            }
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


        private string FormatPhoneNumber(string number)
        {
            return number.Length >= 10
                ? $"({number.Substring(0, 3)}) {number.Substring(3, 3)}-{number.Substring(6, 4)}"
                : number;
        }

        #endregion

        #region Reset
        private void ResetFormFields()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            maskedTextBox2.Text = "";
            maskedTextBox1.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";

            radioButton1.Checked = false;
            radioButton2.Checked= false;
            richTextBox1.Clear();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        #endregion
    }
}
