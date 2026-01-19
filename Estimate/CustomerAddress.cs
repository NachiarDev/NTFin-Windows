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
    public partial class CustomerAddress : Form
    {
        public CustomerAddress()
        {
            InitializeComponent();
        }
        private void CustomerAddress_Load(object sender, EventArgs e)
        {
            LoadCountries();
            LoadAddressType();
            textBox2.Visible = false;
        }
        #region ComboBox
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
        private void LoadAddressType()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, AddressType FROM Tbl_AddressType WHERE Active = 1";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox3.DataSource = dt;
                        comboBox3.DisplayMember = "AddressType";  // Show country name
                        comboBox3.ValueMember = "ID";            // Store country ID
                        comboBox3.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Address Type: " + ex.Message);
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

                    // If textBox2 has ID value => Update
                    if (!string.IsNullOrWhiteSpace(textBox2.Text))
                    {
                        int addressId = Convert.ToInt32(textBox2.Text);

                        string updateQuery = @"
UPDATE Customer_Address 
SET 
    Address = @Address,
    AddressTypeID = @AddressTypeID,
    ZipCode = @ZipCode,
    CountryID = @CountryID,
    StateID = @StateID,
    ModifiedBy = @ModifiedBy,
    ModifiedDate = GETDATE()
WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", addressId);
                            cmd.Parameters.AddWithValue("@Address", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@AddressTypeID", comboBox3.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@ZipCode", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@ModifiedBy", 1);

                            await cmd.ExecuteNonQueryAsync();
                            MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else // Insert
                    {
                        string insertQuery = @"
INSERT INTO Customer_Address (CustomerID, Address, AddressTypeID, ZipCode, CountryID, StateID, BnchID, LocID, CreatedBy, CreatedDate) 
VALUES (@CustomerID, @Address, @AddressTypeID, @ZipCode, @CountryID, @StateID, @BnchID, @LocID, @CreatedBy, GETDATE());

SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CustomerID", 1); // Update if you get actual customer ID dynamically
                            cmd.Parameters.AddWithValue("@Address", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@AddressTypeID", comboBox3.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@ZipCode", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);

                            object newID = await cmd.ExecuteScalarAsync();
                            MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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




        //            string insertOrganisationQuery = @"
        //    INSERT INTO Customer_Address (CustomerID, Address, AddressTypeID,ZipCode,CountryID, StateID, BnchID, LocID,CreatedBy,CreatedDate) 
        //    VALUES              (@CustomerID, @Address,@AddressTypeID,@ZipCode,@CountryID, @StateID,@BnchID,@LocID,@CreatedBy,GETDATE());

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertOrganisationQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@ZipCode", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@Address", richTextBox1.Text);
        //                cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@AddressTypeID", comboBox3.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@CustomerID", 1);
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

        #region validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<Control, string>
    {
                {richTextBox1,"CustomerAdddress is Important "},
                { textBox1, "Zip Code is required" }
                

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
            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Address Type is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox3.Focus();
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
            
           

            return true;
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
            comboBox3.SelectedIndex = -1;
            richTextBox1.Text = "";
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }


        #endregion

        
    }
}
