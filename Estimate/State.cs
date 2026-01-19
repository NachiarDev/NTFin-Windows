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
    public partial class State : Form
    {
        public State()
        {
            InitializeComponent();
        }

        private void State_Load(object sender, EventArgs e)
        {
            textBox4.Visible = false;
            LoadCountries();
        }
        #region ComboBox
        private void LoadCountries()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    string query = "SELECT ID, CountryName FROM Tbl_Country WHERE Active = 1 "; // Only active organisations
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        // Bind data to ComboBox
                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "CountryName";  // What the user sees
                        comboBox1.ValueMember = "ID";      // The actual value
                        comboBox1.SelectedIndex = -1;      // No default selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Countries: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                if (string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("State Code is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox1.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBox2.Text))
                {
                    MessageBox.Show("State Name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox3.Focus();
                    return;
                }
                if (string.IsNullOrWhiteSpace(textBox3.Text))
                {
                    MessageBox.Show("State GSTTIN is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox3.Focus();
                    return;
                }
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    int activeStatus = radioButton1.Checked ? 1 : 0;
                    string stateName = textBox2.Text;
                    object CountryId = comboBox1.SelectedValue ?? DBNull.Value;

                    if (!string.IsNullOrWhiteSpace(textBox4.Text)) // UPDATE logic
                    {
                        int stateId = int.Parse(textBox4.Text);

                        string updateQuery = @"
                    UPDATE Tbl_State 
                    SET CountryID = @CountryID, StateName = @StateName,StateCode = @StateCode,GSTTIN = @GSTTIN, Active = @Active
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CountryID", CountryId);
                            cmd.Parameters.AddWithValue("@StateName", stateName);
                            cmd.Parameters.AddWithValue("@StateCode", textBox1.Text);
                            cmd.Parameters.AddWithValue("@GSTTIN", textBox3.Text);
                            cmd.Parameters.AddWithValue("@Active", activeStatus);
                            
                            cmd.Parameters.AddWithValue("@ID", stateId);

                            int rowsAffected = await cmd.ExecuteNonQueryAsync();

                            if (rowsAffected > 0)
                                MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("No matching record found to update.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else // INSERT logic
                    {
                        string checkQuery = "SELECT COUNT(*) FROM Tbl_State WHERE CountryID = @CountryID AND StateName = @StateName";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@CountryID", CountryId);
                            checkCmd.Parameters.AddWithValue("@StateName", stateName);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This StateName already Existing for the Country. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                    INSERT INTO Tbl_State (CountryID, StateName,StateCode,Active,GSTTIN) 
                    VALUES (@CountryID, @StateName,@StateCode, @Active,@GSTTIN);

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CountryID", CountryId);
                            cmd.Parameters.AddWithValue("@StateName", stateName);
                            cmd.Parameters.AddWithValue("@StateCode", textBox1.Text);
                            cmd.Parameters.AddWithValue("@GSTTIN", textBox3.Text);
                            cmd.Parameters.AddWithValue("@Active", activeStatus);                            
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
            comboBox1.SelectedIndex = -1;

            radioButton1.Checked = false;
            radioButton2.Checked = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }


        #endregion

        #region Validation
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar))
            {
                // Suppress the key press
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        #endregion


    }
}
