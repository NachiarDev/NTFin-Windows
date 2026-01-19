using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Estimate
{
    public partial class Location : Form
    {
        public Location()
        {
            InitializeComponent();
        }
       
        private void Location_Load(object sender, EventArgs e)
        {   
            textBox2.Visible=false;
            LoadOrganisations();
        }

        #region Load ComboBox
        private void LoadOrganisations()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                   
                    string query = "SELECT ID, Name FROM Organisation WHERE Active = 1 "; // Only active organisations
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        // Bind data to ComboBox
                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "Name";  // What the user sees
                        comboBox1.ValueMember = "ID";      // The actual value
                        comboBox1.SelectedIndex = -1;      // No default selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading organisations: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    int activeStatus = radioButton1.Checked ? 1 : 0;
                    string locationName = textBox1.Text;
                    object orgId = comboBox1.SelectedValue ?? DBNull.Value;

                    if (!string.IsNullOrWhiteSpace(textBox2.Text)) // UPDATE logic
                    {
                        int locationId = int.Parse(textBox2.Text);

                        string updateQuery = @"
                    UPDATE Tbl_Location 
                    SET OrgID = @OrgID, LocationName = @LocationName, Active = @Active, ModifiedBy = @ModifiedBy, ModifiedDate = GETDATE()
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@OrgID", orgId);
                            cmd.Parameters.AddWithValue("@LocationName", locationName);
                            cmd.Parameters.AddWithValue("@Active", activeStatus);
                            cmd.Parameters.AddWithValue("@ModifiedBy", 1);
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
                        string checkQuery = "SELECT COUNT(*) FROM Tbl_Location WHERE OrgID = @OrgID AND LocationName = @LocationName";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@OrgID", orgId);
                            checkCmd.Parameters.AddWithValue("@LocationName", locationName);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This Organisation already located. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                    INSERT INTO Tbl_Location (OrgID, LocationName, Active, CreatedBy, CreateDate) 
                    VALUES (@OrgID, @LocationName, @Active, @CreatedBy, GETDATE());

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@OrgID", orgId);
                            cmd.Parameters.AddWithValue("@LocationName", locationName);
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);
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

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
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
            comboBox1.SelectedIndex = -1;
            textBox1.Text = "";
            radioButton1.Checked = false;
            radioButton2.Checked = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        #endregion

    }
}
