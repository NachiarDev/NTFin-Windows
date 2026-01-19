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
    public partial class SubCategory : Form
    {
        public SubCategory()
        {
            InitializeComponent();
        }

        private void SubCategory_Load(object sender, EventArgs e)
        {
            textBox2.Visible = false;
            LoadCategories();
        }
        #region Load ComboBox
        private void LoadCategories()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    string query = "SELECT ID, Name FROM Category"; // Only active organisations
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
                MessageBox.Show("Error loading Categories: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    
                    string SubCategoryName = textBox1.Text;
                    object CategoryID = comboBox1.SelectedValue ?? DBNull.Value;

                    if (!string.IsNullOrWhiteSpace(textBox2.Text)) // UPDATE logic
                    {
                        int subcategoryId = int.Parse(textBox2.Text);

                        string updateQuery = @"
                    UPDATE SubCategory 
                    SET CategoryID = @CategoryID, SubCategoryName = @SubCategoryName
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CategoryID", CategoryID);
                            cmd.Parameters.AddWithValue("@SubCategoryName", SubCategoryName);
                            cmd.Parameters.AddWithValue("@ID", subcategoryId);

                            int rowsAffected = await cmd.ExecuteNonQueryAsync();

                            if (rowsAffected > 0)
                                MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("No matching record found to update.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else // INSERT logic
                    {
                        string checkQuery = "SELECT COUNT(*) FROM SubCategory WHERE CategoryID = @CategoryID AND SubCategoryName = @SubCategoryName";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@CategoryID", CategoryID);
                            checkCmd.Parameters.AddWithValue("@SubCategoryName", SubCategoryName);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This Sub category already Existing. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                    INSERT INTO SubCategory (CategoryID, SubCategoryName) 
                    VALUES (@CategoryID, @SubCategoryName);

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CategoryID", CategoryID);
                            cmd.Parameters.AddWithValue("@SubCategoryName", SubCategoryName);
                           

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
            comboBox1.SelectedIndex = -1;
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        #endregion


    }
}
