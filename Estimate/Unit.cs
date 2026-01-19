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

namespace Estimate
{
    public partial class Unit : Form
    {
        public Unit()
        {
            InitializeComponent();
        }

        private void Unit_Load(object sender, EventArgs e)
        {
            textBox2.Visible= false;
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

                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();



                    if (!string.IsNullOrWhiteSpace(textBox2.Text)) // UPDATE logic
                    {
                        int addressId = int.Parse(textBox2.Text);

                        string updateQuery = @"
                    UPDATE Unit 
                    SET Types = @Types
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@Types", textBox1.Text);

                            cmd.Parameters.AddWithValue("@ID", addressId);

                            int rowsAffected = await cmd.ExecuteNonQueryAsync();

                            if (rowsAffected > 0)
                                MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("No matching record found to update.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else // INSERT logic
                    {
                        string checkQuery = "SELECT COUNT(*) FROM Unit WHERE Types = @Types";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@Types", textBox1.Text);

                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This Unit Type Existing. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                    INSERT INTO Unit (Types) 
                    VALUES (@Types);

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {

                            cmd.Parameters.AddWithValue("@Types", textBox1.Text);


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
        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        private void ResetFormFields()
        {
            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textbox)
                {
                    textBox1.Text = "";
                    textBox2.Text = "";
                }
            }

        }
        #endregion

        #region Validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {

         { textBox1, "Unit Type is required" },

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

        #endregion

        
    }
}
