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
    public partial class TransportType : Form
    {
        public TransportType()
        {
            InitializeComponent();
        }

        private void TransportType_Load(object sender, EventArgs e)
        {
            textBox3.Visible = false;
        }



        #region Validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {

         { textBox1, "Transport Name is required" },
         { textBox2, "Vehicle Number is required" },

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



                    if (!string.IsNullOrWhiteSpace(textBox3.Text)) // UPDATE logic
                    {
                        int Transport = int.Parse(textBox3.Text);

                        string updateQuery = @"
                    UPDATE TransportType 
                    SET TransportName = @TransportName,VehicleNo=@VehicleNo
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@TransportName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@VehicleNo", textBox2.Text);
                            cmd.Parameters.AddWithValue("@ID", Transport);

                            int rowsAffected = await cmd.ExecuteNonQueryAsync();

                            if (rowsAffected > 0)
                                MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("No matching record found to update.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else // INSERT logic
                    {
                        string checkQuery = "SELECT COUNT(*) FROM TransportType WHERE TransportName = @TransportName AND VehicleNo = @VehicleNo";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@TransportName", textBox1.Text);
                            checkCmd.Parameters.AddWithValue("@VehicleNo", textBox2.Text);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This record already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                    INSERT INTO TransportType (TransportName,VehicleNo) 
                    VALUES (@TransportName,@VehicleNo);

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {

                            cmd.Parameters.AddWithValue("@TransportName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@VehicleNo", textBox2.Text);


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

        #region
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
                    textBox3.Text = "";

                }
            }

        }
        #endregion
    }
}
