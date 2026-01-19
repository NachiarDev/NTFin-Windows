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
    
    public partial class ProductSpecification : Form
    {
        Dictionary<string, int> sizeNameToID = new Dictionary<string, int>();
        CheckedListBox checkedListBox;
        Panel dropdownPanel;
        TextBox txtMultiSelect;
        public ProductSpecification()
        {
            InitializeComponent();
        }


        private void ProductSpecification_Load(object sender, EventArgs e)
        { 
            LoadProductName();
            InitializeMultiSelectUnderLabel5();
            textBox2.Visible = false;
         
        }
        #region Multi select
        private void InitializeMultiSelectUnderLabel5()
        {
            // Create the TextBox under label5
            txtMultiSelect = new TextBox();
            txtMultiSelect.ReadOnly = true;
            txtMultiSelect.Width = 200;
            txtMultiSelect.Location = new Point(label5.Left, label5.Bottom + 5);
            txtMultiSelect.Click += TxtMultiSelect_Click;
            this.Controls.Add(txtMultiSelect);

            // Create the dropdown Panel
            dropdownPanel = new Panel();
            dropdownPanel.BorderStyle = BorderStyle.FixedSingle;
            dropdownPanel.Width = txtMultiSelect.Width;
            dropdownPanel.Height = 100;
            dropdownPanel.Location = new Point(txtMultiSelect.Left, txtMultiSelect.Bottom);
            dropdownPanel.Visible = false;
            this.Controls.Add(dropdownPanel);

            // Create the CheckedListBox
            checkedListBox = new CheckedListBox();
            checkedListBox.Dock = DockStyle.Fill;
            checkedListBox.CheckOnClick = true;
            checkedListBox.SelectedIndexChanged += CheckedListBox_SelectedIndexChanged;
            dropdownPanel.Controls.Add(checkedListBox);

            LoadSizesFromDatabase();
        }
        private void TxtMultiSelect_Click(object sender, EventArgs e)
        {
            dropdownPanel.Visible = !dropdownPanel.Visible;
        }

        private void CheckedListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedNames = checkedListBox.CheckedItems.Cast<string>();
            txtMultiSelect.Text = string.Join(", ", selectedNames);
        }


        private void LoadSizesFromDatabase()
        {
            using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
            {
                string query = "SELECT ID, SizeName FROM Size WHERE Active = 1";
                SqlCommand cmd = new SqlCommand(query, conn);

                try
                {
                    conn.Open();
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        string sizeName = reader["SizeName"].ToString();
                        int sizeID = Convert.ToInt32(reader["ID"]);

                        checkedListBox.Items.Add(sizeName); // Add name to UI
                        sizeNameToID[sizeName] = sizeID;    // Store ID behind the scenes
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading sizes: " + ex.Message);
                }
            }
        }
        private string GetSelectedColors()
        {
            List<string> selectedColors = new List<string>();

            if (checkBox1.Checked) selectedColors.Add("1");
            if (checkBox2.Checked) selectedColors.Add("2");
            if (checkBox3.Checked) selectedColors.Add("3");

            return string.Join(",", selectedColors);
        }
        private string GetSelectedSizeIDs()
        {
            var selectedNames = checkedListBox.CheckedItems.Cast<string>();
            var selectedIDs = selectedNames.Select(name => sizeNameToID[name].ToString());
            return string.Join(",", selectedIDs);
        }

        #endregion

        #region Combo box
        private void LoadProductName()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, ProductName FROM Product";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "ProductName";  // Show country name
                        comboBox1.ValueMember = "ID";            // Store country ID
                        comboBox1.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Product Name: " + ex.Message);
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

                    bool isUpdate = !string.IsNullOrWhiteSpace(textBox2.Text);
                    int id = isUpdate ? Convert.ToInt32(textBox2.Text) : 0;

                    if (!isUpdate)
                    {
                        // Check duplicate only when inserting
                        string checkQuery = "SELECT COUNT(*) FROM Pro_Specification WHERE ProductID = @ProductID AND ModelNo = @ModelNo";

                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@ProductID", comboBox1.SelectedValue ?? DBNull.Value);
                            checkCmd.Parameters.AddWithValue("@ModelNo", textBox1.Text);

                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("This ProductName in this model already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                    }

                    string query;
                    if (isUpdate)
                    {
                        query = @"
                    UPDATE Pro_specification 
                    SET ProductID = @ProductID, ModelNo = @ModelNo, General = @General, 
                        Color = @Color, Size = @Size, BnchID = @BnchID, LocID = @LocID
                    WHERE ID = @ID";
                    }
                    else
                    {
                        query = @"
                    INSERT INTO Pro_specification (ProductID, ModelNo, General, Color, Size, BnchID, LocID) 
                    VALUES (@ProductID, @ModelNo, @General, @Color, @Size, @BnchID, @LocID);

                    SELECT SCOPE_IDENTITY();";
                    }

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ProductID", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@ModelNo", textBox1.Text);
                        cmd.Parameters.AddWithValue("@General", richTextBox1.Text);
                        cmd.Parameters.AddWithValue("@Color", GetSelectedColors());
                        cmd.Parameters.AddWithValue("@Size", GetSelectedSizeIDs());
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        if (isUpdate)
                        {
                            cmd.Parameters.AddWithValue("@ID", id);
                            await cmd.ExecuteNonQueryAsync();
                        }
                        else
                        {
                            object newID = await cmd.ExecuteScalarAsync(); // optional use
                        }
                    }

                    string message = isUpdate ? "Updated Successfully" : "Inserted Successfully";
                    MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        //            string checkQuery = "SELECT COUNT(*) FROM Pro_Specification WHERE ProductID = @ProductID AND ModelNo = @ModelNo";

        //            using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
        //            {
        //                checkCmd.Parameters.AddWithValue("@ProductID", comboBox1.SelectedValue ?? DBNull.Value);
        //                checkCmd.Parameters.AddWithValue("@ModelNo", textBox1.Text);

        //                int count = (int)await checkCmd.ExecuteScalarAsync();

        //                if (count > 0)
        //                {
        //                    MessageBox.Show("This ProductName in this model already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                    return;
        //                }
        //            }



        //            string insertProductspecificationQuery = @"
        //    INSERT INTO Pro_specification (ProductID, ModelNo, General, Color, Size,BnchID,LocID) 
        //    VALUES              (@ProductID, @ModelNo, @General, @Color, @Size,@BnchID,@LocID);

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertProductspecificationQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@ProductID", comboBox1.SelectedValue ?? DBNull.Value);

        //                cmd.Parameters.AddWithValue("@ModelNo", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@General", richTextBox1.Text);
        //                cmd.Parameters.AddWithValue("@Color", GetSelectedColors());
        //                cmd.Parameters.AddWithValue("@Size", GetSelectedSizeIDs());

        //                cmd.Parameters.AddWithValue("@BnchID", 1);
        //                cmd.Parameters.AddWithValue("@LocID", 1);
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


        #region Validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Model Number is required" },
        { richTextBox1, "General Data is required" },

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



            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Product Name is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }



            return true;
        }


        #endregion

        #region Reset
        private void ResetFormFields()
        {
            // Clear all TextBoxes
            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textBox)
                {
                    textBox.Text = "";
                }
            }

            // Reset ComboBox and RichTextBox
            comboBox1.SelectedIndex = -1;
            richTextBox1.Text = "";

            // ✅ Uncheck all CheckedListBox items
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, false);
            }

            // ✅ Clear multi-select TextBox and hide dropdown
            txtMultiSelect.Text = "";
            dropdownPanel.Visible = false;

            // ✅ Optional: reset checkbox colors
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();

        }
        #endregion
    }
}
