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
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            LoadOrganisations();
            LoadCountry();
            Loadziplength();
            LoadzipType();


        }
        #region ComboBox fetch
        public void LoadzipType()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, CharType FROM ZipType";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        // Bind data to ComboBox
                        comboBox4.DataSource = dt;
                        comboBox4.DisplayMember = "CharType";  // What the user sees
                        comboBox4.ValueMember = "ID";      // The actual value
                        comboBox4.SelectedIndex = -1;      // No default selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading ZipType: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void Loadziplength()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, CharSize FROM ziplength";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        // Bind data to ComboBox
                        comboBox3.DataSource = dt;
                        comboBox3.DisplayMember = "CharSize";  // What the user sees
                        comboBox3.ValueMember = "ID";      // The actual value
                        comboBox3.SelectedIndex = -1;      // No default selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading ziplength: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
      
        public void LoadCountry()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, CountryName, CurrencyName FROM Country_Settings";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        // Bind data to ComboBox
                        comboBox2.DataSource = dt;
                        comboBox2.DisplayMember = "CountryName";  // What the user sees
                        comboBox2.ValueMember = "ID";             // The actual value
                        comboBox2.SelectedIndex = -1;             // No default selection

                        // Store currency data for later retrieval
                        comboBox2.Tag = dt;
                        comboBox2.SelectedIndexChanged += comboBox2_SelectedIndexChanged;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Country: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != -1)
            {
                DataTable dt = comboBox2.Tag as DataTable;
                if (dt != null)
                {
                    DataRow[] rows = dt.Select("ID = " + comboBox2.SelectedValue);
                    if (rows.Length > 0)
                    {
                        string currencyName = rows[0]["CurrencyName"].ToString();
                        textBox1.Text = currencyName;  // Display currency in a TextBox
                    }
                }
            }
        }
        private void LoadOrganisations()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, Name FROM Organisation";

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
        private async void InsertData()
        {
            try
            {
               
                if (string.IsNullOrWhiteSpace(comboBox1.Text) ||
                    string.IsNullOrWhiteSpace(textBox1.Text) ||
                    comboBox1.SelectedValue == null ||
                    comboBox2.SelectedValue == null ||
                    comboBox3.SelectedValue == null ||
                    comboBox4.SelectedValue == null)
                {
                    MessageBox.Show("All fields except TextBox2 are required. Please fill in all details.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Ensure that at least one radio button is selected
                if (!radioButton1.Checked && !radioButton2.Checked)
                {
                    MessageBox.Show("Please select Active status.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }


                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    await conn.OpenAsync();

                    string checkQuery = "SELECT COUNT(*) FROM Settingspage WHERE OrganisationName = @OrganisationName";
                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                    {
                        if (comboBox1.SelectedValue == null)
                        {
                            MessageBox.Show("Please select a valid organisation.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        checkCmd.Parameters.AddWithValue("@OrganisationName", comboBox1.SelectedValue);

                        int count = (int)await checkCmd.ExecuteScalarAsync();

                        if (count > 0)
                        {
                            MessageBox.Show("An organisation with this ID already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    int activeStatus = radioButton1.Checked ? 1 : 0;

                    string insertSettingsQuery = @"
            INSERT INTO Settingspage (OrganisationName, Country, Currency, Zip, Zip_Format, Active) 
                              VALUES (@OrganisationName, @Country, @Currency, @Zip, @Zip_Format, @Active);

            SELECT SCOPE_IDENTITY();";
                    
                    using (SqlCommand cmd = new SqlCommand(insertSettingsQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Currency", textBox1.Text);                   
                        cmd.Parameters.AddWithValue("@OrganisationName", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Country", comboBox2.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Zip", comboBox3.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Zip_Format", comboBox4.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Active", activeStatus);

                        object newIDObj = await cmd.ExecuteScalarAsync();
                    }

                    MessageBox.Show("Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ResetFormFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            InsertData();
        }
        #endregion

        #region Reset
        private void ResetFormFields()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            radioButton1.Checked=false;
            radioButton2.Checked=false;
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
