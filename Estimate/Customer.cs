using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Estimate
{
    public partial class Customer : Form
    {
        private List<string> imagePaths = new List<string>();
        private int currentIndex = 0;
        private bool isImageZoomed = false;
        private Size originalSize;
        private Point originalLocation;

        private System.Windows.Forms.Button btnNext = new System.Windows.Forms.Button();
        private System.Windows.Forms.Button btnPrev = new System.Windows.Forms.Button();
        private System.Windows.Forms.Button closeButton = new System.Windows.Forms.Button();

        public Customer()
        {
            InitializeComponent();
            
            SetupNavigationButtons();

        }

       

        #region UploadImage
        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select up to 3 Images";
                openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (openFileDialog.FileNames.Length > 3)
                    {
                        MessageBox.Show("You can upload only up to 3 images.", "Image Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string customerFolder = @"C:\Malathi\Estimate\Customer";
                    Directory.CreateDirectory(customerFolder);

                    imagePaths.Clear();

                    foreach (string selectedFilePath in openFileDialog.FileNames)
                    {
                        string fileName = Path.GetFileName(selectedFilePath);
                        string destinationPath = Path.Combine(customerFolder, fileName);

                        File.Copy(selectedFilePath, destinationPath, true);
                        imagePaths.Add(destinationPath);
                    }


                    if (imagePaths.Count > 0)
                    {
                        pictureBox1.Image = System.Drawing.Image.FromFile(imagePaths[0]);
                        pictureBox1.SizeMode = PictureBoxSizeMode.Zoom; // Optional: Fit image
                    }

                   
                }
            }
        }
        #endregion
        private void Customer_Load(object sender, EventArgs e)
        {
            LoadCountries();
            LoadGSTType();
            LoadGSTAreaCode();
            textBox3.Visible = false;
           
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
        private void LoadGSTType()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, GSTType FROM Tbl_GSTType WHERE Active = 1";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox3.DataSource = dt;
                        comboBox3.DisplayMember = "GSTType";  // Show country name
                        comboBox3.ValueMember = "ID";            // Store country ID
                        comboBox3.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading GST Type: " + ex.Message);
            }
        }
        private void LoadGSTAreaCode()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, StateName,GSTTIN FROM Tbl_State WHERE Active = 1";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        dt.Columns.Add("DisplayText", typeof(string), "StateName + ' - ' + GSTTIN");
                        comboBox4.DataSource = dt;
                        comboBox4.DisplayMember = "DisplayText";  // Show country name
                        comboBox4.ValueMember = "ID";            // Store country ID
                        comboBox4.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Area Code: " + ex.Message);
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

                    string imagesForDb = string.Join("|", imagePaths);
                    string formattedPrimaryContact = $"({maskedTextBox1.Text.Substring(0, 3)}) {maskedTextBox1.Text.Substring(3, 3)}-{maskedTextBox1.Text.Substring(6, 4)}";
                    string formattedAlternateContact = $"({maskedTextBox2.Text.Substring(0, 3)}) {maskedTextBox2.Text.Substring(3, 3)}-{maskedTextBox2.Text.Substring(6, 4)}";

                    // Check if ID (textBox3) has value
                    if (!string.IsNullOrWhiteSpace(textBox3.Text)) // ====> UPDATE block
                    {
                        int customerId = int.Parse(textBox3.Text);

                        string updateQuery = @"
                UPDATE Customer SET
                    Code = @Code,
                    CustomerName = @CustomerName,
                    ContactNo = @ContactNo,
                    MobileNo = @MobileNo,
                    Email = @Email,
                    CountryID = @CountryID,
                    StateID = @StateID,
                    GSTTypeID = @GSTTypeID,
                    GSTAreaCode = @GSTAreaCode,
                    GSTNumber = @GSTNumber,
                    image = @image,
                    ModifiedBy = @ModifiedBy,
                    ModifiedDate = GETDATE()
                WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", customerId);
                            cmd.Parameters.AddWithValue("@Code", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CustomerName", textBox2.Text);
                            cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@MobileNo", formattedAlternateContact);
                            cmd.Parameters.AddWithValue("@Email", textBox5.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTTypeID", comboBox3.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTAreaCode", comboBox4.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTNumber", textBox6.Text);
                            cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@ModifiedBy", 1);
                            await cmd.ExecuteNonQueryAsync();
                            MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else // ====> INSERT block
                    {
                        string checkQuery = "SELECT COUNT(*) FROM Customer WHERE CustomerName = @CustomerName AND ContactNo=@ContactNo";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@CustomerName", textBox2.Text);
                            checkCmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
                            int count = (int)await checkCmd.ExecuteScalarAsync();

                            if (count > 0)
                            {
                                MessageBox.Show("A Customer with this Mobile Number already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        string insertQuery = @"
                INSERT INTO Customer (Code, CustomerName, ContactNo, MobileNo, Email, CountryID, StateID, GSTTypeID, GSTAreaCode, GSTNumber, image, BnchID, LocID, CreatedBy, CreatedDate) 
                VALUES (@Code, @CustomerName, @ContactNo, @MobileNo, @Email, @CountryID, @StateID, @GSTTypeID, @GSTAreaCode, @GSTNumber, @image, @BnchID, @LocID, @CreatedBy, GETDATE());
                SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@Code", textBox1.Text);
                            cmd.Parameters.AddWithValue("@CustomerName", textBox2.Text);
                            cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
                            cmd.Parameters.AddWithValue("@MobileNo", formattedAlternateContact);
                            cmd.Parameters.AddWithValue("@Email", textBox5.Text);
                            cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTTypeID", comboBox3.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTAreaCode", comboBox4.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@GSTNumber", textBox6.Text);
                            cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);
                            object newIDObj = await cmd.ExecuteScalarAsync();
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

        //            // Check if Organisation Name and City already exist together
        //            //string checkQuery = "SELECT COUNT(*) FROM Organisation WHERE Name = @Name AND City = @City";
        //            string checkQuery = "SELECT COUNT(*) FROM Customer WHERE CustomerName = @CustomerName AND ContactNo=@ContactNo";
        //            string imagesForDb = string.Join("|", imagePaths);
        //            string formattedPrimaryContact = $"({maskedTextBox1.Text.Substring(0, 3)}) {maskedTextBox1.Text.Substring(3, 3)}-{maskedTextBox1.Text.Substring(6, 4)}";
        //            string formattedAlternateContact = $"({maskedTextBox2.Text.Substring(0, 3)}) {maskedTextBox2.Text.Substring(3, 3)}-{maskedTextBox2.Text.Substring(6, 4)}";
        //            using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
        //            {
        //                checkCmd.Parameters.AddWithValue("@CustomerName", textBox2.Text);
        //                checkCmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
        //                int count = (int)await checkCmd.ExecuteScalarAsync();

        //                if (count > 0)
        //                {
        //                    MessageBox.Show("An Customer with this Mobile Number already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                    return;
        //                }
        //            }




        //            string insertOrganisationQuery = @"
        //    INSERT INTO Customer (Code, CustomerName, ContactNo, MobileNo, Email, CountryID, StateID,GSTTypeID,GSTAreaCode,GSTNumber,image,BnchID,LocID,CreatedBy,CreatedDate) 
        //    VALUES              (@Code, @CustomerName, @ContactNo, @MobileNo, @Email,  @CountryID, @StateID, @GSTTypeID,@GSTAreaCode,@GSTNumber,@image,@BnchID,@LocID,@CreatedBy,GETDATE());

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertOrganisationQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@Code", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@CustomerName", textBox2.Text);
        //                cmd.Parameters.AddWithValue("@ContactNo", formattedPrimaryContact);
        //                cmd.Parameters.AddWithValue("@MobileNo", formattedAlternateContact);
        //                cmd.Parameters.AddWithValue("@Email", textBox5.Text);
        //                cmd.Parameters.AddWithValue("@CountryID", comboBox1.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@StateID", comboBox2.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@GSTTypeID", comboBox3.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@GSTAreaCode", comboBox4.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@GSTNumber", textBox6.Text);

        //                cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
        //                //cmd.Parameters.AddWithValue("@image", image ?? (object)DBNull.Value);
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
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Customer Code is required" },
        { textBox2, "Customer Name is required" },
        { maskedTextBox1, "primary contact Number is required" },
        { maskedTextBox2, "Alternate contact Number is required" },
        { textBox5, "Email is required" },
        

        
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
            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("GST Type is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox3.Focus();
                return false;
            }
            if (comboBox4.SelectedIndex == -1)
            {
                MessageBox.Show("GST Area Code is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox4.Focus();
                return false;
            }

            return true;
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {
            string gst = textBox6.Text.Trim();
            string gstpattern = @"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[0-9]{1}[A-Z]{1}[0-9]{1}$";
            if (!Regex.IsMatch(gst, gstpattern))
            {
                MessageBox.Show("Invaid GST Format!Please enter a valid GST Number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox6.Focus();
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
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            maskedTextBox2.Text = "";
            maskedTextBox1.Text = "";

            pictureBox1.Visible = false;

            // Reset and hide zoom buttons if visible
            closeButton.Visible = false;
            btnNext.Visible = false;
            btnPrev.Visible = false;

            // Reset zoom flag
            isImageZoomed = false;





            pictureBox1.Image = null;






        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();

        }
        #endregion
       
        
        #region Picture
        private void SetupNavigationButtons()
        {

            closeButton.Text = "X";
            closeButton.Size = new Size(30, 30);
            closeButton.Visible = false;
            closeButton.Click += CloseButton_Click;
            this.Controls.Add(closeButton);


            btnNext.Text = ">";
            btnNext.Size = new Size(40, 40);
            btnNext.Visible = false;
            btnNext.Click += BtnNext_Click;
            this.Controls.Add(btnNext);


            btnPrev.Text = "<";
            btnPrev.Size = new Size(40, 40);
            btnPrev.Visible = false;
            btnPrev.Click += BtnPrev_Click;
            this.Controls.Add(btnPrev);
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (imagePaths.Count == 0) return;

            if (!isImageZoomed)
            {
                
                originalSize = pictureBox1.Size;
                originalLocation = pictureBox1.Location;

              
                pictureBox1.Size = new Size(467, 440);
                pictureBox1.Location = new Point(
                    (this.ClientSize.Width - pictureBox1.Width) / 2,
                    (this.ClientSize.Height - pictureBox1.Height) / 2
                );
                pictureBox1.BringToFront();
                pictureBox1.BorderStyle = BorderStyle.FixedSingle;

                // Close Button
                closeButton.Location = new Point(
                    pictureBox1.Right - closeButton.Width - 10,
                    pictureBox1.Top + 10
                );
                closeButton.BringToFront();
                closeButton.Visible = true;

                // Prev Button
                btnPrev.Location = new Point(pictureBox1.Left + 10, pictureBox1.Top + (pictureBox1.Height / 2) - 20);
                btnPrev.BringToFront();
                btnPrev.Visible = true;

                // Next Button
                btnNext.Location = new Point(pictureBox1.Right - btnNext.Width - 10, pictureBox1.Top + (pictureBox1.Height / 2) - 20);
                btnNext.BringToFront();
                btnNext.Visible = true;

                isImageZoomed = true;
            }
        }
        private void BtnNext_Click(object sender, EventArgs e)
        {
            if (currentIndex < imagePaths.Count - 1)
            {
                currentIndex++;
                ShowImageAtCurrentIndex();
            }
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            if (currentIndex > 0)
            {
                currentIndex--;
                ShowImageAtCurrentIndex();
            }
        }
        private void ShowImageAtCurrentIndex()
        {
            if (currentIndex >= 0 && currentIndex < imagePaths.Count)
            {
                pictureBox1.Image = System.Drawing.Image.FromFile(imagePaths[currentIndex]);
                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            // Reset image
            pictureBox1.Size = originalSize;
            pictureBox1.Location = originalLocation;
            pictureBox1.BorderStyle = BorderStyle.None;

            // Hide buttons
            closeButton.Visible = false;
            btnNext.Visible = false;
            btnPrev.Visible = false;

            isImageZoomed = false;
        }


        #endregion




    }
}
