using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Estimate
{
    public partial class BranchDocument : Form
    {
        private List<string> imagePaths = new List<string>();
        private int currentIndex = 0;
        private bool isImageZoomed = false;
        private Size originalSize;
        private Point originalLocation;

        private Button btnNext = new Button();
        private Button btnPrev = new Button();
        private Button closeButton = new Button();
        public BranchDocument()
        {
            InitializeComponent();
            SetupNavigationButtons();
        }

        private void BranchDocument_Load(object sender, EventArgs e)
        {
            LoadBranch();
            textBox2.Visible= false;

        }
        private void LoadBranch()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, Name FROM Branch ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "Name";  // Show country name
                        comboBox1.ValueMember = "ID";            // Store country ID
                        comboBox1.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Error loading BranchName " + ex.Message);
            }
        }




        #region Picture
        private void button1_Click(object sender, EventArgs e)
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

                    string branchDocumentFolder = @"C:\Malathi\Estimate\BranchDocument";
                    Directory.CreateDirectory(branchDocumentFolder);

                    imagePaths.Clear();

                    foreach (string selectedFilePath in openFileDialog.FileNames)
                    {
                        string fileName = Path.GetFileName(selectedFilePath);
                        string destinationPath = Path.Combine(branchDocumentFolder, fileName);

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
        private void CloseButton_Click(object sender, EventArgs e)
        {
            
            pictureBox1.Size = originalSize;
            pictureBox1.Location = originalLocation;
            pictureBox1.BorderStyle = BorderStyle.None;

            closeButton.Visible = false;
            btnNext.Visible = false;
            btnPrev.Visible = false;

            isImageZoomed = false;
        }



        #endregion

        





        #region Insert
        private void button2_Click(object sender, EventArgs e)
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
                    int.TryParse(textBox2.Text.Trim(), out int docID); // this will be used for update

                    if (docID > 0)
                    {
                        // UPDATE existing record
                        string updateQuery = @"
                    UPDATE BnchDoc 
                    SET 
                        FileName = @FileName,
                        docFile = @image,
                        BnchID = @BnchID,
                        LocID = @LocID,
                        ModifiedBy = @ModifiedBy,
                        ModifiedDate = GETDATE()
                    WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@FileName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@BnchID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@LocID", 1);
                            cmd.Parameters.AddWithValue("@ModifiedBy", 1);
                            cmd.Parameters.AddWithValue("@ID", docID);

                            await cmd.ExecuteNonQueryAsync();
                        }

                        MessageBox.Show("Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // INSERT new record
                        string insertQuery = @"
                    INSERT INTO BnchDoc (FileName, docFile, BnchID, LocID, CreatedBy, CreatedDate) 
                    VALUES              (@FileName, @image, @BnchID, @LocID, @CreatedBy, GETDATE());

                    SELECT SCOPE_IDENTITY();";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@FileName", textBox1.Text);
                            cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@BnchID", comboBox1.SelectedValue ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@LocID", 1);
                            cmd.Parameters.AddWithValue("@CreatedBy", 1);

                            object newIDObj = await cmd.ExecuteScalarAsync();
                            textBox2.Text = newIDObj.ToString(); // store the new ID in textBox2
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

        //private async void InsertData()
        //{
        //    try
        //    {

        //        if (!ValidateFields()) return;

        //        using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
        //        {
        //            await conn.OpenAsync();


        //            string imagesForDb = string.Join("|", imagePaths);






        //            string insertOrganisationQuery = @"
        //    INSERT INTO BnchDoc (FileName, docFile, BnchID,LocID,CreatedBy, CreatedDate) 
        //    VALUES              (@FileName,@image,@BnchID,@LocID,@CreatedBy,GETDATE());

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertOrganisationQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@FileName", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
        //                cmd.Parameters.AddWithValue("@BnchID", comboBox1.SelectedValue ?? DBNull.Value);
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
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Branch Name is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }
            if (pictureBox1.Image == null)
            {
                MessageBox.Show("Please upload an image before submitting.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                pictureBox1.Focus();
                return false;
            }
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "File Name is required" },
        


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

            pictureBox1.Visible = false;

            // Reset and hide zoom buttons if visible
            closeButton.Visible = false;
            btnNext.Visible = false;
            btnPrev.Visible = false;

            
            isImageZoomed = false;



            comboBox1.SelectedIndex = -1;
           

            pictureBox1.Image = null;



            


        }

        private void button3_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }

        #endregion


    }
}
