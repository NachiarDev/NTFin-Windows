using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

namespace Estimate
{
   
    public partial class Expenses : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
        DataTable productTable;
        DataSet ds;
        private List<string> imagePaths = new List<string>();
        private int currentIndex = 0;
        private bool isImageZoomed = false;
        private Size originalSize;
        private Point originalLocation;

        private Button btnNext = new Button();
        private Button btnPrev = new Button();
        private Button closeButton = new Button();
        public Expenses()
        {
            InitializeComponent();
            SetupNavigationButtons();
        }

       

        private void Expenses_Load(object sender, EventArgs e)
        {
            LoadPaymentMethod();
            LoadStatus();
          textBox2.Visible=false;
        }
        #region Combobox 
        private void LoadPaymentMethod()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT ID, MethodName FROM PaymentMethod";
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                DataTable dt = new DataTable();
                da.Fill(dt);

                comboBox1.DisplayMember = "MethodName";
                comboBox1.ValueMember = "ID";
                comboBox1.DataSource = dt;
            }

          
        }
        private void LoadStatus()
        {
            using (SqlConnection con = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
            {
                string query = "SELECT ID, StatusName FROM Status";
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                DataTable dt = new DataTable();
                da.Fill(dt);

                comboBox2.DisplayMember = "StatusName";
                comboBox2.ValueMember = "ID";
                comboBox2.DataSource = dt;
            }


        }


        #endregion

        #region Picture
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
                    string query;
                    bool isUpdate = !string.IsNullOrWhiteSpace(textBox2.Text); // ID field (not ExpensesID)

                    if (isUpdate)
                    {
                        // Update existing record using ID
                        query = @"
                    UPDATE Expenses
                    SET 
                        Amount = @Amount,
                        PaymentMethod = @PaymentMethod,
                        Status = @Status,
                        Notes = @Notes,
                        Image = @image,
                        BnchID = @BnchID,
                        LocID = @LocID,
                        CreatedBy = @CreatedBy,
                        CreatedDate = GETDATE(),
                        ApprovedBy = @ApprovedBy,
                        ApprovedDate = GETDATE()
                    WHERE ID = @ID;";
                    }
                    else
                    {
                        // Insert new record with a random ExpensesID
                        Random rand = new Random();
                        int expensesId = rand.Next(100000, 1000000);

                        query = @"
                    INSERT INTO Expenses 
                    (ExpensesID, Date, Amount, PaymentMethod, Status, Notes, CreatedBy, CreatedDate, ApprovedBy, ApprovedDate, Image, BnchID, LocID) 
                    VALUES 
                    (@ExpensesID, GETDATE(), @Amount, @PaymentMethod, @Status, @Notes, @CreatedBy, GETDATE(), @ApprovedBy, GETDATE(), @image, @BnchID, @LocID);";
                    }

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        if (isUpdate)
                        {
                            cmd.Parameters.AddWithValue("@ID", Convert.ToInt32(textBox2.Text)); // use ID for update
                        }
                        else
                        {
                            Random rand = new Random();
                            int expensesId = rand.Next(100000, 1000000);
                            cmd.Parameters.AddWithValue("@ExpensesID", expensesId); // use random ID for insert
                        }

                        cmd.Parameters.AddWithValue("@Amount", textBox1.Text);
                        cmd.Parameters.AddWithValue("@PaymentMethod", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Status", comboBox2.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Notes", richTextBox1.Text);
                        cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@BnchID", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@LocID", 1);
                        cmd.Parameters.AddWithValue("@CreatedBy", 1);
                        cmd.Parameters.AddWithValue("@ApprovedBy", 1);

                        await cmd.ExecuteNonQueryAsync();
                    }

                    MessageBox.Show(isUpdate ? "Updated Successfully" : "Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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


        //            Random rand = new Random();
        //            int expensesId = rand.Next(100000, 1000000);



        //            string insertOrganisationQuery = @"
        //    INSERT INTO Expenses (ExpensesID, Date, Amount,PaymentMethod,Status,Notes,CreatedBy,CreatedDate,ApprovedBy,ApprovedDate,Image,BnchID,LocID) 
        //    VALUES              (@ExpensesID,GETDATE(),@Amount,@PaymentMethod,@Status,@Notes,@CreatedBy,GETDATE(),@ApprovedBy,GETDATE(),@image,@BnchID,@LocID);

        //    SELECT SCOPE_IDENTITY();";

        //            using (SqlCommand cmd = new SqlCommand(insertOrganisationQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@ExpensesID", expensesId);
        //                cmd.Parameters.AddWithValue("@Amount", textBox1.Text);
        //                cmd.Parameters.AddWithValue("@PaymentMethod", comboBox1.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@Status", comboBox2.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@Notes", richTextBox1.Text);
        //                cmd.Parameters.AddWithValue("@image", imagesForDb ?? (object)DBNull.Value);
        //                cmd.Parameters.AddWithValue("@BnchID", comboBox1.SelectedValue ?? DBNull.Value);
        //                cmd.Parameters.AddWithValue("@LocID", 1);
        //                cmd.Parameters.AddWithValue("@CreatedBy", 1);
        //                cmd.Parameters.AddWithValue("@ApprovedBy", 1);
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
        { textBox1, "Amount is required" },
                {richTextBox1,"Description is required" }
    };

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Payment method is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }

            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Payment Status is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Focus();
                return false;
            }
            if (pictureBox1.Image == null)
            {
                MessageBox.Show("At least one image must be uploaded", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                pictureBox1.Focus();
                return false;
            }
            return true;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }


        #endregion

       
    }
}
