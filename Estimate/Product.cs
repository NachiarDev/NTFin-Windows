using AForge.Video;
using AForge.Video.DirectShow;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Estimate
{
    public partial class Product : Form
    {
        public Product()
        {
            InitializeComponent();
            SetupNavigationButtons();
        }
        Dictionary<string, int> categoryDict = new Dictionary<string, int>();
        Dictionary<string, int> subCategoryDict = new Dictionary<string, int>();
        private List<string> imagePaths = new List<string>();
        private int currentIndex = 0;
        private bool isImageZoomed = false;
        private Size originalSize;
        private Point originalLocation;

        private Button btnNext = new Button();
        private Button btnPrev = new Button();
        private Button closeButton = new Button();
        CheckedListBox clbCategory = new CheckedListBox();
        TextBox txtCategory = new TextBox();
        Button btnDropCategory = new Button();
        CheckedListBox clbSubCategory = new CheckedListBox();
        TextBox txtSubCategory = new TextBox();
        Button btnDropSubCategory = new Button();
        private VideoCaptureDevice videoSource;
        private FilterInfoCollection videoDevices;
        private List<Bitmap> capturedImages = new List<Bitmap>();
        private int capturedCount = 0;
        private Bitmap lastFrame = null;
        private void Product_Load(object sender, EventArgs e)
        {
            textBox10.Visible=false;
            LoadUnitType();
            LoadTaxType();
            LoadCategories();
            SetupSubCategoryDropdown();
            textBox3.KeyPress += OnlyAllowDigits;
            textBox4.KeyPress += OnlyAllowDigits;
            textBox5.KeyPress += OnlyAllowDigits;
            textBox6.KeyPress += OnlyAllowDigits;
            textBox7.KeyPress += OnlyAllowDigits;
            textBox8.KeyPress += OnlyAllowDigits;
            int startX = label13.Left;
            int startY = label13.Bottom + 5;

            txtCategory.Location = new Point(startX, startY);
            txtCategory.Width = 150;
            this.Controls.Add(txtCategory);

            btnDropCategory.Text = "▼";
            btnDropCategory.Location = new Point(txtCategory.Right, startY);
            btnDropCategory.Width = 30;
            this.Controls.Add(btnDropCategory);
            btnDropCategory.Click += BtnDropCategory_Click;

            clbCategory.Location = new Point(startX, txtCategory.Bottom);
            clbCategory.Width = txtCategory.Width + btnDropCategory.Width;
            clbCategory.Height = 100;
            clbCategory.Visible = false;
            clbCategory.CheckOnClick = true;

            
            clbCategory.ItemCheck += ClbCategory_ItemCheck;
            this.Controls.Add(clbCategory);

        }
     
        
        #region Category
        private void BtnDropCategory_Click(object sender, EventArgs e)
        {
            clbCategory.Visible = !clbCategory.Visible;
        }

        private void ClbCategory_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate {
                List<string> selectedItems = new List<string>();

                foreach (var item in clbCategory.CheckedItems)
                {
                    selectedItems.Add(item.ToString());
                }

                // Also include the current item being (un)checked
                string currentItem = clbCategory.Items[e.Index].ToString();

                if (e.NewValue == CheckState.Checked && !selectedItems.Contains(currentItem))
                    selectedItems.Add(currentItem);
                else if (e.NewValue == CheckState.Unchecked && selectedItems.Contains(currentItem))
                    selectedItems.Remove(currentItem);

                txtCategory.Text = string.Join(", ", selectedItems);
                LoadSubCategoriesBySelectedCategories();
            });
        }
        private void SetupSubCategoryDropdown()
        {
            int startX = label14.Left;
            int startY = label14.Bottom + 5;

            // TextBox for showing selected subcategories
            txtSubCategory.Location = new Point(startX, startY);
            txtSubCategory.Width = 150;
            this.Controls.Add(txtSubCategory);

            // Button for dropdown arrow
            btnDropSubCategory.Text = "▼";
            btnDropSubCategory.Location = new Point(txtSubCategory.Right, startY);
            btnDropSubCategory.Width = 30;
            this.Controls.Add(btnDropSubCategory);
            btnDropSubCategory.Click += BtnDropSubCategory_Click;

            // CheckedListBox for multiple selection
            clbSubCategory.Location = new Point(startX, txtSubCategory.Bottom);
            clbSubCategory.Width = txtSubCategory.Width + btnDropSubCategory.Width;
            clbSubCategory.Height = 100;
            clbSubCategory.Visible = false;
            clbSubCategory.CheckOnClick = true;
            clbSubCategory.ItemCheck += ClbSubCategory_ItemCheck;

            this.Controls.Add(clbSubCategory);
        }

        private void BtnDropSubCategory_Click(object sender, EventArgs e)
        {
            clbSubCategory.Visible = !clbSubCategory.Visible;
        }

        private void ClbSubCategory_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Ensure that the index is valid and within bounds
            if (e.Index >= 0 && e.Index < clbSubCategory.Items.Count)
            {
                var item = clbSubCategory.Items[e.Index];
                // Proceed with the rest of your logic
            }
            else
            {
                // Handle invalid index scenario (perhaps log or display a message)
                MessageBox.Show("Invalid index in CheckedListBox.");
            }
        }

        private void LoadSubCategoriesBySelectedCategories()
        {
            clbSubCategory.Items.Clear();
            subCategoryDict.Clear(); // ✅ Clear previous mappings

            List<int> selectedCategoryIDs = new List<int>();

            using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
            {
                try
                {
                    conn.Open();

                    foreach (var selectedCat in clbCategory.CheckedItems)
                    {
                        string name = selectedCat.ToString();

                        if (categoryDict.ContainsKey(name))
                            selectedCategoryIDs.Add(categoryDict[name]);
                    }

                    if (selectedCategoryIDs.Count == 0) return;

                    string inClause = string.Join(",", selectedCategoryIDs);

                    string subCatQuery = $"SELECT ID, SubCategoryName FROM SubCategory WHERE CategoryID IN ({inClause})";
                    using (SqlCommand subCatCmd = new SqlCommand(subCatQuery, conn))
                    {
                        SqlDataReader reader = subCatCmd.ExecuteReader();
                        while (reader.Read())
                        {
                            string subName = reader["SubCategoryName"].ToString();
                            int subId = Convert.ToInt32(reader["ID"]);
                            clbSubCategory.Items.Add(subName);
                            subCategoryDict[subName] = subId; // ✅ Add to dictionary
                        }
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading subcategories: " + ex.Message);
                }
            }
        }
        private void LoadCategories()
        {
            clbCategory.Items.Clear();        // ✅ clear items
            categoryDict.Clear();             // ✅ clear dictionary to avoid key conflict

            using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("SELECT ID, Name FROM Category", conn))
                {
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        string name = reader["Name"].ToString();
                        int id = Convert.ToInt32(reader["ID"]);

                        if (!categoryDict.ContainsKey(name)) // prevent duplicate keys
                        {
                            clbCategory.Items.Add(name);      // ✅ add only once
                            categoryDict[name] = id;
                        }
                    }
                }
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
                    string checkQuery = "SELECT COUNT(*) FROM Product WHERE ProductName = @ProductName";
                    string imagesForDb = string.Join("|", imagePaths);
                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                    {
                        checkCmd.Parameters.AddWithValue("@ProductName", textBox9.Text);
                        int count = (int)await checkCmd.ExecuteScalarAsync();

                        if (count > 0)
                        {
                            MessageBox.Show("This ProductName already exists. Please enter different details.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    string selectedCategoryIDs = string.Join(",", clbCategory.CheckedItems
    .OfType<string>()
    .Where(name => categoryDict.ContainsKey(name))
    .Select(name => categoryDict[name]));

                    // Convert selected subcategories to comma-separated IDs
                    string selectedSubCategoryIDs = string.Join(",", clbSubCategory.CheckedItems
                        .OfType<string>()
                        .Where(name => subCategoryDict.ContainsKey(name))
                        .Select(name => subCategoryDict[name]));


                    string insertProductQuery = @"
            INSERT INTO Product (Code, ProductName, Unit, HSNCode, Qty, Rate, Tax,Discount,MRP,TotalQty,TaxType,Image,SelectedCategory,selectedSubCategory,BnchID,LocID) 
            VALUES              (@Code, @ProductName, @Unit, @HSNCode, @Qty,  @Rate, @Tax, @Discount,@MRP,@TotalQty,@TaxType,@Image,@SelectedCategory,@selectedSubCategory,@BnchID,@LocID);

            SELECT SCOPE_IDENTITY();";

                    using (SqlCommand cmd = new SqlCommand(insertProductQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Code", textBox1.Text);
                        cmd.Parameters.AddWithValue("@ProductName", textBox9.Text);
                        cmd.Parameters.AddWithValue("@Unit", comboBox2.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@SelectedCategory", selectedCategoryIDs);
                        cmd.Parameters.AddWithValue("@selectedSubCategory", selectedSubCategoryIDs);
                        cmd.Parameters.AddWithValue("@HSNCode", textBox2.Text);
                        cmd.Parameters.AddWithValue("@Qty", textBox3.Text);
                        cmd.Parameters.AddWithValue("@Rate", textBox4.Text);
                        cmd.Parameters.AddWithValue("@Discount", textBox5.Text);
                        cmd.Parameters.AddWithValue("@Tax", textBox6.Text);
                        cmd.Parameters.AddWithValue("@MRP", textBox7.Text);
                        cmd.Parameters.AddWithValue("@TotalQty", textBox8.Text);
                        cmd.Parameters.AddWithValue("@TaxType", comboBox3.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Image", imagesForDb ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);
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

        #endregion

        #region Combobox
        private void LoadUnitType()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, Types FROM Unit";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox2.DataSource = dt;
                        comboBox2.DisplayMember = "Types";  // Show country name
                        comboBox2.ValueMember = "ID";            // Store country ID
                        comboBox2.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Unit Type: " + ex.Message);
            }
        }
        private void LoadTaxType()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, gsttype FROM Tax";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox3.DataSource = dt;
                        comboBox3.DisplayMember = "gsttype";  // Show country name
                        comboBox3.ValueMember = "ID";            // Store country ID
                        comboBox3.SelectedIndex = -1;           // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Unit Type: " + ex.Message);
            }
        }
        #endregion

        #region Validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Product Code is required" },
        { textBox9, "Product Name is required" },
       { textBox2, "HSN Code is required" },
        { textBox3, "Quantity is required" },
        { textBox4, "Price is required" },
        { textBox5, "Discount is required" },
        { textBox6, "Tax is required" },
        { textBox7, "MRP is required" },
        { textBox8, "Total Quantity is required" },


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
          

           
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Unit Type is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Focus();
                return false;
            }

            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Tax is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox3.Focus();
                return false;
            }
           

            return true;
        }
        private void OnlyAllowDigits(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
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
            for (int i = 0; i < clbCategory.Items.Count; i++)
            {
                clbCategory.SetItemChecked(i, false);
            }

            for (int i = 0; i < clbSubCategory.Items.Count; i++)
            {
                clbSubCategory.SetItemChecked(i, false);
            }
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;

            pictureBox1.Visible = false;
            closeButton.Visible = false;
            btnNext.Visible = false;
            btnPrev.Visible = false;

            isImageZoomed = false;
            pictureBox1.Image = null;

        }



        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }




        #endregion

        #region UploadImage

        private void button4_Click(object sender, EventArgs e)
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

                    string customerFolder = @"C:\Malathi\Estimate\Product";
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
       
        
        #region Picture
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



        #region Take pic
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                // Use the class-level field instead of declaring a local variable
                videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

                if (videoDevices.Count == 0)
                {
                    MessageBox.Show("No video devices found.");
                    return;
                }

                // Start the video capture with the first available video device
                videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
                videoSource.NewFrame += videoSource_NewFrame;
                videoSource.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error starting camera: " + ex.Message);
            }
        }
        private void videoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            try
            {
                // Dispose previous frame if it exists
                lastFrame?.Dispose();

                // Clone the new frame
                lastFrame = (Bitmap)eventArgs.Frame.Clone();

                // Optional: Show it in a PictureBox (for preview)
                pictureBox2.Invoke(new Action(() =>
                {
                    pictureBox2.Image?.Dispose();
                    pictureBox2.Image = (Bitmap)lastFrame.Clone();
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error capturing frame: " + ex.Message);
            }
        }


        

        private void button5_Click(object sender, EventArgs e)
        {
            if (lastFrame == null)
            {
                MessageBox.Show("No image captured from camera.");
                return;
            }

            // Define the customer folder where images will be stored
            string customerFolder = @"C:\Malathi\Estimate\Product";
            Directory.CreateDirectory(customerFolder);

            // Generate a unique file name for the captured image
            string fileName = "CapturedImage_" + Guid.NewGuid().ToString() + ".jpg";
            string destinationPath = Path.Combine(customerFolder, fileName);

            try
            {
                // Save the captured image to the folder
                lastFrame.Save(destinationPath, System.Drawing.Imaging.ImageFormat.Jpeg);

                // Add the path of the saved image to the imagePaths list
                imagePaths.Add(destinationPath);

                // Display the captured image in pictureBox1
                pictureBox1.Image?.Dispose(); // Dispose of any previous image
                pictureBox1.Image = (Bitmap)lastFrame.Clone();
                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;

                MessageBox.Show("Image captured and saved successfully!", "Image Capture", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while saving the image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Product_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (videoSource != null && videoSource.IsRunning)
            {
                videoSource.SignalToStop();
                videoSource.WaitForStop();
            }
        }
        #endregion
    }
}
