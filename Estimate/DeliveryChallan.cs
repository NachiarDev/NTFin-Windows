using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Estimate
{
    public partial class DeliveryChallan : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
        DataTable productTable;
        DataSet ds;

        public DeliveryChallan()
        {
            InitializeComponent();
        }

        private void DeliveryChallan_Load(object sender, EventArgs e)
        {
            LoadCompanyDetails();
            LoadCustomer();
            LoadtransportType();
            LoadProductsIntoGridComboBox();
            textBox2.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit;
        }
        #region Combobox
        private void LoadCustomer()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, CustomerName FROM Customer ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "CustomerName";  // Show Organisation Name
                        comboBox1.ValueMember = "ID";      // Store Organisation ID
                        comboBox1.SelectedIndex = -1;      // No pre-selection
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Customer: " + ex.Message);
            }
        }

        private void LoadtransportType()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
                {
                    conn.Open();
                    string query = "SELECT ID, TransportName FROM TransportType";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        comboBox2.DataSource = dt;
                        comboBox2.DisplayMember = "TransportName";
                        comboBox2.ValueMember = "ID";
                        comboBox2.SelectedIndex = -1;  // No pre-selection
                    }
                }

                // 👉 After loading data, NOW attach the event
                comboBox2.SelectedIndexChanged += comboBox2_SelectedIndexChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Transport Type: " + ex.Message);
            }
        }


        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex >= 0 && comboBox2.SelectedValue != null && int.TryParse(comboBox2.SelectedValue.ToString(), out int selectedId))
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Encrypt=True;Password=N@tsys@2024;TrustServerCertificate=True"))
                    {
                        conn.Open();
                        string query = "SELECT VehicleNo FROM TransportType WHERE ID = @ID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", selectedId);

                            object result = cmd.ExecuteScalar();
                            if (result != null)
                            {
                                textBox14.Text = result.ToString();
                            }
                            else
                            {
                                textBox14.Text = "";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error fetching Vehicle No: " + ex.Message);
                }
            }
            else
            {
                textBox14.Text = "";
            }
        }

        #endregion

        #region CompanyLoad
        private void LoadCompanyDetails()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True"))
                {
                    conn.Open();
                    string query = "SELECT Name, Address, Contact, GSTNumber FROM Company WHERE ID = 1";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            textBox3.Text = reader["Name"].ToString();
                            richTextBox1.Text = reader["Address"].ToString();
                            textBox8.Text = reader["Contact"].ToString();
                            textBox6.Text = reader["GSTNumber"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Company Details: " + ex.Message);
            }
        }
        #endregion

        #region Insert

        private string FormatPhoneNumber(string number)
        {
            return number.Length >= 10
                ? $"({number.Substring(0, 3)}) {number.Substring(3, 3)}-{number.Substring(6, 4)}"
                : number;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            InsertData();
        }

        private async void InsertData()
        {
            try
            {
                if (!ValidateFields())
                    return;

               
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();

                    string formattedrecevContact = FormatPhoneNumber(maskedTextBox1.Text);
                    string formatteddrivrContact = FormatPhoneNumber(maskedTextBox2.Text);

                    string insertDeliveryChallanQuery = @"
            INSERT INTO DeliveryChallan (SName,SContact,SAddress,SGSTIN, RName, RContact, RAddress, RGSTIN,ChallanNo, ChallanDate, EWayNo,TransportName,TransportID,VehicleNumber,GTSNo,Terms_Condition,ChalNo,DriverName,driverContact,Transaporttype,Customer,BnchID,LocID) 
                                VALUES (@SName,@SContact,@SAddress,@SGSTIN, @RName, @RContact, @RAddress, @RGSTIN, @ChallanNo, @ChallanDate, @EWayNo, @TransportName,@TransportID,@VehicleNumber,@GTSNo,@Terms_Condition,1,@DriverName,@driverContact,@Transporttype,@Customer,@BnchID,@LocID);
            SELECT SCOPE_IDENTITY();";

                    int refNo;
                    
                    using (SqlCommand cmd = new SqlCommand(insertDeliveryChallanQuery, conn))
                    {
                        
                        cmd.Parameters.AddWithValue("@SName", textBox3.Text);
                        cmd.Parameters.AddWithValue("@SContact", textBox8.Text);
                        cmd.Parameters.AddWithValue("@SAddress", richTextBox1.Text);
                        cmd.Parameters.AddWithValue("@SGSTIN", textBox6.Text);
                        cmd.Parameters.AddWithValue("@RName",textBox7.Text);
                        cmd.Parameters.AddWithValue("@RContact", formattedrecevContact);
                        cmd.Parameters.AddWithValue("@RAddress", richTextBox2.Text);
                        cmd.Parameters.AddWithValue("@RGSTIN", textBox10.Text);
                        cmd.Parameters.AddWithValue("@ChallanNo", textBox1.Text);
                        cmd.Parameters.AddWithValue("@ChallanDate", textBox2.Text);
                        cmd.Parameters.AddWithValue("@EWayNo", textBox4.Text);
                        cmd.Parameters.AddWithValue("@TransportName", textBox12.Text);
                        cmd.Parameters.AddWithValue("@TransportID", textBox11.Text);
                        cmd.Parameters.AddWithValue("@VehicleNumber", textBox14.Text);
                        cmd.Parameters.AddWithValue("@GTSNo", textBox15.Text);
                        cmd.Parameters.AddWithValue("@Terms_Condition", richTextBox3.Text);
                        cmd.Parameters.AddWithValue("@DriverName",textBox13.Text);
                        cmd.Parameters.AddWithValue("@driverContact", formatteddrivrContact);
                        cmd.Parameters.AddWithValue("@Transporttype", comboBox2.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Customer", comboBox1.SelectedValue ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@BnchID", 1);
                        cmd.Parameters.AddWithValue("@LocID", 1);

                        object newIDObj = cmd.ExecuteScalar();
                        refNo = Convert.ToInt32(newIDObj);
                    }





                    await Task.Delay(10);



                    string updateDelivertchallanQuery = "UPDATE DeliveryChallan SET ChalNo = @ChalNo WHERE ID = @ID";
                    using (SqlCommand updateCmd = new SqlCommand(updateDelivertchallanQuery, conn))
                    {
                        updateCmd.Parameters.AddWithValue("@ChalNo", refNo);
                        updateCmd.Parameters.AddWithValue("@ID", refNo);
                        updateCmd.ExecuteNonQuery();
                    }


                    if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows.Cast<DataGridViewRow>().All(r => r.IsNewRow))
                    {
                        MessageBox.Show("No data available in the grid to insert into DeliveryChallan Details!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null)
                            continue;
                        string productName = row.Cells[0].Value?.ToString() ?? "";

                        decimal qty = Convert.ToDecimal(row.Cells[2].Value ?? 0);

                        // 1. Insert into SalesReturn_Details
                        string insertDeliverychallanQuery = @"
        INSERT INTO Challan_details (CNo,ProductName, TotalQty,Code,BnchID, LocID)                              
        VALUES (@CNo, @ProductName, @TotalQty, @Code,@BnchID, @LocID);";

                        using (SqlCommand cmd = new SqlCommand(insertDeliverychallanQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@CNo", refNo);
                            cmd.Parameters.AddWithValue("@ProductName", productName);

                            cmd.Parameters.AddWithValue("@TotalQty", qty);

                            cmd.Parameters.AddWithValue("@Code", row.Cells[0].Value?.ToString() ?? "");
                            cmd.Parameters.AddWithValue("@BnchID", 1);
                            cmd.Parameters.AddWithValue("@LocID", 1);

                            cmd.ExecuteNonQuery();
                        }

                       
                    }

                    MessageBox.Show("Inserted Successfully");
                }
                ResetFormFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

        }


        #endregion
        #region validation
        private bool ValidateFields()
        {
            var fields = new Dictionary<System.Windows.Forms.Control, string>
    {
        { textBox1, "Challan Number is required" },
        { textBox7, "Receiver Name is required" },
        { maskedTextBox1, "Receiver Contact Number is required" },
         { richTextBox1, "Receiver Address is required" },
         { textBox11, "Transport Number is required" },
         { textBox12, "Transport Name is required" },
         { textBox13, "Driver Name is required" },
         { maskedTextBox2, "Driver Contact Number is required" },
         { textBox4, "Eway Number is required" },
        
    };

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Customer Name is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show(" Transport Type is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Focus();
                return false;
            }
            return true;

        }
        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != '\'' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Ignore the key press
            }
        }

        private void textBox10_Leave(object sender, EventArgs e)
        {
            string gst = textBox10.Text.Trim();
            string gstpattern = @"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[0-9]{1}[A-Z]{1}[0-9]{1}$";
            if (!Regex.IsMatch(gst, gstpattern))
            {
                MessageBox.Show("Invaid GST Format!Please enter a valid GST Number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox10.Focus();
            }
        }
        #endregion


        #region Reset
        private void ResetFormFields()
        {

            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox textBox && textBox != textBox2&&textBox!= textBox3 && textBox != textBox8 && textBox != textBox6)
                {
                    textBox.Text = "";
                }
                if (ctrl is System.Windows.Forms.RichTextBox richtextbox && richtextbox != richTextBox1)
                {
                    richTextBox1.Text = "";
                }
                

            }


            maskedTextBox1.Text = "";
            maskedTextBox2.Text = "";
            

            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;

            


            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (i != 13)
                        {
                            row.Cells[i].Value = null;
                        }
                    }
                }
            }


            comboBox1.Focus();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ResetFormFields();
        }
        #endregion






        private void LoadProductsIntoGridComboBox()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                {
                    conn.Open();
                    string query = "SELECT ID, ProductName,Code FROM Product";
                    SqlDataAdapter da = new SqlDataAdapter(query, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dataGridView1.Columns["Column1"] is DataGridViewComboBoxColumn comboBoxColumn)
                    {
                        comboBoxColumn.DataSource = dt;
                        comboBoxColumn.DisplayMember = "ProductName"; // Show Product Name
                        comboBoxColumn.ValueMember = "ID"; // Store Product ID
                        dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                       
                    }
                    else
                    {
                        MessageBox.Show("ComboBox column 'Column1' not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                    if (e.ColumnIndex == dataGridView1.Columns["Column1"].Index)
                    {
                        if (row.Cells["Column1"].Value == null)
                            return;

                        string selectedProductName = row.Cells["Column1"].Value.ToString();

                        using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                        {
                            conn.Open();
                            string query = "SELECT Code FROM Product WHERE ID = @ID";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@ID", selectedProductName);
                                SqlDataReader reader = cmd.ExecuteReader();

                                if (reader.Read())
                                {
                                    row.Cells["Column2"].Value = reader["Code"].ToString();
                                   
                                }
                            }
                        }



                    }




                    

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                    // If user edited Column3 (Qty)
                    if (e.ColumnIndex == dataGridView1.Columns["Column3"].Index)
                    {
                        if (row.Cells["Column1"].Value == null) return;
                        if (row.Cells["Column3"].Value == null) return;

                        string selectedProductName = row.Cells["Column1"].Value.ToString();
                        int enteredQty = 0;
                        int.TryParse(row.Cells["Column3"].Value.ToString(), out enteredQty);

                        if (enteredQty <= 0) return;

                        int totalQtyFromDB = 0;

                        using (SqlConnection conn = new SqlConnection("Data Source=45.64.104.2;Initial Catalog=kalviaga_NTFin_Test;User ID=TestDBA;Password=N@tsys@2024;Encrypt=True;TrustServerCertificate=True;"))
                        {
                            conn.Open();
                            string query = "SELECT TotalQty FROM Product WHERE ID = @ID";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@ID", selectedProductName);
                                SqlDataReader reader = cmd.ExecuteReader();

                                if (reader.Read())
                                {
                                    totalQtyFromDB = Convert.ToInt32(reader["TotalQty"]);
                                }
                            }
                        }

                        if (enteredQty > totalQtyFromDB)
                        {
                            MessageBox.Show($"Entered Qty ({enteredQty}) is more than available stock ({totalQtyFromDB}).", "Qty Not Available", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            row.Cells["Column3"].Value = null; // Clear wrong input
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == dataGridView1.Columns["Column4"].Index)
            {

                if (dataGridView1.IsCurrentRowDirty)
                {
                    dataGridView1.EndEdit();


                    DialogResult result = MessageBox.Show("Are you sure you want to delete this row?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {

                        if (!dataGridView1.Rows[e.RowIndex].IsNewRow)
                        {
                            dataGridView1.Rows.RemoveAt(e.RowIndex);
                        }
                    }
                }
            }
        }
    }


}