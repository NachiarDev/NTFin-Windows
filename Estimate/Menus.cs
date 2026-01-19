using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Estimate
{
    public partial class Menus : Form
    {
        public Menus()
        {
            InitializeComponent();
        }
        CustomerList objcus=new CustomerList();
        SupplierList objsup=new SupplierList(); 
        StockINList objstockin=new StockINList();
        ExpensesList objexpenses = new ExpensesList();
        EstimateList objestimate=new EstimateList();
        DeliveryChallanList objdelivery=new DeliveryChallanList();
        SalesReturnList objsales=new SalesReturnList();
        PurchaseList objpurchase=new PurchaseList();
        CountryList objcountry=new CountryList();
        StateList objstate=new StateList();
        CompanyList objcompany=new CompanyList();
        BillAddressList objbilladdress=new BillAddressList();
        UnitList objunit=new UnitList();
        ZipTypeList objzip=new ZipTypeList();
        ZipLengthList objziplength=new ZipLengthList();
        TransportTypeList objtransport=new TransportTypeList();
        CategoryList objcategory=new CategoryList();
        SubCategoryList objsubcategory=new SubCategoryList();
        CountrySettingList objcountrysettings=new CountrySettingList();
        MOdeofReturnList objmodeofreturn=new MOdeofReturnList();
        PaymentMethodList objpaymentmethod=new PaymentMethodList();
        StatusList objstatus=new StatusList();
        TaxList objtax=new TaxList();   
        AddressTypeList objaddresstype=new AddressTypeList(); 
        GSTTypeList objgstype=new GSTTypeList();
        SizeList objsize=new SizeList();
        CBEntryList objCbentry=new CBEntryList();
        private void customerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objcus.MdiParent = this;
            objcus.Show();
        }

        private void supplierToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objsup.MdiParent = this;
            objsup.Show();


        }

        private void stockINToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objstockin.MdiParent = this;
            objstockin.Show();
        }

        private void dToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objexpenses.MdiParent= this;
            objexpenses.Show();
        }

        private void estimationToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            objestimate.MdiParent = this;
            objestimate.Show();

        }
        private void invoiceToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            objestimate.MdiParent = this;
            objestimate.Show();
        }

               
        private void countryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objcountry.MdiParent = this;
            objcountry.Show();
        }

        private void stateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objstate.MdiParent = this;
            objstate.Show();
        }

        private void companyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objcompany.MdiParent= this;
            objcompany.Show();
        }

        private void billAddressToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objbilladdress.MdiParent = this;
            objbilladdress.Show();
        }

        private void unitToolStripMenuItem_Click(object sender, EventArgs e)
        {   objunit.MdiParent = this;
            objunit.Show();

        }

        private void zipTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objzip.MdiParent = this;
            objzip.Show();
        }

        private void zipLengthToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objziplength.MdiParent = this;
            objziplength.Show();
        }

        private void transportTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objtransport.MdiParent = this;
            objtransport.Show();
        }

        private void Menus_Load(object sender, EventArgs e)
        {
            AddCalculatorToMDI();
        }
        private void AddCalculatorToMDI()
        {
            Button calcButton = new Button();
            calcButton.Size = new Size(30, 30);
            calcButton.Font = new Font("Segoe UI Symbol", 14, FontStyle.Regular);
            calcButton.Text = "🧮"; // Unicode icon

            // Make button look flat and transparent
            calcButton.FlatStyle = FlatStyle.Flat;
            calcButton.FlatAppearance.BorderSize = 0;
            calcButton.FlatAppearance.MouseOverBackColor = Color.Transparent;
            calcButton.FlatAppearance.MouseDownBackColor = Color.Transparent;
            calcButton.BackColor = Color.Transparent;
            calcButton.UseVisualStyleBackColor = true; // Important for transparency
            calcButton.Cursor = Cursors.Hand;

            // Position at top-right
            calcButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            calcButton.Location = new Point(this.ClientSize.Width - 40, menuStrip1.Height + 4);
            calcButton.BringToFront();

            calcButton.Click += (s, e) =>
            {
                try { System.Diagnostics.Process.Start("calc"); }
                catch (Exception ex) { MessageBox.Show("Cannot open Calculator: " + ex.Message); }
            };

            this.Controls.Add(calcButton);

            // Reposition on resize
            this.Resize += (s, e) =>
            {
                calcButton.Location = new Point(this.ClientSize.Width - 40, menuStrip1.Height + 4);
            };
        }

        private void categoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objcategory.MdiParent = this;
            objcategory.Show();

        }

        private void subCategoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objsubcategory.MdiParent = this;
            objsubcategory.Show();
        }

        private void countrySettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {objcountrysettings.MdiParent = this;
            objcountrysettings.Show();

        }

        private void modeOfReturnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objmodeofreturn.MdiParent = this;
            objmodeofreturn.Show();
        }

        private void paymentMethodToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objpaymentmethod.MdiParent = this;  
            objpaymentmethod.Show();
        }

        private void statusToolStripMenuItem_Click(object sender, EventArgs e)
        {   
            objstatus.MdiParent = this;
            objstatus.Show();

        }

        private void taxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objtax.MdiParent = this;
            objtax.Show();

        }

        private void addressTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {objaddresstype.MdiParent = this;
            objaddresstype.Show();

        }

        private void gSTTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objgstype.MdiParent = this;
            objgstype.Show();
        }

        private void sizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objsize.MdiParent = this;
            objsize.Show();
        }

        private void stcokOutToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void purchaseReturnToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            objpurchase.MdiParent = this;
            objpurchase.Show();
        }

        private void salesReturnToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            objsales.MdiParent = this;
            objsales.Show();
        }

        private void deliveryToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            objdelivery.MdiParent = this;
            objdelivery.Show();

        }

        private void cashBankEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            objCbentry.MdiParent = this;
            objCbentry.Show();
        }
    }
}
