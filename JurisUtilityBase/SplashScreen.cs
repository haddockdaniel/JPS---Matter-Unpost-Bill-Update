using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JurisUtilityBase
{
    public partial class SplashScreen : Form
    {
        public string Status
        {
            get { return labelStatus.Text; }
            set
            {
                labelStatus.Text = value; 
                Refresh();
            }
        }

        public SplashScreen()
        {
            InitializeComponent();
        }

        private void SplashScreen_Load(object sender, EventArgs e)
        {
            this.labelVersion.Text = @"Version " + Application.ProductVersion;
            this.labelCopyright.Text = @"Copyright © 1996-" + DateTime.Now.Year;
            this.labelAppName.Text = "Matter Unpost Bill Update";
            this.labelCompany.Text = Application.CompanyName;
            this.Refresh();
        }
    }
}
