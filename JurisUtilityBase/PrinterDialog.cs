using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace JurisUtilityBase
{
    public partial class PrinterDialog : Form
    {
        public PrinterDialog()
        {
            InitializeComponent();
            var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;

            foreach (String s in printers)
            {
                listBox1.Items.Add(s);
            }
        }

        public string printerName = "";

        private void buttonBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            string port = getPortFromRegistry(listBox1.GetItemText(listBox1.SelectedItem));
            printerName = listBox1.GetItemText(listBox1.SelectedItem) + " on " + port;
            this.Close();
        }

        private string getPortFromRegistry(string printer)
        {
            string port = "";
            var devices = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices"); //Read-accessible even when using a locked-down account

            try
            {

                foreach (string name in devices.GetValueNames())
                {
                    if (name.Equals(printer))
                    {
                        port = (String)devices.GetValue(name);
                        port = port.Replace("winspool,", "");
                        break;
                    }
                }
            }
            catch
            {
                throw;
            }




            return port;
        }

    }
}
