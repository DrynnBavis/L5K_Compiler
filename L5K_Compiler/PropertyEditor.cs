using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace L5K_Compiler
{
    public partial class PropertyEditor : Form
    {
        TreeNode selectedNode = Form1.currentNode;
        public PropertyEditor()
        {
            InitializeComponent();
            InitializeProperties();
        }

        public void InitializeProperties()
        {
            var properties = selectedNode.Tag as LocalCard;
            if (properties.type == null)
            {
                MessageBox.Show("Error: Please select a processor before editing properties.", "No Processor Selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Form1.confirmedAdd = false;
                this.Close();
                return;
            }

            else if (properties.type == "proc")
            {
                nameTxtBox.Enabled = true;
                slotTxtBox.Enabled = true;
                ipAddressControl1.Enabled = false;
            }
            else if (properties.type == "local")
            {
                nameTxtBox.Enabled = true;
                slotTxtBox.Enabled = true;
                ipAddressControl1.Enabled = true;
            }
            else if (properties.type == "drive")
            {
                nameTxtBox.Enabled = true;
                slotTxtBox.Enabled = false;
                ipAddressControl1.Enabled = true;
            }
            if (nameTxtBox.Enabled)
            {
                nameTxtBox.Text = properties.name;
            }
            if(slotTxtBox.Enabled)
            {
                slotTxtBox.Text = properties.slot.ToString();
            }
            if(ipAddressControl1.Enabled)
            {
                ipAddressControl1.Text = properties.ipAdress;
            }
            Form1.confirmedAdd = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1.confirmedAdd = false;
            this.Close();
        }

    private void button1_Click(object sender, EventArgs e)
        {
            var properties = selectedNode.Tag as LocalCard;
            if (nameTxtBox.Enabled)
            {
                properties.name = nameTxtBox.Text;
            }
            if (slotTxtBox.Enabled)
            {
                properties.slot = Convert.ToInt32(slotTxtBox.Text);
            }
            if (ipAddressControl1.Enabled)
            {
                properties.ipAdress = ipAddressControl1.Text;
            }
            this.Close();
        }
    }
}
