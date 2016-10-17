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
                nameLbl.Enabled = true;
                slotTxtBox.Enabled = true;
                slotLbl.Enabled = true;
                ipAddressControl1.Enabled = false;
                ipLbl.Enabled = false;
            }
            else if (properties.type == "local")
            {
                nameTxtBox.Enabled = true;
                nameLbl.Enabled = true;
                slotTxtBox.Enabled = true;
                slotLbl.Enabled = true;
                ipAddressControl1.Enabled = true;
                ipLbl.Enabled = true;
            }
            else if (properties.type == "drive" || properties.type == "ioBlock")
            {
                nameTxtBox.Enabled = true;
                nameLbl.Enabled = true;
                slotTxtBox.Enabled = false;
                slotLbl.Enabled = false;
                ipAddressControl1.Enabled = true;
                ipLbl.Enabled = true;
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
                ipAddressControl1.Text = properties.ipAddress;
            }
            Form1.confirmedAdd = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1.confirmedAdd = false;
            Form1.confirmedEdit = false;
            this.Close();
        }

    private void button1_Click(object sender, EventArgs e)
        {
            bool nameIsOkay = false;
            bool slotIsOkay = false;
            bool ipIsOkay = false;
            var properties = selectedNode.Tag as LocalCard;
            int newSlotNumber = 0;
            try
            {
                if (nameTxtBox.Enabled)
                {
                    if (!char.IsNumber(nameTxtBox.Text.ToString().FirstOrDefault()) && char.IsLetter(nameTxtBox.Text.ToString().FirstOrDefault()))
                    {
                        nameIsOkay = true;
                    }
                    else
                        MessageBox.Show("Error: Names must begin with a letter or '_'.", "Invlid Entry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (slotTxtBox.Enabled)
                {
                    newSlotNumber = Convert.ToInt32(slotTxtBox.Text);
                    if (newSlotNumber == 0)
                        MessageBox.Show("Error: Slot 0 is reserved for processor.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (Form1.localSlots[newSlotNumber] == false && (newSlotNumber < Form1.chassisSize && newSlotNumber >= 1))
                    {
                        slotIsOkay = true;
                    }
                    else if (properties.slot == newSlotNumber)
                    {
                        slotIsOkay = true;
                    }
                    else
                        MessageBox.Show("Error: Slot is either currently in use, or outside of rack range.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (ipAddressControl1.Enabled)
                {
                    ipIsOkay = true;
                }
                bool procIsOkay = (nameTxtBox.Enabled && nameIsOkay && slotTxtBox.Enabled && slotIsOkay);
                bool localIsOkay = (nameTxtBox.Enabled && slotTxtBox.Enabled && ipAddressControl1.Enabled && nameIsOkay && slotIsOkay && ipIsOkay);
                bool driveIsOkay = (nameTxtBox.Enabled && ipAddressControl1.Enabled && nameIsOkay && ipIsOkay);
                if (procIsOkay || localIsOkay || driveIsOkay)
                    Form1.confirmedEdit = true;
                if (procIsOkay && properties.type == "proc")
                {
                    properties.name = nameTxtBox.Text.Replace(" ", "_");
                    this.Close();
                }
                else if (localIsOkay && properties.type == "local")
                {
                    properties.name = nameTxtBox.Text.Replace(" ", "_");
                    properties.ipAddress = ipAddressControl1.Text;
                    if (properties.slot != null)
                        Form1.localSlots[Convert.ToInt32(properties.slot)] = false;
                    Form1.localSlots[newSlotNumber] = true;
                    properties.slot = newSlotNumber;
                    Form1.slotChanged = newSlotNumber;
                    this.Close();
                }
                else if (driveIsOkay && (properties.type == "drive" || properties.type == "ioBlock"))
                {
                    properties.name = nameTxtBox.Text.Replace(" ", "_");
                    properties.ipAddress = ipAddressControl1.Text;
                    if (properties.slot != null)
                        Form1.localSlots[Convert.ToInt32(properties.slot)] = false;
                    Form1.localSlots[newSlotNumber] = true;
                    properties.slot = newSlotNumber;
                    Form1.slotChanged = newSlotNumber;
                    this.Close();
                }
            }

            catch
            {
                MessageBox.Show("Error: Missing property values detected", "Properties Incomplete", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Form1.confirmedEdit = false;
            }
        }
    }
}
