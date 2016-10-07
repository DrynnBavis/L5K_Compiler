using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace L5K_Compiler
{
    public partial class ListSelector : Form
    {
        private string windowName = Form1.typeOfModuleAdded;

        public ListSelector()
        {
            InitializeComponent();
            InitializeList();
            this.Text = windowName;
        }

        private void InitializeList()
        {
            if (windowName == "Drive")
            {
                string[] listDrives = { "ACS880", "test" };
                listBox1.Items.AddRange(listDrives);
            }
            else if (windowName == "IOBlock")
            {
                string[] listIOBlock = { "AENTR", "test" };
                listBox1.Items.AddRange(listIOBlock);
            }
            else if (windowName == "Local Card")
            {
                string[] listLocals = { "1756-EN2T", "test" };
                listBox1.Items.AddRange(listLocals);
            }
            else if (windowName == "Processor")
            {
                string[] listProcessors = { "1756-L71S", "test" };
                listBox1.Items.AddRange(listProcessors);
            }
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Form1.confirmed = false;
            this.Close();
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                Form1.SetModule(listBox1.SelectedItem.ToString());
                Form1.confirmed = true; 
                this.Close();
            }
            else
            {
                MessageBox.Show("No module selected!");
            }
        }
    }
}
