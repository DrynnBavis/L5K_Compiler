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
        IOModule cardAdded = null;

        public ListSelector(List<IOModule> cardList, List<IOModule> cardListADDED)
        {
            InitializeComponent();
            this.Text = windowName + " Selector";
            this._cardList = cardList;
            this._cardListADDED = cardListADDED;
            InitializeList();
        }

        private List<IOModule> _cardList;
        private List<IOModule> _cardListADDED;

        private void InitializeList()
        {
            if (windowName == "Drive")
            {
                string[] listDrives = { "ACS880", "test" };
                listBox1.Items.AddRange(listDrives);
            }

            else if (windowName == "IOBlock")
            {
                if (!_cardList.Any())
                {
                    MessageBox.Show("No cards loaded! Please import cards from IO List.", "Error Empty Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }
                foreach (IOModule card in _cardList)
                {
                    cardAdded = card;
                    listBox1.Items.Add(card.name);
                }
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
            Form1.confirmedAdd = false;
            this.Close();
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                Form1.SetModule(listBox1.SelectedItem.ToString());
                Form1.confirmedAdd = true; 
                this.Close();
            }
            if (cardAdded != null)
            {
                IOModule removedCard = null;
                foreach (IOModule card in _cardList)
                {
                    if (listBox1.SelectedItem.ToString() == card.name)
                        removedCard = card;
                }
                _cardList.Remove(removedCard);
                _cardListADDED.Add(removedCard);
                Form1.confirmedAdd = true;

            }
            else if(listBox1.SelectedItem == null)
            {
                MessageBox.Show("No module selected!");
            }
        }
    }
}
