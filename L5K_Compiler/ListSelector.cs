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
        LocalCard cardAdded = null;

        public ListSelector(List<LocalCard> cardList, List<LocalCard> cardListADDED, List<string> ioNeedsAdding)
        {
            InitializeComponent();
            this.Text = windowName + " Selector";
            this._cardList = cardList;
            this._cardListADDED = cardListADDED;
            this.ioToAdd = ioNeedsAdding;
            InitializeList();
        }

        private List<LocalCard> _cardList;
        private List<LocalCard> _cardListADDED;
        private List<string> ioToAdd;

        private void InitializeList()
        {
            if (windowName == "Drive")
            {
                string[] listDrives = { "PowerFlex 753-ENETR" , "PowerFlex 525-EENET" };
                listBox1.Items.AddRange(listDrives);
                listBox1.SelectionMode = SelectionMode.One;
            }

            else if (windowName == "IOBlock")
            {
                if (!_cardList.Any())
                {
                    MessageBox.Show("No cards loaded! Please import cards from IO List.", "Error Empty Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                    listBox1.SelectionMode = SelectionMode.MultiExtended;
                }
                foreach (LocalCard card in _cardList)
                {
                    cardAdded = card;
                    listBox1.Items.Add(card.name);
                }
            }
            else if (windowName == "Local Card")
            {
                string[] listLocals = { "1756-EN2T" };
                listBox1.Items.AddRange(listLocals);
                listBox1.SelectionMode = SelectionMode.One;
            }
            else if (windowName == "Processor")
            {
                string[] listProcessors = { "1756-L71S", "1756-L72S" };
                listBox1.Items.AddRange(listProcessors);
                listBox1.SelectionMode = SelectionMode.One;
            }
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Form1.confirmedAdd = false;
            this.Close();
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            foreach (Object item in listBox1.SelectedItems)
                addCard(item);
        }

        private void addCard(Object selectedItem)
        {
            if (listBox1.SelectedItem != null)
            {
                Form1.selectedModule = selectedItem.ToString();
                if (windowName == "IOBlock")
                    ioToAdd.Add(selectedItem.ToString());
                Form1.confirmedAdd = true;
                this.Close();
            }
            if (cardAdded != null)
            {
                LocalCard removedCard = null;
                foreach (LocalCard card in _cardList)
                {
                    if (selectedItem.ToString() == card.name)
                        removedCard = card;
                }
                _cardList.Remove(removedCard);
                _cardListADDED.Add(removedCard);
                Form1.confirmedAdd = true;
            }
            else if (selectedItem == null)
            {
                MessageBox.Show("No module selected!");
            }
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = this.listBox1.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                addCard(listBox1.Items[index]);
            }
        }
    }
}
