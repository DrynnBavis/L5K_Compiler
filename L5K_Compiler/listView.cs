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
    public partial class listView : Form
    {
        public listView(List<string> data)
        {
            this._data = data;
            InitializeComponent();
            InitializeForm();
        }

        private List<string> _data;

        private void InitializeForm()
        {
            ColumnHeader header = new ColumnHeader();
            header.Text = "";
            header.Name = "col1";
            listView1.HeaderStyle = ColumnHeaderStyle.None;
            listView1.Columns.Add(header);
            listView1.CheckBoxes = true;
            listView1.AllowColumnReorder = false;
            listView1.LabelEdit = false;
            listView1.View = View.Details;
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.Head‌​erSize);
            foreach (string card in _data)
                listView1.Items.Add(card);
        }
    }
}
