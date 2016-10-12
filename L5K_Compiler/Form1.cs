using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Resources;
using Excel = Microsoft.Office.Interop.Excel;
using Cards = L5K_Compiler.Properties.Resources;


namespace L5K_Compiler
{
    public partial class Form1 : Form
    {
        string outputPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
        string excelPath;
        TreeNode procNode = new TreeNode("Please Select a Processor");
        // Steps to add a new module to the list:
        // 1. create a string array titled as the module's catalog number (exempt any chars other than letters) with a
        //    preceding m.
        // 2. Set it equal to @"" and with your cursor between the two quotes paste the MODULE info from the reference
        //    l5k file. DON'T forget to include the tab before the first 'MODULE'!
        // 3. Go through the new paste and add a second quote beside each existing quote in the pasted content. For
        //    example "Local" will become ""Local"".
        // 4. At the end of the paste add: .Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
        // 5. Lastly go through and replace all areas with needed changes with a '~' symbol. For example the description
        //    "	MODULE _xxxx_01M02 (Description := ""VFD Speed Control"", has a copied over description that will need
        //    to be changed for future modules. So the corrected version will look like this:
        //    "	MODULE _xxxx_01M02 (Description := ""~"",. This is the character that will be find-and-replaced with the
        //    values from the excel document.
        public List<IOModule> ioList = new List<IOModule>();
        public List<IOModule> ioListADDED = new List<IOModule>();
        int[] numMods = Enumerable.Repeat(1, 1000).ToArray();
        public ContextMenuStrip procRightClick;
        public ContextMenuStrip localRightClick;
        public ContextMenuStrip driveRightClick;
        public ContextMenuStrip ioblockRightClick;
        static public string selectedModule = null;
        static public string typeOfModuleAdded = null;
        static public bool confirmedAdd = false;
        static public bool confirmedEdit = false;
        static public bool[] localSlots = new bool[20];
        int[] subLocalCnt = new int[100];
        static public TreeNode currentNode;
        static public int slotChanged;
        static public int chassisSize;
        static public bool chassisSizeSelected = false;
        static public bool processorSelected = false;

        public Form1()
        {
            InitializeComponent();
            InitializeTreeView();
            savePathLbl.Text += outputPath;
            chassisDropSelect.Items.Add("1756-A7      7-Slot ControlLogix Chassis");
            chassisDropSelect.Items.Add("1756-A10    10-Slot ControlLogix Chassis");
            chassisDropSelect.Items.Add("1756-A13    13-Slot ControlLogix Chassis");
            chassisDropSelect.Items.Add("1756-A17    17-Slot ControlLogix Chassis");
        }

        public TreeView form1Tree { get { return treeIO; } }

        private void InitializeTreeView()
        {
            treeIO.Nodes.Add(procNode);
            treeIO.Nodes[0].Tag = new LocalCard();
            var rootTag = procNode.Tag as LocalCard;
            rootTag.slot = 0;
            treeIO.NodeMouseClick += (sender, args) => treeIO.SelectedNode = args.Node;

            ToolStripMenuItem delete = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.delete_40x };
            delete.Text = "Delete";
            delete.Click += new EventHandler(delete_Click);
            ToolStripMenuItem delete2 = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.delete_40x };
            delete2.Text = "Delete";
            delete2.Click += new EventHandler(delete_Click);
            ToolStripMenuItem properties = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.tools_40x };
            properties.Text = "Properties";
            properties.Click += new EventHandler(properties_Click);
            ToolStripMenuItem properties2 = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.tools_40x };
            properties2.Text = "Properties";
            properties2.Click += new EventHandler(properties_Click);
            ToolStripMenuItem properties3 = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.tools_40x };
            properties3.Text = "Properties";
            properties3.Click += new EventHandler(properties_Click);
            ToolStripMenuItem addLocalCard = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.add_40x };
            addLocalCard.Text = "Add Local Card";
            addLocalCard.Click += new EventHandler(addLocalCard_Click);
            ToolStripMenuItem editProc = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.highlight_40x };
            editProc.Text = "Change Processor";
            editProc.Click += new EventHandler(editProc_Click);
            ToolStripMenuItem editLocal = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.highlight_40x };
            editLocal.Text = "Change Local";
            editLocal.Click += new EventHandler(editLocal_Click);
            ToolStripMenuItem addDrive = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.add_40x };
            addDrive.Text = "Add Drive";
            addDrive.Click += new EventHandler(addDrive_Click);
            ToolStripMenuItem addIOBlock = new ToolStripMenuItem() { Image = L5K_Compiler.Properties.Resources.add_40x };
            addIOBlock.Text = "Add IO Block";
            addIOBlock.Click += new EventHandler(addIOBlock_Click);

            procRightClick = new ContextMenuStrip();
            procRightClick.Items.AddRange(new ToolStripMenuItem[] { addLocalCard, editProc, properties, delete });

            localRightClick = new ContextMenuStrip();
            localRightClick.Items.AddRange(new ToolStripMenuItem[] { addDrive, addIOBlock, editLocal, properties2, delete });

            driveRightClick = new ContextMenuStrip();
            driveRightClick.Items.AddRange(new ToolStripMenuItem[] { delete2, properties3 });
        }

        //right click option functions:
        //  Processor
        void addLocalCard_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            if (processorSelected == true && chassisSizeSelected == true)
            {
                typeOfModuleAdded = "Local Card";
                ListSelector test = new ListSelector(ioList, ioListADDED);
                test.ShowDialog();
                if (confirmedAdd)
                {
                    TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                    tn.Text = ("[?]" + selectedModule);
                    tn.Tag = new LocalCard();
                    var tag = tn.Tag as LocalCard;
                    tag.type = "local";
                    treeIO.SelectedNode.Expand();
                    simulatePropertiesClick();
                }
            }
            else
                MessageBox.Show("Error: You need a processor and chassis size selected first.", "Invalid Card Index Detected", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        void editProc_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Processor";
            ListSelector test = new ListSelector(ioList, ioListADDED);
            test.ShowDialog();
            if (confirmedAdd)
            {
                var procTag = treeIO.Nodes[0].Tag as LocalCard;
                procTag.type = "proc";
                treeIO.SelectedNode.Text = "[0]" + selectedModule;
                processorSelected = true;
                simulatePropertiesClick();
            }
        }
        // Local
        void editLocal_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Local Card";
            ListSelector test = new ListSelector(ioList, ioListADDED);
            test.ShowDialog();
            if (confirmedAdd)
            {
                treeIO.SelectedNode.Text = "[?]" + selectedModule;
                simulatePropertiesClick();
            }
        }

        void addDrive_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Drive";
            ListSelector test = new ListSelector(ioList, ioListADDED);
            test.ShowDialog();
            if (confirmedAdd)
            {
                TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                tn.Tag = new LocalCard();
                var tag = tn.Tag as LocalCard;
                tag.type = "drive";
                treeIO.SelectedNode.Expand();
            }
        }

        void addIOBlock_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "IOBlock";
            ListSelector test = new ListSelector(ioList, ioListADDED);
            try
            {
                test.ShowDialog();
            }
            catch { confirmedAdd = false; }
            if (confirmedAdd)
            {
                TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                tn.Tag = new LocalCard();
                var tag = tn.Tag as LocalCard;
                tag.type = "ioBlock";
                treeIO.SelectedNode.Expand();
            }
        }

        void delete_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            var tag = treeIO.SelectedNode.Tag as LocalCard;
            if (tag.type == "IOBlock") //updates lists of used/unused io modules
            {
                IOModule cardToBeSwapped = null;
                foreach (IOModule card in ioListADDED)
                {
                    if (treeIO.SelectedNode.Text == card.name)
                        cardToBeSwapped = card;
                }
                ioListADDED.Remove(cardToBeSwapped);
                ioList.Add(cardToBeSwapped);
            }
            treeIO.Nodes.Remove(treeIO.SelectedNode);
        }

        void properties_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            currentNode = treeIO.SelectedNode;
            var properties = currentNode.Tag as LocalCard;
            PropertyEditor editor = new PropertyEditor();
            if (confirmedAdd)
            {
                editor.ShowDialog();
                if (confirmedEdit)
                {
                    if (currentNode.Level == 2)
                        currentNode.Text = selectedModule + " " + properties.name;
                    else
                        currentNode.Text = "[" + properties.slot + "]" + selectedModule + " " + properties.name;
                }
            }
        }

        void simulatePropertiesClick()
        {
            currentNode = treeIO.SelectedNode;
            var properties = currentNode.Tag as LocalCard;
            PropertyEditor editor = new PropertyEditor();
            if (confirmedAdd)
            {
                editor.ShowDialog();
                if (confirmedEdit)
                {
                    if (currentNode.Level == 2)
                        currentNode.Text = selectedModule + " " + properties.name;
                    else
                        currentNode.Text = "[" + properties.slot + "]" + selectedModule + " " + properties.name;
                }
            }
        }

        void treeIO_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                treeIO.SelectedNode = e.Node;
            }

            if (e.Node.Level == 0)
            {
                e.Node.ContextMenuStrip = procRightClick;
            }
            else if (e.Node.Level == 1)
            {
                e.Node.ContextMenuStrip = localRightClick;
            }
            else if (e.Node.Level == 2)
            {
                e.Node.ContextMenuStrip = driveRightClick;
            }
        }

        public void PrintNodesRecursive(TreeNode oParentNode)
        {
            MessageBox.Show(oParentNode.Text);

            // Start recursion on all subnodes.
            foreach (TreeNode oSubNode in oParentNode.Nodes)
            {
                PrintNodesRecursive(oSubNode);
            }
        }

        private void ChoosePath()
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            folderBrowser.Reset();
            folderBrowser.Description = "Please select a path for the L5K file to be saved";
            folderBrowser.ShowNewFolderButton = false;
            folderBrowser.RootFolder = Environment.SpecialFolder.MyComputer;
            folderBrowser.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\";
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            outputPath = folderBrowser.SelectedPath + "\\";
            savePathLbl.Text = "Current Save Path: " + outputPath;
        }


        private void extractExcelData()
        {
            if (string.IsNullOrWhiteSpace(panelNameBox.Text) || string.IsNullOrWhiteSpace(plcModuleBox.Text))
            {
                MessageBox.Show("Please ensure all boxes are filled properly\n and that the IE version is valid.");
                return;
            }
            int panelNameColumn = int.Parse(panelNameBox.Text);
            int plcModuleColumn = int.Parse(plcModuleBox.Text);
            OpenFileDialog folderBrowser = new OpenFileDialog();
            folderBrowser.Reset();
            folderBrowser.Filter = "Excel Files (.xlsx)|*.xlsx";
            folderBrowser.FilterIndex = 1;
            folderBrowser.Multiselect = false;
            bool? userClickedOK = folderBrowser.ShowDialog() == DialogResult.OK;
            if (userClickedOK == true)
            {
                Cursor.Current = Cursors.WaitCursor;
                changePathBtn.Enabled = false;
                compileBtn.Enabled = false;
                importExcelBtn.Enabled = false;
                excelPath = folderBrowser.FileName;
                Excel.Application app = new Excel.Application();
                if (app == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }
                app.DisplayAlerts = false;
                SplashScreen.ShowSplashScreen();
                SplashScreen.SetStatus("Loading Excel Data. Please wait...");
                Excel.Workbook wb = app.Workbooks.Open(excelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];
                int numRows = ws.UsedRange.Rows.Count;
                int i = 1;
                int countAENTR = 0;
                while (i <= numRows)
                {
                    string cardName = (string)(ws.Cells[i, plcModuleColumn] as Excel.Range).Value;
                    if (cardName == null)
                    {
                        i++;
                        continue;
                    }
                    int cardCount = 0;
                    while (cardName != "1734-AENTR" && i <= (numRows))//starts looking through values under the AENTR
                    {
                        cardName = (string)(ws.Cells[i, plcModuleColumn] as Excel.Range).Value;
                        SplashScreen.SetProgress((int)(i * 100.00 / numRows));
                        if (cardName == null)
                        {
                            i++;
                            continue;
                        }
                        else if (cardName == "1734-IB8S")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IB8S", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-OB8S")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-OB8S", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IB4D")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IB4D", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount });
                            cardCount++;
                            numMods[countAENTR]++;
                        }

                        else if (cardName == "1734-OB4E")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-OB4E", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IE2C")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IE2C", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-OE2C")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-OE2C", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IR2")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IR2", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        i++;
                    }
                    if (i <= numRows && cardName.Contains("AENTR"))
                    {
                        int x = i;
                        while ((string)(ws.Cells[x, 1] as Excel.Range).Value == null)
                        {
                            x++;
                        }
                        ioList.Add(new IOModule { name = (string)(ws.Cells[x, panelNameColumn] as Excel.Range).Value });
                        countAENTR++;
                        i++;
                    }
                }
                wb.Close();
                app.Quit();
                SplashScreen.CloseForm();
                changePathBtn.Enabled = true;
                compileBtn.Enabled = true;
                importExcelBtn.Enabled = true;
            }
        }
        /*
        private void CompileL5K()
        {
            int modSlotCount = 0;
            int aentrCount = 0;
            int etrCount = 0;
            if (!moduleList.Any())
            {
                MessageBox.Show("Error no data had been loaded yet! Please import a properly formatted excel document and try again!", "Error Empty Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(panelNameBox.Text) || string.IsNullOrWhiteSpace(plcModuleBox.Text) || string.IsNullOrWhiteSpace(chassisDropSelect.Text))
            {
                MessageBox.Show("Please ensure all boxes are filled properly\n and that the IE version is valid.");
                return;
            }
            string fileName = Microsoft.VisualBasic.Interaction.InputBox("Please enter a name for the L5K file:",
                "Enter File Name", "NewFile");
            string finalOutput = Cards.header.Replace("@IEVER@", "2.15");
            if (chassisDropSelect.Text.ToString().Contains("L71S"))
            {
                finalOutput += Cards.m1756L71S;
            }

            finalOutput += Environment.NewLine;
            finalOutput += Environment.NewLine;
            SplashScreen.ShowSplashScreen();
            SplashScreen.SetStatus("Compiling your file. Please wait...");
            int numCards = moduleList.Count;
            while (moduleList.Any())
            {
                SplashScreen.SetProgress((int)((numCards - moduleList.Count) * 100.00 / numCards));
                if (moduleList[0].name.Contains("1756"))
                {
                    etrCount++;
                    string newCard = Cards.m1756EN2T.Replace("@SLOT@", etrCount.ToString());
                    newCard = newCard.Replace("@ETHERNUM@", etrCount.ToString());
                    finalOutput = finalOutput + newCard;
                    finalOutput += Environment.NewLine;
                    finalOutput += Environment.NewLine;
                    moduleList.RemoveAt(0);
                    modSlotCount = 0;
                }
                else if (moduleList[0].name.Contains("AENTR"))
                {
                    aentrCount++;
                    modSlotCount = 0;
                    string newCard = Cards.m1734AENTR.Replace("@SLOT@", modSlotCount.ToString());
                    newCard = newCard.Replace("@SIZE@", numMods[aentrCount].ToString());
                    newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                    newCard = newCard.Replace("@ETHERNUM@", etrCount.ToString());
                    finalOutput = finalOutput + newCard;
                    finalOutput += Environment.NewLine;
                    finalOutput += Environment.NewLine;
                    moduleList.RemoveAt(0);
                    modSlotCount++;
                }
                while (moduleList.Any() && !moduleList[0].name.Contains("AENTR") && !moduleList[0].name.Contains("1756"))
                {
                    if (moduleList[0].name == "1734-IB8S")
                    {
                        string newCard = Cards.m1734IB8S.Replace("@SLOT@", modSlotCount.ToString());
                        newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        moduleList.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (moduleList[0].name == "1734-OB8S")
                    {
                        string newCard = Cards.m1734OB8S.Replace("@SLOT@", modSlotCount.ToString());
                        newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        moduleList.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (moduleList[0].name == "1734-IB4D")
                    {
                        string newCard = Cards.m1734IB4D.Replace("@SLOT@", modSlotCount.ToString());
                        newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        moduleList.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (moduleList[0].name == "1734-OB4E")
                    {
                        string newCard = Cards.m1734OB4E.Replace("@SLOT@", modSlotCount.ToString());
                        newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        moduleList.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (moduleList[0].name == "1734-IE2C")
                    {
                        string newCard = Cards.m1734IE2C.Replace("@SLOT@", modSlotCount.ToString());
                        newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        moduleList.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (moduleList[0].name == "1734-OE2C")
                    {
                        string newCard = Cards.m1734OE2C.Replace("@SLOT@", modSlotCount.ToString());
                        newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        moduleList.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (moduleList[0].name == "1734-IR2")
                    {
                        string newCard = Cards.m1734IR2.Replace("@SLOT@", modSlotCount.ToString());
                        newCard = newCard.Replace("@AENTRNUM@", aentrCount.ToString());
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        moduleList.RemoveAt(0);
                        modSlotCount++;
                    }
                }
            }
            finalOutput += Cards.tail;
            File.WriteAllText(outputPath + fileName + ".l5k", finalOutput);
            SplashScreen.CloseForm();
        }*/

        static Form1 frm1 = new Form1();
        static public void SetModule(string newModule)
        {
            selectedModule = newModule;
        }

        private void changePathBtn_Click(object sender, EventArgs e)
        {
            ChoosePath();
        }

        private void compileBtn_Click(object sender, EventArgs e)
        {
            //CompileL5K();
        }

        private void importExcelBtn_Click(object sender, EventArgs e)
        {
            extractExcelData();
        }
        private void numberOnly_KeyPress(object sender, KeyPressEventArgs e) //event only allows numbers
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void commitTreeBtn_Click(object sender, EventArgs e)
        {
            PrintNodesRecursive(treeIO.Nodes[0]);
        }

        private void ComboBox1_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            chassisSizeSelected = true;
            int? highestSlot = 0;
            foreach (TreeNode node in treeIO.Nodes[0].Nodes)
            {
                var properties = node.Tag as LocalCard;
                if (properties.slot != null && properties.slot > highestSlot)
                    highestSlot = properties.slot;
            }
            ComboBox chassisSelect = (ComboBox)sender;
            if (chassisSelect.Text.Contains("A7"))
            {
                if (highestSlot < 7)
                    chassisSize = 7;
                else
                    MessageBox.Show("Error: You have cards in slots that are outside of the requested backplane's bounds.", "Invalid Card Index Detected", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (chassisSelect.Text.Contains("A10"))
            {
                if (highestSlot < 10)
                    chassisSize = 10;
                else
                    MessageBox.Show("Error: You have cards in slots that are outside of the requested backplane's bounds.", "Invalid Card Index Detected", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (chassisSelect.Text.Contains("A13"))
            {
                if (highestSlot < 13)
                    chassisSize = 13;
                else
                    MessageBox.Show("Error: You have cards in slots that are outside of the requested backplane's bounds.", "Invalid Card Index Detected", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (chassisSelect.Text.Contains("A17"))
            {
                if (highestSlot < 17)
                    chassisSize = 17;
                else
                    MessageBox.Show("Error: You have cards in slots that are outside of the requested backplane's bounds.", "Invalid Card Index Detected", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                MessageBox.Show("something impossible has happened");
        }
    }
    public class Module
    {
        public string type = null;
        public string name = null;
        public int? slot = null;
        public string[] chdesc = new string[8];
        public string[] tag = new string[8];
        public string[] address = new string[8];
    }

    public class IOModule
    {
        public List<Module> moduleList = new List<Module>();
        public string name = null;
    }

    public class LocalCard
    {
        public string type = null;
        public string name = null;
        public int? slot = null;
        public string ipAdress = null;
    }
}