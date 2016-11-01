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
        public List<LocalCard> ioList = new List<LocalCard>();
        public List<LocalCard> ioListCOPY = new List<LocalCard>();
        public List<LocalCard> ioListADDED = new List<LocalCard>();
        public List<LocalCard> extractedCards = new List<LocalCard>();
        int[] numMods = Enumerable.Repeat(1, 1000).ToArray();
        public ContextMenuStrip procRightClick;
        public ContextMenuStrip localRightClick;
        public ContextMenuStrip driveRightClick;
        public ContextMenuStrip ioblockRightClick;
        static public string selectedModule = null;
        public List<string> ioToAdd = new List<string>();
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
        public string IOModuleUsed = null;
        List<string> missingCards = new List<string>();
        string cardsNotAddedString;

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
                ListSelector test = new ListSelector(ioList, ioListADDED, ioToAdd);
                test.ShowDialog();
                if (confirmedAdd)
                {
                    TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                    tn.Text = ("[?]" + selectedModule);
                    tn.Tag = new LocalCard();
                    var tag = tn.Tag as LocalCard;
                    tag.type = "local";
                    tag.code = selectedModule;
                    treeIO.SelectedNode.Expand();
                    treeIO.SelectedNode = tn;
                    treeIO.Focus();
                }
            }
            else
                MessageBox.Show("Error: You need a processor and chassis size selected first.", "Invalid Card Index Detected", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        void editProc_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Processor";
            ListSelector test = new ListSelector(ioList, ioListADDED, ioToAdd);
            test.ShowDialog();
            if (confirmedAdd)
            {
                var procTag = treeIO.Nodes[0].Tag as LocalCard;
                procTag.type = "proc";
                procTag.code = selectedModule;
                treeIO.SelectedNode.Text = "[0]" + selectedModule;
                processorSelected = true;
            }
        }
        // Local
        void editLocal_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            TreeNode tn = treeIO.SelectedNode;
            if (tn.Tag != null)
            {
                var clear = tn.Tag as LocalCard;
                if (clear.slot != null)
                    localSlots[Convert.ToInt32(clear.slot)] = false;
            }
            tn.Tag = new LocalCard();
            var tag = tn.Tag as LocalCard;
            typeOfModuleAdded = "Local Card";
            ListSelector test = new ListSelector(ioList, ioListADDED, ioToAdd);
            test.ShowDialog();
            if (confirmedAdd)
            {
                treeIO.SelectedNode.Text = "[?]" + selectedModule;
                tag.type = "local";
                tag.code = selectedModule;
                foreach (TreeNode child in tn.Nodes)
                {
                    var childTag = child.Tag as LocalCard;
                    childTag.parent = tag.name;
                }
            }
        }

        void addDrive_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Drive";
            ListSelector test = new ListSelector(ioList, ioListADDED, ioToAdd);
            test.ShowDialog();
            if (confirmedAdd)
            {
                TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                var parent = treeIO.SelectedNode.Tag as LocalCard;
                tn.Tag = new LocalCard();
                var tag = tn.Tag as LocalCard;
                tag.type = "drive";
                tag.parent = parent.name;
                tag.code = selectedModule;
                treeIO.SelectedNode = tn;
                treeIO.Focus();
                treeIO.SelectedNode.Expand();
            }
        }

        void addIOBlock_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            addIOBlock();
        }

        public void addIOBlock()
        {
            typeOfModuleAdded = "IOBlock";
            ListSelector test = new ListSelector(ioList, ioListADDED, ioToAdd);
            try
            {
                test.ShowDialog();
            }
            catch { confirmedAdd = false; }
            if (confirmedAdd)
            {
                TreeNode parentNode = treeIO.SelectedNode;
                foreach (string ioBlock in ioToAdd)
                {
                    TreeNode tn = parentNode.Nodes.Add(ioBlock);
                    treeIO.ShowNodeToolTips = true;
                    var parent = treeIO.SelectedNode.Tag as LocalCard;
                    tn.Tag = new LocalCard();
                    var tag = tn.Tag as LocalCard;
                    tag.type = "ioBlock";
                    tag.name = ioBlock;
                    tag.code = IOModuleUsed;
                    tag.parent = parent.name;
                    foreach (LocalCard ioMod in ioListADDED)
                    {
                        if (tag.name == ioMod.name)
                            tag.moduleList = ioMod.moduleList;
                    }
                    foreach (Module ioCard in tag.moduleList)
                    {
                        if (tn.ToolTipText == "")
                            tn.ToolTipText += ioCard.code;
                        else
                            tn.ToolTipText += "\n" + ioCard.code;
                    }
                    treeIO.SelectedNode = tn;
                    treeIO.Focus();
                    tn.Text = "1734-AENTR " + tag.name;
                   treeIO.SelectedNode.Expand();
                }
                ioToAdd.Clear();
            }
        }

        void delete_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            simulateDelete(treeIO.SelectedNode);
        }

        void simulateDelete(TreeNode nodeToDelete)
        {
            var tag = nodeToDelete.Tag as LocalCard;
            if (tag.type == "ioBlock") //updates lists of used/unused io modules
            {
                LocalCard cardToBeSwapped = null;
                foreach (LocalCard card in ioListADDED)
                {
                    if (tag.name == card.name)
                        cardToBeSwapped = card;
                }
                ioListADDED.Remove(cardToBeSwapped);
                ioList.Add(cardToBeSwapped);
            }
            else if (tag.type == "local")
            {
                while (nodeToDelete.Nodes.Count != 0)
                {
                    simulateDelete(nodeToDelete.Nodes[0]);
                }
                localSlots[Convert.ToInt32(tag.slot)] = false;
            }
            treeIO.Nodes.Remove(nodeToDelete);
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
                        currentNode.Text = properties.code + " " + properties.name;
                    else
                        currentNode.Text = "[" + properties.slot + "]" + properties.code + " " + properties.name;
                }
            }
            foreach (TreeNode child in treeIO.SelectedNode.Nodes)
            {
                var childTag = child.Tag as LocalCard;
                childTag.parent = properties.name;
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

        public void ExtractNodesRecursive(TreeNode oParentNode)
        {
            var properties = oParentNode.Tag as LocalCard;
            extractedCards.Add(properties);

            // Start recursion on all subnodes.
            foreach (TreeNode oSubNode in oParentNode.Nodes)
            {
                ExtractNodesRecursive(oSubNode);
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
                string cardName = null;
                while (i <= numRows)
                {
                    cardName = (string)(ws.Cells[i, plcModuleColumn] as Excel.Range).Value;
                    if (cardName == null)
                    {
                        i++;
                        continue;
                    }
                    if(IOModuleUsed == null && !cardName.Contains("1734"))
                    {
                        i++;
                        continue;
                    }
                    if (IOModuleUsed == null)
                        IOModuleUsed = cardName;
                    int cardCount = 0;
                    while (cardName != IOModuleUsed && i <= (numRows))//starts looking through values under the AENTR
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
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IB8S", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-OB8S")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-OB8S", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IB4D")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IB4D", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }

                        else if (cardName == "1734-OB4E")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-OB4E", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IE2C")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IE2C", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-OE2C")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-OE2C", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IR2")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-IR2", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-OB8E")
                        {
                            ioList[countAENTR - 1].moduleList.Add(new Module { type = "1734-OB8E", name = ws.Cells[i, plcModuleColumn + 3].value, slot = cardCount, code = cardName });
                            cardCount++;
                            numMods[countAENTR]++;
                        }
                        else if (cardName.Contains("1734") && cardName != IOModuleUsed && cardName != "1734-CTM" && cardName != "1734-EP24DC" && cardName != "1734-FPD")
                        {
                            missingCards.Add("IO Module: " + ws.Cells[i, panelNameColumn].value + " Card: " + cardName + " Address: " + ws.Cells[i, plcModuleColumn + 4].value + "\n");
                        }
                        i++;
                    }
                    if (i <= numRows && cardName == IOModuleUsed)
                    {
                        int x = i;
                        while ((string)(ws.Cells[x, 1] as Excel.Range).Value == null)
                        {
                            x++;
                        }
                        ioList.Add(new LocalCard { name = (string)(ws.Cells[x, panelNameColumn] as Excel.Range).Value, code = IOModuleUsed });
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
                foreach (LocalCard item in ioList) //manual deep copy because c# doesn't have an easy way of doing this
                {
                    ioListCOPY.Add(item);
                }
            }
        }

        private void CompileL5K()
        {
            bool allLocalsNamed = true;
            foreach (TreeNode localNode in treeIO.Nodes[0].Nodes)
            {
                var tag = localNode.Tag as LocalCard;
                if (tag.name == null)
                {
                    allLocalsNamed = false;
                    break;
                }
            }
            if(!allLocalsNamed)
            {
                MessageBox.Show("Error: Local Card(s) found to be nameless. Please make sure all Local Cards are named and try again.", "Names Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int modSlotCount = 0;
            int aentrCount = 0;
            int etrCount = 0;
            extractedCards.Clear();
            ExtractNodesRecursive(treeIO.Nodes[0]);
            List<LocalCard> cardsForOutput = new List<LocalCard>();
            foreach (LocalCard card in extractedCards)
            {
                cardsForOutput.Add(card);
                if (card.type == "ioBlock")
                {
                    foreach (LocalCard ioBlock in ioListCOPY)
                    {
                        if (card.name == ioBlock.name)
                        {
                            int ioCardCount = 0;
                            foreach (Module ioCard in ioBlock.moduleList)
                            {
                                ioCardCount++;
                                string name = ioBlock.name + "_" + ioCardCount;
                                cardsForOutput.Add(new LocalCard { code = ioCard.code, type = ioCard.type, name = name, parent = ioBlock.name});
                            }
                        }
                    }
                }
            }
            if (!extractedCards.Any())
            {
                MessageBox.Show("Error no data had been loaded yet! Please create a tree to compile first.", "Error Empty Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(panelNameBox.Text) || string.IsNullOrWhiteSpace(plcModuleBox.Text) || string.IsNullOrWhiteSpace(chassisDropSelect.Text))
            {
                MessageBox.Show("Please ensure all boxes are filled properly\n before compiling.");
                return;
            }
            string fileName = Microsoft.VisualBasic.Interaction.InputBox("Please enter a name for the L5K file:",
                "Enter File Name", "NewFile");
            if (fileName == "")
                return;
            string finalOutput = Cards.header.Replace("@IEVER@", "2.15");
            finalOutput += Environment.NewLine;
            finalOutput += Environment.NewLine;
            try
            {
                if (cardsForOutput[0].code == "1756-L71S")
                {
                    finalOutput += Cards.m1756L71S.Replace("@SIZE@", chassisSize.ToString());
                    if (cardsForOutput[0].name != null)
                        finalOutput = finalOutput.Replace("@NAME@", cardsForOutput[0].name);
                    else
                        finalOutput = finalOutput.Replace("@NAME@", "DefaultName");
                    if (cardsForOutput[0].slot != null)
                        finalOutput = finalOutput.Replace("@SLOT@", cardsForOutput[0].slot.ToString());
                    else
                        finalOutput = finalOutput.Replace("@SLOT@", "0");
                    cardsForOutput.RemoveAt(0);
                }
                else if (cardsForOutput[0].code == "1756-L72S")
                {
                    finalOutput += Cards.m1756L72S.Replace("@SIZE@", chassisSize.ToString());
                    if (cardsForOutput[0].name != null)
                        finalOutput = finalOutput.Replace("@NAME@", cardsForOutput[0].name);
                    else
                        finalOutput = finalOutput.Replace("@NAME@", "DefaultName");
                    if (cardsForOutput[0].slot != null)
                        finalOutput = finalOutput.Replace("@SLOT@", cardsForOutput[0].slot.ToString());
                    else
                        finalOutput = finalOutput.Replace("@SLOT@", "0");
                    cardsForOutput.RemoveAt(0);
                }
            }
            catch
            {
                MessageBox.Show("Error: No recognizable processor was found in tree.", "Invalid or Missing processor", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            finalOutput += Environment.NewLine;
            finalOutput += Environment.NewLine;
            SplashScreen.ShowSplashScreen();
            SplashScreen.SetStatus("Compiling your file. Please wait...");
            int numCards = cardsForOutput.Count;
            while (cardsForOutput.Any())
            {
                SplashScreen.SetProgress((int)((numCards - cardsForOutput.Count) * 100.00 / numCards));
                if (cardsForOutput[0].code == IOModuleUsed)
                {
                    aentrCount++;
                    string replaceSLOT = "";
                    string replaceNAME = "";
                    string replaceIP = "192.168.0.1";
                    string replacedesc = "";
                    if (cardsForOutput[0].slot != null)
                        replaceSLOT = cardsForOutput[0].slot.ToString();
                    if (cardsForOutput[0].name != null)
                        replaceNAME = cardsForOutput[0].name;
                    if (cardsForOutput[0].ipAddress != null)
                        replaceIP = cardsForOutput[0].ipAddress;
                    if (cardsForOutput[0].desc != null)
                        replacedesc = cardsForOutput[0].desc;
                    string newCard = Cards.m1734AENTR.Replace("@SLOT@", replaceSLOT);
                    newCard = newCard.Replace("@SIZE@", numMods[aentrCount].ToString());
                    newCard = newCard.Replace("@NAME@", replaceNAME);
                    newCard = newCard.Replace("@IP@", replaceIP);
                    newCard = newCard.Replace("@PARENT@", cardsForOutput[0].parent);
                    newCard = newCard.Replace("@DESC@", replacedesc);
                    finalOutput = finalOutput + newCard;
                    finalOutput += Environment.NewLine;
                    finalOutput += Environment.NewLine;
                    cardsForOutput.RemoveAt(0);
                }
                else if (cardsForOutput[0].code == "PowerFlex 525-EENET")
                    {
                        string replaceSLOT = "";
                        string replaceNAME = "DefaultName";
                        string replaceIP = "192.168.0.1";
                        string replacedesc = "DefaultDesc";
                        if (cardsForOutput[0].slot != null)
                            replaceSLOT = cardsForOutput[0].slot.ToString();
                        if (cardsForOutput[0].name != null)
                            replaceNAME = cardsForOutput[0].name;
                        if (cardsForOutput[0].ipAddress != null)
                            replaceIP = cardsForOutput[0].ipAddress;
                        if (cardsForOutput[0].desc != null)
                            replacedesc = cardsForOutput[0].desc;
                        string newCard = Cards.mPowerFlex525EENET.Replace("@NAME@", replaceNAME);
                        newCard = newCard.Replace("@IP@", replaceIP);
                        newCard = newCard.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@DESC@", replacedesc);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                    }
                else if (cardsForOutput[0].code == "PowerFlex 753-ENETR")
                {
                    string replaceSLOT = "";
                    string replaceNAME = "DefaultName";
                    string replaceIP = "192.168.0.1";
                    string replacedesc = "DefaultDesc";
                    if (cardsForOutput[0].slot != null)
                        replaceSLOT = cardsForOutput[0].slot.ToString();
                    if (cardsForOutput[0].name != null)
                        replaceNAME = cardsForOutput[0].name;
                    if (cardsForOutput[0].ipAddress != null)
                        replaceIP = cardsForOutput[0].ipAddress;
                    if (cardsForOutput[0].desc != null)
                        replacedesc = cardsForOutput[0].desc;
                    string newCard = Cards.mPowerFlex753ENETR.Replace("@NAME@", replaceNAME);
                    newCard = newCard.Replace("@IP@", replaceIP);
                    newCard = newCard.Replace("@PARENT@", cardsForOutput[0].parent);
                    newCard = newCard.Replace("@DESC@", replacedesc);
                    finalOutput = finalOutput + newCard;
                    finalOutput += Environment.NewLine;
                    finalOutput += Environment.NewLine;
                    cardsForOutput.RemoveAt(0);
                }
                while (cardsForOutput.Any() && (cardsForOutput[0].code != IOModuleUsed && cardsForOutput[0].code != "PowerFlex 525-EENET" && cardsForOutput[0].code != "PowerFlex 753-ENETR"))
                {
                    if (cardsForOutput[0].code == "1756-EN2T")
                    {
                        etrCount++;
                        string replaceSLOT = "1";
                        string replaceNAME = "";
                        string replaceIP = "192.168.0.1";
                        if (cardsForOutput[0].slot != null)
                            replaceSLOT = cardsForOutput[0].slot.ToString();
                        if (cardsForOutput[0].name != null)
                            replaceNAME = cardsForOutput[0].name;
                        if (cardsForOutput[0].name != null)
                            replaceIP = cardsForOutput[0].ipAddress;
                        string newCard = Cards.m1756EN2T.Replace("@SLOT@", replaceSLOT);
                        newCard = newCard.Replace("@NAME@", replaceNAME);
                        newCard = newCard.Replace("@IP@", replaceIP);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount = 0;
                    }
                    else if (cardsForOutput[0].code == "1734-IB8S")
                    {
                        string newCard = Cards.m1734IB8S.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_", ""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (cardsForOutput[0].code == "1734-OB8S")
                    {
                        string newCard = Cards.m1734OB8S.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_", ""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (cardsForOutput[0].code == "1734-IB4D")
                    {
                        string newCard = Cards.m1734IB4D.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_", ""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (cardsForOutput[0].code == "1734-OB4E")
                    {
                        string newCard = Cards.m1734OB4E.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_", ""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (cardsForOutput[0].code == "1734-IE2C")
                    {
                        string newCard = Cards.m1734IE2C.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_", ""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (cardsForOutput[0].code == "1734-OE2C")
                    {
                        string newCard = Cards.m1734OE2C.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_", ""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (cardsForOutput[0].code == "1734-IR2")
                    {
                        string newCard = Cards.m1734IR2.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_", ""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                    else if (cardsForOutput[0].code == "1734-OB8E")
                    {
                        string newCard = Cards.m1734OB8E.Replace("@PARENT@", cardsForOutput[0].parent);
                        newCard = newCard.Replace("@SLOT@", cardsForOutput[0].name.Substring(cardsForOutput[0].name.Length - 2).Replace("_",""));
                        newCard = newCard.Replace("@NAME@", cardsForOutput[0].name);
                        finalOutput = finalOutput + newCard;
                        finalOutput += Environment.NewLine;
                        finalOutput += Environment.NewLine;
                        cardsForOutput.RemoveAt(0);
                        modSlotCount++;
                    }
                }
            }
            finalOutput += Cards.tail;
            File.WriteAllText(outputPath + fileName + ".l5k", finalOutput);
            SplashScreen.CloseForm();
            MessageBox.Show("compilation done!");
            if (missingCards.Any())
            {
                listView cardLoader = new listView(missingCards);
                cardLoader.Show();
            }
        }

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
            CompileL5K();
        }

        private void importExcelBtn_Click(object sender, EventArgs e)
        {
            extractExcelData();
        }

        private void showMissingCardsBtn_Click(object sender, EventArgs e)
        {
            listView cardLoader = new listView(missingCards);
            cardLoader.Show();
        }

        private void numberOnly_KeyPress(object sender, KeyPressEventArgs e) //event only allows numbers
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void dupeDriveBtn_Click(object sender, EventArgs e)
        {
            var selectedTag = treeIO.SelectedNode.Tag as LocalCard;
            if(selectedTag.type == "drive")
            {
                if (selectedTag.ipAddress == null || selectedTag.name == null)
                {
                    MessageBox.Show("Error either the IP Address or Name of the Drive selected is invalid. Please corret any mistakes and try again.");
                    return;
                }
                TreeNode newNode = new TreeNode();
                newNode.Text = treeIO.SelectedNode.Text;
                newNode.Tag = new LocalCard();
                var newTag = newNode.Tag as LocalCard;
                newTag.code = selectedTag.code;
                newTag.desc = selectedTag.desc;
                newTag.ipAddress = selectedTag.ipAddress;
                newTag.moduleList = selectedTag.moduleList;
                newTag.name = selectedTag.name;
                newTag.parent = selectedTag.parent;
                newTag.slot = selectedTag.slot;
                newTag.type = selectedTag.type;
                newTag.ipAddress = IncreaseIpIndex(newTag.ipAddress, 1);
                int motorIndex = Convert.ToInt32(selectedTag.name.Substring(selectedTag.name.Length - 2));
                motorIndex++;
                if (motorIndex < 10)
                    newTag.name = selectedTag.name.Substring(0, (selectedTag.name.Length - 2)) + "0" + motorIndex.ToString();
                else
                    newTag.name = selectedTag.name.Substring(0, (selectedTag.name.Length - 2)) + motorIndex.ToString();
                treeIO.SelectedNode.Parent.Nodes.Add(newNode);
                newNode.Text = newTag.code + " " + newTag.name;
                treeIO.SelectedNode = newNode;
                treeIO.Focus();
            }
            else
                MessageBox.Show("Error: Cannot duplicate a non-drive node. Please Select a drive-node to use this feature.", "Invalid Node Type", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string IncreaseIpIndex(string ipGiven, int increaseVal)
        {
            string[] splitOctets = ipGiven.ToString().Split('.');
            int fourth = Convert.ToInt32(splitOctets[3]) + increaseVal;
            if (fourth > 255)
            {
                MessageBox.Show("The increased IP index exceeded 255. Value changed to 255.");
                fourth = 255;
            }
            string ipReturn = (splitOctets[0] + "." + splitOctets[1] + "." + splitOctets[2] + "." + fourth.ToString());
            return ipReturn;
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
        public string code = null;
        public int? slot = null;
    }

    public class LocalCard
    {
        public List<Module> moduleList = new List<Module>();
        public string desc = null;
        public string type = null;
        public string name = null;
        public int? slot = null;
        public string ipAddress = null;
        public string code = null;
        public string parent = null;
    }
}