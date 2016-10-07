﻿using System;
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
        List<Module> moduleList = new List<Module>();
        int[] numMods = Enumerable.Repeat(1, 1000).ToArray();
        public ContextMenuStrip procRightClick;
        public ContextMenuStrip localRightClick;
        public ContextMenuStrip driveRightClick;
        public ContextMenuStrip ioblockRightClick;
        static public string selectedModule = null;
        static public string typeOfModuleAdded = null;
        static public bool confirmed = false;
        int localCnt = 0;
        int[] subLocalCnt = new int[100];

        public Form1()
        {
            InitializeComponent();
            InitializeTreeView();
            savePathLbl.Text += outputPath;
            chassisDropSelect.Items.Add("1756-A7     7-Slot ControlLogix Chassis");
            chassisDropSelect.Items.Add("1756-A10    10-Slot ControlLogix Chassis");
            chassisDropSelect.Items.Add("1756-A13    13-Slot ControlLogix Chassis");
            chassisDropSelect.Items.Add("1756-A17    17-Slot ControlLogix Chassis");
        }

        private void InitializeTreeView()
        {
            treeIO.Nodes.Add(procNode);
            treeIO.NodeMouseClick += (sender, args) => treeIO.SelectedNode = args.Node;

            ToolStripMenuItem delete = new ToolStripMenuItem();
            delete.Text = "Delete";
            delete.Click += new EventHandler(delete_Click);
            ToolStripMenuItem delete2 = new ToolStripMenuItem();
            delete2.Text = "Delete";
            delete2.Click += new EventHandler(delete_Click);
            ToolStripMenuItem properties = new ToolStripMenuItem();
            properties.Text = "Properties";
            properties.Click += new EventHandler(properties_Click);
            ToolStripMenuItem properties2 = new ToolStripMenuItem();
            properties2.Text = "Properties";
            properties2.Click += new EventHandler(properties_Click);
            ToolStripMenuItem properties3 = new ToolStripMenuItem();
            properties3.Text = "Properties";
            properties3.Click += new EventHandler(properties_Click);
            ToolStripMenuItem addLocalCard = new ToolStripMenuItem();
            addLocalCard.Text = "Add Local Card";
            addLocalCard.Click += new EventHandler(addLocalCard_Click);
            ToolStripMenuItem editProc = new ToolStripMenuItem();
            editProc.Text = "Change Processor";
            editProc.Click += new EventHandler(editProc_Click);
            ToolStripMenuItem editLocal = new ToolStripMenuItem();
            editLocal.Text = "Change Local";
            editLocal.Click += new EventHandler(editLocal_Click);
            ToolStripMenuItem addDrive = new ToolStripMenuItem();
            addDrive.Text = "Add Drive";
            addDrive.Click += new EventHandler(addDrive_Click);
            ToolStripMenuItem addIOBlock = new ToolStripMenuItem();
            addIOBlock.Text = "Add IO Block";
            addIOBlock.Click += new EventHandler(addIOBlock_Click);

            procRightClick = new ContextMenuStrip();
            procRightClick.Items.AddRange(new ToolStripMenuItem[]{addLocalCard, editProc, delete, properties });
            //procNode.ContextMenuStrip = procRightClick;

            localRightClick = new ContextMenuStrip();
            localRightClick.Items.AddRange(new ToolStripMenuItem[] { editLocal, addDrive, addIOBlock, properties2, delete });
            localRightClick.Items.ToString();

            driveRightClick = new ContextMenuStrip();
            driveRightClick.Items.AddRange(new ToolStripMenuItem[] { delete2, properties3 });
        }

        //right click option functions:
        //  Processor
        void addLocalCard_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Local Card";
            ListSelector test = new ListSelector();
            test.ShowDialog();
            if (confirmed)
            {
                TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                tn.Text = ("[" + tn.Index + "]" + selectedModule);
                tn.Tag = new LocalCard();
                treeIO.SelectedNode.Expand();
            }
        }

        void editProc_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Processor";
            ListSelector test = new ListSelector();
            test.ShowDialog();
            if (confirmed)
                treeIO.SelectedNode.Text = selectedModule;
        }
        // Local
        void editLocal_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Local Card";
            ListSelector test = new ListSelector();
            test.ShowDialog();
            if (confirmed)
            {
                treeIO.SelectedNode.Text = "[" + treeIO.SelectedNode.Index + "]" + selectedModule;
                treeIO.SelectedNode.Tag = new LocalCard();
            }
        }

        void addDrive_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "Drive";
            ListSelector test = new ListSelector();
            test.ShowDialog();
            if (confirmed)
            {
                TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                tn.Tag = new LocalCard();
                treeIO.SelectedNode.Expand();
            }
        }
        
        void addIOBlock_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            typeOfModuleAdded = "IOBlock";
            ListSelector test = new ListSelector();
            test.ShowDialog();
            if (confirmed)
            {
                TreeNode tn = treeIO.SelectedNode.Nodes.Add(selectedModule);
                tn.Tag = new LocalCard();
                treeIO.SelectedNode.Expand();
            }
        }

        void delete_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
            treeIO.Nodes.Remove(treeIO.SelectedNode);
            foreach (TreeNode node in treeIO.Nodes[0].Nodes)
            {
                if (node.Index >= 9)
                    node.Text = "[" + node.Index + "]" + node.Text.ToString().Substring(4);
                else
                    node.Text = "[" + node.Index + "]" + node.Text.ToString().Substring(3);
            }
        }

        void properties_Click(object sender, EventArgs e)
        {
            ToolStripItem clickedItem = sender as ToolStripItem;
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
                    string cardName = (string)(ws.Cells[i, 2] as Excel.Range).Value;
                    while ((cardName == null || (cardName != "1734-AENTR" && cardName != "1756-EN2T")) && i <= numRows)//starts looking through values under the AENTR
                    {
                        cardName = (string)(ws.Cells[i, 2] as Excel.Range).Value;
                        SplashScreen.SetProgress((int)(i * 100.00 / numRows));
                        if (cardName == null)
                        {
                            i++;
                            continue;
                        }
                        else if (cardName == "1734-IB8S")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value, (string)(ws.Cells[i + 4, 3] as Excel.Range).Value, (string)(ws.Cells[i + 5, 3] as Excel.Range).Value, (string)(ws.Cells[i + 6, 3] as Excel.Range).Value, (string)(ws.Cells[i + 7, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value, (string)(ws.Cells[i + 4, 4] as Excel.Range).Value, (string)(ws.Cells[i + 5, 4] as Excel.Range).Value, (string)(ws.Cells[i + 6, 4] as Excel.Range).Value, (string)(ws.Cells[i + 7, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value, (string)(ws.Cells[i + 4, 5] as Excel.Range).Value, (string)(ws.Cells[i + 5, 5] as Excel.Range).Value, (string)(ws.Cells[i + 6, 5] as Excel.Range).Value, (string)(ws.Cells[i + 7, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { type = "1734-IB8S", name = "", modDesc = "8-CH Safety Rated Input Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-OB8S")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value, (string)(ws.Cells[i + 4, 3] as Excel.Range).Value, (string)(ws.Cells[i + 5, 3] as Excel.Range).Value, (string)(ws.Cells[i + 6, 3] as Excel.Range).Value, (string)(ws.Cells[i + 7, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value, (string)(ws.Cells[i + 4, 4] as Excel.Range).Value, (string)(ws.Cells[i + 5, 4] as Excel.Range).Value, (string)(ws.Cells[i + 6, 4] as Excel.Range).Value, (string)(ws.Cells[i + 7, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value, (string)(ws.Cells[i + 4, 5] as Excel.Range).Value, (string)(ws.Cells[i + 5, 5] as Excel.Range).Value, (string)(ws.Cells[i + 6, 5] as Excel.Range).Value, (string)(ws.Cells[i + 7, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { type = "1734-OB8S", name = "", modDesc = "8-CH Safety Rated Output Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IB4D")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { type = "1734-IB4D", name = "", modDesc = "4-CH Diagnostic Input Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }

                        else if (cardName == "1734-OB4E")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { type = "1734-OB4E", name = "", modDesc = "4-CH Output Module, Protected", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IE2C")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { type = "1734-IE2C", name = "", modDesc = "2-CH, Analog I Input Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-OE2C")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { type = "1734-OE2C", name = "", modDesc = "2-CH, Analog I Output Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if (cardName == "1734-IR2")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { type = "1734-IR2", name = "", modDesc = "2-CH RTD Input Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        i++;
                    }
                    if (cardName == "1756-EN2T")
                    {
                        moduleList.Add(new Module { name = "1756-EN2T", modDesc = "Expansion" });
                        i++;
                    }
                    else if(cardName.Contains("AENTR"))
                    {
                        int x = i;
                        while ((string)(ws.Cells[x, 1] as Excel.Range).Value == null)
                        {
                            x++;
                        }
                        moduleList.Add(new Module { type = "1734-AENTR", name = (string)(ws.Cells[x, 1] as Excel.Range).Value, modDesc = "Expansion" });
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
                finalOutput += Cards.m1756L71S;
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

    }
    public class Module
    {
        public string type = "";
        public string name = "";
        public string modDesc = "";
        public string[] chdesc = new string[8];
        public string[] tag = new string[8];
        public string[] address = new string[8];
    }

    public class LocalCard
    {
        public string type = null;
        public string name = null;
        public int? slot = null;
        public string ipAdress = null;
    }
}