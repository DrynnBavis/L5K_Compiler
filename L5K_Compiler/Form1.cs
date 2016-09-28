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


namespace L5K_Compiler
{
    public partial class Form1 : Form
    {
        string outputPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
        string excelPath;
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

        public Form1()
        {
            InitializeComponent();
            procTypeDrop.Items.Add("1756-L71S");
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
        }

        private void CompileL5K()
        {
            if (!moduleList.Any())
            {
                MessageBox.Show("Error no data had been loaded yet! Please import a properly formatted excel document and try again!", "Error Empty Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            int x = 0;
            while (moduleList.Any())
            {
                if (moduleList[x].name == "1734-AENTR")
                {
                    while (moduleList[x].name != "1734-AENTR")
                    {
                        x++;
                    }
                    //send the count of io cards (AKA x) to chasis when you figure that out
                    for (int z = 0; z < x; z++)
                    {
                        string[] textToOutput = @"(*********************************************

  Import-Export
  Version   := RSLogix 5000 v27.00
  Owner     := Gyptech, Gyptech
  Exported  := Fri Sep 23 16:47:53 2016

  Note:  File encoded in UTF-8.  Only edit file in a program 
         which supports UTF-8 (like Notepad, not Wordpad).

**********************************************)
IE_VER := 2.18;

CONTROLLER modTest (ProcessorType := ""1756 - L71S"",
                    Major:= 27,
                    TimeSlice:= 20,
                    ShareUnusedTimeSlice:= 1,
                    RedundancyEnabled:= 0,
                    KeepTestEditsOnSwitchOver:= 0,
                    DataTablePadPercentage:= 50,
                    SecurityCode:= 0,
                    ChangesToDetect:= 16#ffff_ffff_ffff_ffff,
                    SFCExecutionControl:= ""CurrentActive"",
                    SFCRestartPosition:= ""MostRecent"",
                    SFCLastScan:= ""DontScan"",
                    SerialNumber:= 16#0000_0000,
                    MatchProjectToController:= No,
                    CanUseRPIFromProducer:= No,
                    SafetyLocked:= No,
                    SignatureRunModeProtect:= No,
                    ConfigureSafetyIOAlways:= No,
                    InhibitAutomaticFirmwareUpdate:= 0,
                    PassThroughConfiguration:= EnabledWithAppend,
                    DownloadProjectDocumentationAndExtendedProperties:= Yes,
                    ReportMinorOverflow:= 0)
	MODULE Local (Parent := ""Local"",
	              ParentModPortId:= 1,
	              CatalogNumber:= ""1756-L71S"",
	              Vendor:= 1,
	              ProductType:= 14,
	              ProductCode:= 158,
	              Major:= 27,
	              Minor:= 11,
	              PortLabel:= ""RxBACKPLANE"",
	              ChassisSize:= 10,
	              Slot:= 0,
	              Mode:= 2#0000_0000_0000_0001,
	              CompatibleModule:= 0,
	              KeyMask:= 2#0000_0000_0001_1111,
	              SafetyNetwork:= 16#0000_3fd1_0454_fec4)
	END_MODULE

    MODULE modTest: Partner(Parent:= ""Local"",
                             ParentModPortId:= 1,
                             CatalogNumber:= ""1756-L7SP"",
                             Vendor:= 1,
                             ProductType:= 14,
                             ProductCode:= 146,
                             Major:= 27,
                             Minor:= 11,
                             PortLabel:= ""RxBACKPLANE"",
                             Slot:= 1,
                             Mode:= 2#0000_0000_0000_0000,
	                        CompatibleModule:= 0,
                             KeyMask:= 2#0000_0000_0001_1111,
	                        SafetyNetwork:= 16#0000_0000_0000_0000)
	END_MODULE
 

     MODULE name(Parent:= ""Local"",
                  ParentModPortId:= 1,
                  CatalogNumber:= ""1756-EN2TR"",
                  Vendor:= 1,
                  ProductType:= 12,
                  ProductCode:= 200,
                  Major:= 10,
                  Minor:= 1,
                  PortLabel:= ""RxBACKPLANE"",
                  Slot:= 2,
                  Mode:= 2#0000_0000_0000_0000,
	             CompatibleModule:= 1,
                  KeyMask:= 2#0000_0000_0001_1111)
			ExtendedProp:= [[[___ <public><ConfigID>4325481</ConfigID></public>___]]]
	END_MODULE".Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    }
                }
                moduleList.RemoveAt(0);
            }
            /*
            string fileName = Microsoft.VisualBasic.Interaction.InputBox("Please enter a name for the L5K file:", 
                "Enter File Name", "NewFile");
            m1756EN2T[0] = m1756EN2T[0].Replace("~", "INSERT NAME OF DRIVE HERE");
            File.WriteAllLines(outputPath + fileName + ".l5k", m1756EN2T);
            */
        }

        private void extractExcelData()
        {
            OpenFileDialog folderBrowser = new OpenFileDialog();
            folderBrowser.Reset();
            folderBrowser.Filter = "Excel Files (.xlsx)|*.xlsx";
            folderBrowser.FilterIndex = 1;
            folderBrowser.Multiselect = false;
            bool? userClickedOK = folderBrowser.ShowDialog() == DialogResult.OK;
            if (userClickedOK == true)
            {
                progBar.Enabled = true;
                progBar.Visible = true;
                this.progBar.Maximum = 100;
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
                Excel.Workbook wb = app.Workbooks.Open(excelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];
                int numRows = ws.UsedRange.Rows.Count;
                int i = 1;
                int countAENTR = 0;
                while (i <= numRows)
                {
                    while (((string)(ws.Cells[i, 2] as Excel.Range).Value == null || (string)(ws.Cells[i, 2] as Excel.Range).Value != "1734-AENTR") && i <= numRows)//starts looking through values under the AENTR
                    {
                        this.progBar.Value = (int)(i * 100.00 / numRows);
                        if ((string)(ws.Cells[i, 2] as Excel.Range).Value == null)
                        {
                            i++;
                            continue;
                        }
                        else if ((string)(ws.Cells[i, 2] as Excel.Range).Value == "1734-IB8S")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value, (string)(ws.Cells[i + 4, 3] as Excel.Range).Value, (string)(ws.Cells[i + 5, 3] as Excel.Range).Value, (string)(ws.Cells[i + 6, 3] as Excel.Range).Value, (string)(ws.Cells[i + 7, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value, (string)(ws.Cells[i + 4, 4] as Excel.Range).Value, (string)(ws.Cells[i + 5, 4] as Excel.Range).Value, (string)(ws.Cells[i + 6, 4] as Excel.Range).Value, (string)(ws.Cells[i + 7, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value, (string)(ws.Cells[i + 4, 5] as Excel.Range).Value, (string)(ws.Cells[i + 5, 5] as Excel.Range).Value, (string)(ws.Cells[i + 6, 5] as Excel.Range).Value, (string)(ws.Cells[i + 7, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { name = "1734-IB8S", modDesc = "8-CH Safety Rated Input Module" , address = xAdress, chdesc = xDesc, tag = xTag});
                            numMods[countAENTR]++;
                        }
                        else if ((string)(ws.Cells[i, 2] as Excel.Range).Value == "1734-OB8S")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value, (string)(ws.Cells[i + 4, 3] as Excel.Range).Value, (string)(ws.Cells[i + 5, 3] as Excel.Range).Value, (string)(ws.Cells[i + 6, 3] as Excel.Range).Value, (string)(ws.Cells[i + 7, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value, (string)(ws.Cells[i + 4, 4] as Excel.Range).Value, (string)(ws.Cells[i + 5, 4] as Excel.Range).Value, (string)(ws.Cells[i + 6, 4] as Excel.Range).Value, (string)(ws.Cells[i + 7, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value, (string)(ws.Cells[i + 4, 5] as Excel.Range).Value, (string)(ws.Cells[i + 5, 5] as Excel.Range).Value, (string)(ws.Cells[i + 6, 5] as Excel.Range).Value, (string)(ws.Cells[i + 7, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { name = "1734-OB8S", modDesc = "8-CH Safety Rated Output Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if ((string)(ws.Cells[i, 2] as Excel.Range).Value == "1734-IB4D")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value};
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value};
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value};
                            moduleList.Add(new Module { name = "1734-IB4D", modDesc = "4-CH Diagnostic Input Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                            
                        else if ((string)(ws.Cells[i, 2] as Excel.Range).Value == "1734-OB4E")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value, (string)(ws.Cells[i + 2, 3] as Excel.Range).Value, (string)(ws.Cells[i + 3, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value, (string)(ws.Cells[i + 2, 4] as Excel.Range).Value, (string)(ws.Cells[i + 3, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value, (string)(ws.Cells[i + 2, 5] as Excel.Range).Value, (string)(ws.Cells[i + 3, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { name = "1734-OB4E", modDesc = "4-CH Output Module, Protected", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if ((string)(ws.Cells[i, 2] as Excel.Range).Value == "1734-IE2C")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value};
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value};
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value};
                            moduleList.Add(new Module { name = "1734-IE2C", modDesc = "2-CH, Analog I Input Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if ((string)(ws.Cells[i, 2] as Excel.Range).Value == "1734-OE2C")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { name = "1734-OE2C", modDesc = "2-CH, Analog I Output Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        else if ((string)(ws.Cells[i, 2] as Excel.Range).Value == "1734-IR2")
                        {
                            string[] xAdress = { (string)(ws.Cells[i, 3] as Excel.Range).Value, (string)(ws.Cells[i + 1, 3] as Excel.Range).Value };
                            string[] xTag = { (string)(ws.Cells[i, 4] as Excel.Range).Value, (string)(ws.Cells[i + 1, 4] as Excel.Range).Value };
                            string[] xDesc = { (string)(ws.Cells[i, 5] as Excel.Range).Value, (string)(ws.Cells[i + 1, 5] as Excel.Range).Value };
                            moduleList.Add(new Module { name = "1734-IR2", modDesc = "2-CH RTD Input Module", address = xAdress, chdesc = xDesc, tag = xTag });
                            numMods[countAENTR]++;
                        }
                        //numMods[countAENTR]++;
                        i++;
                    }
                    moduleList.Add(new Module { name = "1734-AENTR", modDesc = "Expansion"});
                    countAENTR++;
                    i++;
                }
                wb.Close();
                app.Quit();
                this.Cursor = Cursors.Default;
                changePathBtn.Enabled = true;
                compileBtn.Enabled = true;
                importExcelBtn.Enabled = true;
                progBar.Enabled = false;
                progBar.Visible = false;
                string cat = L5K_Compiler.Properties.Resources.header;
            }
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
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e) //event allows numbers and one decimal place
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }
    }
    public class Module
    {
        public string name = "";
        public string modDesc = "";
        public string[] chdesc = new string[8];
        public string[] tag = new string[8];
        public string[] address = new string[8];
    }
}
