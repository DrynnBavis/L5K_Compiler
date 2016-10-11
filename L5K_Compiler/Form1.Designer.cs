namespace L5K_Compiler
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.importExcelBtn = new System.Windows.Forms.Button();
            this.changePathBtn = new System.Windows.Forms.Button();
            this.compileBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.chassisDropSelect = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.panelNameBox = new System.Windows.Forms.TextBox();
            this.plcModuleBox = new System.Windows.Forms.TextBox();
            this.savePathLbl = new System.Windows.Forms.Label();
            this.treeIO = new System.Windows.Forms.TreeView();
            this.commitTreeBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // importExcelBtn
            // 
            this.importExcelBtn.Location = new System.Drawing.Point(236, 68);
            this.importExcelBtn.Margin = new System.Windows.Forms.Padding(1);
            this.importExcelBtn.Name = "importExcelBtn";
            this.importExcelBtn.Size = new System.Drawing.Size(137, 23);
            this.importExcelBtn.TabIndex = 0;
            this.importExcelBtn.Text = "Import Excel";
            this.importExcelBtn.UseVisualStyleBackColor = true;
            this.importExcelBtn.Click += new System.EventHandler(this.importExcelBtn_Click);
            // 
            // changePathBtn
            // 
            this.changePathBtn.Location = new System.Drawing.Point(13, 377);
            this.changePathBtn.Margin = new System.Windows.Forms.Padding(1);
            this.changePathBtn.Name = "changePathBtn";
            this.changePathBtn.Size = new System.Drawing.Size(117, 23);
            this.changePathBtn.TabIndex = 1;
            this.changePathBtn.Text = "Change Save Path";
            this.changePathBtn.UseVisualStyleBackColor = true;
            this.changePathBtn.Click += new System.EventHandler(this.changePathBtn_Click);
            // 
            // compileBtn
            // 
            this.compileBtn.BackColor = System.Drawing.SystemColors.Control;
            this.compileBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.compileBtn.Location = new System.Drawing.Point(13, 402);
            this.compileBtn.Margin = new System.Windows.Forms.Padding(1);
            this.compileBtn.Name = "compileBtn";
            this.compileBtn.Size = new System.Drawing.Size(494, 23);
            this.compileBtn.TabIndex = 2;
            this.compileBtn.Text = "Compile L5K";
            this.compileBtn.UseVisualStyleBackColor = false;
            this.compileBtn.Click += new System.EventHandler(this.compileBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Chassis Size:";
            // 
            // chassisDropSelect
            // 
            this.chassisDropSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.chassisDropSelect.FormattingEnabled = true;
            this.chassisDropSelect.Location = new System.Drawing.Point(12, 25);
            this.chassisDropSelect.Name = "chassisDropSelect";
            this.chassisDropSelect.Size = new System.Drawing.Size(201, 21);
            this.chassisDropSelect.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(230, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Declare Columns:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(236, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Panel Name";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(308, 27);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "PLC Module";
            // 
            // panelNameBox
            // 
            this.panelNameBox.Location = new System.Drawing.Point(236, 44);
            this.panelNameBox.MaxLength = 2;
            this.panelNameBox.Name = "panelNameBox";
            this.panelNameBox.Size = new System.Drawing.Size(62, 20);
            this.panelNameBox.TabIndex = 11;
            this.panelNameBox.Text = "1";
            this.panelNameBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numberOnly_KeyPress);
            // 
            // plcModuleBox
            // 
            this.plcModuleBox.Location = new System.Drawing.Point(308, 44);
            this.plcModuleBox.MaxLength = 2;
            this.plcModuleBox.Name = "plcModuleBox";
            this.plcModuleBox.Size = new System.Drawing.Size(62, 20);
            this.plcModuleBox.TabIndex = 12;
            this.plcModuleBox.Text = "5";
            this.plcModuleBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numberOnly_KeyPress);
            // 
            // savePathLbl
            // 
            this.savePathLbl.AutoSize = true;
            this.savePathLbl.Location = new System.Drawing.Point(134, 382);
            this.savePathLbl.Name = "savePathLbl";
            this.savePathLbl.Size = new System.Drawing.Size(100, 13);
            this.savePathLbl.TabIndex = 13;
            this.savePathLbl.Text = "Current Save Path: ";
            // 
            // treeIO
            // 
            this.treeIO.Location = new System.Drawing.Point(12, 52);
            this.treeIO.Name = "treeIO";
            this.treeIO.Size = new System.Drawing.Size(201, 320);
            this.treeIO.TabIndex = 14;
            this.treeIO.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeIO_NodeMouseClick);
            // 
            // commitTreeBtn
            // 
            this.commitTreeBtn.Location = new System.Drawing.Point(220, 348);
            this.commitTreeBtn.Name = "commitTreeBtn";
            this.commitTreeBtn.Size = new System.Drawing.Size(75, 23);
            this.commitTreeBtn.TabIndex = 15;
            this.commitTreeBtn.Text = "Commit Tree";
            this.commitTreeBtn.UseVisualStyleBackColor = true;
            this.commitTreeBtn.Click += new System.EventHandler(this.commitTreeBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 431);
            this.Controls.Add(this.commitTreeBtn);
            this.Controls.Add(this.treeIO);
            this.Controls.Add(this.savePathLbl);
            this.Controls.Add(this.plcModuleBox);
            this.Controls.Add(this.panelNameBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chassisDropSelect);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.compileBtn);
            this.Controls.Add(this.changePathBtn);
            this.Controls.Add(this.importExcelBtn);
            this.Icon = global::L5K_Compiler.Properties.Resources.icon;
            this.Margin = new System.Windows.Forms.Padding(1);
            this.Name = "Form1";
            this.Text = "Gyptech L5K Compiler";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button importExcelBtn;
        private System.Windows.Forms.Button changePathBtn;
        private System.Windows.Forms.Button compileBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox chassisDropSelect;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox panelNameBox;
        private System.Windows.Forms.TextBox plcModuleBox;
        private System.Windows.Forms.Label savePathLbl;
        private System.Windows.Forms.Button commitTreeBtn;
        public System.Windows.Forms.TreeView treeIO;
    }
}

