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
            this.ieVerBox = new System.Windows.Forms.TextBox();
            this.ieVerLbl = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.procTypeDrop = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // importExcelBtn
            // 
            this.importExcelBtn.Location = new System.Drawing.Point(13, 10);
            this.importExcelBtn.Margin = new System.Windows.Forms.Padding(1);
            this.importExcelBtn.Name = "importExcelBtn";
            this.importExcelBtn.Size = new System.Drawing.Size(117, 23);
            this.importExcelBtn.TabIndex = 0;
            this.importExcelBtn.Text = "Import Excel";
            this.importExcelBtn.UseVisualStyleBackColor = true;
            this.importExcelBtn.Click += new System.EventHandler(this.importExcelBtn_Click);
            // 
            // changePathBtn
            // 
            this.changePathBtn.Location = new System.Drawing.Point(13, 42);
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
            this.compileBtn.Location = new System.Drawing.Point(13, 74);
            this.compileBtn.Margin = new System.Windows.Forms.Padding(1);
            this.compileBtn.Name = "compileBtn";
            this.compileBtn.Size = new System.Drawing.Size(117, 23);
            this.compileBtn.TabIndex = 2;
            this.compileBtn.Text = "Compile L5K";
            this.compileBtn.UseVisualStyleBackColor = true;
            this.compileBtn.Click += new System.EventHandler(this.compileBtn_Click);
            // 
            // ieVerBox
            // 
            this.ieVerBox.Location = new System.Drawing.Point(321, 12);
            this.ieVerBox.MaxLength = 4;
            this.ieVerBox.Name = "ieVerBox";
            this.ieVerBox.Size = new System.Drawing.Size(100, 20);
            this.ieVerBox.TabIndex = 4;
            this.ieVerBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // ieVerLbl
            // 
            this.ieVerLbl.AutoSize = true;
            this.ieVerLbl.Location = new System.Drawing.Point(267, 15);
            this.ieVerLbl.Name = "ieVerLbl";
            this.ieVerLbl.Size = new System.Drawing.Size(48, 13);
            this.ieVerLbl.TabIndex = 5;
            this.ieVerLbl.Text = "IE_VER:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(209, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Ctrl Proc Type:";
            // 
            // procTypeDrop
            // 
            this.procTypeDrop.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.procTypeDrop.FormattingEnabled = true;
            this.procTypeDrop.Location = new System.Drawing.Point(300, 52);
            this.procTypeDrop.Name = "procTypeDrop";
            this.procTypeDrop.Size = new System.Drawing.Size(121, 21);
            this.procTypeDrop.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(433, 142);
            this.Controls.Add(this.procTypeDrop);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ieVerLbl);
            this.Controls.Add(this.ieVerBox);
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
        private System.Windows.Forms.TextBox ieVerBox;
        private System.Windows.Forms.Label ieVerLbl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox procTypeDrop;
    }
}

