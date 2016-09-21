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
            this.SuspendLayout();
            // 
            // importExcelBtn
            // 
            this.importExcelBtn.Location = new System.Drawing.Point(14, 12);
            this.importExcelBtn.Name = "importExcelBtn";
            this.importExcelBtn.Size = new System.Drawing.Size(313, 55);
            this.importExcelBtn.TabIndex = 0;
            this.importExcelBtn.Text = "Import Excel";
            this.importExcelBtn.UseVisualStyleBackColor = true;
            // 
            // changePathBtn
            // 
            this.changePathBtn.Location = new System.Drawing.Point(14, 89);
            this.changePathBtn.Name = "changePathBtn";
            this.changePathBtn.Size = new System.Drawing.Size(313, 55);
            this.changePathBtn.TabIndex = 1;
            this.changePathBtn.Text = "Change Save Path";
            this.changePathBtn.UseVisualStyleBackColor = true;
            this.changePathBtn.Click += new System.EventHandler(this.changePathBtn_Click);
            // 
            // compileBtn
            // 
            this.compileBtn.Location = new System.Drawing.Point(14, 163);
            this.compileBtn.Name = "compileBtn";
            this.compileBtn.Size = new System.Drawing.Size(313, 55);
            this.compileBtn.TabIndex = 2;
            this.compileBtn.Text = "Compile L5K";
            this.compileBtn.UseVisualStyleBackColor = true;
            this.compileBtn.Click += new System.EventHandler(this.compileBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(16F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1442, 684);
            this.Controls.Add(this.compileBtn);
            this.Controls.Add(this.changePathBtn);
            this.Controls.Add(this.importExcelBtn);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button importExcelBtn;
        private System.Windows.Forms.Button changePathBtn;
        private System.Windows.Forms.Button compileBtn;
    }
}

