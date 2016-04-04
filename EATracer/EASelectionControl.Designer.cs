namespace EATracer
{
    partial class EASelectionControl
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
                this.Repository.Exit();
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.fileTextBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.modelCombo = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboElements = new System.Windows.Forms.ComboBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.label4 = new System.Windows.Forms.Label();
            this.comboPackage = new System.Windows.Forms.ComboBox();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "EA File:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "EA Files|*.eap";
            // 
            // fileTextBox
            // 
            this.fileTextBox.Location = new System.Drawing.Point(83, 14);
            this.fileTextBox.Name = "fileTextBox";
            this.fileTextBox.ReadOnly = true;
            this.fileTextBox.Size = new System.Drawing.Size(262, 20);
            this.fileTextBox.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(358, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(53, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Model:";
            // 
            // modelCombo
            // 
            this.modelCombo.FormattingEnabled = true;
            this.modelCombo.Location = new System.Drawing.Point(83, 48);
            this.modelCombo.Name = "modelCombo";
            this.modelCombo.Size = new System.Drawing.Size(328, 21);
            this.modelCombo.TabIndex = 4;
            this.modelCombo.SelectedIndexChanged += new System.EventHandler(this.modelCombo_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 134);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(48, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Element:";
            // 
            // comboElements
            // 
            this.comboElements.FormattingEnabled = true;
            this.comboElements.Location = new System.Drawing.Point(83, 130);
            this.comboElements.Name = "comboElements";
            this.comboElements.Size = new System.Drawing.Size(328, 21);
            this.comboElements.TabIndex = 6;
            this.comboElements.SelectedIndexChanged += new System.EventHandler(this.comboElements_SelectedIndexChanged);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 389);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(430, 22);
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "Ready!";
            // 
            // statusLabel
            // 
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(42, 17);
            this.statusLabel.Text = "Ready!";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(18, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Package:";
            // 
            // comboPackage
            // 
            this.comboPackage.FormattingEnabled = true;
            this.comboPackage.Location = new System.Drawing.Point(83, 90);
            this.comboPackage.Name = "comboPackage";
            this.comboPackage.Size = new System.Drawing.Size(328, 21);
            this.comboPackage.TabIndex = 9;
            this.comboPackage.SelectedIndexChanged += new System.EventHandler(this.comboPackage_SelectedIndexChanged);
            // 
            // EASelectionControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.comboPackage);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.comboElements);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.modelCombo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.fileTextBox);
            this.Controls.Add(this.label1);
            this.Name = "EASelectionControl";
            this.Size = new System.Drawing.Size(430, 411);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox fileTextBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox modelCombo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboElements;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboPackage;
    }
}
