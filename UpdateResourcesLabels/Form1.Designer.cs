namespace UpdateResourcesLabels
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
            this.button1 = new System.Windows.Forms.Button();
            this.rtbMissedKeys = new System.Windows.Forms.RichTextBox();
            this.lbProcess = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(136, 71);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // rtbMissedKeys
            // 
            this.rtbMissedKeys.Location = new System.Drawing.Point(17, 136);
            this.rtbMissedKeys.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.rtbMissedKeys.Name = "rtbMissedKeys";
            this.rtbMissedKeys.Size = new System.Drawing.Size(1105, 509);
            this.rtbMissedKeys.TabIndex = 1;
            this.rtbMissedKeys.Text = "";
            // 
            // lbProcess
            // 
            this.lbProcess.AutoSize = true;
            this.lbProcess.Location = new System.Drawing.Point(22, 104);
            this.lbProcess.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbProcess.Name = "lbProcess";
            this.lbProcess.Size = new System.Drawing.Size(53, 13);
            this.lbProcess.TabIndex = 2;
            this.lbProcess.Text = "lbProcess";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1146, 670);
            this.Controls.Add(this.lbProcess);
            this.Controls.Add(this.rtbMissedKeys);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RichTextBox rtbMissedKeys;
        private System.Windows.Forms.Label lbProcess;
    }
}

