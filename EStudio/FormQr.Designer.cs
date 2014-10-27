namespace EStudio
{
    partial class FormQr
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
            this.pictureBoxQR = new System.Windows.Forms.PictureBox();
            this.btnQrExit = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxQR)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBoxQR
            // 
            this.pictureBoxQR.Location = new System.Drawing.Point(24, 47);
            this.pictureBoxQR.Name = "pictureBoxQR";
            this.pictureBoxQR.Size = new System.Drawing.Size(285, 207);
            this.pictureBoxQR.TabIndex = 0;
            this.pictureBoxQR.TabStop = false;
            // 
            // btnQrExit
            // 
            this.btnQrExit.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnQrExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnQrExit.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQrExit.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.btnQrExit.Location = new System.Drawing.Point(24, 12);
            this.btnQrExit.Name = "btnQrExit";
            this.btnQrExit.Size = new System.Drawing.Size(60, 29);
            this.btnQrExit.TabIndex = 1;
            this.btnQrExit.Text = "<<返回";
            this.btnQrExit.UseVisualStyleBackColor = false;
            this.btnQrExit.Click += new System.EventHandler(this.btnQrExit_Click);
            // 
            // FormQr
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.HighlightText;
            this.ClientSize = new System.Drawing.Size(334, 266);
            this.Controls.Add(this.btnQrExit);
            this.Controls.Add(this.pictureBoxQR);
            this.MaximizeBox = false;
            this.Name = "FormQr";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "扫一扫";
            this.Load += new System.EventHandler(this.FormQr_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxQR)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBoxQR;
        private System.Windows.Forms.Button btnQrExit;
    }
}