namespace Lean.Scanning
{
    partial class Scanning_Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Scanning_Login));
            this.uiImageButton1 = new Sunny.UI.UIImageButton();
            ((System.ComponentModel.ISupportInitialize)(this.uiImageButton1)).BeginInit();
            this.SuspendLayout();
            // 
            // lblSubText
            // 
            this.lblSubText.Location = new System.Drawing.Point(376, 421);
            // 
            // uiImageButton1
            // 
            this.uiImageButton1.BackColor = System.Drawing.Color.Transparent;
            this.uiImageButton1.BackgroundImage = global::Lean.Scanning.Properties.Resources.serial;
            this.uiImageButton1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.uiImageButton1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiImageButton1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiImageButton1.Location = new System.Drawing.Point(158, 192);
            this.uiImageButton1.Name = "uiImageButton1";
            this.uiImageButton1.Size = new System.Drawing.Size(128, 128);
            this.uiImageButton1.TabIndex = 10;
            this.uiImageButton1.TabStop = false;
            this.uiImageButton1.Text = null;
            // 
            // Scanning_Login
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(750, 450);
            this.Controls.Add(this.uiImageButton1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.LoginImage = Sunny.UI.UILoginForm.UILoginImage.Login6;
            this.Name = "Scanning_Login";
            this.Password = "Scanning";
            this.Text = "Scanning_Login";
            this.UserName = "bukan05";
            this.ButtonLoginClick += new System.EventHandler(this.Scanning_Login_ButtonLoginClick);
            this.ButtonCancelClick += new System.EventHandler(this.Scanning_Login_ButtonCancelClick);
            this.Load += new System.EventHandler(this.Scanning_Login_Load);
            this.Controls.SetChildIndex(this.lblTitle, 0);
            this.Controls.SetChildIndex(this.lblSubText, 0);
            this.Controls.SetChildIndex(this.uiPanel1, 0);
            this.Controls.SetChildIndex(this.uiImageButton1, 0);
            ((System.ComponentModel.ISupportInitialize)(this.uiImageButton1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Sunny.UI.UIImageButton uiImageButton1;
    }
}