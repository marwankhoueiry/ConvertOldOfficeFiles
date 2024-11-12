
namespace ConvertOldOfficeFiles
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.tbPath = new System.Windows.Forms.ToolStripTextBox();
            this.btSelectPath = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btConvert = new System.Windows.Forms.ToolStripButton();
            this.btCheck = new System.Windows.Forms.ToolStripButton();
            this.tbOutput = new System.Windows.Forms.TextBox();
            this.chkIncludeSubFolders = new System.Windows.Forms.CheckBox();
            this.statusStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabel,
            this.progressBar});
            this.statusStrip1.Location = new System.Drawing.Point(0, 369);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 10, 0);
            this.statusStrip1.Size = new System.Drawing.Size(927, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(39, 17);
            this.statusLabel.Text = "Ready";
            // 
            // progressBar
            // 
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(75, 16);
            this.progressBar.Visible = false;
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel1,
            this.tbPath,
            this.btSelectPath,
            this.toolStripSeparator1,
            this.btConvert,
            this.btCheck});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(927, 25);
            this.toolStrip1.TabIndex = 3;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(58, 22);
            this.toolStripLabel1.Text = "Directory:";
            // 
            // tbPath
            // 
            this.tbPath.Name = "tbPath";
            this.tbPath.Size = new System.Drawing.Size(301, 25);
            // 
            // btSelectPath
            // 
            this.btSelectPath.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btSelectPath.Image = ((System.Drawing.Image)(resources.GetObject("btSelectPath.Image")));
            this.btSelectPath.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btSelectPath.Name = "btSelectPath";
            this.btSelectPath.Size = new System.Drawing.Size(23, 22);
            this.btSelectPath.Text = "...";
            this.btSelectPath.Click += new System.EventHandler(this.btSelectPath_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // btConvert
            // 
            this.btConvert.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btConvert.Image = ((System.Drawing.Image)(resources.GetObject("btConvert.Image")));
            this.btConvert.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btConvert.Name = "btConvert";
            this.btConvert.Size = new System.Drawing.Size(53, 22);
            this.btConvert.Text = "C&onvert";
            this.btConvert.Click += new System.EventHandler(this.btConvert_Click);
            // 
            // btCheck
            // 
            this.btCheck.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btCheck.Image = ((System.Drawing.Image)(resources.GetObject("btCheck.Image")));
            this.btCheck.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btCheck.Name = "btCheck";
            this.btCheck.Size = new System.Drawing.Size(44, 22);
            this.btCheck.Text = "&Check";
            this.btCheck.Click += new System.EventHandler(this.btCheck_Click);
            // 
            // tbOutput
            // 
            this.tbOutput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbOutput.Font = new System.Drawing.Font("Consolas", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbOutput.Location = new System.Drawing.Point(0, 25);
            this.tbOutput.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbOutput.Multiline = true;
            this.tbOutput.Name = "tbOutput";
            this.tbOutput.ReadOnly = true;
            this.tbOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbOutput.Size = new System.Drawing.Size(927, 344);
            this.tbOutput.TabIndex = 4;
            // 
            // chkIncludeSubFolders
            // 
            this.chkIncludeSubFolders.AutoSize = true;
            this.chkIncludeSubFolders.Location = new System.Drawing.Point(504, 3);
            this.chkIncludeSubFolders.Name = "chkIncludeSubFolders";
            this.chkIncludeSubFolders.Size = new System.Drawing.Size(114, 17);
            this.chkIncludeSubFolders.TabIndex = 5;
            this.chkIncludeSubFolders.Text = "&Include Subfolders";
            this.chkIncludeSubFolders.UseVisualStyleBackColor = true;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(927, 391);
            this.Controls.Add(this.chkIncludeSubFolders);
            this.Controls.Add(this.tbOutput);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.statusStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmMain";
            this.Text = "ConvertOfficeFiles";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Shown += new System.EventHandler(this.frmMain_Shown);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
        private System.Windows.Forms.ToolStripProgressBar progressBar;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripTextBox tbPath;
        private System.Windows.Forms.ToolStripButton btConvert;
        private System.Windows.Forms.ToolStripButton btSelectPath;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.TextBox tbOutput;
        private System.Windows.Forms.ToolStripButton btCheck;
        private System.Windows.Forms.CheckBox chkIncludeSubFolders;
    }
}

