namespace DocumentVariablesExample
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
            if (disposing && (components != null)) {
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
            this.progressBarControl1 = new DevExpress.XtraEditors.ProgressBarControl();
            this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
            this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
            this.richEditControl1 = new DevExpress.XtraRichEdit.RichEditControl();
            this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
            this.richEditControl2 = new DevExpress.XtraRichEdit.RichEditControl();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.btn_ShowHiddenText = new DevExpress.XtraEditors.CheckButton();
            this.btn_ShowCodes = new DevExpress.XtraEditors.CheckButton();
            this.btnMailMerge = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
            this.xtraTabControl1.SuspendLayout();
            this.xtraTabPage1.SuspendLayout();
            this.xtraTabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // progressBarControl1
            // 
            this.progressBarControl1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBarControl1.Location = new System.Drawing.Point(0, 544);
            this.progressBarControl1.Name = "progressBarControl1";
            this.progressBarControl1.Size = new System.Drawing.Size(784, 18);
            this.progressBarControl1.TabIndex = 1;
            // 
            // xtraTabControl1
            // 
            this.xtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.xtraTabControl1.Location = new System.Drawing.Point(0, 37);
            this.xtraTabControl1.Name = "xtraTabControl1";
            this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
            this.xtraTabControl1.Size = new System.Drawing.Size(784, 507);
            this.xtraTabControl1.TabIndex = 2;
            this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
            this.xtraTabPage1,
            this.xtraTabPage2});
            this.xtraTabControl1.Selected += new DevExpress.XtraTab.TabPageEventHandler(this.xtraTabControl1_Selected);
            // 
            // xtraTabPage1
            // 
            this.xtraTabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.xtraTabPage1.Controls.Add(this.richEditControl1);
            this.xtraTabPage1.Name = "xtraTabPage1";
            this.xtraTabPage1.Size = new System.Drawing.Size(778, 481);
            this.xtraTabPage1.Text = "Template";
            // 
            // richEditControl1
            // 
            this.richEditControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControl1.Location = new System.Drawing.Point(0, 0);
            this.richEditControl1.Name = "richEditControl1";
            this.richEditControl1.Options.Fields.HighlightMode = DevExpress.XtraRichEdit.FieldsHighlightMode.Always;
            this.richEditControl1.Size = new System.Drawing.Size(778, 481);
            this.richEditControl1.TabIndex = 0;
            this.richEditControl1.Text = "richEditControl1";
            this.richEditControl1.MailMergeStarted += new DevExpress.XtraRichEdit.MailMergeStartedEventHandler(this.richEditControl1_MailMergeStarted);
            this.richEditControl1.MailMergeFinished += new DevExpress.XtraRichEdit.MailMergeFinishedEventHandler(this.richEditControl1_MailMergeFinished);
            this.richEditControl1.MailMergeRecordStarted += new DevExpress.XtraRichEdit.MailMergeRecordStartedEventHandler(this.richEditControl1_MailMergeRecordStarted);
            this.richEditControl1.CalculateDocumentVariable += new DevExpress.XtraRichEdit.CalculateDocumentVariableEventHandler(this.eventHandler_CalculateDocumentVariable);
            this.richEditControl1.MailMergeRecordFinished += new DevExpress.XtraRichEdit.MailMergeRecordFinishedEventHandler(this.richEditControl1_MailMergeRecordFinished);
            // 
            // xtraTabPage2
            // 
            this.xtraTabPage2.Controls.Add(this.richEditControl2);
            this.xtraTabPage2.Name = "xtraTabPage2";
            this.xtraTabPage2.Size = new System.Drawing.Size(778, 481);
            this.xtraTabPage2.Text = "Result";
            // 
            // richEditControl2
            // 
            this.richEditControl2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControl2.Location = new System.Drawing.Point(0, 0);
            this.richEditControl2.Name = "richEditControl2";
            this.richEditControl2.Size = new System.Drawing.Size(778, 481);
            this.richEditControl2.TabIndex = 0;
            this.richEditControl2.Text = "richEditControl2";
            // 
            // panelControl1
            // 
            this.panelControl1.Controls.Add(this.btn_ShowHiddenText);
            this.panelControl1.Controls.Add(this.btn_ShowCodes);
            this.panelControl1.Controls.Add(this.btnMailMerge);
            this.panelControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelControl1.Location = new System.Drawing.Point(0, 0);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(784, 37);
            this.panelControl1.TabIndex = 3;
            // 
            // btn_ShowHiddenText
            // 
            this.btn_ShowHiddenText.Location = new System.Drawing.Point(360, 9);
            this.btn_ShowHiddenText.Name = "btn_ShowHiddenText";
            this.btn_ShowHiddenText.Size = new System.Drawing.Size(104, 23);
            this.btn_ShowHiddenText.TabIndex = 4;
            this.btn_ShowHiddenText.Text = "Show Hidden Text";
            this.btn_ShowHiddenText.CheckedChanged += new System.EventHandler(this.btn_ShowHiddenText_CheckedChanged);
            // 
            // btn_ShowCodes
            // 
            this.btn_ShowCodes.Location = new System.Drawing.Point(242, 9);
            this.btn_ShowCodes.Name = "btn_ShowCodes";
            this.btn_ShowCodes.Size = new System.Drawing.Size(95, 23);
            this.btn_ShowCodes.TabIndex = 3;
            this.btn_ShowCodes.Text = "Show Field Codes";
            this.btn_ShowCodes.CheckedChanged += new System.EventHandler(this.btn_ShowCodes_Click);
            // 
            // btnMailMerge
            // 
            this.btnMailMerge.Location = new System.Drawing.Point(12, 8);
            this.btnMailMerge.Name = "btnMailMerge";
            this.btnMailMerge.Size = new System.Drawing.Size(111, 23);
            this.btnMailMerge.TabIndex = 0;
            this.btnMailMerge.Text = "Mail Merge";
            this.btnMailMerge.Click += new System.EventHandler(this.btnMailMerge_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 562);
            this.Controls.Add(this.xtraTabControl1);
            this.Controls.Add(this.progressBarControl1);
            this.Controls.Add(this.panelControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
            this.xtraTabControl1.ResumeLayout(false);
            this.xtraTabPage1.ResumeLayout(false);
            this.xtraTabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.ProgressBarControl progressBarControl1;
        private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
        private DevExpress.XtraRichEdit.RichEditControl richEditControl1;
        private DevExpress.XtraRichEdit.RichEditControl richEditControl2;
        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraEditors.SimpleButton btnMailMerge;
        private DevExpress.XtraEditors.CheckButton btn_ShowHiddenText;
        private DevExpress.XtraEditors.CheckButton btn_ShowCodes;

    }
}

