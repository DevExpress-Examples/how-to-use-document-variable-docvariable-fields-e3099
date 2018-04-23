using System;
using System.Drawing;
using System.Windows.Forms;
#region #usings
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Services;
#endregion #usings

namespace DocumentVariablesExample {
    public partial class Form1 : DevExpress.XtraEditors.XtraForm {
        RichEditControl richEdit;

        public Form1() {
            InitializeComponent();
            richEditControl1.LoadDocument("Docs\\invitation.docx");

            LockFields();

            richEditControl1.Options.MailMerge.DataSource = new SampleData();
            richEditControl2.Document.CalculateDocumentVariable += new CalculateDocumentVariableEventHandler(eventHandler_CalculateDocumentVariable);
            this.richEdit = richEditControl1;
        }

        private void LockFields() {
            #region #lockfields
            richEditControl1.Document.Fields[0].Locked = true;
            richEditControl2.Options.Fields.UpdateLockedFields = UpdateLockedFields.Never;
            #endregion #lockfields
        }

        private void btnMailMerge_Click(object sender, EventArgs e) {
            MailMergeOptions myMergeOptions = richEditControl1.Document.CreateMailMergeOptions();
            myMergeOptions.MergeMode = MergeMode.NewSection;
            this.Cursor = Cursors.WaitCursor;
            richEditControl1.Document.MailMerge(myMergeOptions, richEditControl2.Document);
            this.Cursor = Cursors.Default;
            xtraTabControl1.SelectedTabPageIndex = 1;
        }
        #region #servicesubst
        private void richEditControl1_MailMergeStarted(object sender, MailMergeStartedEventArgs e) {
            richEditControl1.RemoveService(typeof(IProgressIndicationService));
            richEditControl1.AddService(typeof(IProgressIndicationService),
                new MyProgressIndicatorService(richEditControl1, this.progressBarControl1));
        }
        #endregion #servicesubst

        private void richEditControl1_MailMergeFinished(object sender, MailMergeFinishedEventArgs e) {
            richEditControl1.RemoveService(typeof(IProgressIndicationService));
        }

        #region #mailmergerecordstarted
        private void richEditControl1_MailMergeRecordStarted(object sender, MailMergeRecordStartedEventArgs e) {
            DocumentRange _range = e.RecordDocument.InsertText(e.RecordDocument.Range.Start,
String.Format("Created on {0:G}\n\n", DateTime.Now));
            CharacterProperties cp = e.RecordDocument.BeginUpdateCharacters(_range);
            cp.FontSize = 8;
            cp.ForeColor = Color.Red;
            cp.Hidden = true;
            e.RecordDocument.EndUpdateCharacters(cp);
        }
        #endregion #mailmergerecordstarted

        #region #mailmergerecordfinished
        private void richEditControl1_MailMergeRecordFinished(object sender, MailMergeRecordFinishedEventArgs e) {
            e.RecordDocument.AppendDocumentContent("Docs\\bungalow.docx", DocumentFormat.OpenXml);
        }
        #endregion #mailmergerecordfinished

        #region #calculatedocumentvariable
        void eventHandler_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e) {
            string location = e.Arguments[0].Value.ToString();
            if ((location.Trim() == String.Empty) || (location.Contains("<"))) {
                e.Value = " ";
                e.Handled = true;
                return;
            }

            switch (e.VariableName) {
                case "Weather":
                    Conditions conditions = new Conditions();
                    conditions = Weather.GetCurrentConditions(location);
                    e.Value = String.Format("Weather for {0}: \nConditions: {1}\nTemperature (C) :{2}\nHumidity: {3}\nWind: {4}\n",
                        location, conditions.Condition, conditions.TempC, conditions.Humidity, conditions.Wind);
                    break;
                case "LOCATION":
                    if (location == "DO NOT CHANGE!") e.Value = DocVariableValue.Current;
                    break;
            }
            e.Handled = true;
        }
        #endregion #calculatedocumentvariable

        private void btnUpdateDocVariableFields_Click(object sender, EventArgs e) {
            #region #updatedocvariablefields
            richEditControl2.Options.Fields.UpdateLockedFields = UpdateLockedFields.DocVariableOnly;
            richEditControl2.Document.Fields.Update();
            #endregion #updatedocvariablefields
        }

        private void xtraTabControl1_Selected(object sender, DevExpress.XtraTab.TabPageEventArgs e) {
            switch (e.PageIndex) {
                case 0:
                    richEdit = richEditControl1;
                    this.btnMailMerge.Enabled = true;
                    break;
                case 1:
                    richEdit = richEditControl2;
                    this.btnMailMerge.Enabled = false;
                    break;
            }
        }

        private void btn_ShowCodes_Click(object sender, EventArgs e) {
            Document doc = richEdit.Document;
            bool showCodes = btn_ShowCodes.Checked;
            doc.BeginUpdate();
            foreach (Field f in doc.Fields) f.ShowCodes = showCodes;
            doc.EndUpdate();
        }

        private void btn_ShowHiddenText_CheckedChanged(object sender, EventArgs e) {
            richEdit.Options.FormattingMarkVisibility.HiddenText = btn_ShowHiddenText.Checked ? RichEditFormattingMarkVisibility.Visible : RichEditFormattingMarkVisibility.Hidden;
        }

    }
}