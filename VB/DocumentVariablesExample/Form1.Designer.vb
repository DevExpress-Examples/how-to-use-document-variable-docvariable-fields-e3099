Namespace DocumentVariablesExample
    Partial Public Class Form1
        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        #Region "Windows Form Designer generated code"

        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Me.progressBarControl1 = New DevExpress.XtraEditors.ProgressBarControl()
            Me.xtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
            Me.xtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControl1 = New DevExpress.XtraRichEdit.RichEditControl()
            Me.xtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControl2 = New DevExpress.XtraRichEdit.RichEditControl()
            Me.panelControl1 = New DevExpress.XtraEditors.PanelControl()
            Me.btnUpdateDocVariableFields = New DevExpress.XtraEditors.SimpleButton()
            Me.btn_ShowHiddenText = New DevExpress.XtraEditors.CheckButton()
            Me.btn_ShowCodes = New DevExpress.XtraEditors.CheckButton()
            Me.btnMailMerge = New DevExpress.XtraEditors.SimpleButton()
            Me.defaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel(Me.components)
            DirectCast(Me.progressBarControl1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            DirectCast(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.xtraTabControl1.SuspendLayout()
            Me.xtraTabPage1.SuspendLayout()
            Me.xtraTabPage2.SuspendLayout()
            DirectCast(Me.panelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.panelControl1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' progressBarControl1
            ' 
            Me.progressBarControl1.Dock = System.Windows.Forms.DockStyle.Bottom
            Me.progressBarControl1.Location = New System.Drawing.Point(0, 544)
            Me.progressBarControl1.Name = "progressBarControl1"
            Me.progressBarControl1.Size = New System.Drawing.Size(784, 18)
            Me.progressBarControl1.TabIndex = 1
            ' 
            ' xtraTabControl1
            ' 
            Me.xtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.xtraTabControl1.Location = New System.Drawing.Point(0, 37)
            Me.xtraTabControl1.Name = "xtraTabControl1"
            Me.xtraTabControl1.SelectedTabPage = Me.xtraTabPage1
            Me.xtraTabControl1.Size = New System.Drawing.Size(784, 507)
            Me.xtraTabControl1.TabIndex = 2
            Me.xtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() { Me.xtraTabPage1, Me.xtraTabPage2})
            ' 
            ' xtraTabPage1
            ' 
            Me.xtraTabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
            Me.xtraTabPage1.Controls.Add(Me.richEditControl1)
            Me.xtraTabPage1.Name = "xtraTabPage1"
            Me.xtraTabPage1.Size = New System.Drawing.Size(782, 482)
            Me.xtraTabPage1.Text = "Template"
            ' 
            ' richEditControl1
            ' 
            Me.richEditControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControl1.EnableToolTips = True
            Me.richEditControl1.Location = New System.Drawing.Point(0, 0)
            Me.richEditControl1.Name = "richEditControl1"
            Me.richEditControl1.Options.Bookmarks.AllowNameResolution = False
            Me.richEditControl1.Options.Fields.HighlightMode = DevExpress.XtraRichEdit.FieldsHighlightMode.Always
            Me.richEditControl1.Size = New System.Drawing.Size(782, 482)
            Me.richEditControl1.TabIndex = 0
            Me.richEditControl1.Text = "richEditControl1"
            ' 
            ' xtraTabPage2
            ' 
            Me.xtraTabPage2.Controls.Add(Me.richEditControl2)
            Me.xtraTabPage2.Name = "xtraTabPage2"
            Me.xtraTabPage2.Size = New System.Drawing.Size(778, 479)
            Me.xtraTabPage2.Text = "Result"
            ' 
            ' richEditControl2
            ' 
            Me.richEditControl2.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControl2.EnableToolTips = True
            Me.richEditControl2.Location = New System.Drawing.Point(0, 0)
            Me.richEditControl2.Name = "richEditControl2"
            Me.richEditControl2.Options.Bookmarks.AllowNameResolution = False
            Me.richEditControl2.Size = New System.Drawing.Size(778, 479)
            Me.richEditControl2.TabIndex = 0
            Me.richEditControl2.Text = "richEditControl2"
            ' 
            ' panelControl1
            ' 
            Me.panelControl1.Controls.Add(Me.btnUpdateDocVariableFields)
            Me.panelControl1.Controls.Add(Me.btn_ShowHiddenText)
            Me.panelControl1.Controls.Add(Me.btn_ShowCodes)
            Me.panelControl1.Controls.Add(Me.btnMailMerge)
            Me.panelControl1.Dock = System.Windows.Forms.DockStyle.Top
            Me.panelControl1.Location = New System.Drawing.Point(0, 0)
            Me.panelControl1.Name = "panelControl1"
            Me.panelControl1.Size = New System.Drawing.Size(784, 37)
            Me.panelControl1.TabIndex = 3
            ' 
            ' btnUpdateDocVariableFields
            ' 
            Me.btnUpdateDocVariableFields.Location = New System.Drawing.Point(569, 8)
            Me.btnUpdateDocVariableFields.Name = "btnUpdateDocVariableFields"
            Me.btnUpdateDocVariableFields.Size = New System.Drawing.Size(190, 23)
            Me.btnUpdateDocVariableFields.TabIndex = 5
            Me.btnUpdateDocVariableFields.Text = "Update Locked DOCVARIABLE Fields"
            ' 
            ' btn_ShowHiddenText
            ' 
            Me.btn_ShowHiddenText.Location = New System.Drawing.Point(360, 9)
            Me.btn_ShowHiddenText.Name = "btn_ShowHiddenText"
            Me.btn_ShowHiddenText.Size = New System.Drawing.Size(104, 23)
            Me.btn_ShowHiddenText.TabIndex = 4
            Me.btn_ShowHiddenText.Text = "Show Hidden Text"
            ' 
            ' btn_ShowCodes
            ' 
            Me.btn_ShowCodes.Location = New System.Drawing.Point(242, 9)
            Me.btn_ShowCodes.Name = "btn_ShowCodes"
            Me.btn_ShowCodes.Size = New System.Drawing.Size(95, 23)
            Me.btn_ShowCodes.TabIndex = 3
            Me.btn_ShowCodes.Text = "Show Field Codes"
            ' 
            ' btnMailMerge
            ' 
            Me.btnMailMerge.Location = New System.Drawing.Point(12, 8)
            Me.btnMailMerge.Name = "btnMailMerge"
            Me.btnMailMerge.Size = New System.Drawing.Size(111, 23)
            Me.btnMailMerge.TabIndex = 0
            Me.btnMailMerge.Text = "Mail Merge"
            ' 
            ' defaultLookAndFeel1
            ' 
            Me.defaultLookAndFeel1.LookAndFeel.SkinName = "Office 2013"
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(784, 562)
            Me.Controls.Add(Me.xtraTabControl1)
            Me.Controls.Add(Me.progressBarControl1)
            Me.Controls.Add(Me.panelControl1)
            Me.Name = "Form1"
            Me.Text = "Form1"
            DirectCast(Me.progressBarControl1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            DirectCast(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.xtraTabControl1.ResumeLayout(False)
            Me.xtraTabPage1.ResumeLayout(False)
            Me.xtraTabPage2.ResumeLayout(False)
            DirectCast(Me.panelControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.panelControl1.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub

        #End Region

        Private progressBarControl1 As DevExpress.XtraEditors.ProgressBarControl
        Private WithEvents xtraTabControl1 As DevExpress.XtraTab.XtraTabControl
        Private xtraTabPage1 As DevExpress.XtraTab.XtraTabPage
        Private xtraTabPage2 As DevExpress.XtraTab.XtraTabPage
        Private WithEvents richEditControl1 As DevExpress.XtraRichEdit.RichEditControl
        Private richEditControl2 As DevExpress.XtraRichEdit.RichEditControl
        Private panelControl1 As DevExpress.XtraEditors.PanelControl
        Private WithEvents btnMailMerge As DevExpress.XtraEditors.SimpleButton
        Private WithEvents btn_ShowHiddenText As DevExpress.XtraEditors.CheckButton
        Private WithEvents btn_ShowCodes As DevExpress.XtraEditors.CheckButton
        Private WithEvents btnUpdateDocVariableFields As DevExpress.XtraEditors.SimpleButton
        Private defaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    End Class
End Namespace

