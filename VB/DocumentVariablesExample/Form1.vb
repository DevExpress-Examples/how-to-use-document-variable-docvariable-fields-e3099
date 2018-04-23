Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native
#Region "#usings"
Imports DevExpress.XtraRichEdit
Imports DevExpress.Services
#End Region ' #usings
Imports DevExpress.XtraRichEdit.Commands

Namespace DocumentVariablesExample
	Partial Public Class Form1
		Inherits Form
		Private richEdit As RichEditControl

		Public Sub New()
			InitializeComponent()
			richEditControl1.LoadDocument("Docs\invitation.docx")
			richEditControl1.Options.MailMerge.DataSource = New SampleData()
			AddHandler richEditControl2.Document.CalculateDocumentVariable, AddressOf eventHandler_CalculateDocumentVariable
			Me.richEdit = richEditControl1
		End Sub

		Private Sub btnMailMerge_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnMailMerge.Click
			Dim myMergeOptions As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
			myMergeOptions.MergeMode = MergeMode.NewSection
			'myMergeOptions.FirstRecordIndex = 1;
			'myMergeOptions.LastRecordIndex = 3;
			Me.Cursor = Cursors.WaitCursor
			richEditControl1.Document.MailMerge(myMergeOptions, richEditControl2.Document)
			Me.Cursor = Cursors.Default
			xtraTabControl1.SelectedTabPageIndex = 1
			'richEditControl2.ActiveView.ZoomFactor = 0.4f;
		End Sub
#Region "#servicesubst"
		Private Sub richEditControl1_MailMergeStarted(ByVal sender As Object, ByVal e As MailMergeStartedEventArgs) Handles richEditControl1.MailMergeStarted
			richEditControl1.RemoveService(GetType(IProgressIndicationService))
			richEditControl1.AddService(GetType(IProgressIndicationService), New MyProgressIndicatorService(richEditControl1, Me.progressBarControl1))
		End Sub
  #End Region ' #servicesubst

		Private Sub richEditControl1_MailMergeFinished(ByVal sender As Object, ByVal e As MailMergeFinishedEventArgs) Handles richEditControl1.MailMergeFinished
			richEditControl1.RemoveService(GetType(IProgressIndicationService))
		End Sub

#Region "#mailmergerecordstarted"
		Private Sub richEditControl1_MailMergeRecordStarted(ByVal sender As Object, ByVal e As MailMergeRecordStartedEventArgs) Handles richEditControl1.MailMergeRecordStarted
			Dim _range As DocumentRange = e.RecordDocument.InsertText(e.RecordDocument.Range.Start, String.Format("Created on {0:G}" & Constants.vbLf + Constants.vbLf, DateTime.Now))
			Dim cp As CharacterProperties = e.RecordDocument.BeginUpdateCharacters(_range)
			cp.FontSize = 8
			cp.ForeColor = Color.Red
			cp.Hidden = True
			e.RecordDocument.EndUpdateCharacters(cp)
		End Sub
#End Region ' #mailmergerecordstarted

#Region "#mailmergerecordfinished"
		Private Sub richEditControl1_MailMergeRecordFinished(ByVal sender As Object, ByVal e As MailMergeRecordFinishedEventArgs) Handles richEditControl1.MailMergeRecordFinished
			e.RecordDocument.AppendDocumentContent("Docs\bungalow.docx", DocumentFormat.OpenXml)
		End Sub
#End Region ' #mailmergerecordfinished

#Region "#calculatedocumentvariable"
		Private Sub eventHandler_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs) Handles richEditControl1.CalculateDocumentVariable
			Dim location As String = e.Arguments(0).Value.ToString()

			Console.WriteLine(e.VariableName & " " & location)

			If (location.Trim() = String.Empty) OrElse (location.Contains("<")) Then
				e.Value = " "
				e.Handled = True
				Return
			End If

			Select Case e.VariableName
				Case "Weather"
					Dim conditions As New Conditions()
					conditions = Weather.GetCurrentConditions(location)
					e.Value = String.Format("Forecast for {0}: " & Constants.vbLf & "Conditions: {1}" & Constants.vbLf & "Temperature (C) :{2}" & Constants.vbLf & "Humidity: {3}" & Constants.vbLf & "Wind: {4}" & Constants.vbLf, conditions.City, conditions.Condition, conditions.TempC, conditions.Humidity, conditions.Wind)
				Case "Location"
					Dim loc() As GeoLocation = GeoLocation.GeocodeAddress(location)
					e.Value = String.Format(" {0}" & Constants.vbLf & "Latitude: {1}" & Constants.vbLf & "Longitude: {2}" & Constants.vbLf, loc(0).Address, loc(0).Latitude.ToString(), loc(0).Longitude.ToString())
			End Select
			e.Handled = True
		End Sub
#End Region ' #calculatedocumentvariable
		Private Sub xtraTabControl1_Selected(ByVal sender As Object, ByVal e As DevExpress.XtraTab.TabPageEventArgs) Handles xtraTabControl1.Selected
			Select Case e.PageIndex
				Case 0
					richEdit = richEditControl1
					Me.btnMailMerge.Enabled = True
				Case 1
					richEdit = richEditControl2
					Me.btnMailMerge.Enabled = False
			End Select
		End Sub

		Private Sub btn_ShowCodes_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_ShowCodes.CheckedChanged
			Dim doc As Document = richEdit.Document
			Dim showCodes As Boolean = btn_ShowCodes.Checked
			doc.BeginUpdate()
			For Each f As Field In doc.Fields
				f.ShowCodes = showCodes
			Next f
			doc.EndUpdate()
		End Sub

		Private Sub btn_ShowHiddenText_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles btn_ShowHiddenText.CheckedChanged
			If btn_ShowHiddenText.Checked Then
				richEdit.Options.FormattingMarkVisibility.HiddenText = RichEditFormattingMarkVisibility.Visible
			Else
				richEdit.Options.FormattingMarkVisibility.HiddenText = RichEditFormattingMarkVisibility.Hidden
			End If
		End Sub
	End Class
End Namespace