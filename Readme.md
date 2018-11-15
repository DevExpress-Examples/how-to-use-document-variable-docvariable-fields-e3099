<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/DocumentVariablesExample/Form1.cs) (VB: [Form1.vb](./VB/DocumentVariablesExample/Form1.vb))
* [MyProgressIndicator.cs](./CS/DocumentVariablesExample/MyProgressIndicator.cs) (VB: [MyProgressIndicator.vb](./VB/DocumentVariablesExample/MyProgressIndicator.vb))
* [Program.cs](./CS/DocumentVariablesExample/Program.cs) (VB: [Program.vb](./VB/DocumentVariablesExample/Program.vb))
* [SampleData.cs](./CS/DocumentVariablesExample/SampleData.cs) (VB: [SampleData.vb](./VB/DocumentVariablesExample/SampleData.vb))
* [Weather.cs](./CS/DocumentVariablesExample/Weather.cs) (VB: [Weather.vb](./VB/DocumentVariablesExample/Weather.vb))
<!-- default file list end -->
# How to use document variable (DOCVARIABLE) fields


<p>This example illustrates the use of a <strong>DOCVARIABLE </strong>field to provide additional information which is dependent on the value of a merged field. This technique is implemented so each merged document contains a weather report for a location that corresponds to the current data record.</p>
<p>NB: We do not provide code for retrieving weather information. You can implement a custom weather information provider.</p>
<p>The location is represented by a merge field. It is included as an argument within the DOCVARIABLE field. When the DOCVARIABLE field is updated, the <strong>DevExpress.XtraRichEdit.API.Native.Document.CalculateDocumentVariable</strong> event is triggered. A code within the event handler obtains the information on weather. It uses <u>e.VariableName</u> to get the name of the variable within the field, <u>e.Arguments</u> to get the location and returns the calculated result in <u>e.Value</u> property.<br /> The <strong>MailMergeRecordStarted</strong> event is handled to insert a hidden text indicating when the document is created. <br /> The <strong>MyProgressIndicatorService</strong> class is implemented and registered as a service to allow progress indication using the ProgressBar control.</p>


<h3>Description</h3>

Illustrates the use of&nbsp;<a href="http://help.devexpress.com/#WindowsForms/CustomDocument17168">locked field</a>&nbsp;feature to retain&nbsp;<strong>DOCVARIABLE </strong>field value after mail merge. Subsequently the field can be updated as your need dictates.

<br/>


