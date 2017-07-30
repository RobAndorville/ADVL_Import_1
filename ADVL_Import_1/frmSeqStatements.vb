Public Class frmSeqStatements
    'This form is used to add processing statements to the sequence shown on the Import Sequence form.

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------
#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '------------------------------------------------------------------------------------------------------------
#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Main.Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Main.Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Left = Settings.<FormSettings>.<Left>.Value
            End If

            If Settings.<FormSettings>.<Top>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Top = Settings.<FormSettings>.<Top>.Value
            End If

            If Settings.<FormSettings>.<Height>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Height = Settings.<FormSettings>.<Height>.Value
            End If

            If Settings.<FormSettings>.<Width>.Value = Nothing Then
                'Form setting not saved.
            Else
                Me.Width = Settings.<FormSettings>.<Width>.Value
            End If

            'Add code to read other saved setting here:

        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '----------------------------------------------------------------------------------------------------------------------------

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings

        cmbCommandLine.Items.Add("<ExitLoopIf>No_more_input_files</ExitLoopIf>")
        cmbCommandLine.Items.Add("<ExitLoopIf>At_end_of_file</ExitLoopIf>")

        cmbCommandLine.Items.Add("<ProcessingCommand>OpenDatabase</ProcessingCommand>")
        cmbCommandLine.Items.Add("<ProcessingCommand>RunRegExList</ProcessingCommand>")
        cmbCommandLine.Items.Add("<ProcessingCommand>ProcessMatches</ProcessingCommand>")
        cmbCommandLine.Items.Add("<ProcessingCommand>WriteToDatabase</ProcessingCommand>")
        cmbCommandLine.Items.Add("<ProcessingCommand>CloseDatabase</ProcessingCommand>")

        cmbCommandLine.Items.Add("<ReadTextCommand>OpenFirstFile</ReadTextCommand>")
        cmbCommandLine.Items.Add("<ReadTextCommand>ReadNextLine</ReadTextCommand>")
        cmbCommandLine.Items.Add("<ReadTextCommand>OpenNextFile</ReadTextCommand>")

        cmbCommandLine.Items.Add("<!---->")
        cmbCommandLine.Items.Add("<InputDateFormat>yyyyMMdd</InputDateFormat>")
        cmbCommandLine.Items.Add("<OutputDateFormat>dd MMMM yyyy</OutputDateFormat>")

        cmbExitLoopStatus.Items.Add("No_more_input_files")
        cmbExitLoopStatus.Items.Add("At_end_of_file")

        cmbShowValue.Items.Add("Number of selected text files")
        cmbShowValue.Items.Add("Selected text file number")
        cmbShowValue.Items.Add("Text file path")
        cmbShowValue.Items.Add("Sequence status")

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if form is minimised.
        End If
    End Sub



#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------
#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnAddLoop_Click(sender As Object, e As EventArgs) Handles btnAddLoop.Click
        'Add a loop to the processing sequence:

        Main.Sequence.rtbSequence.SelectedText = vbCrLf & "<Loop description=""" & txtLoopDescr.Text & """>" & vbCrLf & "</Loop>" & vbCrLf
        Main.Sequence.FormatXmlText()

    End Sub

    Private Sub LoopExit_Click(sender As Object, e As EventArgs) Handles LoopExit.Click
        'Add an Exit Loop statement to the processing sequence
        Main.Sequence.rtbSequence.SelectedText = "<ExitLoopIf>" & cmbExitLoopStatus.Text & "</ExitLoopIf>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnAddGroup_Click(sender As Object, e As EventArgs) Handles btnAddGroup.Click
        'Add a group to the processing sequence:
        Main.Sequence.rtbSequence.SelectedText = vbCrLf & "<Group description=""" & txtGroupDescr.Text & """>" & vbCrLf & "</Group>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnGroupExit_Click(sender As Object, e As EventArgs) Handles btnGroupExit.Click
        'Add an Exit Group statement to the processing sequence
        Main.Sequence.rtbSequence.SelectedText = "<ExitGroupIf>" & txtExitGroupStatusText.Text & "</ExitGroupIf>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnAddComment_Click(sender As Object, e As EventArgs) Handles btnAddComment.Click
        'Add a comment
        Main.Sequence.rtbSequence.SelectedText = "<!--" & txtComment.Text & "-->" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnShowMessageLine_Click(sender As Object, e As EventArgs) Handles btnShowMessageLine.Click
        'Insert statements to show the line of text in the message window:
        Main.Sequence.rtbSequence.SelectedText = "<ShowMessageLine>" & txtMessageLine.Text & "</ShowMessageLine>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnShowMessageString_Click(sender As Object, e As EventArgs) Handles btnShowMessageString.Click
        'Insert statement to show the text in the message window without appending CrLf characters:
        Main.Sequence.rtbSequence.SelectedText = "<ShowMessageString>" & txtMessageString.Text & "</ShowMessageString>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnShowValue_Click(sender As Object, e As EventArgs) Handles btnShowValue.Click
        'Insert a statement to show the selected values in the Messages window.
        'Dim ValueString As String

        Select Case cmbShowValue.Text
            Case "Number of selected text files"
                Main.Sequence.rtbSequence.SelectedText = "<ShowValue>Number of selected text files</ShowValue>" & vbCrLf
                Main.Sequence.FormatXmlText()
            Case "Selected text file number"
                Main.Sequence.rtbSequence.SelectedText = "<ShowValue>Selected text file number</ShowValue>" & vbCrLf
                Main.Sequence.FormatXmlText()
            Case "Text file path"
                Main.Sequence.rtbSequence.SelectedText = "<ShowValue>Text file path</ShowValue>" & vbCrLf
                Main.Sequence.FormatXmlText()
            Case "Sequence status"
                Main.Sequence.rtbSequence.SelectedText = "<ShowValue>Sequence status</ShowValue>" & vbCrLf
                Main.Sequence.FormatXmlText()
            Case ""
                'Main.MessageColor = Color.Red
                'Main.AddMessage("Show Value button has been pressed with no value selected." & vbCrLf)
                Main.Message.AddWarning("Show Value button has been pressed with no value selected." & vbCrLf)
            Case Else
                'TDS_Import.ShowMessage("Show Value button has been pressed with unrecognised value selected: " & cmbShowValue.Text & vbCrLf, Color.Red)
                'Main.MessageColor = Color.Red
                'Main.AddMessage("Show Value button has been pressed with unrecognised value selected: " & cmbShowValue.Text & vbCrLf)
                Main.Message.AddWarning("Show Value button has been pressed with unrecognised value selected: " & cmbShowValue.Text & vbCrLf)
        End Select
    End Sub

    Private Sub btnCommandLine_Click(sender As Object, e As EventArgs) Handles btnCommandLine.Click
        'Add a Command line to the processing sequence:
        'Command lines include:
        '<ExitLoopIf>No_more_input_files</ExitLoopIf>
        '<ExitLoopIf>At_end_of_file</ExitLoopIf>
        '<ProcessingCommand>WriteToDatabase</ProcessingCommand>
        '<!----> (XML blank commment line)
        '<InputDateFormat>yyyyMMdd</InputDateFormat>
        '<OutputDateFormat>dd MMMM yyyy</OutputDateFormat>

        Main.Sequence.rtbSequence.SelectedText = cmbCommandLine.Text & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnRunGroup_Click(sender As Object, e As EventArgs) Handles btnRunGroup.Click
        'Insert a statement to run the Group statements:
        Main.Sequence.rtbSequence.SelectedText = "<RunGroup>" & txtRunGroup.Text & "</RunGroup>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnSelectGroup_Click(sender As Object, e As EventArgs) Handles btnSelectGroup.Click
        'Insert a statement to select a Statment Group:
        Main.Sequence.rtbSequence.SelectedText = "<SelectGroup>" & txtSelectGroup.Text & "</SelectGroup>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnRunGroupIf_Click(sender As Object, e As EventArgs) Handles btnRunGroupIf.Click
        'Insert a statement to run a Statment Group if the Status string matches:
        Main.Sequence.rtbSequence.SelectedText = "<RunGroupIf>" & txtRunGroupIf.Text & "</RunGroupIf>" & vbCrLf
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnCurrentDate_Click(sender As Object, e As EventArgs) Handles btnCurrentDate.Click
        'Add current date
        Main.Sequence.rtbSequence.SelectedText = Format(Now, "d-MMM-yyyy")
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnCurrentTime_Click(sender As Object, e As EventArgs) Handles btnCurrentTime.Click
        'Add current time
        Main.Sequence.rtbSequence.SelectedText = Format(Now, "H:mm:ss")
        Main.Sequence.FormatXmlText()
    End Sub

    Private Sub btnCurrentDateTime_Click(sender As Object, e As EventArgs) Handles btnCurrentDateTime.Click
        'Add current DateTime
        Main.Sequence.rtbSequence.SelectedText = Format(Now, "d-MMM-yyyy H:mm:ss")
        Main.Sequence.FormatXmlText()
    End Sub



#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Events - Events that can be triggered by this form." '--------------------------------------------------------------------------------------------------------------------------
#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class