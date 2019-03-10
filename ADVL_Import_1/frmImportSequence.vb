Public Class frmImportSequence
    'The Import Sequence form is used to display and edit an Import Sequence.

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

    'Declare Forms used by this form:
    Public WithEvents SeqStatements As frmSeqStatements

    'XDocument version:
    Dim xmlSequence As System.Xml.Linq.XDocument
    Dim xmlPathSeq As System.Xml.XPath.XPathDocument
    Dim childList As IEnumerable(Of XElement)

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
                               <RecordSequence><%= chkRecordSteps.Checked.ToString %></RecordSequence>
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
            If Settings.<FormSettings>.<RecordSequence>.Value = Nothing Then
                chkRecordSteps.Checked = False
            Else
                chkRecordSteps.Checked = Settings.<FormSettings>.<RecordSequence>.Value
            End If

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

    'Private Sub frmTemplate_Load(sender As Object, e As EventArgs) Handles Me.Load
    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings

        txtName.Text = Main.Import.ImportSequenceName
        txtDescription.Text = Main.Import.ImportSequenceDescription

        Dim xmlSeq As System.Xml.Linq.XDocument

        Main.Project.ReadXmlData(Main.Import.ImportSequenceName, xmlSeq)

        If xmlSeq Is Nothing Then
            Exit Sub
        End If

        rtbSequence.Text = xmlSeq.ToString
        FormatXmlText()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Main.RecordSequence = False
        Me.Close() 'Close the form
    End Sub

    'Private Sub frmTemplate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if form is minimised.
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------

    Private Sub btnStatements_Click(sender As Object, e As EventArgs) Handles btnStatements.Click
        'Open the Sequence Statements form:
        If IsNothing(SeqStatements) Then
            SeqStatements = New frmSeqStatements
            SeqStatements.Show()
        Else
            SeqStatements.Show()
        End If
    End Sub

    Private Sub SeqStatements_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SeqStatements.FormClosed
        SeqStatements = Nothing
    End Sub

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Public Sub FormatXmlText()
        'Format the XML text in rtbSequence rich text box control:

        Dim Posn As Integer
        Dim SelLen As Integer
        Posn = rtbSequence.SelectionStart
        SelLen = rtbSequence.SelectionLength

        'Set colour of the start tag names (for a tag without attributes):
        Dim RegExString2 As String = "(?<=<)([A-Za-z\d]+)(?=>)"
        Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExString2)
        Dim myMatches2 As System.Text.RegularExpressions.MatchCollection
        myMatches2 = myRegEx2.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Crimson
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of the start tag names (for a tag with attributes):
        Dim RegExString2b As String = "(?<=<)([A-Za-z\d]+)(?= [A-Za-z\d]+=""[A-Za-z\d ]+"">)"
        Dim myRegEx2b As New System.Text.RegularExpressions.Regex(RegExString2b)
        Dim myMatches2b As System.Text.RegularExpressions.MatchCollection
        myMatches2b = myRegEx2b.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2b
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Crimson
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of the attribute names (for a tag with attributes):
        Dim RegExString2c As String = "(?<=<[A-Za-z\d]+ )([A-Za-z\d]+)(?==""[A-Za-z\d ]+"">)"
        Dim myRegEx2c As New System.Text.RegularExpressions.Regex(RegExString2c)
        Dim myMatches2c As System.Text.RegularExpressions.MatchCollection
        myMatches2c = myRegEx2c.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2c
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Crimson
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of the attribute values (for a tag with attributes):
        Dim RegExString2d As String = "(?<=<[A-Za-z\d]+ [A-Za-z\d]+="")([A-Za-z\d ]+)(?="">)"
        Dim myRegEx2d As New System.Text.RegularExpressions.Regex(RegExString2d)
        Dim myMatches2d As System.Text.RegularExpressions.MatchCollection
        myMatches2d = myRegEx2d.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2d
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Black
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Bold)
        Next

        'Set colour of the end tag names:
        Dim RegExString3 As String = "(?<=</)([A-Za-z\d]+)(?=>)"
        Dim myRegEx3 As New System.Text.RegularExpressions.Regex(RegExString3)
        Dim myMatches3 As System.Text.RegularExpressions.MatchCollection
        myMatches3 = myRegEx3.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches3
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Crimson
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of comments:
        Dim RegExString4 As String = "(?<=<!--)([A-Za-z\d \.,:]+)(?=-->)"
        Dim myRegEx4 As New System.Text.RegularExpressions.Regex(RegExString4)
        Dim myMatches4 As System.Text.RegularExpressions.MatchCollection
        myMatches4 = myRegEx4.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches4
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Gray
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of "<" and ">" characters to blue
        Dim RegExString As String = "</|<!--|-->|<|>"
        Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExString)
        Dim myMatches As System.Text.RegularExpressions.MatchCollection
        myMatches = myRegEx.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Blue
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set tag contents (between ">" and "</") to black, bold
        Dim RegExString5 As String = "(?<=>)([A-Za-z\d \.,\:\-\&\;\\_]+)(?=</)"
        Dim myRegEx5 As New System.Text.RegularExpressions.Regex(RegExString5)
        Dim myMatches5 As System.Text.RegularExpressions.MatchCollection
        myMatches5 = myRegEx5.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches5
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectionColor = Color.Black
            Dim f As Font = rtbSequence.SelectionFont
            rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Bold)
        Next

        'Remove blank lines
        Dim RegExString6 As String = "(?<=\n)\ *\n"
        Dim myRegEx6 As New System.Text.RegularExpressions.Regex(RegExString6)
        Dim myMatches6 As System.Text.RegularExpressions.MatchCollection
        myMatches6 = myRegEx6.Matches(rtbSequence.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches6
            rtbSequence.Select(aMatch.Index, aMatch.Length)
            rtbSequence.SelectedText = ""
        Next

        rtbSequence.SelectionStart = Posn
        rtbSequence.SelectionLength = SelLen

    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        'Open a processing sequence file:

        Dim SelectedFileName As String = ""

        SelectedFileName = Main.Project.SelectDataFile("Sequence", "Sequence")
        Main.Message.Add("Selected Import Sequence: " & SelectedFileName & vbCrLf)

        txtName.Text = SelectedFileName

        Dim xmlSeq As System.Xml.Linq.XDocument

        Main.Project.ReadXmlData(SelectedFileName, xmlSeq)

        If xmlSeq Is Nothing Then
            Exit Sub
        End If

        rtbSequence.Text = xmlSeq.ToString
        FormatXmlText()

        Main.Import.ImportSequenceName = SelectedFileName
        Main.Import.ImportSequenceDescription = xmlSeq.<ProcessingSequence>.<Description>.Value
        txtDescription.Text = Main.Import.ImportSequenceDescription

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        'Create a new processing sequence.

        'NOTE: CHECK IF THIS CODE CAN BE UPDATED TO USE A LOCAL VARIABLE INSTEAD OF xmlSequence!!!

        If rtbSequence.Text = "" Then
            'Current Processing Sequence is blank. OK to create a new one.
        Else
            'Current Processing Sequence contains data.
            'Check if it is OK to overwrite:
            If MessageBox.Show("Overwrite existing file?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.No Then
                Exit Sub
            End If
        End If

        xmlSequence = <?xml version="1.0" encoding="utf-8"?>
                      <!---->
                      <!--Processing Sequence generated by the Signalworks ADVL_Import_1 application.-->
                      <ProcessingSequence>
                          <CreationDate><%= Format(Now, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                          <Description><%= Trim(txtDescription.Text) %></Description>
                          <!---->
                          <Process>
                              <!--Insert processing sequence code here:-->
                          </Process>
                      </ProcessingSequence>

        rtbSequence.Text = xmlSequence.Document.ToString
        FormatXmlText()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Save the Import Sequence in a file:

        Try
            Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(rtbSequence.Text)

            Dim SequenceFileName As String = ""

            If Trim(txtName.Text).EndsWith(".Sequence") Then
                SequenceFileName = Trim(txtName.Text)
            Else
                SequenceFileName = Trim(txtName.Text) & ".Sequence"
            End If
            Main.Project.SaveXmlData(SequenceFileName, xmlSeq)
            Main.Message.Add("Import Sequence saved OK")
        Catch ex As Exception
            Main.Message.AddWarning(ex.Message)
        End Try

    End Sub

    Private Sub chkRecordSteps_CheckedChanged(sender As Object, e As EventArgs) Handles chkRecordSteps.CheckedChanged
        If chkRecordSteps.Checked Then
            Main.RecordSequence = True
        Else
            Main.RecordSequence = False
        End If
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click

        Dim XDoc As New System.Xml.XmlDocument
        XDoc.LoadXml(rtbSequence.Text)

        Dim SequenceStatus As String

        Main.Import.RunXSequence(XDoc)

    End Sub

    Private Sub btnStatusCheck_Click(sender As Object, e As EventArgs) Handles btnStatusCheck.Click

    End Sub





#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class