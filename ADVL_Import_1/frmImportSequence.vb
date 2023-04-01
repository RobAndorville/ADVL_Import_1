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
            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
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


        XmlHtmDisplay1.AllowDrop = True

        XmlHtmDisplay1.WordWrap = False

        XmlHtmDisplay1.Settings.ClearAllTextTypes()
        'Default message display settings:
        XmlHtmDisplay1.Settings.AddNewTextType("Warning")
        XmlHtmDisplay1.Settings.TextType("Warning").FontName = "Arial"
        XmlHtmDisplay1.Settings.TextType("Warning").Bold = True
        XmlHtmDisplay1.Settings.TextType("Warning").Color = Color.Red
        XmlHtmDisplay1.Settings.TextType("Warning").PointSize = 12

        XmlHtmDisplay1.Settings.AddNewTextType("Default")
        XmlHtmDisplay1.Settings.TextType("Default").FontName = "Arial"
        XmlHtmDisplay1.Settings.TextType("Default").Bold = False
        XmlHtmDisplay1.Settings.TextType("Default").Color = Color.Black
        XmlHtmDisplay1.Settings.TextType("Default").PointSize = 10

        XmlHtmDisplay1.Settings.XValue.Bold = True

        XmlHtmDisplay1.Settings.UpdateFontIndexes()
        XmlHtmDisplay1.Settings.UpdateColorIndexes()




        Dim xmlSeq As System.Xml.Linq.XDocument

        Main.Project.ReadXmlData(Main.Import.ImportSequenceName, xmlSeq)

        If xmlSeq Is Nothing Then
            Exit Sub
        End If

        'rtbSequence.Text = xmlSeq.ToString
        'rtbSequence.Text = xmlSeq.ToString
        'FormatXmlText()
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlDoc, True)
        XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlSeq, True)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Main.RecordSequence = False

        'Close the Statements form if it is open:
        If IsNothing(SeqStatements) Then
            'The form is already closed.
        Else
            SeqStatements.Close() 'Close the form.
        End If

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

    'Public Sub FormatXmlText()
    '    'Format the XML text in rtbSequence rich text box control:

    '    Dim Posn As Integer
    '    Dim SelLen As Integer
    '    Posn = rtbSequence.SelectionStart
    '    SelLen = rtbSequence.SelectionLength

    '    'Set colour of the start tag names (for a tag without attributes):
    '    Dim RegExString2 As String = "(?<=<)([A-Za-z\d]+)(?=>)"
    '    Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExString2)
    '    Dim myMatches2 As System.Text.RegularExpressions.MatchCollection
    '    myMatches2 = myRegEx2.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches2
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Crimson
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
    '    Next

    '    'Set colour of the start tag names (for a tag with attributes):
    '    Dim RegExString2b As String = "(?<=<)([A-Za-z\d]+)(?= [A-Za-z\d]+=""[A-Za-z\d ]+"">)"
    '    Dim myRegEx2b As New System.Text.RegularExpressions.Regex(RegExString2b)
    '    Dim myMatches2b As System.Text.RegularExpressions.MatchCollection
    '    myMatches2b = myRegEx2b.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches2b
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Crimson
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
    '    Next

    '    'Set colour of the attribute names (for a tag with attributes):
    '    Dim RegExString2c As String = "(?<=<[A-Za-z\d]+ )([A-Za-z\d]+)(?==""[A-Za-z\d ]+"">)"
    '    Dim myRegEx2c As New System.Text.RegularExpressions.Regex(RegExString2c)
    '    Dim myMatches2c As System.Text.RegularExpressions.MatchCollection
    '    myMatches2c = myRegEx2c.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches2c
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Crimson
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
    '    Next

    '    'Set colour of the attribute values (for a tag with attributes):
    '    Dim RegExString2d As String = "(?<=<[A-Za-z\d]+ [A-Za-z\d]+="")([A-Za-z\d ]+)(?="">)"
    '    Dim myRegEx2d As New System.Text.RegularExpressions.Regex(RegExString2d)
    '    Dim myMatches2d As System.Text.RegularExpressions.MatchCollection
    '    myMatches2d = myRegEx2d.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches2d
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Black
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Bold)
    '    Next

    '    'Set colour of the end tag names:
    '    Dim RegExString3 As String = "(?<=</)([A-Za-z\d]+)(?=>)"
    '    Dim myRegEx3 As New System.Text.RegularExpressions.Regex(RegExString3)
    '    Dim myMatches3 As System.Text.RegularExpressions.MatchCollection
    '    myMatches3 = myRegEx3.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches3
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Crimson
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
    '    Next

    '    'Set colour of comments:
    '    Dim RegExString4 As String = "(?<=<!--)([A-Za-z\d \.,:]+)(?=-->)"
    '    Dim myRegEx4 As New System.Text.RegularExpressions.Regex(RegExString4)
    '    Dim myMatches4 As System.Text.RegularExpressions.MatchCollection
    '    myMatches4 = myRegEx4.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches4
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Gray
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
    '    Next

    '    'Set colour of "<" and ">" characters to blue
    '    Dim RegExString As String = "</|<!--|-->|<|>"
    '    Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExString)
    '    Dim myMatches As System.Text.RegularExpressions.MatchCollection
    '    myMatches = myRegEx.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Blue
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
    '    Next

    '    'Set tag contents (between ">" and "</") to black, bold
    '    Dim RegExString5 As String = "(?<=>)([A-Za-z\d \.,\:\-\&\;\\_]+)(?=</)"
    '    Dim myRegEx5 As New System.Text.RegularExpressions.Regex(RegExString5)
    '    Dim myMatches5 As System.Text.RegularExpressions.MatchCollection
    '    myMatches5 = myRegEx5.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches5
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectionColor = Color.Black
    '        Dim f As Font = rtbSequence.SelectionFont
    '        rtbSequence.SelectionFont = New Font(f.Name, f.Size, FontStyle.Bold)
    '    Next

    '    'Remove blank lines
    '    Dim RegExString6 As String = "(?<=\n)\ *\n"
    '    Dim myRegEx6 As New System.Text.RegularExpressions.Regex(RegExString6)
    '    Dim myMatches6 As System.Text.RegularExpressions.MatchCollection
    '    myMatches6 = myRegEx6.Matches(rtbSequence.Text)
    '    For Each aMatch As System.Text.RegularExpressions.Match In myMatches6
    '        rtbSequence.Select(aMatch.Index, aMatch.Length)
    '        rtbSequence.SelectedText = ""
    '    Next

    '    rtbSequence.SelectionStart = Posn
    '    rtbSequence.SelectionLength = SelLen

    'End Sub



    Public Sub FormatXmlText()
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(XmlHtmDisplay1.Text, True)
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(XmlHtmDisplay1.FixXmlText(XmlHtmDisplay1.Text), True)
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(FixXmlText(XmlHtmDisplay1.Text), True)

        'Main.Message.Add(FixXmlText(XmlHtmDisplay1.Text) & vbCrLf)
        'Main.Message.Add(XmlHtmDisplay1.Text & vbCrLf)
        'Main.Message.Add(FixXmlText(XmlHtmDisplay1.Text & vbCrLf) & vbCrLf)

        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(FixXmlText(XmlHtmDisplay1.Text & vbCrLf), True)

        XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(XmlHtmDisplay1.FixXmlText(XmlHtmDisplay1.Text & vbCrLf), True)


    End Sub


    Public Function FixXmlText(XmlText As String) As String
        'Fix an XML string so that it can be loaded correcly using the LoadXml method of a System.Xml.XmlDocument
        'Replace "<" in an element value with "&lt;"
        'Replace ">" in an element value with "&gt;"

        'XML Terminology:
        'XML declaration <?xml version="1.0" encoding="UTF-8"?>
        'Comments begin with <!-- and end with -->.
        'Start-tag <Element>
        'End-tag </Element>
        'Empty-element tag (Element />
        'Content  The characters between the start-tag and end-tag, if any, are the element's content, and may contain markup, including other elements, which are called child elements.
        'Predefined entities:
        '&lt; represents "<";
        '&gt; represents ">";
        '&amp; represents "&";
        '&apos; represents "'";
        '&quot; represents '"'.

        Dim FixedXmlText As New System.Text.StringBuilder
        Dim StartPos As Integer
        Dim EndPos As Integer
        Dim ScanPos As Integer = 0
        Dim LastPos As Integer = XmlText.Length

        If XmlText.Trim.StartsWith("<?xml") Then
            StartPos = XmlText.IndexOf("<?xml")
            EndPos = XmlText.IndexOf("?>", StartPos)
            FixedXmlText.Append(XmlText.Substring(StartPos, EndPos - StartPos + 2))
            ScanPos = EndPos + 2
        End If
        FixedXmlText.Append(ProcessContent(XmlText, ScanPos, LastPos))
        Return FixedXmlText.ToString
    End Function

    Private Function ProcessContent(ByRef XmlText As String, FromIndex As Integer, ToIndex As Integer) As String
        'Process the XML content in the XmlText string between FromIndex and ToIndex.
        'THIS VERSION SEARCHES FOR MATCHING End-Tags
        '
        'Content alternatives:
        'Content only
        '<!---->                                        One or more comments
        '<Element />                                    One or more empty element tags
        '<Element></Element>                            One or more empty elements
        '<Element>Content</Element>                     One or more elements containing content
        '<Element>                                      One or more elements containing child elements
        '  <ChildElement>ChildContent</ChildElement>
        '</Element>

        Dim StartScan As Integer = FromIndex 'The start of the current content scan
        Dim ScanIndex As Integer = FromIndex 'The current scan position
        Dim LtCharIndex As Integer 'The index position of the next < character
        Dim GtCharIndex As Integer 'The index position of the next > character
        Dim FixedXmlText As New System.Text.StringBuilder 'This is used to build the fixed XML text for the content if it contains XML tags
        Dim StartTagText As String = "" 'The text of a found Start-tag. The text may include attributes following the name.
        Dim EndNameIndex As Integer 'The index position of the end of the StartTagName. If the StartTagText contains attributes, the StartTagName will be followed by a space then the attributes.
        Dim StartTagName As String = "" 'The name of a found Start-tag
        Dim EndTagIndex As Integer 'The index of an End-tag
        Dim StartTagCount As Integer = 1 'The nesting level of the StartTag
        Dim EndTagCount As Integer = 1 'The nesting level of the EndTag
        Dim StartSearch As Integer 'StartSearch index used for counting other Start-Tags named StartTagName
        Dim NextSearch As Integer 'Search for the next Start-Tag named StartTagName
        Dim SearchIndex As Integer
        Dim Match As Boolean
        Dim TagLevelMatch As Boolean 'If True, the Start-Tag and End-Tag have matching levels.
        Dim SearchEndTagFrom As Integer 'The index to start the End-tag search from
        Dim ElementFound As Boolean = False 'True if an Element or a Comment was found.
        Dim EndSearch As Boolean 'If True, End the Search to find the End-Tag

        'While ScanIndex <= ToIndex
        While ScanIndex < ToIndex
            'Find the first pair of < > characters
            LtCharIndex = XmlText.IndexOf("<", ScanIndex) 'Find the start of the next Element
            If LtCharIndex = -1 Then '< char not found
                If ToIndex - ScanIndex = 2 Then
                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                        Exit While
                    End If
                End If
                'The characters between FromIndex and ToIndex are Content
                'NOTE: StartScan and FromIndex should be the same here: StartScan only advances if the Content contains one or more comments or elements.
                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                FixedXmlText.Append(Content)
                ScanIndex = ToIndex + 1
            ElseIf LtCharIndex >= ToIndex Then
                'Check if the remaining characters are CrLf:
                If ToIndex - ScanIndex = 2 Then
                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                        Exit While
                    End If
                End If
                'Check if the remaining characters are blank:
                If XmlText.Substring(ScanIndex, ToIndex - ScanIndex).Trim = "" Then
                    Exit While
                End If
                'Check if the remaining characters with blanks removed are CrLf:
                'Check if the remaining characters are blank:
                If XmlText.Substring(ScanIndex, ToIndex - ScanIndex).Trim = vbCrLf Then
                    Exit While
                End If
                'The characters between FromIndex and ToIndex are Content
                'NOTE: StartScan and FromIndex should be the same here
                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                FixedXmlText.Append(Content)
                ScanIndex = ToIndex + 1
            Else
                'The < character is within the Content range
                'Search for a > character
                GtCharIndex = XmlText.IndexOf(">", LtCharIndex + 1)
                If GtCharIndex = -1 Then '> char not found
                    If ToIndex - ScanIndex = 2 Then
                        If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                            Exit While
                        End If
                    End If
                    'The characters between FromIndex and ToIndex are Content
                    'NOTE: StartScan and FromIndex should be the same here
                    Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                    FixedXmlText.Append(Content)
                    ScanIndex = ToIndex + 1
                ElseIf GtCharIndex > ToIndex Then
                    If ToIndex - ScanIndex = 2 Then
                        If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                            Exit While
                        End If
                    End If
                    'The characters between FromIndex and ToIndex are Content
                    'NOTE: StartScan and FromIndex should be the same here
                    Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                    FixedXmlText.Append(Content)
                    ScanIndex = ToIndex + 1
                Else
                    'A start-tagChar and end-tagChar <> pair has been found.
                    'The <> characters will contain a comment, an element name or be part of the element content.
                    If XmlText.Substring(LtCharIndex, 4) = "<!--" Then 'This is the start of a comment
                        If XmlText.Substring(GtCharIndex - 2, 3) = "-->" Then 'This is the end of a comment ---------------------  <--Comment-->  ---------------------------------
                            FixedXmlText.Append(XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1) & vbCrLf) 'Add the Comment to the Fixed XML Text
                            ScanIndex = GtCharIndex + 1
                            StartScan = GtCharIndex + 1
                        Else
                            'This is not a comment.
                            'The whole content must be the content of a single element
                            'The characters between FromIndex and ToIndex are Content
                            'NOTE: StartScan and FromIndex should be the same here
                            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                            FixedXmlText.Append(Content)
                            ScanIndex = ToIndex + 1
                        End If
                    Else 'This is a start-tag, empty element or content of a single element
                        If XmlText.Chars(GtCharIndex - 1) = "/" Then 'This is an empty element
                            FixedXmlText.Append(XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1))
                            ScanIndex = GtCharIndex + 1
                            StartScan = GtCharIndex + 1
                        Else
                            StartTagCount = 1
                            EndTagCount = 0
                            StartSearch = GtCharIndex + 1
                            EndSearch = False
                            While StartTagCount > EndTagCount And EndSearch = False
                                'Continue searching for StartTag-EndTag tag pairs with the name StartTagName until matching tags are found (StartTagCount = EndTagCount).
                                StartTagText = XmlText.Substring(LtCharIndex + 1, GtCharIndex - LtCharIndex - 1) 'This is the text of the Start-tag
                                EndNameIndex = StartTagText.IndexOf(" ")
                                If EndNameIndex = -1 Then 'There is no space in StartTagText so it contains no attributes.
                                    StartTagName = StartTagText
                                Else
                                    StartTagName = StartTagText.Substring(0, EndNameIndex)
                                End If
                                'Find the matching End-tag - The matching End-tag must have a matching TagName and a matching level.
                                TagLevelMatch = False
                                SearchEndTagFrom = GtCharIndex + 1
                                While TagLevelMatch = False
                                    EndTagIndex = XmlText.IndexOf("</" & StartTagName & ">", SearchEndTagFrom)
                                    If EndTagIndex = -1 Then 'There is no matching End-tag
                                        'This is not an element.
                                        'The whole content must be the content of a single element
                                        'The characters between FromIndex and ToIndex are Content
                                        'NOTE: StartScan and FromIndex should be the same here
                                        Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                                        FixedXmlText.Append(Content)
                                        ScanIndex = ToIndex + 1
                                        EndSearch = True 'End the search for the End-Tag
                                        Exit While 'There is no matching End-tag!
                                    ElseIf EndTagIndex > ToIndex - StartTagName.Length - 1 Then 'The matching tag is outside of the Content
                                        'This is not an element.
                                        'The whole content must be the content of a single element
                                        'The characters between FromIndex and ToIndex are Content
                                        'NOTE: StartScan and FromIndex should be the same here
                                        Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                                        FixedXmlText.Append(Content)
                                        ScanIndex = ToIndex + 1
                                        EndSearch = True 'End the search for the End-Tag
                                        Exit While 'There is no matching End-tag withing the Content range.
                                    Else 'Matching End-tag found at EndTagIndex. 
                                        EndTagCount += 1 'Increment the End Tag Count
                                        'Search for any other Start-Tags named StartTagName between LtCharIndex and EndTagIndex
                                        Match = True
                                        NextSearch = StartSearch 'Search for <StartTagName> (without attributes)
                                        While Match = True 'Search for Start-Tags of the form: <StartTagName>
                                            SearchIndex = XmlText.IndexOf("<" & StartTagName & ">", NextSearch, EndTagIndex - NextSearch)
                                            If SearchIndex = -1 Then
                                                Match = False
                                            Else
                                                NextSearch = SearchIndex + StartTagName.Length
                                                StartTagCount += 1
                                            End If
                                        End While
                                        Match = True
                                        NextSearch = StartSearch 'Set NextSearch back to StartSearch to search the same chars for <StartTagName ...(with attributes)
                                        While Match = True 'Search for Start-Tags of the form: <StartTagName ...> (Start-Tag with attributes)
                                            SearchIndex = XmlText.IndexOf("<" & StartTagName & " ", NextSearch, EndTagIndex - NextSearch)
                                            If SearchIndex = -1 Then
                                                Match = False
                                            Else
                                                NextSearch = SearchIndex + StartTagName.Length
                                                StartTagCount += 1
                                            End If
                                        End While
                                        StartSearch = EndTagIndex + 1 'All Start-Tags named StartTagName have been found to EndTagIndex : Update StartSearch - If more searches are needed, they will start from here.
                                        If StartTagCount = EndTagCount Then
                                            TagLevelMatch = True
                                            FixedXmlText.Append("<" & StartTagText & ">" & ProcessContent(XmlText, GtCharIndex + 1, EndTagIndex) & "</" & StartTagName & ">" & vbCrLf)
                                            ScanIndex = EndTagIndex + StartTagName.Length + 3
                                            ElementFound = True
                                        Else
                                            SearchEndTagFrom = EndTagIndex + StartTagName.Length + 3
                                        End If
                                    End If
                                End While
                            End While
                        End If
                    End If
                End If
            End If
        End While
        If ElementFound Then
            Return vbCrLf & FixedXmlText.ToString
        Else
            Return FixedXmlText.ToString
        End If
    End Function

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        'Open a processing sequence file:

        Dim SelectedFileName As String = ""

        SelectedFileName = Main.Project.SelectDataFile("Sequence", "Sequence")
        Main.Message.Add("Selected Import Sequence: " & SelectedFileName & vbCrLf)

        txtName.Text = SelectedFileName

        Dim xmlSeq As System.Xml.Linq.XDocument
        'Dim xmlDoc As New System.Xml.XmlDocument

        Main.Project.ReadXmlData(SelectedFileName, xmlSeq)
        'Main.Project.ReadXmlDocData(SelectedFileName, xmlDoc)

        If xmlSeq Is Nothing Then
            'If xmlDoc Is Nothing Then
            Exit Sub
        End If

        'rtbSequence.Text = xmlSeq.ToString
        'rtbSequence.Text = xmlSeq.ToString
        'FormatXmlText()
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlDoc, True)
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlDoc, True)

        'Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" '& vbCrLf

        XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlSeq.ToString, True)
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(XmlHeader & xmlSeq.ToString, True)

        Main.Import.ImportSequenceName = SelectedFileName
        Main.Import.ImportSequenceDescription = xmlSeq.<ProcessingSequence>.<Description>.Value
        'Main.Import.ImportSequenceDescription = xmlDoc.<ProcessingSequence>.<Description>.Value
        txtDescription.Text = Main.Import.ImportSequenceDescription

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        'Create a new processing sequence.

        'NOTE: CHECK IF THIS CODE CAN BE UPDATED TO USE A LOCAL VARIABLE INSTEAD OF xmlSequence!!!

        'If rtbSequence.Text = "" Then
        If XmlHtmDisplay1.Text = "" Then
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

        'rtbSequence.Text = xmlSequence.Document.ToString
        'XmlHtmDisplay1.Text = xmlSequence.Document.ToString

        'FormatXmlText()
        'XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlDoc, True)
        XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlSequence, True)

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Save the Import Sequence in a file:

        Try
            'Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(rtbSequence.Text)
            'Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(XmlHtmDisplay1.Text)
            Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(FixXmlText(XmlHtmDisplay1.Text & vbCrLf))

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
        'XDoc.LoadXml(rtbSequence.Text)
        XDoc.LoadXml(XmlHtmDisplay1.Text)

        'Dim SequenceStatus As String

        Main.Import.RunXSequence(XDoc)

    End Sub



    Private Sub btnCancelImport_Click(sender As Object, e As EventArgs) Handles btnCancelImport.Click
        'Cancel the Import process.
        Main.Import.CancelImport = True
        Main.Message.Add("CancelImport = " & Main.Import.CancelImport.ToString & vbCrLf)
    End Sub

    Private Sub SeqStatements_AddLine(NewLine As String) Handles SeqStatements.AddLine
        'A New Line has been received from the Sequence Statements form.
        'This will be added to the sequence.
        XmlHtmDisplay1.SelectedText = NewLine
        FormatXmlText()
    End Sub

    Private Sub btnFormat_Click(sender As Object, e As EventArgs) Handles btnFormat.Click
        FormatXmlText()
    End Sub

    Private Sub XmlHtmDisplay1_Message(Msg As String) Handles XmlHtmDisplay1.Message
        Main.Message.Add(Msg)
    End Sub

    Private Sub XmlHtmDisplay1_ErrorMessage(Msg As String) Handles XmlHtmDisplay1.ErrorMessage
        Main.Message.AddWarning(Msg)
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class