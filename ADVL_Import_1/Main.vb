'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
''WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Class Main
    'The ADVL_Import_1 is a general purpose data import application.

#Region " Coding Notes - Notes on the code used in this class." '------------------------------------------------------------------------------------------------------------------------------

    'ADD THE SYSTEM UTILITIES REFERENCE: ==========================================================================================
    'The following references are required by this software: 
    'Project \ Add Reference... \ ADVL_Utilities_Library_1.dll
    'The Utilities Library is used for Project Management, Archive file management, running XSequence files and running XMessage files.
    'If there are problems with a reference, try deleting it from the references list and adding it again.

    'ADD THE SERVICE REFERENCE: ===================================================================================================
    'A service reference to the Message Service must be added to the source code before this service can be used.
    'This is used to connect to the Application Network.

    'Adding the service reference to a project that includes the WcfMsgServiceLib project: -----------------------------------------
    'Project \ Add Service Reference
    'Press the Discover button.
    'Expand the items in the Services window and select IMsgService.
    'Press OK.
    '------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------
    'Adding the service reference to other projects that dont include the WcfMsgServiceLib project: -------------------------------
    'Run the ADVL_Application_Network_1 application to start the Application Network message service.
    'In Microsoft Visual Studio select: Project \ Add Service Reference
    'Enter the address: http://localhost:8733/ADVLService
    'Press the Go button.
    'MsgService is found.
    'Press OK to add ServiceReference1 to the project.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE MsgServiceCallback CODE: =============================================================================================
    'This is used to connect to the Application Network.
    'In Microsoft Visual Studio select: Project \ Add Class
    'MsgServiceCallback.vb
    'Add the following code to the class:
    'Imports System.ServiceModel
    'Public Class MsgServiceCallback
    '    Implements ServiceReference1.IMsgServiceCallback
    '    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
    '        'A message has been received.
    '        'Set the InstrReceived property value to the message (usually in XMessage format). This will also apply the instructions in the XMessage.
    '        Main.InstrReceived = message
    '    End Sub
    'End Class
    '------------------------------------------------------------------------------------------------------------------------------

#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables and class objects used in this form and this application." '-------------------------------------------------------------------------------------------------

    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.

    'Declare Forms used by the application:
    'Public WithEvents TemplateForm As frmTemplate
    Public WithEvents ShowTextFile As frmShowTextFile
    Public WithEvents Sequence As frmImportSequence
    Public WithEvents ClipboardWindow As frmClipboardWindow

    'Declare objects used to connect to the Application Network:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppName As String = "" 'The name of the client requesting service
    Dim ClientAppLocn As String = "" 'The location in the Client application requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocn As XElement 'The location part of the xmessage.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    Public WithEvents Import As New Import 'The class used to import data.

    Dim RegExIndex As Integer 'Tracks the current RegEx index (The first RegEx has RegExIndex = 0)

    'Database Locations varlables:
    Structure strucGridProp 'Regular Expression entry structure
        Dim Mult As Single
        Dim Status As String
    End Structure

    Dim GridProp(,) As strucGridProp 'Array used to store properties for corresponding cells in DataGridView1

    Dim DbDestIndex As Integer

    Dim ListChanged As Boolean = False

    Dim ModifyValueType As Import.ModifyValuesTypes


#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '------------------------------------------------------------------------------------------------------------

    Private _connectionHashcode As Integer 'The Application Network connection hashcode. This is used to identify a connection in the Application Netowrk when reconnecting.
    Property ConnectionHashcode As Integer
        Get
            Return _connectionHashcode
        End Get
        Set(value As Integer)
            _connectionHashcode = value
        End Set
    End Property


    Private _connectedToAppNet As Boolean = False  'True if the application is connected to the Application Network.
    Property ConnectedToAppnet As Boolean
        Get
            Return _connectedToAppNet
        End Get
        Set(value As Boolean)
            _connectedToAppNet = value
        End Set
    End Property

    Private _instrReceived As String = "" 'Contains Instructions received from the Application Network message service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value

                'Add the message to the XMessages window:
                Message.Color = Color.Blue
                Message.FontStyle = FontStyle.Bold
                Message.XAdd("Message received: " & vbCrLf)
                Message.SetNormalStyle()
                Message.XAdd(_instrReceived & vbCrLf & vbCrLf)

                If _instrReceived.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
                    Try
                        'Inititalise the reply message:
                        Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                        MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                        xmessage = New XElement("XMsg")
                        xlocn = New XElement("Main") 'Initially set the location in the Client App to Main.

                        'Run the received message:
                        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                        XDoc.LoadXml(XmlHeader & vbCrLf & _instrReceived)
                        XMsg.Run(XDoc, Status)
                    Catch ex As Exception
                        Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
                    End Try

                    'XMessage has been run.
                    'Reply to this message:
                    'Add the message reply to the XMessages window:
                    'Complete the MessageXDoc:
                    xmessage.Add(xlocn) 'Add the location part of the reply message. (This also contains the reply instructions.)
                    MessageXDoc.Add(xmessage)
                    MessageText = MessageXDoc.ToString

                    If ClientAppName = "" Then
                        'No client to send a message to!
                    Else
                        If MessageText = "" Then
                            'No message to send!
                        Else
                            Message.Color = Color.Red
                            Message.FontStyle = FontStyle.Bold
                            Message.XAdd("Message sent to " & ClientAppName & ":" & vbCrLf)
                            Message.SetNormalStyle()
                            Message.XAdd(MessageText & vbCrLf & vbCrLf)
                            'SendMessage sends the contents of MessageText to MessageDest.
                            SendMessage() 'This subroutine triggers the timer to send the message after a short delay.
                        End If
                    End If
                Else

                End If
            End If

        End Set
    End Property


    Private _recordSequence As Boolean 'If True then processing sequences manually applied using the forms will be recorded in the processing sequence.
    Property RecordSequence As Boolean
        Get
            Return _recordSequence
        End Get
        Set(value As Boolean)
            _recordSequence = value
        End Set
    End Property


#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                               <!---->
                               <!--Match Text - RegEx Grid Settings-->
                               <RegExNameColumnWidth><%= DataGridView1.Columns(0).Width %></RegExNameColumnWidth>
                               <RegExDescrColumnWidth><%= DataGridView1.Columns(1).Width %></RegExDescrColumnWidth>
                               <!---->
                               <!--Database Locations - Grid Settings-->
                               <RegExVariableColumnWidth><%= DataGridView2.Columns(0).Width %></RegExVariableColumnWidth>
                               <VariableTypeColumnWidth><%= DataGridView2.Columns(1).Width %></VariableTypeColumnWidth>
                               <DestTableColumnWidth><%= DataGridView2.Columns(2).Width %></DestTableColumnWidth>
                               <DestFieldColumnWidth><%= DataGridView2.Columns(3).Width %></DestFieldColumnWidth>
                               <StatusFieldColumnWidth><%= DataGridView2.Columns(4).Width %></StatusFieldColumnWidth>
                               <SplitterDistance><%= SplitContainer2.SplitterDistance %></SplitterDistance>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value
            'Restore Match Text Tab Settings:
            If Settings.<FormSettings>.<RegExNameColumnWidth>.Value <> Nothing Then DataGridView1.Columns(0).Width = Settings.<FormSettings>.<RegExNameColumnWidth>.Value
            If Settings.<FormSettings>.<RegExDescrColumnWidth>.Value <> Nothing Then DataGridView1.Columns(1).Width = Settings.<FormSettings>.<RegExDescrColumnWidth>.Value
            'Restore Database Locations Tab Settings:
            If Settings.<FormSettings>.<RegExVariableColumnWidth>.Value <> Nothing Then DataGridView2.Columns(0).Width = Settings.<FormSettings>.<RegExVariableColumnWidth>.Value
            If Settings.<FormSettings>.<VariableTypeColumnWidth>.Value <> Nothing Then DataGridView2.Columns(1).Width = Settings.<FormSettings>.<VariableTypeColumnWidth>.Value
            If Settings.<FormSettings>.<DestTableColumnWidth>.Value <> Nothing Then DataGridView2.Columns(2).Width = Settings.<FormSettings>.<DestTableColumnWidth>.Value
            If Settings.<FormSettings>.<DestFieldColumnWidth>.Value <> Nothing Then DataGridView2.Columns(3).Width = Settings.<FormSettings>.<DestFieldColumnWidth>.Value
            If Settings.<FormSettings>.<StatusFieldColumnWidth>.Value <> Nothing Then DataGridView2.Columns(4).Width = Settings.<FormSettings>.<StatusFieldColumnWidth>.Value
            If Settings.<FormSettings>.<SplitterDistance>.Value <> Nothing Then SplitContainer2.SplitterDistance = Settings.<FormSettings>.<SplitterDistance>.Value

            'If Settings.<FormSettings>.<InputFileDirectory>.Value = Nothing Then
            'Else
            '    txtInputFileDir.Text = Settings.<FormSettings>.<InputFileDirectory>.Value
            '    FillLstTextFiles()
            'End If

            'Add code to read other saved setting here:

        End If
    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        ApplicationInfo.Name = "ADVL_Import_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath

        ApplicationInfo.Description = "General purpose data import application."
        ApplicationInfo.CreationDate = "21-Aug-2016 12:00:00"

        'Author -----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Author.Name = "Signalworks Pty Ltd"
        ApplicationInfo.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        ApplicationInfo.Author.Contact = "http://www.andorville.com.au/"

        'File Associations: -----------------------------------------------------------------------------------------------
        'Add any file associations here.
        'The file extension and a description of files that can be opened by this application are specified.
        'The example below specifies a coordinate system parameter file type with the file extension .ADVLCoord.
        'Dim Assn1 As New ADVL_System_Utilities.FileAssociation
        'Assn1.Extension = "ADVLCoord"
        'Assn1.Description = "Andorville (TM) software coordinate system parameter file"
        'ApplicationInfo.FileAssociations.Add(Assn1)

        'Version ----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        'Copyright --------------------------------------------------------------------------------------------------------
        'Add your copyright information here.
        ApplicationInfo.Copyright.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.Copyright.PublicationYear = "2016"

        'Trademarks -------------------------------------------------------------------------------------------------------
        'Add your trademark information here.
        Dim Trademark1 As New ADVL_Utilities_Library_1.Trademark
        Trademark1.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark1.Text = "Andorville"
        Trademark1.Registered = False
        Trademark1.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark1)
        Dim Trademark2 As New ADVL_Utilities_Library_1.Trademark
        Trademark2.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark2.Text = "AL-H7"
        Trademark2.Registered = False
        Trademark2.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark2)

        'License -------------------------------------------------------------------------------------------------------
        'Add your license information here.
        ApplicationInfo.License.CopyrightOwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.License.PublicationYear = "2016"

        'License Links:
        'http://choosealicense.com/
        'http://www.apache.org/licenses/
        'http://opensource.org/

        'Apache License 2.0 ---------------------------------------------
        ApplicationInfo.License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
        ApplicationInfo.License.Notice = ApplicationInfo.License.ApacheLicenseNotice 'Get the pre-defined Aapche license notice.
        ApplicationInfo.License.Text = ApplicationInfo.License.ApacheLicenseText     'Get the pre-defined Apache license text.

        'Code to use other pre-defined license types is shown below:

        'GNU General Public License, version 3 --------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.GNU_GPL_V3_0
        'ApplicationInfo.License.Notice = 'Add the License Notice to ADVL_Utilities_Library_1 License class.
        'ApplicationInfo.License.Text = 'Add the License Text to ADVL_Utilities_Library_1 License class.

        'The MIT License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.MIT_License
        'ApplicationInfo.License.Notice = ApplicationInfo.License.MITLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.MITLicenseText

        'No License Specified -------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.None
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'The Unlicense --------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.The_Unlicense
        'ApplicationInfo.License.Notice = ApplicationInfo.License.UnLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.UnLicenseText

        'Unknown License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.Unknown
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'Source Code: --------------------------------------------------------------------------------------------------
        'Add your source code information here if required.
        'THIS SECTION WILL BE UPDATED TO ALLOW A GITHUB LINK.
        ApplicationInfo.SourceCode.Language = "Visual Basic 2015"
        ApplicationInfo.SourceCode.FileName = ""
        ApplicationInfo.SourceCode.FileSize = 0
        ApplicationInfo.SourceCode.FileHash = ""
        ApplicationInfo.SourceCode.WebLink = ""
        ApplicationInfo.SourceCode.Contact = ""
        ApplicationInfo.SourceCode.Comments = ""

        'ModificationSummary: -----------------------------------------------------------------------------------------
        'Add any source code modification here is required.
        ApplicationInfo.ModificationSummary.BaseCodeName = ""
        ApplicationInfo.ModificationSummary.BaseCodeDescription = ""
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Major = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Minor = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Build = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Revision = 0
        ApplicationInfo.ModificationSummary.Description = "This is the first released version of the application. No earlier base code used."

        'Library List: ------------------------------------------------------------------------------------------------
        'Add the ADVL_Utilties_Library_1 library:
        Dim NewLib As New ADVL_Utilities_Library_1.LibrarySummary
        NewLib.Name = "ADVL_System_Utilities"
        NewLib.Description = "System Utility classes used in Andorville (TM) software development system applications"
        NewLib.CreationDate = "7-Jan-2016 12:00:00"
        NewLib.LicenseNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598" & vbCrLf &
                               vbCrLf &
                               "Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf &
                               "you may not use this file except in compliance with the License." & vbCrLf &
                               "You may obtain a copy of the License at" & vbCrLf &
                               vbCrLf &
                               "http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf &
                               vbCrLf &
                               "Unless required by applicable law or agreed to in writing, software" & vbCrLf &
                               "distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf &
                               "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf &
                               "See the License for the specific language governing permissions and" & vbCrLf &
                               "limitations under the License." & vbCrLf

        NewLib.CopyrightNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598"

        NewLib.Version.Major = 1
        NewLib.Version.Minor = 0
        NewLib.Version.Build = 1
        NewLib.Version.Revision = 0

        NewLib.Author.Name = "Signalworks Pty Ltd"
        NewLib.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        NewLib.Author.Contact = "http://www.andorville.com.au/"

        Dim NewClass1 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass1.Name = "ZipComp"
        NewClass1.Description = "The ZipComp class is used to compress files into and extract files from a zip file."
        NewLib.Classes.Add(NewClass1)
        Dim NewClass2 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass2.Name = "XSequence"
        NewClass2.Description = "The XSequence class is used to run an XML property sequence (XSequence) file. XSequence files are used to record and replay processing sequences in Andorville (TM) software applications."
        NewLib.Classes.Add(NewClass2)
        Dim NewClass3 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass3.Name = "XMessage"
        NewClass3.Description = "The XMessage class is used to read an XML Message (XMessage). An XMessage is a simplified XSequence used to exchange information between Andorville (TM) software applications."
        NewLib.Classes.Add(NewClass3)
        Dim NewClass4 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass4.Name = "Location"
        NewClass4.Description = "The Location class consists of properties and methods to store data in a location, which is either a directory or archive file."
        NewLib.Classes.Add(NewClass4)
        Dim NewClass5 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass5.Name = "Project"
        NewClass5.Description = "An Andorville (TM) software application can store data within one or more projects. Each project stores a set of related data files. The Project class contains properties and methods used to manage a project."
        NewLib.Classes.Add(NewClass5)
        Dim NewClass6 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass6.Name = "ProjectSummary"
        NewClass6.Description = "ProjectSummary stores a summary of a project."
        NewLib.Classes.Add(NewClass6)
        Dim NewClass7 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass7.Name = "DataFileInfo"
        NewClass7.Description = "The DataFileInfo class stores information about a data file."
        NewLib.Classes.Add(NewClass7)
        Dim NewClass8 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass8.Name = "Message"
        NewClass8.Description = "The Message class contains text properties and methods used to display messages in an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass8)
        Dim NewClass9 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass9.Name = "ApplicationSummary"
        NewClass9.Description = "The ApplicationSummary class stores a summary of an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass9)
        Dim NewClass10 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass10.Name = "LibrarySummary"
        NewClass10.Description = "The LibrarySummary class stores a summary of a software library used by an application."
        NewLib.Classes.Add(NewClass10)
        Dim NewClass11 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass11.Name = "ClassSummary"
        NewClass11.Description = "The ClassSummary class stores a summary of a class contained in a software library."
        NewLib.Classes.Add(NewClass11)
        Dim NewClass12 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass12.Name = "ModificationSummary"
        NewClass12.Description = "The ModificationSummary class stores a summary of any modifications made to an application or library."
        NewLib.Classes.Add(NewClass12)
        Dim NewClass13 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass13.Name = "ApplicationInfo"
        NewClass13.Description = "The ApplicationInfo class stores information about an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass13)
        Dim NewClass14 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass14.Name = "Version"
        NewClass14.Description = "The Version class stores application, library or project version information."
        NewLib.Classes.Add(NewClass14)
        Dim NewClass15 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass15.Name = "Author"
        NewClass15.Description = "The Author class stores information about an Author."
        NewLib.Classes.Add(NewClass15)
        Dim NewClass16 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass16.Name = "FileAssociation"
        NewClass16.Description = "The FileAssociation class stores the file association extension and description. An application can open files on its file association list."
        NewLib.Classes.Add(NewClass16)
        Dim NewClass17 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass17.Name = "Copyright"
        NewClass17.Description = "The Copyright class stores copyright information."
        NewLib.Classes.Add(NewClass17)
        Dim NewClass18 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass18.Name = "License"
        NewClass18.Description = "The License class stores license information."
        NewLib.Classes.Add(NewClass18)
        Dim NewClass19 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass19.Name = "SourceCode"
        NewClass19.Description = "The SourceCode class stores information about the source code for the application."
        NewLib.Classes.Add(NewClass19)
        Dim NewClass20 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass20.Name = "Usage"
        NewClass20.Description = "The Usage class stores information about application or project usage."
        NewLib.Classes.Add(NewClass20)
        Dim NewClass21 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass21.Name = "Trademark"
        NewClass21.Description = "The Trademark class stored information about a trademark used by the author of an application or data."
        NewLib.Classes.Add(NewClass21)

        ApplicationInfo.Libraries.Add(NewLib)

        'Add other library information here: --------------------------------------------------------------------------

    End Sub


    'Save the form settings if the form is being minimised:
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub SaveProjectSettings()
        'Save the project settings in an XML file.
        'Add any Project Settings to be saved into the settingsData XDocument.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Project settings for ADVL_Coordinates_1 application.-->
                           <ProjectSettings>
                           </ProjectSettings>

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)

    End Sub

    Private Sub RestoreProjectSettings()
        'Restore the project settings from an XML document.

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore a Project Setting example:
            'If Settings.<ProjectSettings>.<Setting1>.Value = Nothing Then
            '    'Project setting not saved.
            '    'Setting1 = ""
            'Else
            '    'Setting1 = Settings.<ProjectSettings>.<Setting1>.Value
            'End If

            'Continue restoring saved settings.

        End If

    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '----------------------------------------------------------------------------------------------------------------------------

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Write the startup messages in a stringbuilder object.
        'Messages cannot be written using Message.Add until this is set up later in the startup sequence.
        Dim sb As New System.Text.StringBuilder
        sb.Append("------------------- Starting Application: ADVL Import Application ---------------------------------------------------------------- " & vbCrLf)

        'Set the Application Directory path: ------------------------------------------------
        Project.ApplicationDir = My.Application.Info.DirectoryPath.ToString

        'Read the Application Information file: ---------------------------------------------
        ApplicationInfo.ApplicationDir = My.Application.Info.DirectoryPath.ToString 'Set the Application Directory property

        If ApplicationInfo.ApplicationLocked Then
            MessageBox.Show("The application is locked. If the application is not already in use, remove the 'Application_Info.lock file from the application directory: " & ApplicationInfo.ApplicationDir, "Notice", MessageBoxButtons.OK)
            Dim dr As System.Windows.Forms.DialogResult
            dr = MessageBox.Show("Press 'Yes' to unlock the application", "Notice", MessageBoxButtons.YesNo)
            If dr = System.Windows.Forms.DialogResult.Yes Then
                ApplicationInfo.UnlockApplication()
            Else
                Application.Exit()
                'System.Windows.Forms.Application.Exit()
            End If
        End If

        ReadApplicationInfo()
        ApplicationInfo.LockApplication()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()
        sb.Append("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#0.##") & " hours" & vbCrLf)

        'Restore Project information: -------------------------------------------------------
        Project.ApplicationName = ApplicationInfo.Name
        Project.ReadLastProjectInfo()
        Project.ReadProjectInfoFile()
        Project.Usage.StartTime = Now

        Project.ReadProjectInfoFile()

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Import.SettingsLocn = Project.SettingsLocn
        Import.DataLocn = Project.DataLocn
        Import.RestoreSettings()

        'Set up the Message object:
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn

        'Initialise all the tab forms:
        InitialiseTabs()



        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        RestoreFormSettings() 'Restore the form settings
        RestoreProjectSettings() 'Restore the Project settings

        'Show the project information: ------------------------------------------------------
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select
        txtCreationDate.Text = Format(Project.Usage.FirstUsed, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsLocationPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataLocationPath.Text = Project.DataLocn.Path


        sb.Append("------------------- Started OK ------------------------------------------------------------------------------------------------------------------------ " & vbCrLf & vbCrLf)
        Me.Show() 'Show this form before showing the Message form
        Message.Add(sb.ToString)




    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        DisconnectFromAppNet() 'Disconnect from the Application Network.

        'SaveFormSettings() 'Save the settings of this form.
        'SaveProjectSettings() 'Save project settings.
        Import.SaveSettings()

        ApplicationInfo.WriteFile() 'Update the Application Information file.
        ApplicationInfo.UnlockApplication()

        Project.SaveLastProjectInfo() 'Save information about the last project used.

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.

        Application.Exit()

    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Save the form settings if the form state is normal. (A minimised form will have the incorrect size and location.)
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------

    'Private Sub btnOpenTemplateForm_Click(sender As Object, e As EventArgs) Handles btnOpenTemplateForm.Click
    '    'Open the Template form:
    '    If IsNothing(TemplateForm) Then
    '        TemplateForm = New frmTemplate
    '        TemplateForm.Show()
    '    Else
    '        TemplateForm.Show()
    '    End If
    'End Sub

    'Private Sub TemplateForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles TemplateForm.FormClosed
    '    TemplateForm = Nothing
    'End Sub

    Private Sub btnMessages_Click(sender As Object, e As EventArgs) Handles btnMessages.Click
        'Show the Messages form.
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()
        Message.MessageForm.BringToFront()
    End Sub

    Private Sub btnShowFile_Click(sender As Object, e As EventArgs) Handles btnShowFile.Click
        'Show the ShowTextFile form:
        If IsNothing(ShowTextFile) Then
            ShowTextFile = New frmShowTextFile
            ShowTextFile.Show()
        Else
            ShowTextFile.Show()
        End If

        ShowTextFile.txtFilePath.Text = lstTextFiles.SelectedItem.ToString
        ShowTextFile.ShowSelectedTextFile()
    End Sub

    Private Sub ShowTextFile_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ShowTextFile.FormClosed
        ShowTextFile = Nothing
    End Sub

    Private Sub btnView_Click(sender As Object, e As EventArgs) Handles btnView.Click
        'Show the Import Sequence form:
        If IsNothing(Sequence) Then
            Sequence = New frmImportSequence
            Sequence.Show()
        Else
            Sequence.Show()
        End If
    End Sub

    Private Sub Sequence_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Sequence.FormClosed
        Sequence = Nothing
    End Sub

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub InitialiseTabs()
        'Initialise all the tab forms.

        'Initialise Import Sequence Tab: --------------------------------------------------------------------------------------------------------------------------------------------------
        txtName.Text = Import.ImportSequenceName
        txtDescription.Text = Import.ImportSequenceDescription


        'Initialise Input Files Tab: ------------------------------------------------------------------------------------------------------------------------------------------------------

        rbManual.Checked = True 'Select Manual file selection as default

        txtInputFileDir.Text = Import.TextFileDir
        If txtInputFileDir.Text <> "" Then
            FillLstTextFiles()
        End If

        If Import.SelectFileMode = "Manual" Then
            rbManual.Checked = True
        Else
            rbSelectionFile.Checked = True
        End If

        'Highlight selected text files in the list box:
        Dim I As Integer
        Dim FoundPosn As Integer
        For I = 1 To Import.SelectedFileCount
            Debug.Print("I = " & I.ToString)
            Debug.Print("ImportTextIntoDatabase.SelTextFiles(I - 1).ToString = " & Import.SelectedFiles(I - 1).ToString)
            FoundPosn = lstTextFiles.FindStringExact(Import.SelectedFiles(I - 1).ToString)
            Debug.Print("FoundPosn = " & FoundPosn.ToString)
            If FoundPosn = -1 Then
                'Path not found in the list
            Else
                lstTextFiles.SetSelected(FoundPosn, True) 'Highlight the path
            End If
        Next I

        'Initialise Output Files Tab: -----------------------------------------------------------------------------------------------------------------------------------------------------
        'Import.ImportSequenceName
        'Import.DatabasePath
        cmbDatabaseType.Items.Add("Access2007To2013")
        cmbDatabaseType.SelectedIndex = 0 'Select the first item
        If Import.DatabasePath = "" Then

        Else
            txtDatabasePath.Text = Import.DatabasePath
            FillLstTables()
        End If


        'Initialise Match Text Tab: -------------------------------------------------------------------------------------------------------------------------------------------------------
        'Set up the DataGrid:
        DataGridView1.ColumnCount = 2
        DataGridView1.RowCount = 1
        DataGridView1.Columns(0).HeaderText = "RegEx Name"
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(0).ToolTipText = "The name of the Regular Expression"
        DataGridView1.Columns(1).HeaderText = "Description"
        DataGridView1.Columns(1).Width = 120
        DataGridView1.Columns(1).ToolTipText = "A description of the Regular Expression"

        RefreshMatchText()

        'Initialise Locations Tab: --------------------------------------------------------------------------------------------------------------------------------------------------------

        'Set up the Database Locations grid:
        DataGridView2.ColumnCount = 5
        DataGridView2.RowCount = 1
        DataGridView2.Columns(0).HeaderText = "RegEx Variable"
        DataGridView2.Columns(0).Width = 120
        DataGridView2.Columns(0).ToolTipText = "The name of the Regular Expression variable"
        DataGridView2.Columns(1).HeaderText = "Variable Type"
        DataGridView2.Columns(1).Width = 120
        DataGridView2.Columns(1).ToolTipText = "The type of variable (Single Value or Multiple Value)"
        DataGridView2.Columns(2).HeaderText = "Destination Table"
        DataGridView2.Columns(2).Width = 120
        DataGridView2.Columns(2).ToolTipText = "The destination table for the variable matched by the regular expression"
        DataGridView2.Columns(3).HeaderText = "Destination Field"
        DataGridView2.Columns(3).Width = 120
        DataGridView2.Columns(3).ToolTipText = "The destination field for the variable matched by the regular expression"
        DataGridView2.Columns(4).HeaderText = "Status Field"
        DataGridView2.Columns(4).Width = 120
        DataGridView2.Columns(4).ToolTipText = "(Optional) The field used to store the status of the value (eg OK or N/A)"
        DataGridView2.AllowUserToResizeColumns = True
        DataGridView2.AllowUserToAddRows = False

        'Set up the Values grid:
        DataGridView3.ColumnCount = 1
        DataGridView3.RowCount = 1
        DataGridView3.Columns(0).HeaderText = "Value 1"
        DataGridView3.Columns(0).Width = 120
        DataGridView3.Columns(0).ToolTipText = "The values of the variables matched by the regular expression"
        DataGridView3.AllowUserToAddRows = False

        cmbVariable.Left = 50
        cmbVariable.Width = 120
        cmbType.Left = 50 + 120
        cmbType.Width = 120
        cmbTable.Left = 50 + 120 + 120
        cmbTable.Width = 120
        cmbField.Left = 50 + 120 + 120 + 120
        cmbField.Width = 120
        cmbStatus.Left = 50 + 120 + 120 + 120 + 120
        cmbStatus.Width = 120

        DbDestIndex = 0 'Set the selected Database Destination Index number to 0 (the first row)
        ListChanged = False
        'RefreshForm()
        RefreshLocations()

        'Initialise Modify Tab: -----------------------------------------------------------------------------------------------------------------------------------------------------------
        txtRegExVariable.Text = Import.ModifyValuesRegExVariable
        txtInputDateFormat.Text = Import.ModifyValuesInputDateFormat
        txtOutputDateFormat.Text = Import.ModifyValuesOutputDateFormat
        txtCharsToReplace.Text = Import.ModifyValuesCharsToReplace
        txtReplacementChars.Text = Import.ModifyValuesReplacementChars
        txtFixedValue.Text = Import.ModifyValuesFixedValue

        Select Case Import.ModifyValuesType
            Case Import.ModifyValuesTypes.ConvertDate
                rbConvertDate.Checked = True
            'Case Import.ModifyValuesTypes.CurrentDate
            Case Import.ModifyValuesTypes.AppendCurrentDate
                rbAppendCurrentDate.Checked = True
            'Case Import.ModifyValuesTypes.CurrentDateTime
            Case Import.ModifyValuesTypes.AppendCurrentDateTime
                rbAppendCurrentDateTime.Checked = True
            'Case Import.ModifyValuesTypes.CurrentTime
            Case Import.ModifyValuesTypes.AppendCurrentTime
                rbAppendCurrentDateTime.Checked = True
            'Case Import.ModifyValuesTypes.FileDir
            Case Import.ModifyValuesTypes.AppendFileDir
                rbAppendTextFileDirectory.Checked = True
            'Case Import.ModifyValuesTypes.FileName
            Case Import.ModifyValuesTypes.AppendFileName
                rbAppendTextFileName.Checked = True
            'Case Import.ModifyValuesTypes.FilePath
            Case Import.ModifyValuesTypes.AppendFilePath
                rbAppendTextFilePath.Checked = True
            'Case Import.ModifyValuesTypes.FixedValue
            Case Import.ModifyValuesTypes.AppendFixedValue
                rbClearValue.Checked = True
            Case Import.ModifyValuesTypes.ReplaceChars
                rbReplaceChars.Checked = True
            Case Import.ModifyValuesTypes.AppendRegExVarValue
                rbAppendRegExVar.Checked = True
            Case Import.ModifyValuesTypes.ClearValue
                rbClearValue.Checked = True
            Case Else
                rbConvertDate.Checked = True
        End Select


        'Initialise Multipliers Tab: ------------------------------------------------------------------------------------------------------------------------------------------------------


        DataGridView4.ColumnCount = 3
        DataGridView4.RowCount = 1
        DataGridView4.Columns(0).HeaderText = "RegEx Multiplier"
        DataGridView4.Columns(0).Width = 120
        DataGridView4.Columns(0).ToolTipText = "The name of the Regular Expression Multiplier variable"
        DataGridView4.Columns(1).HeaderText = "Multiplier Code"
        DataGridView4.Columns(1).Width = 120
        DataGridView4.Columns(1).ToolTipText = "The code representing a multiplier value"
        DataGridView4.Columns(2).HeaderText = "Multiplier Value"
        DataGridView4.Columns(2).Width = 120
        DataGridView4.Columns(2).ToolTipText = "The multiplier value corresponding to the code"

        DataGridView4.AllowUserToAddRows = True

        'Fill Database Destination grid:
        Dim MaxRow As Integer
        'MaxRow = TextToDatabase.MultiplierCodeCount
        MaxRow = Import.MultiplierCodeCount
        DataGridView4.EditMode = DataGridViewEditMode.EditProgrammatically
        Dim I2 As Integer
        For I2 = 1 To MaxRow
            DataGridView4.Rows.Add(1)
            'DataGridView4.Rows(I - 1).Cells(0).Value = TextToDatabase.MultiplierCode(I - 1).RegExMultiplierVariable
            DataGridView4.Rows(I2 - 1).Cells(0).Value = Import.MultiplierCode(I2 - 1).RegExMultiplierVariable
            'DataGridView4.Rows(I - 1).Cells(1).Value = TextToDatabase.MultiplierCode(I - 1).MultiplierCode
            DataGridView4.Rows(I2 - 1).Cells(1).Value = Import.MultiplierCode(I2 - 1).MultiplierCode
            'DataGridView4.Rows(I - 1).Cells(2).Value = TextToDatabase.MultiplierCode(I - 1).MultiplierValue
            DataGridView4.Rows(I2 - 1).Cells(2).Value = Import.MultiplierCode(I2 - 1).MultiplierValue
        Next
        DataGridView4.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2

        DataGridView4.Rows.Item(0).Selected = True

        'ReadFormSettingsXmlFile()

        'Fill cmbVariable from RegEx array:

        Dim J As Integer
        Dim VarName As String
        Dim Match As Boolean
        'For I = 1 To TextToDatabase.DbDestCount
        For I2 = 1 To Import.DbDestCount
            'If TextToDatabase.DbDest(I - 1).Type = "Multiplier" Then
            If Import.DbDest(I2 - 1).Type = "Multiplier" Then
                'VarName = TextToDatabase.DbDest(I - 1).RegExVariable
                VarName = Import.DbDest(I2 - 1).RegExVariable
                If cmbMultVariable.Items.Count > 0 Then
                    Match = False
                    For J = 0 To cmbMultVariable.Items.Count - 1
                        If VarName = cmbMultVariable.Items(J) Then
                            Match = True
                            Exit For
                        End If
                    Next
                    If Match = False Then
                        cmbMultVariable.Items.Add(VarName)
                    End If
                Else
                    cmbMultVariable.Items.Add(VarName)
                End If
            End If
        Next

    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

#Region " Project Information Tab" '-----------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Import.SaveSettings() 'Save the current settings. If a new project is created or selected, new settings will be used.
        Project.SelectProject()
        InitialiseTabs() 'Initialise all the tab forms using the Import settings in the selected project.
    End Sub

    'Project Events:
    Private Sub Project_Message(Msg As String) Handles Project.Message
        'Display the Project message:
        Message.Add(Msg & vbCrLf)
    End Sub

    Private Sub Project_ErrorMessage(Msg As String) Handles Project.ErrorMessage
        'Display the Project error message:
        'Message.SetWarningStyle()
        'Message.Add(Msg & vbCrLf)
        'Message.SetNormalStyle()
        Message.AddWarning(Msg & vbCrLf)
    End Sub

    Private Sub Project_Closing() Handles Project.Closing
        'The current project is closing.

        SaveFormSettings() 'Save the form settings - they are saved in the Project before is closes.
        SaveProjectSettings() 'Update this subroutine if project settings need to be saved.

        'Save the current project usage information:
        Project.Usage.SaveUsageInfo()
    End Sub

    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Import.SettingsLocn = Project.SettingsLocn
        Import.DataLocn = Project.DataLocn
        'Import.Cle
        Import.RestoreSettings()

        RestoreFormSettings() 'This restores the Main form settings for the selected project.
        Project.ReadProjectInfoFile()
        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn

        'Restore the new project settings:
        RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        'Show the project information:
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select

        txtCreationDate.Text = Format(Project.CreationDate, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsLocationPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataLocationPath.Text = Project.DataLocn.Path

    End Sub

#End Region 'Project Information Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Import Sequence Tab" '-------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnOpenSequence_Click(sender As Object, e As EventArgs) Handles btnOpenSequence.Click
        Dim SelectedFileName As String = ""

        SelectedFileName = Project.SelectDataFile("Sequence", "Sequence")
        Message.Add("Selected Import Sequence: " & SelectedFileName & vbCrLf)

        If SelectedFileName = "" Then

        Else
            txtName.Text = SelectedFileName

            Dim xmlSeq As System.Xml.Linq.XDocument

            Project.ReadXmlData(SelectedFileName, xmlSeq)

            If xmlSeq Is Nothing Then
                Exit Sub
            End If

            'rtbSequence.Text = xmlSeq.ToString
            'FormatXmlText()

            Import.ImportSequenceName = SelectedFileName
            Import.ImportSequenceDescription = xmlSeq.<ProcessingSequence>.<Description>.Value
            txtDescription.Text = Import.ImportSequenceDescription
        End If

    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click

    End Sub

    Private Sub Import_ErrorMessage(Message As String) Handles Import.ErrorMessage
        Me.Message.AddWarning(Message & vbCrLf)
    End Sub

    Private Sub Import_Message(Message As String) Handles Import.Message
        Me.Message.Add(Message & vbCrLf)
    End Sub

#End Region 'Import Sequence Tab ------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Input Files Tab" '-----------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub lstTextFiles_LostFocus(sender As Object, e As EventArgs) Handles lstTextFiles.LostFocus
        'Text file selections have been made: update the selected file list:
        UpdateSelTextFileList()
    End Sub

    Private Sub UpdateSelTextFileList()
        'Update the list of selected text files:

        Import.SelTextFilesClear() 'Clear the current list of selected text files.

        Dim I As Integer 'Loop index.
        For I = 1 To lstTextFiles.SelectedItems.Count
            Import.SelTextFileAppend(lstTextFiles.SelectedItems(I - 1))
        Next
    End Sub

    Private Sub rbManual_CheckedChanged(sender As Object, e As EventArgs) Handles rbManual.CheckedChanged
        'Save the file selection mode:
        If rbManual.Checked = True Then
            Import.SelectFileMode = "Manual"
        ElseIf rbSelectionFile.Checked = True Then
            Import.SelectFileMode = "SelectionFile"
        End If
    End Sub

    Private Sub rbSelectionFile_CheckedChanged(sender As Object, e As EventArgs) Handles rbSelectionFile.CheckedChanged
        'Save the file selection mode:
        If rbManual.Checked = True Then
            Import.SelectFileMode = "Manual"
        ElseIf rbSelectionFile.Checked = True Then
            Import.SelectFileMode = "SelectionFile"
        End If
    End Sub

    Private Sub btnTextFileDir_Click(sender As Object, e As EventArgs) Handles btnTextFileDir.Click
        If txtInputFileDir.Text <> "" Then
            FolderBrowserDialog1.SelectedPath = txtInputFileDir.Text
        End If
        FolderBrowserDialog1.ShowDialog()
        txtInputFileDir.Text = FolderBrowserDialog1.SelectedPath
        Import.TextFileDir = FolderBrowserDialog1.SelectedPath
        FillLstTextFiles()

    End Sub

    Public Sub FillLstTextFiles()
        'Fill lstTextFiles with the names of all .txt files in the directory shown in the txtTextFileDir textbox

        'If System.IO.File.Exists(txtTextFileDir.Text) Then
        If System.IO.Directory.Exists(txtInputFileDir.Text) Then
            Dim FileList() As String = System.IO.Directory.GetFiles(txtInputFileDir.Text, "*.*")
            Dim I As Integer 'Loop index

            lstTextFiles.Items.Clear()

            For I = 0 To FileList.Count - 1
                lstTextFiles.Items.Add(FileList(I))
            Next
        Else
            'MessageBox.Show("The directory: " & txtInputFileDir.Text & " doesnt exist!", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("The directory: " & txtInputFileDir.Text & " doesnt exist!" & vbCrLf)
        End If
    End Sub

    Private Sub btnAddInputFilesToSeq_Click(sender As Object, e As EventArgs) Handles btnAddInputFilesToSeq.Click
        'Save the Input Files settings in the current Processing Sequence

        Dim I As Integer

        'If IsNothing(TDS_Import.Sequence) Then
        If IsNothing(Sequence) Then
            'MessageBox.Show("The Sequencing form is not open.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("The Sequencing form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            'Posn = TDS_Import.Sequence.rtbSequence.SelectionStart
            Posn = Sequence.rtbSequence.SelectionStart
            'Debug.Print("Posn = " & Str(Posn))

            'TDS_Import.Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Input Data: The text file directory and selected files:-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Input Data: The text file directory and selected files:-->" & vbCrLf

            'TDS_Import.Sequence.rtbSequence.SelectedText = "<InputData>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<InputData>" & vbCrLf
            'TDS_Import.Sequence.rtbSequence.SelectedText = "  <TextFileDirectory>" & Trim(txtTextFileDir.Text) & "</TextFileDirectory>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <TextFileDirectory>" & Trim(txtInputFileDir.Text) & "</TextFileDirectory>" & vbCrLf
            'TDS_Import.Sequence.rtbSequence.SelectedText = "  <TextFilesToProcess>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <TextFilesToProcess>" & vbCrLf

            If rbManual.Checked = True Then
                'TDS_Import.Sequence.rtbSequence.SelectedText = "  <SelectFileMode>" & "Manual" & "</SelectFileMode>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <SelectFileMode>" & "Manual" & "</SelectFileMode>" & vbCrLf
                'TDS_Import.Sequence.rtbSequence.SelectedText = "    <Command>ClearSelectedFileList</Command>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "    <Command>ClearSelectedFileList</Command>" & vbCrLf
                For I = 1 To lstTextFiles.SelectedItems.Count
                    'TDS_Import.Import.SelectedFiles(I - 1) = lstTextFiles.SelectedItems(I - 1)
                    Import.SelectedFiles(I - 1) = lstTextFiles.SelectedItems(I - 1)
                    'TDS_Import.Sequence.rtbSequence.SelectedText = "    <TextFile>" & lstTextFiles.SelectedItems(I - 1) & "</TextFile>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <TextFile>" & lstTextFiles.SelectedItems(I - 1) & "</TextFile>" & vbCrLf
                Next
            Else
                'TDS_Import.Sequence.rtbSequence.SelectedText = "  <SelectFileMode>" & "SelectionFile" & "</SelectFileMode>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <SelectFileMode>" & "SelectionFile" & "</SelectFileMode>" & vbCrLf
            End If
            'TDS_Import.Sequence.rtbSequence.SelectedText = "  <SelectionFilePath>" & Trim(txtSelectionFile.Text) & "</SelectionFilePath>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <SelectionFilePath>" & Trim(txtSelectionFile.Text) & "</SelectionFilePath>" & vbCrLf
            'TDS_Import.Sequence.rtbSequence.SelectedText = "  </TextFilesToProcess>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </TextFilesToProcess>" & vbCrLf

            'TDS_Import.Sequence.rtbSequence.SelectedText = "</InputData>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "</InputData>" & vbCrLf
            'TDS_Import.Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            'TDS_Import.Sequence.FormatXmlText()
            Sequence.FormatXmlText()

        End If
    End Sub

#End Region 'Input Files Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Output Files Tab" '----------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnDatabase_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
        'Select the database file:

        'OpenFileDialog1.InitialDirectory = TDS_Import.ProjectPath
        If Import.DatabasePath <> "" Then
            OpenFileDialog1.InitialDirectory = Import.DatabasePath
        End If
        OpenFileDialog1.Filter = "Access Database |*.accdb"
        OpenFileDialog1.FileName = ""

        If txtDatabasePath.Text <> "" Then
            Dim fInfo As New System.IO.FileInfo(txtDatabasePath.Text)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.FileName = fInfo.Name
        End If

        OpenFileDialog1.ShowDialog()
        txtDatabasePath.Text = OpenFileDialog1.FileName
        'TDS_Import.txtDatabase.Text = txtDatabase.Text
        Import.DatabasePath = txtDatabasePath.Text

        FillLstTables()
    End Sub

    Private Sub FillLstTables()
        'Fill the lstSelectTable listbox with the availalble tables in the selected database.

        If txtDatabasePath.Text = "" Then Exit Sub

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        lstTables.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + txtDatabasePath.Text

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)

        conn.Open()

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstTables
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            lstTables.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub FillLstFields()
        'Fill the lstFields listbox with the availalble fields in the selected table.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If lstTables.SelectedIndex = -1 Then 'No item is selected
            lstFields.Items.Clear()

        Else 'A table has been selected. List its fields:
            lstFields.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtDatabasePath.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            'commandString = "SELECT * FROM " + lstTables.SelectedItem.ToString
            commandString = "SELECT Top 500 * FROM " + lstTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                lstFields.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If
    End Sub

    Private Sub lstTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTables.SelectedIndexChanged
        FillLstFields()
    End Sub

    Private Sub btnAddOutputFilesToSeq_Click(sender As Object, e As EventArgs) Handles btnAddOutputFilesToSeq.Click
        'Save the Output Files settings to the current Processing Sequence

        If IsNothing(Sequence) Then
            Message.AddWarning("The Import Sequence form is not open." & vbCrLf & "Press the View button on the Import Sequence tab to show this form." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Output Files: The destination for the imported data:-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <Database>" & vbCrLf
            'Sequence.rtbSequence.SelectedText = "  <Database>" & Trim(txtDatabasePath.Text) & "</Database>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <Path>" & Trim(txtDatabasePath.Text) & "</Path>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <Type>" & Trim(cmbDatabaseType.SelectedItem) & "</Type>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </Database>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf
            Sequence.FormatXmlText()

        End If

    End Sub

#End Region 'Output Files Tab ---------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Read Tab" '------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnFirstFile_Click(sender As Object, e As EventArgs) Handles btnFirstFile.Click
        'Open the first file in the selected file list:

        If Import.SelectedFileCount = 0 Then
            'MessageBox.Show("No input files have been selected.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("No input files have been selected." & vbCrLf)
        Else
            Import.SelectFirstFile() 'This selects the first file in the list, updates TextFilePath and opens the file.
            txtInputFile.Text = Import.CurrentFilePath

            If RecordSequence = True Then 'Record this step in the processing sequence.
                Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "OpenFirstFile" & "</ReadTextCommand>" & vbCrLf
                Sequence.FormatXmlText()
            End If
            'End If
        End If
    End Sub

    Private Sub btnNextFile_Click(sender As Object, e As EventArgs) Handles btnNextFile.Click
        'Open the next file in the selected file list:

        Import.SelectNextFile()
        If Import.ImportStatusContains("No_more_input_files") Then
            'The end of the input file list has been reached
            txtInputFile.Text = ""
        Else
            txtInputFile.Text = Import.CurrentFilePath
        End If

        If RecordSequence = True Then 'Record this step in the processing sequence.
            Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "OpenNextFile" & "</ReadTextCommand>" & vbCrLf
            Sequence.FormatXmlText()
        End If
    End Sub

    Private Sub btnGoToStart_Click(sender As Object, e As EventArgs) Handles btnGoToStart.Click
        'Go to the start of the current input file:
        rtbReadText.Text = ""

        Import.GoToStartOfText()

        If RecordSequence = True Then 'Record this step in the processing sequence.
            Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "GoToStart" & "</ReadTextCommand>" & vbCrLf
            Sequence.FormatXmlText()
        End If
    End Sub

    Private Sub btnReadAll_Click(sender As Object, e As EventArgs) Handles btnReadAll.Click
        'Read all of the current input file:

        Import.ReadAllText()
        rtbReadText.Text = Import.TextStore

        If RecordSequence = True Then 'Record this step in the processing sequence.
            Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "ReadAll" & "</ReadTextCommand>" & vbCrLf
            Sequence.FormatXmlText()
        End If
    End Sub

    Private Sub btnReadNextLine_Click(sender As Object, e As EventArgs) Handles btnReadNextLine.Click
        'Read the next line in the current input file:

        Import.ReadNextLineOfText()
        rtbReadText.Text = Import.TextStore

        If RecordSequence = True Then 'Record this step in the processing sequence.
            Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "ReadNextLine" & "</ReadTextCommand>" & vbCrLf
            Sequence.FormatXmlText()
        End If
    End Sub

    Private Sub btnReadNLines_Click(sender As Object, e As EventArgs) Handles btnReadNLines.Click
        'Read the next N lines in the current input file:
        Dim NLines As Integer

        If Trim(Import.CurrentFilePath) = "" Then
            'MessageBox.Show("No text file has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("No text file has been specified." & vbCrLf)
            Exit Sub
        End If

        If txtNLines.Text <> "" Then
            NLines = Int(Val(txtNLines.Text))
            Import.ReadNLinesOfText(NLines)
            rtbReadText.Text = Import.TextStore

            If RecordSequence = True Then 'Record this step in the processing sequence.
                Sequence.rtbSequence.SelectedText = "  <ReadTextNLines>" & txtNLines.Text & "</ReadTextNLines>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "ReadNLines" & "</ReadTextCommand>" & vbCrLf
                Sequence.FormatXmlText()
            End If
        Else 'N Lines is not specified
        End If
    End Sub

    Private Sub btnSkipNLines_Click(sender As Object, e As EventArgs) Handles btnSkipNLines.Click
        'Skip the next N lines in the current input file:
        Dim NLines As Integer

        If Trim(Import.CurrentFilePath) = "" Then
            'MessageBox.Show("No text file has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("No text file has been specified." & vbCrLf)
            Exit Sub
        End If

        If txtNLines.Text <> "" Then
            NLines = Int(Val(txtNLines.Text))
            Import.SkipNLinesOfText(NLines)

            If RecordSequence = True Then 'Record this step in the processing sequence.
                Sequence.rtbSequence.SelectedText = "  <ReadTextNLines>" & txtNLines.Text & "</ReadTextNLines>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "SkipNLines" & "</ReadTextCommand>" & vbCrLf
                Sequence.FormatXmlText()
            End If
        Else 'N Lines is not specified
        End If
    End Sub

    Private Sub btnReadToString_Click(sender As Object, e As EventArgs) Handles btnReadToString.Click
        'Read to the specified string 

        Import.ReadTextToString(txtString.Text)
        rtbReadText.Text = Import.TextStore

        If RecordSequence = True Then 'Record this step in the processing sequence.
            Sequence.rtbSequence.SelectedText = "  <ReadTextString>" & txtString.Text & "</ReadTextString>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "ReadToString" & "</ReadTextCommand>" & vbCrLf
            Sequence.FormatXmlText()
        End If
    End Sub

    Private Sub btnSkipPastString_Click(sender As Object, e As EventArgs) Handles btnSkipPastString.Click
        'Skip past the specified string 

        Import.SkipTextPastString(txtString.Text)

        If RecordSequence = True Then
            Sequence.rtbSequence.SelectedText = "  <ReadTextString>" & txtString.Text & "</ReadTextString>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <ReadTextCommand>" & "SkipPastString" & "</ReadTextCommand>" & vbCrLf
            Sequence.FormatXmlText()
        End If
    End Sub

    Private Sub btnOpenClipboardWindow_Click(sender As Object, e As EventArgs) Handles btnOpenClipboardWindow.Click
        'Open the Clipboard Window form:
        If IsNothing(ClipboardWindow) Then
            ClipboardWindow = New frmClipboardWindow
            ClipboardWindow.Show()
        Else
            ClipboardWindow.Show()
        End If
    End Sub

    Private Sub ClipboardWindow_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ClipboardWindow.FormClosed
        ClipboardWindow = Nothing
    End Sub

#End Region 'Read Tab -----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Match Text Tab" '-------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub RefreshMatchText()

        txtRegExList.Text = Import.RegExListName
        txtListDescription.Text = Import.RegExListDescr

        DataGridView1.Rows.Clear()

        'Fill RegEx grid:
        Dim MaxRow As Integer
        MaxRow = Import.RegExCount

        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically
        For I = 1 To MaxRow
            DataGridView1.Rows.Add(1)
            DataGridView1.Rows(I - 1).Cells(0).Value = Import.RegEx(I - 1).Name
            DataGridView1.Rows(I - 1).Cells(1).Value = Import.RegEx(I - 1).Descr
        Next
        DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2

        RegExIndex = 0

        DataGridView1.Rows(RegExIndex).Selected = True

        'Show the selected RegEx entry on the form for editing.
        txtRegExName.Text = Import.RegEx(RegExIndex).Name
        txtRegExDescr.Text = Import.RegEx(RegExIndex).Descr
        txtRegEx.Text = Import.RegEx(RegExIndex).RegEx
        If Import.RegEx(RegExIndex).ExitOnMatch = True Then
            chkMatchExit.Checked = True
        Else
            chkMatchExit.Checked = False
        End If
        If Import.RegEx(RegExIndex).ExitOnNoMatch = True Then
            chkNoMatchExit.Checked = True
        Else
            chkNoMatchExit.Checked = False
        End If

        txtMatchStatus.Text = Import.RegEx(RegExIndex).MatchStatus
        txtNoMatchStatus.Text = Import.RegEx(RegExIndex).NoMatchStatus

        If txtMatchStatus.Text = "" Then
            chkMatchStatus.Checked = False
        Else
            chkMatchStatus.Checked = True
        End If

        If txtNoMatchStatus.Text = "" Then
            chkNoMatchStatus.Checked = False
        Else
            chkNoMatchStatus.Checked = True
        End If

        lblCount.Text = Str(Import.RegExCount)
        lblRegExNo.Text = Str(RegExIndex + 1)

    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        'Open a RegEx list file:

        '''''OpenFileDialog1.InitialDirectory = TDS_Import.ProjectPath
        ''OpenFileDialog1.InitialDirectory = Project.
        'OpenFileDialog1.Filter = "RegEx List |*.Regexlist"
        'OpenFileDialog1.FileName = ""
        'OpenFileDialog1.ShowDialog()
        'txtRegExList.Text = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
        '''''OpenRegExListFile()
        Dim SelectedFileName As String = ""
        SelectedFileName = Project.SelectDataFile("RegEx List", "RegExList")

        Message.Add("Selected RegEx List: " & SelectedFileName & vbCrLf)

        txtRegExList.Text = SelectedFileName

        OpenRegExListFile()

    End Sub

    Public Sub OpenRegExListFile()
        'Opens the Regular Expression List File shown in the txtRegExList textbox.

        'Read the XML file:
        Dim Index As Integer
        'Dim TempRegEx As TDS_Utilities.ImportSys.strucRegEx
        Dim TempRegEx As Import.strucRegEx
        'If System.IO.File.Exists(TDS_Import.ProjectPath & "\" & txtRegExList.Text & ".Regexlist") Then
        If Project.DataFileExists(txtRegExList.Text) Then
            'Dim RegExList As System.Xml.Linq.XDocument = XDocument.Load(TDS_Import.ProjectPath & "\" & txtRegExList.Text & ".Regexlist")
            Dim RegExList As System.Xml.Linq.XDocument
            Project.ReadXmlData(txtRegExList.Text, RegExList)

            txtListDescription.Text = RegExList.<RegularExpressionList>.<Description>.Value
            Import.RegExListDescr = txtListDescription.Text
            Import.RegExListName = txtRegExList.Text
            Import.RegExListCreationDate = CStr(RegExList.<RegularExpressionList>.<CreationDate>.Value)
            Dim RegExs = From item In RegExList.<RegularExpressionList>.<RegularExpression>
            Index = 0
            Import.RegExClear()
            For Each item In RegExs
                TempRegEx.Name = item.<Name>.Value
                TempRegEx.Descr = item.<Descr>.Value
                TempRegEx.RegEx = item.<RegEx>.Value
                If item.<ExitOnMatch>.Value = "true" Then
                    TempRegEx.ExitOnMatch = True
                Else
                    TempRegEx.ExitOnMatch = False
                End If
                If item.<ExitOnNoMatch>.Value = "true" Then
                    TempRegEx.ExitOnNoMatch = True
                Else
                    TempRegEx.ExitOnNoMatch = False
                End If
                TempRegEx.MatchStatus = item.<MatchStatus>.Value
                TempRegEx.NoMatchStatus = item.<NoMatchStatus>.Value

                Import.RegExAppend(TempRegEx)
            Next
        Else

        End If

        'ReFreshForm()
        RefreshMatchText()
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Save the RegEx list with the specified file name

        Dim RegExListFileName As String
        RegExListFileName = Trim(txtRegExList.Text)
        If RegExListFileName.EndsWith(".RegExList") Then
        Else
            RegExListFileName = RegExListFileName & ".RegExList"
            txtRegExList.Text = RegExListFileName
        End If

        'Exit if no RegEx List name has been specified:
        If RegExListFileName = "" Then
            'MessageBox.Show("No RegEx List name has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("No RegEx List name has been specified." & vbCrLf)
            Exit Sub
        End If

        'Exit if there are no RegEx records in the _RegEx array:
        If IsNothing(Import.mRegEx) Then
            'MessageBox.Show("No Regular Expressions have been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("No Regular Expressions have been specified." & vbCrLf)
            Exit Sub
        End If

        'Exit if there is an existing file that we don't want to overwrite:
        'If System.IO.File.Exists(TDS_Import.ProjectPath & "\" & RegExListFileName & ".Regexlist") Then
        If Project.DataFileExists(RegExListFileName & ".RegExList") Then
            If MessageBox.Show("Overwrite existing file?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.No Then
                Exit Sub
            End If
        End If

        'Now save the RegEx list in the specified file:
        Dim RegExList = <?xml version="1.0" encoding="utf-8"?>
                        <!---->
                        <!--Regular Expression List-->
                        <RegularExpressionList>
                            <!--Summmary-->
                            <Name><%= RegExListFileName %></Name>
                            <Description><%= Trim(txtListDescription.Text) %></Description>
                            <CreationDate><%= Format(Now, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                            <!---->
                            <!--Regular Expressions-->
                            <%= From item In Import.mRegEx
                                Select
                                <RegularExpression>
                                    <Name><%= item.Name %></Name>
                                    <Descr><%= item.Descr %></Descr>
                                    <RegEx><%= item.RegEx %></RegEx>
                                    <ExitOnMatch><%= item.ExitOnMatch %></ExitOnMatch>
                                    <ExitOnNoMatch><%= item.ExitOnNoMatch %></ExitOnNoMatch>
                                    <MatchStatus><%= item.MatchStatus %></MatchStatus>
                                    <NoMatchStatus><%= item.NoMatchStatus %></NoMatchStatus>
                                </RegularExpression>
                            %>
                        </RegularExpressionList>

        '  <%= From item In TDS_Import.Import.mRegEx

        'RegExList.Save(TDS_Import.ProjectPath & "\" & RegExListFileName & ".Regexlist")
        'Project.SaveXmlData(RegExListFileName & ".Regexlist", RegExList)
        Project.SaveXmlData(RegExListFileName, RegExList)

    End Sub

    Private Sub btnAddToSeq_Click(sender As Object, e As EventArgs) Handles btnAddToSeq.Click
        'Add the MatchText settings to the current Import Sequence

        If IsNothing(Sequence) Then
            'MessageBox.Show("The Import Sequence form is not open.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("The Import Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Match Text Regular Expression List: The Regular Expression list used to match text in the input text file :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <MatchTextRegExList>" & Trim(txtRegExList.Text) & "</MatchTextRegExList>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()

        End If
    End Sub

    Private Sub btnRunRegExList_Click(sender As Object, e As EventArgs) Handles btnRunRegExList.Click
        'Run the all the Regular Expressions in the list:

        Import.RunRegExList()

        UpdateMatches()

        'TO UPDATE:
        ''Update the matches shown on the DbDestinations form:
        'If IsNothing(TDS_Import.DbDestinations) Then
        'Else
        '    TDS_Import.DbDestinations.UpdateMatches()
        'End If

        If RecordSequence = True Then
            If IsNothing(Sequence) Then
                'The Sequence form is not open.
            Else
                'Write the processing sequence steps:
                Sequence.rtbSequence.SelectedText = "  <ProcessingCommand>" & "RunRegExList" & "</ProcessingCommand>" & vbCrLf
                Sequence.FormatXmlText()

            End If
        End If
    End Sub

    Private Sub btnMoveUp_Click(sender As Object, e As EventArgs) Handles btnMoveUp.Click
        'Move current RegEx up in the list:

        'Dim TempRegEx As TDS_Utilities.ImportSys.strucRegEx
        Dim TempRegEx As Import.strucRegEx

        If RegExIndex > 0 Then
            RegExIndex = RegExIndex - 1
            'TempRegEx = TDS_Import.Import.RegEx(RegExIndex)
            TempRegEx = Import.RegEx(RegExIndex)
            'TDS_Import.Import.RegExModify(RegExIndex + 1, TempRegEx)
            Import.RegExModify(RegExIndex + 1, TempRegEx)

            TempRegEx.Name = Trim(txtRegExName.Text)
            TempRegEx.Descr = Trim(txtRegExDescr.Text)
            TempRegEx.RegEx = Trim(txtRegEx.Text)
            If chkMatchExit.Checked Then
                TempRegEx.ExitOnMatch = True
            Else
                TempRegEx.ExitOnMatch = False
            End If
            If chkNoMatchExit.Checked Then
                TempRegEx.ExitOnNoMatch = True
            Else
                TempRegEx.ExitOnNoMatch = False
            End If
            If chkMatchStatus.Checked Then
                TempRegEx.MatchStatus = Trim(txtMatchStatus.Text)
            Else
                TempRegEx.MatchStatus = ""
            End If
            If chkNoMatchStatus.Checked Then
                TempRegEx.NoMatchStatus = Trim(txtNoMatchStatus.Text)
            Else
                TempRegEx.NoMatchStatus = ""
            End If

            'TDS_Import.Import.RegExModify(RegExIndex, TempRegEx)
            Import.RegExModify(RegExIndex, TempRegEx)

            lblRegExNo.Text = Str(RegExIndex + 1)

        Else
            'Already at the first RegEx string
        End If
    End Sub

    Private Sub btnMoveDown_Click(sender As Object, e As EventArgs) Handles btnMoveDown.Click

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Add a new RegEx entry to the list:

        'Dim TempRegEx As TDS_Utilities.ImportSys.strucRegEx
        Dim TempRegEx As Import.strucRegEx
        Dim TempRow As Integer

        'If TDS_Import.Import.RegExCount > 0 Then
        If Import.RegExCount > 0 Then
            'Save the current RegEx string:
            TempRegEx.Name = Trim(txtRegExName.Text)
            TempRegEx.Descr = Trim(txtRegExDescr.Text)
            TempRegEx.RegEx = Trim(txtRegEx.Text)
            If chkMatchExit.Checked Then
                TempRegEx.ExitOnMatch = True
            Else
                TempRegEx.ExitOnMatch = False
            End If
            If chkNoMatchExit.Checked Then
                TempRegEx.ExitOnNoMatch = True
            Else
                TempRegEx.ExitOnNoMatch = False
            End If
            If chkMatchStatus.Checked Then
                TempRegEx.MatchStatus = Trim(txtMatchStatus.Text)
            Else
                TempRegEx.MatchStatus = ""
            End If
            If chkNoMatchStatus.Checked Then
                TempRegEx.NoMatchStatus = Trim(txtNoMatchStatus.Text)
            Else
                TempRegEx.NoMatchStatus = ""
            End If

            'TDS_Import.Import.RegExModify(RegExIndex, TempRegEx)
            Import.RegExModify(RegExIndex, TempRegEx)

            'Add a new blank RegEx:
            TempRegEx.Name = ""
            txtRegExName.Text = ""
            TempRegEx.Descr = ""
            txtRegExDescr.Text = ""
            TempRegEx.RegEx = ""
            txtRegEx.Text = ""
            TempRegEx.ExitOnMatch = False
            chkMatchExit.Checked = False
            TempRegEx.ExitOnNoMatch = False
            chkNoMatchExit.Checked = False
            TempRegEx.MatchStatus = ""
            txtMatchStatus.Text = ""
            TempRegEx.NoMatchStatus = ""
            txtNoMatchStatus.Text = ""

            RegExIndex = RegExIndex + 1
            'TDS_Import.Import.RegExInsert(RegExIndex, TempRegEx)
            Import.RegExInsert(RegExIndex, TempRegEx)

            'Update DataGridView1:
            TempRow = DataGridView1.FirstDisplayedScrollingRowIndex 'Save the number of the first displayed row
            Dim MaxRow As Integer
            'MaxRow = TDS_Import.Import.RegExCount
            MaxRow = Import.RegExCount
            DataGridView1.Rows.Clear()
            DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically
            For I = 1 To MaxRow
                DataGridView1.Rows.Add(1)
                'DataGridView1.Rows(I - 1).Cells(0).Value = TDS_Import.Import.RegEx(I - 1).Name
                DataGridView1.Rows(I - 1).Cells(0).Value = Import.RegEx(I - 1).Name
                'DataGridView1.Rows(I - 1).Cells(1).Value = TDS_Import.Import.RegEx(I - 1).Descr
                DataGridView1.Rows(I - 1).Cells(1).Value = Import.RegEx(I - 1).Descr
            Next

            'Highlight the current row:
            DataGridView1.ClearSelection()
            DataGridView1.Rows(RegExIndex).Selected = True
            DataGridView1.FirstDisplayedScrollingRowIndex = TempRow 'Restore the number of the first displayed row

            'lblCount.Text = Str(TDS_Import.Import.RegExCount)
            lblCount.Text = Str(Import.RegExCount)
            lblRegExNo.Text = Str(RegExIndex + 1)
        Else
            'Add a new blank RegEx:
            TempRegEx.Name = ""
            txtRegExName.Text = ""
            TempRegEx.Descr = ""
            txtRegExDescr.Text = ""
            TempRegEx.ExitOnMatch = False
            chkMatchExit.Checked = False
            TempRegEx.ExitOnNoMatch = False
            chkNoMatchExit.Checked = False
            TempRegEx.MatchStatus = ""
            txtMatchStatus.Text = ""
            TempRegEx.NoMatchStatus = ""
            txtNoMatchStatus.Text = ""

            RegExIndex = 0
            'Highlight the current row:
            If DataGridView1.RowCount > 0 Then
                DataGridView1.ClearSelection()
                DataGridView1.Rows(RegExIndex).Selected = True
            End If


            'TDS_Import.Import.RegExInsert(RegExIndex, TempRegEx)
            Import.RegExInsert(RegExIndex, TempRegEx)
            'lblCount.Text = Str(TDS_Import.Import.RegExCount)
            lblCount.Text = Str(Import.RegExCount)
            lblRegExNo.Text = Str(RegExIndex + 1)
        End If
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click

        Import.RegExDelete(RegExIndex)
        RefreshMatchText()

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        UpdateRegEx()
    End Sub

    Private Sub UpdateRegEx()
        'Update the Regular Expression

        'RegExIndex points to the current RegEx in DataGridView1

        'Dim TempRegEx As TDS_Utilities.ImportSys.strucRegEx
        Dim TempRegEx As Import.strucRegEx

        'If TDS_Import.Import.RegExCount > 0 Then
        If Import.RegExCount > 0 Then
            'Update DataGridView1:
            DataGridView1.Rows(RegExIndex).Cells(0).Value = Trim(txtRegExName.Text)
            DataGridView1.Rows(RegExIndex).Cells(1).Value = Trim(txtRegExDescr.Text)

            'Save the RegEx settings in TempRegEx:
            TempRegEx.Name = Trim(txtRegExName.Text)
            TempRegEx.Descr = Trim(txtRegExDescr.Text)
            TempRegEx.RegEx = Trim(txtRegEx.Text)
            If chkMatchExit.Checked Then
                TempRegEx.ExitOnMatch = True
            Else
                TempRegEx.ExitOnMatch = False
            End If
            If chkNoMatchExit.Checked Then
                TempRegEx.ExitOnNoMatch = True
            Else
                TempRegEx.ExitOnNoMatch = False
            End If
            If chkMatchStatus.Checked Then
                TempRegEx.MatchStatus = Trim(txtMatchStatus.Text)
            Else
                TempRegEx.MatchStatus = ""
            End If
            If chkNoMatchStatus.Checked Then
                TempRegEx.NoMatchStatus = Trim(txtNoMatchStatus.Text)
            Else
                TempRegEx.NoMatchStatus = ""
            End If

            'Update the RegEx List:
            'TDS_Import.Import.RegExModify(RegExIndex, TempRegEx)
            Import.RegExModify(RegExIndex, TempRegEx)

        Else 'This is the first RegEx to add to the list
            'Update DataGridView1:
            DataGridView1.Rows(RegExIndex).Cells(0).Value = Trim(txtRegExName.Text)
            DataGridView1.Rows(RegExIndex).Cells(1).Value = Trim(txtRegExDescr.Text)

            'Save the RegEx settings in TempRegEx:
            TempRegEx.Name = Trim(txtRegExName.Text)
            TempRegEx.Descr = Trim(txtRegExDescr.Text)
            TempRegEx.RegEx = Trim(txtRegEx.Text)
            If chkMatchExit.Checked Then
                TempRegEx.ExitOnMatch = True
            Else
                TempRegEx.ExitOnMatch = False
            End If
            If chkNoMatchExit.Checked Then
                TempRegEx.ExitOnNoMatch = True
            Else
                TempRegEx.ExitOnNoMatch = False
            End If
            If chkMatchStatus.Checked Then
                TempRegEx.MatchStatus = Trim(txtMatchStatus.Text)
            Else
                TempRegEx.MatchStatus = ""
            End If
            If chkNoMatchStatus.Checked Then
                TempRegEx.NoMatchStatus = Trim(txtNoMatchStatus.Text)
            Else
                TempRegEx.NoMatchStatus = ""
            End If

            'Update the RegEx List:
            'TDS_Import.Import.RegExModify(RegExIndex, TempRegEx)
            Import.RegExModify(RegExIndex, TempRegEx)

            'lblCount.Text = Str(TDS_Import.Import.RegExCount)
            lblCount.Text = Str(Import.RegExCount)
            lblRegExNo.Text = Str(RegExIndex + 1)

        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

    End Sub

    Private Sub btnRunRegEx_Click(sender As Object, e As EventArgs) Handles btnRunRegEx.Click
        'Run the current Regular Expression:
        'RegExIndex points to the current RegEx in DataGridView1.
        '(This is also the RegEx being edited in the Name, Description and RegEx text boxes.)
        Import.RunRegEx(RegExIndex)

        'Update the matches shown on the DbDestinations form:
        'If IsNothing(TDS_Import.DbDestinations) Then

        'Else
        '    TDS_Import.DbDestinations.UpdateMatches()
        'End If
        UpdateMatches()

        Dim NValues As Integer
        NValues = Import.DbDestValues.GetUpperBound(1) + 1 'This is the number of columns in the DbDestValues array
        'TDS_Import.AddMessage("Number of columns in DbDestValues is: " & Str(NValues) & vbCrLf)
        Message.Add("Number of columns in DbDestValues is: " & Str(NValues) & vbCrLf)

        Dim NRows As Integer
        NRows = Import.DbDestValues.GetUpperBound(0) + 1 'This is the number of rows in the DbDestValues array
        'TDS_Import.AddMessage("Number of rows in DbDestValues is: " & Str(NRows) & vbCrLf)
        Message.Add("Number of rows in DbDestValues is: " & Str(NRows) & vbCrLf)

        Dim I As Integer
        For I = 0 To NRows - 1
            'TDS_Import.AddMessage("Row number: " & Str(I) & " Value: " & TDS_Import.Import.DbDestValues(I, 0) & vbCrLf)
            Message.Add("Row number: " & Str(I) & " Value: " & Import.DbDestValues(I, 0) & vbCrLf)
        Next
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'The DataGridView1 has been clicked

        Dim SelRow As Integer
        'Highlight the selected row:
        SelRow = DataGridView1.SelectedCells(0).RowIndex
        DataGridView1.Rows(SelRow).Selected = True
        RegExIndex = SelRow
        lblRegExNo.Text = RegExIndex + 1

        'Show the selected RegEx entry on the form for editing.
        'txtRegExName.Text = TextToDatabase.RegEx(RegExIndex).Name
        txtRegExName.Text = Import.RegEx(RegExIndex).Name
        'txtRegExDescr.Text = TextToDatabase.RegEx(RegExIndex).Descr
        txtRegExDescr.Text = Import.RegEx(RegExIndex).Descr
        'txtRegEx.Text = TextToDatabase.RegEx(RegExIndex).RegEx
        txtRegEx.Text = Import.RegEx(RegExIndex).RegEx
        'If TextToDatabase.RegEx(RegExIndex).ExitOnMatch = True Then
        If Import.RegEx(RegExIndex).ExitOnMatch = True Then
            chkMatchExit.Checked = True
        Else
            chkMatchExit.Checked = False
        End If
        'If TextToDatabase.RegEx(RegExIndex).ExitOnNoMatch = True Then
        If Import.RegEx(RegExIndex).ExitOnNoMatch = True Then
            chkNoMatchExit.Checked = True
        Else
            chkNoMatchExit.Checked = False
        End If

        'txtMatchStatus.Text = TextToDatabase.RegEx(RegExIndex).MatchStatus
        txtMatchStatus.Text = Import.RegEx(RegExIndex).MatchStatus
        'txtNoMatchStatus.Text = TextToDatabase.RegEx(RegExIndex).NoMatchStatus
        txtNoMatchStatus.Text = Import.RegEx(RegExIndex).NoMatchStatus

        If txtMatchStatus.Text = "" Then
            chkMatchStatus.Checked = False
        Else
            chkMatchStatus.Checked = True
        End If

        If txtNoMatchStatus.Text = "" Then
            chkNoMatchStatus.Checked = False
        Else
            chkNoMatchStatus.Checked = True
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub txtRegExName_LostFocus(sender As Object, e As EventArgs) Handles txtRegExName.LostFocus
        UpdateRegEx()
    End Sub

    Private Sub txtRegExDescr_LostFocus(sender As Object, e As EventArgs) Handles txtRegExDescr.LostFocus
        UpdateRegEx()
    End Sub

    Private Sub txtRegEx_LostFocus(sender As Object, e As EventArgs) Handles txtRegEx.LostFocus
        UpdateRegEx()
    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        'Create a new Regular Expression list.

        txtRegExList.Text = ""
        txtListDescription.Text = ""
        DataGridView1.Rows.Clear()
        Import.RegExClear() 'Clear the regex list
        lblRegExNo.Text = 0
        lblCount.Text = 0
        txtRegExName.Text = ""
        txtRegExDescr.Text = ""
        txtRegEx.Text = ""
        txtMatchStatus.Text = ""
        txtNoMatchStatus.Text = ""


    End Sub


#End Region 'Match Text Tab -----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Locations Tab" '--------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnOpenLocationList_Click(sender As Object, e As EventArgs) Handles btnOpenLocationList.Click
        'Open a database location list file:

        Dim SelectedFileName As String = ""
        SelectedFileName = Project.SelectDataFile("Database Location List", "DbLocnList")

        Message.Add("Selected Database Locations List: " & SelectedFileName & vbCrLf)

        txtDbDestList.Text = SelectedFileName
        Import.DbDestListName = SelectedFileName

        OpenDbDestListFile()
        ListChanged = False

    End Sub

    Public Sub OpenDbDestListFile()
        'Opens the Database Destinations List File shown in the txtDbDesList textbox.

        'Read the XML file:
        Dim Index As Integer
        'Dim TempDbDest As TDS_Utilities.ImportSys.strucDbDest
        Dim TempDbDest As Import.strucDbDest
        'Dim TempDbMult As TDS_Utilities.ImportSys.strucMultiplier
        Dim TempDbMult As Import.strucMultiplier
        'If System.IO.File.Exists(TDS_Import.ProjectPath & "\" & txtDbDestList.Text & ".Dbdestlist") Then
        If Project.DataFileExists(txtDbDestList.Text) Then
            'Dim DbDestList As System.Xml.Linq.XDocument = XDocument.Load(TDS_Import.ProjectPath & "\" & txtDbDestList.Text & ".Dbdestlist")
            Dim DbDestList As System.Xml.Linq.XDocument
            Project.ReadXmlData(txtDbDestList.Text, DbDestList)
            'TDS_Import.Import.DbDestListDescription = DbDestList.<DatabaseDestinations>.<Description>.Value
            Import.DbDestListDescription = DbDestList.<DatabaseDestinations>.<Description>.Value
            'txtListDescription.Text = TDS_Import.Import.DbDestListDescription
            txtListDescription.Text = Import.DbDestListDescription

            If DbDestList.<DatabaseDestinations>.<CreationDate>.Value = "" Then

            Else
                'TDS_Import.Import.DbDestListCreationDate = DbDestList.<DatabaseDestinations>.<CreationDate>.Value
                Import.DbDestListCreationDate = DbDestList.<DatabaseDestinations>.<CreationDate>.Value
            End If

            If DbDestList.<DatabaseDestinations>.<LastEditDate>.Value = "" Then

            Else
                Import.DbDestListLastEditDate = DbDestList.<DatabaseDestinations>.<LastEditDate>.Value
            End If

            Dim DbDests = From item In DbDestList.<DatabaseDestinations>.<DestinationList>.<DatabaseDestination>
            Index = 0
            Import.DbDestClear()
            'Read each Database Destination entry:
            For Each item In DbDests
                TempDbDest.RegExVariable = item.<RegExVariable>.Value
                TempDbDest.Type = item.<Type>.Value
                TempDbDest.TableName = item.<TableName>.Value
                TempDbDest.FieldName = item.<FieldName>.Value
                TempDbDest.StatusField = item.<StatusField>.Value
                Import.DbDestAppend(TempDbDest)
            Next
            'Read the Multipliers:
            Dim DbMult = From item In DbDestList.<DatabaseDestinations>.<MultiplierList>.<Multiplier>
            Index = 0
            Import.MultipliersClear()
            'Read each Multiplier entry:
            For Each item In DbMult
                TempDbMult.RegExMultiplierVariable = item.<RegExMultiplier>.Value
                TempDbMult.MultiplierCode = item.<MultiplierCode>.Value
                TempDbMult.MultiplierValue = item.<MultiplierValue>.Value
                Import.MultipliersAppend(TempDbMult)
            Next
            If DbDestList.<DatabaseDestinations>.<UseNullValueString>.Value <> Nothing Then
                If DbDestList.<DatabaseDestinations>.<UseNullValueString>.Value = "true" Then
                    Import.UseNullValueString = True
                Else
                    Import.UseNullValueString = False
                End If
            Else
                Import.UseNullValueString = False
            End If
            If DbDestList.<DatabaseDestinations>.<NullValueString>.Value <> Nothing Then
                Import.NullValueString = DbDestList.<DatabaseDestinations>.<NullValueString>.Value
            Else
                Import.NullValueString = ""
            End If
        Else

        End If

        'RefreshForm()
        RefreshLocations()

    End Sub

    'Public Sub RefreshForm()
    Public Sub RefreshLocations()
        'Refresh the form:

        txtDbDestList.Text = Import.DbDestListName
        'txtListDescription.Text = Import.DbDestListDescription
        txtDbDestListDescr.Text = Import.DbDestListDescription

        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()

        'Fill Database Destination grid:
        Dim MaxRow As Integer
        MaxRow = Import.DbDestCount
        DataGridView2.EditMode = DataGridViewEditMode.EditProgrammatically
        For I = 1 To MaxRow
            DataGridView2.Rows.Add(1)
            DataGridView2.Rows(I - 1).Cells(0).Value = Import.DbDest(I - 1).RegExVariable
            DataGridView2.Rows(I - 1).Cells(1).Value = Import.DbDest(I - 1).Type
            DataGridView2.Rows(I - 1).Cells(2).Value = Import.DbDest(I - 1).TableName
            DataGridView2.Rows(I - 1).Cells(3).Value = Import.DbDest(I - 1).FieldName
            DataGridView2.Rows(I - 1).Cells(4).Value = Import.DbDest(I - 1).StatusField
            DataGridView3.Rows.Add(1)
        Next
        DataGridView2.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2

        'Fill list of variable type options
        cmbType.Items.Clear()
        cmbType.Items.Add("Single Value")
        cmbType.Items.Add("Multiple Value")
        cmbType.Items.Add("Single Multiplier")
        cmbType.Items.Add("Multiple Multiplier")
        cmbType.Items.Add("Currency")

        'Fill list of available destination tables:
        FillCmbTable()

        'Show UseNullValueString settings:
        If Import.UseNullValueString = True Then
            chkUseNullValueString.Checked = True
        Else
            chkUseNullValueString.Checked = False
        End If
        txtNullValueString.Text = Import.NullValueString

    End Sub

    Private Sub FillCmbTable()
        'Fill the cmbTable combobox with the availalble tables in the selected database.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbTable.Items.Clear()

        If Import.DatabasePath = "" Then
            Exit Sub
        End If

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + Import.DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)

        conn.Open()

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstSelectTable
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            cmbTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Public Sub UpdateMatches()
        'Update the matched values in DataGridView2

        Dim NCols As Integer 'The number of values matched - DataGridView2 should have its number of columns adjusted to match this.
        Dim NRows As Integer

        'NCols = Import.DbDestValues.GetUpperBound(1)
        'NRows = Import.DbDestValues.GetUpperBound(0)
        NCols = Import.DbDestValues.GetUpperBound(1) + 1
        NRows = Import.DbDestValues.GetUpperBound(0) + 1

        'DataGridView3.ColumnCount = NCols + 1
        DataGridView3.ColumnCount = NCols
        'DataGridView3.RowCount = NRows + 1 'TEMPORARY CHANGE TO DEBUG AN ERROR - TO BE RESOLVED
        DataGridView3.RowCount = NRows
        Dim RowNo As Integer
        Dim ColNo As Integer

        'For RowNo = 0 To NRows
        For RowNo = 0 To NRows - 1
            'For ColNo = 0 To NCols
            For ColNo = 0 To NCols - 1
                DataGridView3.Rows(RowNo).Cells(ColNo).Value = Import.DbDestValues(RowNo, ColNo)
            Next
        Next

    End Sub

    Private Sub FillCmbField()
        'Fill the cmbField combobox with the availalble fields in the selected table.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If cmbTable.SelectedIndex = -1 Then 'No item is selected
            cmbField.Items.Clear()
        Else 'A table has been selected. List its fields:
            cmbField.Items.Clear()
            cmbStatus.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + Import.DatabasePath

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            commandString = "SELECT * FROM " + cmbTable.SelectedItem.ToString

            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count

            Dim I As Integer
            For I = 0 To NFields - 1
                cmbField.Items.Add(dt.Columns(I).ColumnName.ToString)
                cmbStatus.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If

    End Sub

    Private Sub btnSaveLocationList_Click(sender As Object, e As EventArgs) Handles btnSaveLocationList.Click
        'Save the Database Destinations list with the specified file name.
        'txtDbDestList.Text contains the file name.
        SaveDbDestList()
    End Sub

    Private Sub SaveDbDestList()

        Dim DbDestListFileName As String
        DbDestListFileName = Trim(txtDbDestList.Text)
        If DbDestListFileName.EndsWith(".DbLocnList") Then
        Else
            DbDestListFileName = DbDestListFileName & ".DbLocnList"
            txtDbDestList.Text = DbDestListFileName
        End If

        'Exit if no RegEx List name has been specified:
        If DbDestListFileName = "" Then
            'MessageBox.Show("No Databse Destinations List name has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("No Database Destinations List name has been specified." & vbCrLf)
            Exit Sub
        End If

        'Exit if there are no RegEx records in the _RegEx array:
        If IsNothing(Import.mDbDest) Then
            'MessageBox.Show("No Database Destinations have been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("No Database Destinations have been specified." & vbCrLf)
            Exit Sub
        End If

        If IsNothing(Import.mMultiplierCodes) Then
            'Create a blank multiplier to prevent error creating XML file of database destinations:
            'Dim TempMult As TDS_Utilities.ImportSys.strucMultiplier
            Dim TempMult As Import.strucMultiplier
            TempMult.RegExMultiplierVariable = ""
            TempMult.MultiplierCode = ""
            TempMult.MultiplierValue = 0
            Import.MultipliersAppend(TempMult)
        End If

        'Exit if there is an existing file that we don't want to overwrite:
        'If System.IO.File.Exists(TDS_Import.ProjectPath & "\" & DbDestListFileName & ".Dbdestlist") Then
        If Project.DataFileExists(DbDestListFileName) Then
            If MessageBox.Show("Overwrite existing file?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.No Then
                Exit Sub
            End If
            Import.DbDestListLastEditDate = Now
        Else
            'We are not overwriting an existing file:
            Import.DbDestListCreationDate = Now
            Import.DbDestListLastEditDate = Now
        End If

        'Save the current Database Destination data:
        'Dim TempDbDest As TDS_Utilities.ImportSys.strucDbDest
        '   DataGridView2.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
        Dim TempDbDest As Import.strucDbDest
        TempDbDest.RegExVariable = DataGridView2.Rows(DbDestIndex).Cells(0).Value
        TempDbDest.Type = DataGridView2.Rows(DbDestIndex).Cells(1).Value
        TempDbDest.TableName = DataGridView2.Rows(DbDestIndex).Cells(2).Value
        TempDbDest.FieldName = DataGridView2.Rows(DbDestIndex).Cells(3).Value
        TempDbDest.StatusField = DataGridView2.Rows(DbDestIndex).Cells(4).Value
        Import.DbDestModify(DbDestIndex, TempDbDest)

        Import.DbDestListName = DbDestListFileName
        Import.DbDestListDescription = Trim(txtListDescription.Text)

        Dim DbDestList = <?xml version="1.0" encoding="utf-8"?>
                         <!---->
                         <!--Database Destnations-->
                         <DatabaseDestinations>
                             <!--Summmary-->
                             <Name><%= DbDestListFileName %></Name>
                             <Description><%= Trim(txtListDescription.Text) %></Description>
                             <CreationDate><%= Format(Import.DbDestListCreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                             <LastEditDate><%= Format(Import.DbDestListLastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                             <!---->
                             <!--Destination List:-->
                             <DestinationList>
                                 <%= From item In Import.mDbDest
                                     Select
                                     <DatabaseDestination>
                                         <RegExVariable><%= item.RegExVariable %></RegExVariable>
                                         <Type><%= item.Type %></Type>
                                         <TableName><%= item.TableName %></TableName>
                                         <FieldName><%= item.FieldName %></FieldName>
                                         <StatusField><%= item.StatusField %></StatusField>
                                     </DatabaseDestination>
                                 %>
                             </DestinationList>
                             <!---->
                             <!--Multiplier List:-->
                             <MultiplierList>
                                 <%= From item In Import.mMultiplierCodes
                                     Select
                                     <Multiplier>
                                         <RegExMultiplier><%= item.RegExMultiplierVariable %></RegExMultiplier>
                                         <MultiplierCode><%= item.MultiplierCode %></MultiplierCode>
                                         <MultiplierValue><%= item.MultiplierValue %></MultiplierValue>
                                     </Multiplier>
                                 %>
                             </MultiplierList>
                             <!---->
                             <!--Null Value String-->
                             <UseNullValueString><%= Import.UseNullValueString %></UseNullValueString>
                             <NullValueString><%= Import.NullValueString %></NullValueString>

                         </DatabaseDestinations>

        'DbDestList.Save(TDS_Import.ProjectPath & "\" & DbDestListFileName & ".Dbdestlist")
        Project.SaveXmlData(DbDestListFileName, DbDestList)
        ListChanged = False
    End Sub

    Private Sub btnNewLocationList_Click(sender As Object, e As EventArgs) Handles btnNewLocationList.Click
        'Create a new DbDestinations list

        If ListChanged = True Then
            'Dim dr As Windows.Forms.DialogResult
            Dim dr As DialogResult
            dr = MessageBox.Show("Save the changes to the Database Destinations list?", "Notice", MessageBoxButtons.YesNoCancel)
            'If dr = Windows.Forms.DialogResult.Yes Then 'Save the changes to the Database Destinations list:
            If dr = DialogResult.Yes Then
                'Exit if no RegEx List name has been specified:
                If Trim(txtDbDestList.Text) = "" Then
                    'MessageBox.Show("No Databse Destinations List name has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Message.AddWarning("No Databse Destinations List name has been specified." & vbCrLf)
                    Exit Sub
                End If

                'Exit if there are no RegEx records in the _RegEx array:
                If IsNothing(Import.mDbDest) Then
                    'MessageBox.Show("No Database Destinations have been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Message.AddWarning("No Database Destinations have been specified." & vbCrLf)
                    Exit Sub
                End If
                SaveDbDestList()
                'ElseIf dr = Windows.Forms.DialogResult.No Then
            ElseIf dr = DialogResult.No Then
                'Dont save the Database Destinations list.
            Else 'Cancel
                Exit Sub
            End If
        End If

        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        txtDbDestList.Text = ""
        txtDbDestListDescr.Text = ""
        Import.DbDestClear()
        RegExIndex = 0
    End Sub

    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        DataGridView2.Rows.Clear() 'Clear the Data Destinations grid
        DataGridView3.Rows.Clear() 'Clear the Text Matches grid
        txtDbDestList.Text = ""
        txtListDescription.Text = ""
        Import.DbDestClear()
    End Sub

    Private Sub btnClearMatches_Click(sender As Object, e As EventArgs) Handles btnClearMatches.Click
        ClearMatches()
    End Sub

    Public Sub ClearMatches()
        'Clear all the text matches in DataGridView3

        DataGridView3.ColumnCount = 1
        Dim I As Integer
        For I = 1 To DataGridView3.RowCount
            DataGridView3.Rows(I - 1).Cells(0).Value = ""
        Next
    End Sub

    Private Sub btnWrite_Click(sender As Object, e As EventArgs) Handles btnWrite.Click
        'Write to matches to the database:

        Import.SetMultipliers() 'Process the multiplier rows. Values may require the application of multiplier factors before being written to the database.
        Import.GetTableList()   'Generates a list of destination tables. This list is required before the data is written to the database.
        Import.GetFieldList()   'Generates a list of destination fields. GetTableList must be run before this.
        Import.GetFieldValues() 'Generates a list of field values.

        'Old Write to Database code:
        'Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        'Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        ''Set up the database connection:
        ''http://www.connectionstrings.com/access/
        ''Specify the connection string (Access 2007):
        'connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + _
        '"data source = " + TDS_Import.Import.DatabasePath
        ''Connect to the Access database:
        'conn = New System.Data.OleDb.OleDbConnection(connectionString)
        'conn.Open()

        'Dim da As New System.Data.OleDb.OleDbDataAdapter
        'Dim cmd As New System.Data.OleDb.OleDbCommand

        'da.InsertCommand = cmd
        'da.InsertCommand.Connection = conn

        'TDS_Import.Import.WriteToDatabase(da)

        'da.InsertCommand.Connection.Close()

        'New Write to Database code:
        Import.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0; data source = " + Import.DatabasePath
        Import.OpenDatabase()
        Import.WriteToDatabase()
        Import.CloseDatabase()


        If RecordSequence = True Then 'Add the WriteToDatabase statement to the Sequence:
            If IsNothing(Sequence) Then
                'MessageBox.Show("The Import Sequence form is not open.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Message.AddWarning("The Import Sequence form is not open." & vbCrLf)
            Else
                'Write new instruction to the Sequence text.
                'Dim Posn As Integer
                'Posn = TDS_Import.Sequence.rtbSequence.SelectionStart
                Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Note: Move OpenDatabase and CloseDatabase statements out of loops for more efficent processing.-->" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ProcessingCommand>OpenDatabase</ProcessingCommand>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ProcessingCommand>ProcessMatches</ProcessingCommand>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ProcessingCommand>WriteToDatabase</ProcessingCommand>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ProcessingCommand>CloseDatabase</ProcessingCommand>" & vbCrLf

                Sequence.FormatXmlText()

            End If
        End If
    End Sub


    Private Sub btnAddLocnsToSequence_Click(sender As Object, e As EventArgs) Handles btnAddLocnsToSequence.Click
        'Save the DatabaseDestinations setting in the current Processing Sequence

        If IsNothing(Sequence) Then
            'MessageBox.Show("The Import Sequence form is not open.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("The Import Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Database Destinations: The table and field destinations of each text match in the destination database:-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <DatabaseDestinationsList>" & Trim(txtDbDestList.Text) & "</DatabaseDestinationsList>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()

        End If
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim SelRow As Integer

        'Highlight the selected row:
        SelRow = DataGridView2.SelectedCells(0).RowIndex
        DataGridView2.Rows(SelRow).Selected = True

        If DbDestIndex <> SelRow Then
            'Save the previously selected Database Destination data:
            Dim TempDbDest As Import.strucDbDest
            TempDbDest.RegExVariable = DataGridView2.Rows(DbDestIndex).Cells(0).Value
            TempDbDest.Type = DataGridView2.Rows(DbDestIndex).Cells(1).Value
            TempDbDest.TableName = DataGridView2.Rows(DbDestIndex).Cells(2).Value
            TempDbDest.FieldName = DataGridView2.Rows(DbDestIndex).Cells(3).Value
            TempDbDest.StatusField = DataGridView2.Rows(DbDestIndex).Cells(4).Value
            Import.DbDestModify(DbDestIndex, TempDbDest)
        End If

        DbDestIndex = SelRow
        ListChanged = True
    End Sub


    Private Sub DataGridView2_LostFocus(sender As Object, e As EventArgs) Handles DataGridView2.LostFocus
        'Save the data in the current row:

        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If

        Dim TempDbDest As Import.strucDbDest
        TempDbDest.RegExVariable = DataGridView2.Rows(DbDestIndex).Cells(0).Value
        TempDbDest.Type = DataGridView2.Rows(DbDestIndex).Cells(1).Value
        TempDbDest.TableName = DataGridView2.Rows(DbDestIndex).Cells(2).Value
        TempDbDest.FieldName = DataGridView2.Rows(DbDestIndex).Cells(3).Value
        TempDbDest.StatusField = DataGridView2.Rows(DbDestIndex).Cells(4).Value
        Import.DbDestModify(DbDestIndex, TempDbDest)
        ListChanged = True
    End Sub

    Private Sub cmbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType.SelectedIndexChanged
        'Update the Value Type entry:
        Dim SelRow As Integer
        If DataGridView2.SelectedCells.Count = 0 Then
            'MsgBox("No cell selected in the grid!", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("No cell selected in the grid!" & vbCrLf)
            Beep()
        Else
            SelRow = DataGridView2.SelectedCells.Item(0).RowIndex
            DataGridView2.Rows(SelRow).Cells(1).Value = cmbType.SelectedItem.ToString
        End If
    End Sub

    Private Sub cmbField_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbField.SelectedIndexChanged
        'Update the destination Field entry:

        Dim SelRow As Integer
        If DataGridView2.SelectedCells.Count = 0 Then
            'MsgBox("No cell selected in the grid!", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("No cell selected in the grid!" & vbCrLf)
            Beep()
        Else
            SelRow = DataGridView2.SelectedCells.Item(0).RowIndex
            DataGridView2.Rows(SelRow).Cells(3).Value = cmbField.SelectedItem.ToString
        End If
    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStatus.SelectedIndexChanged
        'Update the Status Field entry:
        Dim SelRow As Integer
        If DataGridView2.SelectedCells.Count = 0 Then
            'MsgBox("No cell selected in the grid!", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("No cell selected in the grid!" & vbCrLf)
            Beep()
        Else
            SelRow = DataGridView2.SelectedCells.Item(0).RowIndex
            DataGridView2.Rows(SelRow).Cells(4).Value = cmbStatus.SelectedItem.ToString
        End If
    End Sub

    Private Sub cmbTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTable.SelectedIndexChanged
        'The selected table has been changed:

        'Update the list of avaialble fields:
        FillCmbField()

        'Update the destination table entry:
        Dim SelRow As Integer
        If DataGridView2.SelectedCells.Count = 0 Then
            'MsgBox("No cell selected in the grid!", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("No cell selected in the grid!" & vbCrLf)
            Beep()
        Else
            SelRow = DataGridView2.SelectedCells.Item(0).RowIndex
            DataGridView2.Rows(SelRow).Cells(2).Value = cmbTable.SelectedItem.ToString
        End If
    End Sub

    Private Sub btnAddRow_Click(sender As Object, e As EventArgs) Handles btnAddRow.Click
        'Add a row to DataGridView2

        Dim SelRow As Integer
        If DataGridView2.Rows.Count = 0 Then
            DataGridView2.Rows.Insert(0)
            DataGridView3.Rows.Insert(0)
            Import.DbDestInsertBlank(0)
        Else
            If DataGridView2.SelectedCells.Count = 0 Then
                'MsgBox("A row must be selected on the Databse Destination grid!", MsgBoxStyle.Information, "Notice")
                Message.AddWarning("A row must be selected on the Databse Destination grid!" & vbCrLf)
            Else
                SelRow = DataGridView2.SelectedCells.Item(0).RowIndex
                If DataGridView2.Rows(SelRow).IsNewRow = True Then 'Uncommited row - new row cannot be appended
                    'MsgBox("Uncommitted row - new row cannot be appended. You may need to enter a RegEx Variable name!", MsgBoxStyle.Information, "Notice")
                    Message.AddWarning("Uncommitted row - new row cannot be appended. You may need to enter a RegEx Variable name!" & vbCrLf)
                Else
                    DataGridView2.Rows.Insert(SelRow + 1)
                    DataGridView3.Rows.Insert(SelRow + 1)
                    'TextToDatabase.DbDestInsertBlank(SelRow + 1)
                    Import.DbDestInsertBlank(SelRow + 1)
                End If
            End If
        End If

    End Sub

    Private Sub btnDeleteRow_Click(sender As Object, e As EventArgs) Handles btnDeleteRow.Click
        'Delete a row from DataGridView2

        Dim SelRow As Integer

        SelRow = DataGridView2.SelectedCells.Item(0).RowIndex
        If DataGridView2.Rows(SelRow).IsNewRow = True Then 'Uncommited row - cannot be deleted
            'MsgBox("Uncommitted row - cannot be deleted", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("Uncommitted row - cannot be deleted" & vbCrLf)
        Else
            DataGridView2.Rows.RemoveAt(SelRow)
            DataGridView3.Rows.RemoveAt(SelRow)
            'Now update the Database Destination data in Import:
            'TextToDatabase.DbDestDelete(SelRow)
            Import.DbDestDelete(SelRow)
        End If

        If DbDestIndex > DataGridView2.RowCount - 1 Then 'The current Index is pointing past the last row in DataGridView2
            DbDestIndex = DataGridView2.RowCount - 1 '
        End If

    End Sub

    'Private Sub DataGridView2_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView2.CellBeginEdit
    '    Debug.Print("DataGridView2.CellBeginEdit")
    'End Sub

    'Private Sub DataGridView2_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView2.RowsAdded
    '    Debug.Print("DataGridView2.RowsAdded")
    'End Sub

    Private Sub chkUseNullValueString_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseNullValueString.CheckedChanged
        'If True, a null value string, such as "N/A" or "--", is used in the FieldValues to indicate a null value.
        If chkUseNullValueString.Checked = True Then
            Import.UseNullValueString = True
        Else
            Import.UseNullValueString = False
        End If
    End Sub

    Private Sub txtNullValueString_LostFocus(sender As Object, e As EventArgs) Handles txtNullValueString.LostFocus
        'Update the NullValueString property in Import.
        Import.NullValueString = txtNullValueString.Text
    End Sub

#End Region 'Locations Tab ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Modify Tab"

    Private Sub btnApplyModify_Click(sender As Object, e As EventArgs) Handles btnApplyModify.Click
        ApplyModifyValues()
    End Sub

    Public Sub ApplyModifyValues()
        'Apply the modification to the Database Destinations table:

        Import.ModifyValuesRegExVariable = txtRegExVariable.Text

        Select Case ModifyValueType
            Case Import.ModifyValuesTypes.AppendCurrentDate
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentDate
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.AppendCurrentDateTime
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentDateTime
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.AppendCurrentTime
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentTime
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.AppendFileDir
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFileDir
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.AppendFileName
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFileName
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.AppendFilePath
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFilePath
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.AppendFixedValue
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFixedValue
                Import.ModifyValuesFixedValue = txtFixedValue.Text
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.AppendRegExVarValue
                Import.ModifyValuesType = Import.ModifyValuesTypes.AppendRegExVarValue
                'Import.ModifyValuesRegExVarToAppend = txtAppendRegExVar.Text
                Import.ModifyValuesRegExVarValFrom = txtAppendRegExVar.Text
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.ClearValue
                Import.ModifyValuesType = Import.ModifyValuesTypes.ClearValue
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.ConvertDate
                Import.ModifyValuesType = Import.ModifyValuesTypes.ConvertDate
                Import.ModifyValuesInputDateFormat = txtInputDateFormat.Text
                Import.ModifyValuesOutputDateFormat = txtOutputDateFormat.Text
                Import.ModifyValuesApply()

            Case Import.ModifyValuesTypes.ReplaceChars
                Import.ModifyValuesType = Import.ModifyValuesTypes.ReplaceChars
                Import.ModifyValuesCharsToReplace = txtCharsToReplace.Text
                Import.ModifyValuesReplacementChars = txtReplacementChars.Text
                Import.ModifyValuesApply()
            Case Else
                Message.AddWarning("Unknown modification type: " & ModifyValueType.ToString)

        End Select

        UpdateMatches()

    End Sub

    Public Sub ApplyModifyValues_Old()
        'Apply the modification to the Database Destinations table:

        'If IsNothing(TextToDatabase.DbDestinations) Then

        'Else
        Dim I As Integer
        Dim strVarName As String 'This is used to hold the RegEx variable name in the current row
        Dim CapCount As Integer
        Dim J As Integer
        'Find matching RegEx Variables in the Database Destinations grid:
        'For I = 1 To TextToDatabase.DbDestinations.DataGridView1.RowCount
        For I = 1 To DataGridView2.RowCount 'Processes each row in the Data Destinations grid.
            'If TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(0).Value = Nothing Then 'No RegEx variable name is specified.
            If DataGridView2.Rows(I - 1).Cells(0).Value = Nothing Then 'No RegEx variable name is specified.
            Else
                'strVarName = TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(0).Value.ToString 'The RegEx variable name in the current grid row.
                strVarName = DataGridView2.Rows(I - 1).Cells(0).Value.ToString 'The RegEx variable name in the current grid row.
                If strVarName = txtRegExVariable.Text Then 'The RegExVariable at the current row matches the required variable to modify
                    If DataGridView2.Rows(I - 1).Cells(1).Value.ToString = "Single Value" Then

                    ElseIf DataGridView2.Rows(I - 1).Cells(1).Value.ToString = "Multiple Value" Then

                    End If
                    'If IsNothing(TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value) Then
                    If IsNothing(DataGridView3.Rows(I - 1).Cells(0).Value) Then
                        'txtTestInputString.Text = ""
                        'There is no text to modify
                    Else
                        'txtTestInputString.Text = TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value.ToString
                        Dim OutputString As String
                        Dim InputString As String
                        'InputString = TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value.ToString
                        InputString = DataGridView3.Rows(I - 1).Cells(0).Value.ToString
                        ConvertDate(txtInputDateFormat.Text, txtOutputDateFormat.Text, InputString, OutputString)
                        'TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value = OutputString
                        DataGridView3.Rows(I - 1).Cells(0).Value = OutputString

                    End If

                End If

            End If
        Next
        'End If
    End Sub

    Private Sub ConvertDate(ByVal InputDateFormat As String, ByVal OutputDateFormat As String, ByVal DateString As String, ByRef OutputDateString As String)
        'Date string conversion

        Dim DateVal As Date
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture

        Try
            DateVal = Date.ParseExact(DateString, InputDateFormat, provider)
            If OutputDateFormat = "" Then
                OutputDateString = DateVal.ToString()
            Else
                OutputDateString = DateVal.ToString(OutputDateFormat)
            End If

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try


    End Sub

    Private Sub btnGetFirstVar_Click(sender As Object, e As EventArgs) Handles btnGetFirstVar.Click
        'Write the RegEx text match in the Database Destinations table:

        'If IsNothing(TextToDatabase.DbDestinations) Then

        'Else
        Dim I As Integer
        Dim strVarName As String 'This is used to hold the RegEx variable name in the current row
        Dim CapCount As Integer
        Dim J As Integer
        'Find matching RegEx Variables in the Database Destinations grid:
        'For I = 1 To TextToDatabase.DbDestinations.DataGridView1.RowCount
        For I = 1 To DataGridView2.RowCount
            'If TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(0).Value = Nothing Then 'No RegEx variable name is specified.
            If DataGridView2.Rows(I - 1).Cells(0).Value = Nothing Then 'No RegEx variable name is specified.
            Else
                'strVarName = TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(0).Value.ToString 'The RegEx variable name in the current grid row.
                strVarName = DataGridView2.Rows(I - 1).Cells(0).Value.ToString 'The RegEx variable name in the current grid row.
                If strVarName = txtRegExVariable.Text Then 'The RegExVariable at the current row matches the required variable to modify
                    'If IsNothing(TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value) Then
                    If IsNothing(DataGridView3.Rows(I - 1).Cells(0).Value) Then
                        txtTestInputString.Text = ""
                    Else
                        'txtTestInputString.Text = TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value.ToString
                        txtTestInputString.Text = DataGridView3.Rows(I - 1).Cells(0).Value.ToString
                    End If

                End If
                'CapCount = myMatch.Groups(strVarName).Captures.Count 'Get the number of captures
                'If CapCount = 1 Then
                '    TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value = myMatch.Groups(strVarName).ToString
                '    rtbMessages.AppendText("RegEx Variable matched: " & strVarName & vbCrLf)
                'ElseIf CapCount > 1 Then 'Multiple matches found.
                '    If TextToDatabase.DbDestinations.DataGridView1.ColumnCount < CapCount + 5 Then
                '        TextToDatabase.DbDestinations.DataGridView1.ColumnCount = CapCount + 5
                '        For J = 1 To CapCount
                '            TextToDatabase.DbDestinations.DataGridView1.Columns(J + 4).HeaderText = "Value" & Str(J)
                '        Next
                '    End If
                '    For J = 1 To CapCount
                '        TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(J + 4).Value = myMatch.Groups(strVarName).Captures.Item(J - 1).ToString
                '    Next
                'Else 'No matches found.

                'End If
            End If
        Next
        'End If
    End Sub

    Private Sub btnGetNextVar_Click(sender As Object, e As EventArgs) Handles btnGetNextVar.Click

    End Sub

    Private Sub btnModifyTest_Click(sender As Object, e As EventArgs) Handles btnModifyTest.Click
        'Test modify values settings:

        If rbConvertDate.Checked = True Then
            TestDateConversion()
        ElseIf rbReplaceChars.Checked = True Then
            TestReplaceChars()
        ElseIf rbClearValue.Checked = True Then
            txtTestOutputString.Text = txtFixedValue.Text
        Else

        End If
    End Sub

    Private Sub TestDateConversion()
        Dim dateString As String
        Dim format As String
        Dim result As Date
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture

        dateString = txtTestInputString.Text
        format = txtInputDateFormat.Text

        Try
            result = Date.ParseExact(dateString, format, provider)
            If txtOutputDateFormat.Text = "" Then
                txtTestOutputString.Text = result.ToString()
            Else
                txtTestOutputString.Text = result.ToString(txtOutputDateFormat.Text)
            End If

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub

    Private Sub TestReplaceChars()
        Dim charsToReplace As String
        Dim replacementChars As String
        Dim InputString As String

        charsToReplace = txtCharsToReplace.Text
        replacementChars = txtReplacementChars.Text

        InputString = txtTestInputString.Text
        txtTestOutputString.Text = InputString.Replace(charsToReplace, replacementChars)


    End Sub

    Private Sub rbConvertDate_Click(sender As Object, e As EventArgs) Handles rbConvertDate.Click
        If rbConvertDate.Checked Then
            'Import.ModifyValuesType =  
            Import.ModifyValuesType = Import.ModifyValuesTypes.ConvertDate
        End If
    End Sub

    Private Sub rbReplaceChars_Click(sender As Object, e As EventArgs) Handles rbReplaceChars.Click
        If rbReplaceChars.Checked Then
            Import.ModifyValuesType = Import.ModifyValuesTypes.ReplaceChars
        End If
    End Sub

    Private Sub rbFixedValue_Click(sender As Object, e As EventArgs) Handles rbClearValue.Click
        If rbClearValue.Checked Then
            'Import.ModifyValuesType = Import.ModifyValuesTypes.FixedValue
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFixedValue
        End If
    End Sub

    Private Sub rbTextFileName_Click(sender As Object, e As EventArgs)
        If rbAppendTextFileName.Checked Then
            'Import.ModifyValuesType = Import.ModifyValuesTypes.FileName
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFileName
        End If
    End Sub

    Private Sub rbTextFileDirectory_Click(sender As Object, e As EventArgs)
        If rbAppendTextFileDirectory.Checked Then
            'Import.ModifyValuesType = Import.ModifyValuesTypes.FileDir
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFileDir
        End If
    End Sub

    Private Sub rbTextFilePath_Click(sender As Object, e As EventArgs)
        If rbAppendTextFilePath.Checked Then
            'Import.ModifyValuesType = Import.ModifyValuesTypes.FilePath
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFilePath
        End If
    End Sub

    Private Sub rbCurrentDate_Click(sender As Object, e As EventArgs)
        If rbAppendCurrentDate.Checked Then
            'Import.ModifyValuesType = Import.ModifyValuesTypes.CurrentDate
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentDate
        End If
    End Sub

    Private Sub rbCurrentTime_Click(sender As Object, e As EventArgs)
        If rbAppendCurrentTime.Checked Then
            'Import.ModifyValuesType = Import.ModifyValuesTypes.CurrentTime
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentTime
        End If
    End Sub

    Private Sub rbCurrentDateTime_Click(sender As Object, e As EventArgs)
        If rbAppendCurrentDateTime.Checked Then
            'Import.ModifyValuesType = Import.ModifyValuesTypes.CurrentDateTime
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentDateTime
        End If
    End Sub

    Private Sub txtRegExVariable_LostFocus(sender As Object, e As EventArgs) Handles txtRegExVariable.LostFocus
        Import.ModifyValuesRegExVariable = txtRegExVariable.Text
    End Sub

    Private Sub txtInputDateFormat_LostFocus(sender As Object, e As EventArgs) Handles txtInputDateFormat.LostFocus
        Import.ModifyValuesInputDateFormat = txtInputDateFormat.Text
    End Sub

    Private Sub txtOutputDateFormat_LostFocus(sender As Object, e As EventArgs) Handles txtOutputDateFormat.LostFocus
        Import.ModifyValuesOutputDateFormat = txtOutputDateFormat.Text
    End Sub

    Private Sub txtCharsToReplace_LostFocus(sender As Object, e As EventArgs) Handles txtCharsToReplace.LostFocus
        Import.ModifyValuesCharsToReplace = txtCharsToReplace.Text
    End Sub

    Private Sub txtReplacementChars_LostFocus(sender As Object, e As EventArgs) Handles txtReplacementChars.LostFocus
        Import.ModifyValuesReplacementChars = txtReplacementChars.Text
    End Sub

    Private Sub txtFixedValue_LostFocus(sender As Object, e As EventArgs)
        Import.ModifyValuesFixedValue = txtFixedValue.Text
    End Sub

    Private Sub btnAddModifyToSeq_Click(sender As Object, e As EventArgs) Handles btnAddModifyToSeq.Click
        'Save the Modify Values setting in the current Processing Sequence

        Dim I As Integer

        If IsNothing(Sequence) Then
            'MessageBox.Show("The Sequence form is not open.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Message.AddWarning("The Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart
            Sequence.rtbSequence.SelectedText = "<ModifyValues>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <RegExVariable>" & Trim(txtRegExVariable.Text) & "</RegExVariable>" & vbCrLf

            If rbClearValue.Checked Then
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Clear_value</ModifyType>" & vbCrLf
            ElseIf rbConvertDate.Checked Then
                Sequence.rtbSequence.SelectedText = "  <InputDateFormat>" & Trim(txtInputDateFormat.Text) & "</InputDateFormat>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <OutputDateFormat>" & Trim(txtOutputDateFormat.Text) & "</OutputDateFormat>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Convert_date</ModifyType>" & vbCrLf
            ElseIf rbReplaceChars.Checked Then
                Sequence.rtbSequence.SelectedText = "  <CharactersToReplace>" & Trim(txtCharsToReplace.Text) & "</CharactersToReplace>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ReplacementCharacters>" & Trim(txtReplacementChars.Text) & "</ReplacementCharacters>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Replace_characters</ModifyType>" & vbCrLf
                'ElseIf rbClearValue.Checked Then
            ElseIf rbAppendFixedValue.Checked Then
                Sequence.rtbSequence.SelectedText = "  <FixedValue>" & txtFixedValue.Text & "</FixedValue>" & vbCrLf
                'Sequence.rtbSequence.SelectedText = "  <ModifyType>Fixed_value</ModifyType>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_fixed_value</ModifyType>" & vbCrLf
            ElseIf rbAppendRegExVar.Checked Then
                'Sequence.rtbSequence.SelectedText = "  <RegExVariableValue>" & Trim(txtAppendRegExVar.Text) & "</RegExVariableValue>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <RegExVariableValueFrom>" & Trim(txtAppendRegExVar.Text) & "</RegExVariableValueFrom>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_RegEx_variable_value</ModifyType>" & vbCrLf
            ElseIf rbAppendTextFileName.Checked Then
                'Sequence.rtbSequence.SelectedText = "  <ModifyType>Text_file_name</ModifyType>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_file_name</ModifyType>" & vbCrLf
            ElseIf rbAppendTextFileDirectory.Checked Then
                'Sequence.rtbSequence.SelectedText = "  <ModifyType>Text_file_directory</ModifyType>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_file_directory</ModifyType>" & vbCrLf
            ElseIf rbAppendTextFilePath.Checked Then
                'Sequence.rtbSequence.SelectedText = "  <ModifyType>Text_file_path</ModifyType>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_file_path</ModifyType>" & vbCrLf
            ElseIf rbAppendCurrentDate.Checked Then
                'Sequence.rtbSequence.SelectedText = "  <ModifyType>Current_date</ModifyType>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_current_date</ModifyType>" & vbCrLf
            ElseIf rbAppendCurrentTime.Checked Then
                'Sequence.rtbSequence.SelectedText = "  <ModifyType>Current_time</ModifyType>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_current_time</ModifyType>" & vbCrLf
            ElseIf rbAppendCurrentDateTime.Checked Then
                'Sequence.rtbSequence.SelectedText = "  <ModifyType>Current_date_time</ModifyType>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_current_date_time</ModifyType>" & vbCrLf
            End If

                Sequence.rtbSequence.SelectedText = "</ModifyValues>" & vbCrLf
                'TDS_Import.Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

                Sequence.FormatXmlText()

            End If
    End Sub

    Private Sub rbClearValue_CheckedChanged(sender As Object, e As EventArgs) Handles rbClearValue.CheckedChanged
        If rbClearValue.Checked Then
            txtModType.Text = "Clear the value"
            ModifyValueType = Import.ModifyValuesTypes.ClearValue
        End If
    End Sub

    Private Sub rbConvertDate_CheckedChanged(sender As Object, e As EventArgs) Handles rbConvertDate.CheckedChanged
        If rbConvertDate.Checked Then
            txtModType.Text = "Convert date"
            ModifyValueType = Import.ModifyValuesTypes.ConvertDate
        End If
    End Sub

    Private Sub rbReplaceChars_CheckedChanged(sender As Object, e As EventArgs) Handles rbReplaceChars.CheckedChanged
        If rbReplaceChars.Checked Then
            txtModType.Text = "Replace characters"
            ModifyValueType = Import.ModifyValuesTypes.ReplaceChars
        End If
    End Sub


    Private Sub rbAppendFixedValue_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendFixedValue.CheckedChanged
        If rbAppendFixedValue.Checked Then
            txtModType.Text = "Append with fixed value"
            ModifyValueType = Import.ModifyValuesTypes.AppendFixedValue
        End If
    End Sub

    Private Sub rbAppendRegExVar_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendRegExVar.CheckedChanged
        If rbAppendRegExVar.Checked Then
            txtModType.Text = "Append RegEx variable value"
            ModifyValueType = Import.ModifyValuesTypes.AppendRegExVarValue
        End If
    End Sub

    Private Sub rbAppendTextFileName_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendTextFileName.CheckedChanged
        If rbAppendTextFileName.Checked Then
            txtModType.Text = "Append file name"
            ModifyValueType = Import.ModifyValuesTypes.AppendFileName
        End If
    End Sub

    Private Sub rbAppendTextFileDirectory_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendTextFileDirectory.CheckedChanged
        If rbAppendTextFileDirectory.Checked Then
            txtModType.Text = "Append file directory"
            ModifyValueType = Import.ModifyValuesTypes.AppendFileDir
        End If
    End Sub

    Private Sub rbAppendTextFilePath_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendTextFilePath.CheckedChanged
        If rbAppendTextFilePath.Checked Then
            txtModType.Text = "Append file path"
            ModifyValueType = Import.ModifyValuesTypes.AppendFilePath
        End If
    End Sub

    Private Sub rbAppendCurrentDate_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendCurrentDate.CheckedChanged
        If rbAppendCurrentDate.Checked Then
            txtModType.Text = "Append current date"
            ModifyValueType = Import.ModifyValuesTypes.AppendCurrentDate
        End If
    End Sub

    Private Sub rbAppendCurrentTime_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendCurrentTime.CheckedChanged
        If rbAppendCurrentTime.Checked Then
            txtModType.Text = "Append current time"
            ModifyValueType = Import.ModifyValuesTypes.AppendCurrentTime
        End If
    End Sub

    Private Sub rbAppendCurrentDateTime_CheckedChanged(sender As Object, e As EventArgs) Handles rbAppendCurrentDateTime.CheckedChanged
        If rbAppendCurrentDateTime.Checked Then
            txtModType.Text = "Append current date/time"
            ModifyValueType = Import.ModifyValuesTypes.AppendCurrentDateTime
        End If
    End Sub


#End Region 'Modify Tab ------------------------------------------------------------------------------------------------------------------------------------------------------------------------



#Region "Multipliers Tab" '------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnMultMoveUp_Click(sender As Object, e As EventArgs) Handles btnMultMoveUp.Click
        'Move grid entry up:

        Dim SelRow As Integer

        Dim TempRow() As String
        Dim I As Integer 'Loop Index
        Dim NCols As Integer 'Number of columns in the GridView

        If DataGridView4.SelectedCells.Count = 0 Then
            'MsgBox("No cell selected in the grid!", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("No cell selected in the grid!" & vbCrLf)
        Else
            SelRow = DataGridView4.SelectedCells.Item(0).RowIndex
            If SelRow = 0 Then 'At first row.
            Else
                'Move row up:
                NCols = DataGridView4.ColumnCount

                ReDim TempRow(0 To NCols - 1)
                For I = 0 To NCols - 1
                    TempRow(I) = DataGridView4.Rows(SelRow - 1).Cells(I).Value
                    DataGridView4.Rows(SelRow - 1).Cells(I).Value = DataGridView4.Rows(SelRow).Cells(I).Value
                    DataGridView4.Rows(SelRow).Cells(I).Value = TempRow(I)
                Next
                SelRow = SelRow - 1
                DataGridView4.ClearSelection()
                DataGridView4.Rows.Item(SelRow).Selected = True
            End If
        End If
    End Sub

    Private Sub btnMultMoveDown_Click(sender As Object, e As EventArgs) Handles btnMultMoveDown.Click
        'Move grid entry down:

        Dim SelRow As Integer
        Dim TempRow() As String
        Dim I As Integer 'Loop Index
        Dim NCols As Integer 'Number of columns in the GridView

        If DataGridView4.SelectedCells.Count = 0 Then
            'MsgBox("No cell selected in the grid!", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("No cell selected in the grid!" & vbCrLf)
        Else
            SelRow = DataGridView4.SelectedCells.Item(0).RowIndex
            If SelRow = DataGridView4.Rows.Count - 1 Then 'At last row.
            Else
                'Move row down:
                NCols = DataGridView4.ColumnCount
                ReDim TempRow(0 To NCols - 1)
                For I = 0 To NCols - 1
                    TempRow(I) = DataGridView4.Rows(SelRow + 1).Cells(I).Value
                    DataGridView4.Rows(SelRow + 1).Cells(I).Value = DataGridView4.Rows(SelRow).Cells(I).Value
                    DataGridView4.Rows(SelRow).Cells(I).Value = TempRow(I)
                Next
                'Change selected row:
                SelRow = SelRow + 1
                DataGridView4.ClearSelection()
                DataGridView4.Rows.Item(SelRow).Selected = True
            End If
        End If
    End Sub

    Private Sub btnMultAddRow_Click(sender As Object, e As EventArgs) Handles btnMultAddRow.Click
        'Add a row to the DataGridView4

        Dim SelRow As Integer

        SelRow = DataGridView4.SelectedCells.Item(0).RowIndex
        If DataGridView4.Rows(SelRow).IsNewRow = True Then 'Uncommited row - new row cannot be appended
            'MsgBox("Uncommitted row - new row cannot be appended", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("Uncommitted row - new row cannot be appended" & vbCrLf)
        Else
            DataGridView4.Rows.Insert(SelRow + 1)
        End If
    End Sub

    Private Sub btnMultDeleteRow_Click(sender As Object, e As EventArgs) Handles btnMultDeleteRow.Click
        'Delete a row from the DataGridView4

        Dim SelRow As Integer

        SelRow = DataGridView4.SelectedCells.Item(0).RowIndex
        If DataGridView4.Rows(SelRow).IsNewRow = True Then 'Uncommited row - cannot be deleted
            'MsgBox("Uncommitted row - cannot be deleted", MsgBoxStyle.Information, "Notice")
            Message.AddWarning("Uncommitted row - cannot be deleted" & vbCrLf)
        Else
            DataGridView4.Rows.RemoveAt(SelRow)
        End If
    End Sub

    Private Sub cmbMultVariable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMultVariable.SelectedIndexChanged
        'Update the Value Type entry:
        Dim SelRow As Integer
        If DataGridView4.SelectedCells.Count = 0 Then
            MsgBox("No cell selected in the grid!", MsgBoxStyle.Information, "Notice")
        Else
            SelRow = DataGridView4.SelectedCells.Item(0).RowIndex
            If DataGridView4.Rows(SelRow).IsNewRow = True Then 'Uncommited row - new row cannot be appended
                DataGridView4.Rows.Insert(SelRow)
            Else
                DataGridView4.Rows.Insert(SelRow)
            End If
            DataGridView4.Rows(SelRow).Cells(0).Value = cmbVariable.SelectedItem.ToString
        End If
    End Sub

    Private Sub DataGridView4_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DataGridView4.ColumnWidthChanged
        cmbMultVariable.Left = 50
        cmbMultVariable.Width = DataGridView4.Columns(0).Width
        cmbMultCode.Left = cmbMultVariable.Left + cmbMultVariable.Width
        cmbMultCode.Width = DataGridView4.Columns(1).Width
    End Sub

#End Region 'Multipliers Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Online/Offline code" '-------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnOnline_Click(sender As Object, e As EventArgs) Handles btnOnline.Click
        'Connect to or disconnect from the Application Network.
        If ConnectedToAppnet = False Then
            ConnectToAppNet()
        Else
            DisconnectFromAppNet()
        End If
    End Sub

    Private Sub ConnectToAppNet()
        'Connect to the Application Network. (Message Exchange)

        Dim Result As Boolean

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.SetWarningStyle()
            Message.Add("client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds

                Result = client.Connect(ApplicationInfo.Name, ServiceReference1.clsConnectionAppTypes.Application, False, False) 'Application Name is "Application_Template"
                'appName, appType, getAllWarnings, getAllMessages

                If Result = True Then
                    Message.Add("Connected to the Application Network as " & ApplicationInfo.Name & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToAppnet = True
                    SendApplicationInfo()
                Else
                    Message.Add("Connection to the Application Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Application Network is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If

    End Sub


    Private Sub DisconnectFromAppNet()
        'Disconnect from the Application Network.

        Dim Result As Boolean

        If IsNothing(client) Then
            Message.Add("Already disconnected from the Application Network." & vbCrLf)
            btnOnline.Text = "Offline"
            btnOnline.ForeColor = Color.Red
            ConnectedToAppnet = False
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted." & vbCrLf)
            Else
                Try
                    Message.Add("Running client.Disconnect(ApplicationName)   ApplicationName = " & ApplicationInfo.Name & vbCrLf)
                    client.Disconnect(ApplicationInfo.Name) 'NOTE: If Application Network has closed, this application freezes at this line! Try Catch EndTry added to fix this.
                    btnOnline.Text = "Offline"
                    btnOnline.ForeColor = Color.Red
                    ConnectedToAppnet = False
                Catch ex As Exception
                    Message.SetWarningStyle()
                    Message.Add("Error disconnecting from Application Network: " & ex.Message & vbCrLf)
                End Try
            End If
        End If
    End Sub

    Private Sub SendApplicationInfo()
        'Send the application information to the Administrator connections.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to send application information.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                Dim applicationInfo As New XElement("ApplicationInfo")
                Dim name As New XElement("Name", Me.ApplicationInfo.Name)
                applicationInfo.Add(name)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)
                client.SendMessage("ApplicationNetwork", doc.ToString)
            End If
        End If

    End Sub

#End Region 'Online/Offline code ------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Process XMessages" '---------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub XMsg_Instruction(Info As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville (TM) applications.
        '
        'An XSequence file is an AL-H7 (TM) Information Vector Sequence stored in an XML format.
        'AL-H7(TM) is the name of a programming system that uses sequences of information and location value pairs to store data items or processing steps.
        'A single information and location value pair is called a knowledge element (or noxel).
        'Any program, mathematical expression or data set can be expressed as an Information Vector Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville(TM) applciations for examples.

        Select Case Locn

            Case "EndOfSequence"
                'End of Information Vector Sequence reached.

            Case Else
                Message.SetWarningStyle()
                Message.Add("Unknown location: " & Locn & vbCrLf)
                Message.SetNormalStyle()

        End Select

    End Sub

    Private Sub SendMessage()
        'Code used to send a message after a timer delay.
        'The message destination is stored in MessageDest
        'The message text is stored in MessageText
        Timer1.Interval = 100 '100ms delay
        Timer1.Enabled = True 'Start the timer.
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If IsNothing(client) Then
            'Message.SetWarningStyle()
            'Message.Add("No client connection available!" & vbCrLf)
            Message.AddWarning("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                'Message.SetWarningStyle()
                'Message.Add("client state is faulted. Message not sent!" & vbCrLf)
                Message.AddWarning("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Try
                    Message.Add("Sending a message. Number of characters: " & MessageText.Length & vbCrLf)
                    'client.SendMessage(MessageDest, MessageText)
                    client.SendMessage(ClientAppName, MessageText)
                    'Message.XAdd(MessageText & vbCrLf) 'NOTE this is displayed in Property InstrReceived
                    MessageText = "" 'Clear the message after it has been sent.
                    ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
                    ClientAppLocn = "" 'Clear the Client Application Location after the message has been sent.
                Catch ex As Exception
                    'Message.SetWarningStyle()
                    'Message.Add("Error sending message: " & ex.Message & vbCrLf)
                    Message.AddWarning("Error sending message: " & ex.Message & vbCrLf)
                End Try
            End If
        End If

        'Stop timer:
        Timer1.Enabled = False
    End Sub

    Private Sub btnGridProp_Click(sender As Object, e As EventArgs) Handles btnGridProp.Click


        ''Show DbDest() array.
        'rtbDebugging.Text = "Contents of the GridProp() array:" & vbCrLf
        'Dim NRows As Integer = Import.DbDestCount
        'rtbDebugging.AppendText("Number of Rows: ")



    End Sub

    Private Sub btnDbDest_Click(sender As Object, e As EventArgs) Handles btnDbDest.Click
        rtbDebugging.Text = Import.ReturnDbDestData
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

    End Sub

    Private Sub txtNullValueString_TextChanged(sender As Object, e As EventArgs) Handles txtNullValueString.TextChanged

    End Sub























































































#End Region 'Process XMessages --------------------------------------------------------------------------------------------------------------------------------------------------------

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------







End Class
