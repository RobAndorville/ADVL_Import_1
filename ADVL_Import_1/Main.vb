﻿'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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

Imports System.Security.Permissions
<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
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
    Public WithEvents ShowTextFile As frmShowTextFile
    Public WithEvents Sequence As frmImportSequence
    Public WithEvents ClipboardWindow As frmClipboardWindow

    Public WithEvents WebPageList As frmWebPageList

    Public WithEvents NewHtmlDisplay As frmHtmlDisplay
    Public HtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents NewWebPage As frmWebPage
    Public WebPageFormList As New ArrayList 'Used for displaying multiple WebView forms.

    'Declare objects used to connect to the Application Network:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppNetName As String = "" 'The name of thge client Application Network requesting service. ADDED 2Feb19.
    Dim ClientAppName As String = "" 'The name of the client requesting service
    Dim ClientConnName As String = "" 'The name of the client connection requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    Public ConnectionName As String = "" 'The name of the connection used to connect this application to the ComNet.
    Public AppNetName As String = "" 'Added 2Feb19

    Public MsgServiceAppPath As String = "" 'The application path of the Message Service application (ComNet). This is where the "Application.Lock" file will be while ComNet is running
    Public MsgServiceExePath As String = "" 'The executable path of the Message Service.

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

    'Main.Load variables:
    Dim ProjectSelected As Boolean = False 'If True, a project has been selected using Command Arguments. Used in Main.Load.
    Dim StartupConnectionName As String = "" 'If not "" the application will be connected to the AppNet using this connection name in  Main.Load.

    'The following variables are used to run JavaScript in Web Pages loaded into the Document View: -------------------
    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence
    'To run an XSequence:
    '  XSeq.RunXSequence(xDoc, Status) 'ImportStatus in Import
    '    Handle events:
    '      XSeq.ErrorMsg
    '      XSeq.Instruction(Info, Locn)

    Private XStatus As New System.Collections.Specialized.StringCollection

    'Variables used to restore Item values on a web page.
    Private FormName As String
    Private ItemName As String
    Private SelectId As String

    'StartProject variables:
    Private StartProject_AppName As String  'The application name
    Private StartProject_ConnName As String 'The connection name
    Private StartProject_ProjID As String   'The project ID

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

    Private _connectedToComNet As Boolean = False  'True if the application is connected to the Communication Network (Message Service).
    Property ConnectedToComNet As Boolean
        Get
            Return _connectedToComNet
        End Get
        Set(value As Boolean)
            _connectedToComNet = value
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
                Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
                If _instrReceived.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
                    Try
                        'Inititalise the reply message:
                        Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                        MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                        xmessage = New XElement("XMsg")
                        xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                        'Run the received message:
                        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                        XDoc.LoadXml(XmlHeader & vbCrLf & _instrReceived)
                        Message.XAddXml(XDoc)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                        XMsg.Run(XDoc, Status)
                    Catch ex As Exception
                        Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
                    End Try

                    'XMessage has been run.
                    'Reply to this message:
                    'Add the message reply to the XMessages window:
                    'Complete the MessageXDoc:
                    xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the last location reply instructions to the message.
                    MessageXDoc.Add(xmessage)
                    MessageText = MessageXDoc.ToString

                    If ClientAppName = "" Then
                        'No client to send a message to!
                    Else
                        If MessageText = "" Then
                            'No message to send!
                        Else
                            Message.XAddText("Message sent to " & ClientAppName & ":" & vbCrLf, "XmlSentNotice")
                            Message.XAddXml(MessageText)
                            Message.XAddText(vbCrLf, "Normal") 'Add extra line
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

    Private _closedFormNo As Integer 'Temporarily holds the number of the form that is being closed. 
    Property ClosedFormNo As Integer
        Get
            Return _closedFormNo
        End Get
        Set(value As Integer)
            _closedFormNo = value
        End Set
    End Property

    Private _startPageFileName As String = "" 'The file name of the html document displayed in the Start Page tab.
    Public Property StartPageFileName As String
        Get
            Return _startPageFileName
        End Get
        Set(value As String)
            _startPageFileName = value
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
                               <MsgServiceAppPath><%= MsgServiceAppPath %></MsgServiceAppPath>
                               <MsgServiceExePath><%= MsgServiceExePath %></MsgServiceExePath>
                               <!---->
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

            If Settings.<FormSettings>.<MsgServiceAppPath>.Value <> Nothing Then MsgServiceAppPath = Settings.<FormSettings>.<MsgServiceAppPath>.Value
            If Settings.<FormSettings>.<MsgServiceExePath>.Value <> Nothing Then MsgServiceExePath = Settings.<FormSettings>.<MsgServiceExePath>.Value

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
            End If
        End If

        ReadApplicationInfo()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()

        'Restore Project information: -------------------------------------------------------
        Project.ApplicationName = ApplicationInfo.Name

        'Set up the Message object:
        Message.ApplicationName = ApplicationInfo.Name

        'Set up a temporary initial settings location:
        Dim TempLocn As New ADVL_Utilities_Library_1.FileLocation
        TempLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        TempLocn.Path = ApplicationInfo.ApplicationDir
        Message.SettingsLocn = TempLocn

        Me.Show() 'Show this form before showing the Message form - This will show the App icon on top in the TaskBar.

        'Start showing messages here - Message system is set up.
        Message.AddText("------------------- Starting Application: ADVL Import ----------------- " & vbCrLf, "Heading")
        Message.AddText("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#.##") & " hours" & vbCrLf, "Normal")

        'https://msdn.microsoft.com/en-us/library/z2d603cy(v=vs.80).aspx#Y550
        'Process any command line arguments:
        Try
            For Each s As String In My.Application.CommandLineArgs
                Message.Add("Command line argument: " & vbCrLf)
                Message.AddXml(s & vbCrLf & vbCrLf)
                InstrReceived = s
            Next
        Catch ex As Exception
            Message.AddWarning("Error processing command line arguments: " & ex.Message & vbCrLf)
        End Try

        If ProjectSelected = False Then
            'Read the Settings Location for the last project used:
            Project.ReadLastProjectInfo()
            'The Last_Project_Info.xml file contains:
            '  Project Name and Description. Settings Location Type and Settings Location Path.
            Message.Add("Last project info has been read." & vbCrLf)
            Message.Add("Project.Type.ToString  " & Project.Type.ToString & vbCrLf)
            Message.Add("Project.Path  " & Project.Path & vbCrLf)

            'At this point read the application start arguments, if any.
            'The selected project may be changed here.

            'Check if the project is locked:
            If Project.ProjectLocked Then
                Message.AddWarning("The project is locked: " & Project.Name & vbCrLf)
                Dim dr As System.Windows.Forms.DialogResult
                dr = MessageBox.Show("Press 'Yes' to unlock the project", "Notice", MessageBoxButtons.YesNo)
                If dr = System.Windows.Forms.DialogResult.Yes Then
                    Project.UnlockProject()
                    Message.AddWarning("The project has been unlocked: " & Project.Name & vbCrLf)
                    'Read the Project Information file: -------------------------------------------------
                    Message.Add("Reading project info." & vbCrLf)
                    Project.ReadProjectInfoFile()   'Read the file in the SettingsLocation: ADVL_Project_Info.xml

                    Project.ReadParameters()
                    Project.ReadParentParameters()
                    If Project.ParentParameterExists("AppNetName") Then
                        Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
                        AppNetName = Project.Parameter("AppNetName").Value
                    Else
                        AppNetName = Project.GetParameter("AppNetName")
                    End If

                    Project.LockProject() 'Lock the project while it is open in this application.
                    'Set the project start time. This is used to track project usage.
                    Project.Usage.StartTime = Now
                    ApplicationInfo.SettingsLocn = Project.SettingsLocn
                    'Set up the Message object:
                    Message.SettingsLocn = Project.SettingsLocn
                Else
                    'Continue without any project selected.
                    Project.Name = ""
                    Project.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Project.Description = ""
                    Project.SettingsLocn.Path = ""
                    Project.DataLocn.Path = ""
                End If
            Else
                'Read the Project Information file: -------------------------------------------------
                Message.Add("Reading project info." & vbCrLf)
                Project.ReadProjectInfoFile()  'Read the file in the SettingsLocation: ADVL_Project_Info.xml

                'ADDED 2Feb19:
                Project.ReadParameters()
                Project.ReadParentParameters()
                If Project.ParentParameterExists("AppNetName") Then
                    Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
                    AppNetName = Project.Parameter("AppNetName").Value
                Else
                    AppNetName = Project.GetParameter("AppNetName")
                End If

                Project.LockProject() 'Lock the project while it is open in this application.
                'Set the project start time. This is used to track project usage.
                Project.Usage.StartTime = Now
                ApplicationInfo.SettingsLocn = Project.SettingsLocn
                'Set up the Message object:
                Message.SettingsLocn = Project.SettingsLocn
            End If
        Else 'Project has been opened using Command Line arguments.
            Project.ReadParameters()
            Project.ReadParentParameters()
            If Project.ParentParameterExists("AppNetName") Then
                Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
                AppNetName = Project.Parameter("AppNetName").Value
            Else
                AppNetName = Project.GetParameter("AppNetName")
            End If

            Project.LockProject() 'Lock the project while it is open in this application.

            ProjectSelected = False 'Reset the Project Selected flag.
        End If

        'START Initialise the form: ===============================================================

        Me.WebBrowser1.ObjectForScripting = Me

        InitialiseForm() 'Initialise the form for a new project.

        'END   Initialise the form: ---------------------------------------------------------------

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        RestoreFormSettings() 'Restore the form settings
        RestoreProjectSettings() 'Restore the Project settings

        ShowProjectInfo() 'Show the project information.

        Message.AddText("------------------- Started OK -------------------------------------------------------------------------- " & vbCrLf & vbCrLf, "Heading")

        Me.Show() 'Show this form before showing the Message form

        If StartupConnectionName = "" Then
            If Project.ConnectOnOpen Then
                ConnectToComNet() 'The Project is set to connect when it is opened.
            ElseIf ApplicationInfo.ConnectOnStartup Then
                ConnectToComNet() 'The Application is set to connect when it is started.
            Else
                'Don't connect to ComNet.
            End If
        Else
            'Connect to ComNet using the connection name StartupConnectionName.
            ConnectToComNet(StartupConnectionName)
        End If

        'Start the timer to keep the connection awake:
        'Timer3.Interval = 10000 '10 seconds - for testing
        Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
        Timer3.Enabled = True
        Timer3.Start()

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.

        OpenStartPage()

        Import.SettingsLocn = Project.SettingsLocn
        Import.DataLocn = Project.DataLocn
        Import.RestoreSettings()
        'Initialise all the tab forms:
        InitialiseTabs()
    End Sub

    Private Sub ShowProjectInfo()
        'Show the project information:

        txtParentProject.Text = Project.ParentProjectName
        txtAppNetName.Text = Project.GetParameter("AppNetName")
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
        txtProjectPath.Text = Project.Path
        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataPath.Text = Project.DataLocn.Path

        Select Case Project.SystemLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSystemLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSystemLocationType.Text = "Archive"
        End Select
        txtSystemPath.Text = Project.SystemLocn.Path

        If Project.ConnectOnOpen Then
            chkConnect.Checked = True
        Else
            chkConnect.Checked = False
        End If

        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        DisconnectFromComNet() 'Disconnect from the Communication Network.

        Import.SaveSettings()

        ApplicationInfo.WriteFile() 'Update the Application Information file.

        Project.SaveLastProjectInfo() 'Save information about the last project used.
        Project.SaveParameters() 'ADDED 3Feb19

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        Project.UnlockProject() 'Unlock the project.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.
        ApplicationInfo.UnlockApplication()

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


    Private Sub btnWebPages_Click(sender As Object, e As EventArgs) Handles btnWebPages.Click
        'Open the Web Pages form.

        If IsNothing(WebPageList) Then
            WebPageList = New frmWebPageList
            WebPageList.Show()
        Else
            WebPageList.Show()
            WebPageList.BringToFront()
        End If
    End Sub

    Private Sub WebPageList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles WebPageList.FormClosed
        WebPageList = Nothing
    End Sub

    Public Function OpenNewWebPage() As Integer
        'Open a new HTML Web View window, or reuse an existing one if avaiable.
        'The new forms index number in WebViewFormList is returned.

        NewWebPage = New frmWebPage
        If WebPageFormList.Count = 0 Then
            WebPageFormList.Add(NewWebPage)
            WebPageFormList(0).FormNo = 0
            WebPageFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in WebViewFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WebPageFormList.Count - 1 'Check if there are closed forms in WebViewFormList. They can be re-used.
                If IsNothing(WebPageFormList(I)) Then
                    WebPageFormList(I) = NewWebPage
                    WebPageFormList(I).FormNo = I
                    WebPageFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in WebViewFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WebViewFormList
                Dim FormNo As Integer
                WebPageFormList.Add(NewWebPage)
                FormNo = WebPageFormList.Count - 1
                WebPageFormList(FormNo).FormNo = FormNo
                WebPageFormList(FormNo).Show
                Return FormNo 'The new WebPage is at position FormNo in WebPageFormList()
            End If
        End If
    End Function

    Public Sub WebPageFormClosed()
        'This subroutine is called when the Web Page form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebPage form.
        'The WebPage form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the WebPage form.
        'This property should be updated by the WebPage form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebPageList should be set to Nothing.

        If WebPageFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebPageFormList
            Exit Sub
        End If

        If IsNothing(WebPageFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            WebPageFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Function OpenNewHtmlDisplayPage() As Integer
        'Open a new HTML display window, or reuse an existing one if avaiable.
        'The new forms index number in HtmlDisplayFormList is returned.

        NewHtmlDisplay = New frmHtmlDisplay
        If HtmlDisplayFormList.Count = 0 Then
            HtmlDisplayFormList.Add(NewHtmlDisplay)
            HtmlDisplayFormList(0).FormNo = 0
            HtmlDisplayFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in HtmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To HtmlDisplayFormList.Count - 1 'Check if there are closed forms in HtmlDisplayFormList. They can be re-used.
                If IsNothing(HtmlDisplayFormList(I)) Then
                    HtmlDisplayFormList(I) = NewHtmlDisplay
                    HtmlDisplayFormList(I).FormNo = I
                    HtmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in HtmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to HtmlDisplayFormList
                Dim FormNo As Integer
                HtmlDisplayFormList.Add(NewHtmlDisplay)
                FormNo = HtmlDisplayFormList.Count - 1
                HtmlDisplayFormList(FormNo).FormNo = FormNo
                HtmlDisplayFormList(FormNo).Show
                Return FormNo 'The new HtmlDisplay is at position FormNo in HtmlDisplayFormList()
            End If
        End If
    End Function


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

    Private Sub ApplicationInfo_RestoreDefaults() Handles ApplicationInfo.RestoreDefaults
        'Restore the default application settings.
        DefaultAppProperties()
    End Sub

    Private Sub ApplicationInfo_UpdateExePath() Handles ApplicationInfo.UpdateExePath
        'Update the Executable Path.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath
    End Sub

    Public Sub UpdateWebPage(ByVal FileName As String)
        'Update the web page in WebPageFormList if the Web file name is FileName.

        Dim NPages As Integer = WebPageFormList.Count
        Dim I As Integer

        For I = 0 To NPages - 1
            If WebPageFormList(I).FileName = FileName Then
                WebPageFormList(I).OpenDocument
            End If
        Next
    End Sub


#Region " Start Page Code" '=========================================================================================================================================

    Public Sub OpenStartPage()
        'Open the StartPage.html file and display in the Start Page tab.

        If Project.DataFileExists("StartPage.html") Then
            StartPageFileName = "StartPage.html"
            DisplayStartPage()
        Else
            CreateStartPage()
            StartPageFileName = "StartPage.html"
            DisplayStartPage()
        End If

    End Sub

    Public Sub DisplayStartPage()
        'Display the StartPage.html file in the Start Page tab.

        'If Project.DataFileExists("StartPage.html") Then
        If Project.DataFileExists(StartPageFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(StartPageFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            WebBrowser1.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & StartPageFileName & vbCrLf)
        End If
    End Sub

    Private Sub CreateStartPage()
        'Create a new default StartPage.html file.

        Dim htmData As New IO.MemoryStream
        Dim sw As New IO.StreamWriter(htmData)
        sw.Write(AppInfoHtmlString("Application Information")) 'Create a web page providing information about the application.
        sw.Flush()
        Project.SaveData("StartPage.html", htmData)
    End Sub

    Public Function AppInfoHtmlString(ByVal DocumentTitle As String) As String
        'Create an Application Information Web Page.

        'This function should be edited to provide a brief description of the Application.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & "Andorville&trade; Import Application" & "</h2>" & vbCrLf & vbCrLf) 'Add the page title.
        sb.Append("<hr>" & vbCrLf) 'Add a horizontal divider line.
        sb.Append("<p>The Import application is used to import information from text files into a database.</p>" & vbCrLf) 'Add an application description.
        sb.Append("<hr>" & vbCrLf & vbCrLf) 'Add a horizontal divider line.

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultJavaScriptString() As String
        'Generate the default JavaScript section of an Andorville(TM) Workflow Web Page.

        Dim sb As New System.Text.StringBuilder

        'Add JavaScript section:
        sb.Append("<script>" & vbCrLf & vbCrLf)

        'START: User defined JavaScript functions ==========================================================================
        'Add functions to implement the main actions performed by this web page.
        sb.Append("//START: User defined JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Add functions to implement the main actions performed by this web page." & vbCrLf & vbCrLf)

        sb.Append("//END:   User defined JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User defined JavaScript functions --------------------------------------------------------------------------


        'START: User modified JavaScript functions ==========================================================================
        'Modify these function to save all required web page settings and process all expected XMessage instructions.
        sb.Append("//START: User modified JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Modify these function to save all required web page settings and process all expected XMessage instructions." & vbCrLf & vbCrLf)

        'Add the SaveSettings function - This is used to save web page settings between sessions.
        sb.Append("//Save the web page settings." & vbCrLf)
        sb.Append("function SaveSettings() {" & vbCrLf)
        sb.Append("  var xSettings = ""<Settings>"" + "" \n"" ; //String containing the web page settings in XML format." & vbCrLf)
        sb.Append("  //Add xml lines to save each setting." & vbCrLf & vbCrLf)
        sb.Append("  xSettings +=    ""</Settings>"" + ""\n"" ; //End of the Settings element." & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  //Save the settings as an XML file in the project." & vbCrLf)
        sb.Append("  window.external.SaveHtmlSettings(xSettings) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Process a single XMsg instruction (Information:Location pair)
        sb.Append("//Process an XMessage instruction:" & vbCrLf)
        sb.Append("function XMsgInstruction(Info, Locn) {" & vbCrLf)
        sb.Append("  switch(Locn) {" & vbCrLf)
        sb.Append("  //Insert case statements here." & vbCrLf)
        sb.Append("  default:" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("  }" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   User modified JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User modified JavaScript functions --------------------------------------------------------------------------

        'START: Required Document Library Web Page JavaScript functions ==========================================================================
        sb.Append("//START: Required Document Library Web Page JavaScript functions ==========================================================================" & vbCrLf & vbCrLf)

        'Add the AddText function - This sends a message to the message window using a named text type.
        sb.Append("//Add text to the Message window using a named txt type:" & vbCrLf)
        sb.Append("function AddText(Msg, TextType) {" & vbCrLf)
        sb.Append("  window.external.AddText(Msg, TextType) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddMessage function - This sends a message to the message window using default black text.
        sb.Append("//Add a message to the Message window using the default black text:" & vbCrLf)
        sb.Append("function AddMessage(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddMessage(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddWarning function - This sends a red, bold warning message to the message window.
        sb.Append("//Add a warning message to the Message window using bold red text:" & vbCrLf)
        sb.Append("function AddWarning(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddWarning(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreSettings function - This is used to restore web page settings.
        sb.Append("//Restore the web page settings." & vbCrLf)
        sb.Append("function RestoreSettings() {" & vbCrLf)
        sb.Append("  window.external.RestoreHtmlSettings() " & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'This line runs the RestoreSettings function when the web page is loaded.
        sb.Append("//Restore the web page settings when the page loads." & vbCrLf)
        sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append(vbCrLf)

        'Restores a single setting on the web page.
        sb.Append("//Restore a web page setting." & vbCrLf)
        sb.Append("  function RestoreSetting(FormName, ItemName, ItemValue) {" & vbCrLf)
        sb.Append("  document.forms[FormName][ItemName].value = ItemValue ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreOption function - This is used to add an option to a Select list.
        sb.Append("//Restore a Select control Option." & vbCrLf)
        sb.Append("function RestoreOption(SelectId, OptionText) {" & vbCrLf)
        sb.Append("  var x = document.getElementById(SelectId) ;" & vbCrLf)
        sb.Append("  var option = document.createElement(""Option"") ;" & vbCrLf)
        sb.Append("  option.text = OptionText ;" & vbCrLf)
        sb.Append("  x.add(option) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   Required Document Library Web Page JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf)
        'END:   Required Document Library Web Page JavaScript functions --------------------------------------------------------------------------

        sb.Append("</script>" & vbCrLf & vbCrLf)

        Return sb.ToString

    End Function


    Public Function DefaultHtmlString(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h1>" & DocumentTitle & "</h1>" & vbCrLf & vbCrLf)

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

#End Region 'Start Page Code ------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods Called by JavaScript - A collection of methods that can be called by JavaScript in a web page shown in WebBrowser1" '==================================
    'These methods are used to display HTML pages in the Document tab.
    'The same methods can be found in the WebView form, which displays web pages on seprate forms.

    Public Sub JSMethodTest1()
        'Test method that is called from JavaScript.
        Message.Add("JSMethodTest1 called OK." & vbCrLf)
    End Sub

    Public Sub JSMethodTest2(ByVal Var1 As String, ByVal Var2 As String)
        'Test method that is called from JavaScript.
        Message.Add("Var1 = " & Var1 & " Var2 = " & Var2 & vbCrLf)
    End Sub

    Public Sub JSDisplayXml(ByRef XDoc As XDocument)
        Message.Add(XDoc.ToString & vbCrLf & vbCrLf)
    End Sub

    Public Sub ShowMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub SaveHtmlSettings(ByVal xSettings As String, ByVal FileName As String)
        'Save the Html settings for a web page.

        'Convert the XSettings to XML format:

        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Project.SaveXmlData(FileName, XDocSettings)

    End Sub

    Public Sub RestoreHtmlSettings_Old(ByVal FileName As String)
        'Restore the Html settings for a web page.

        Dim XDocSettings As New System.Xml.Linq.XDocument
        Project.ReadXmlData(FileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'Message.Add("No HTML Settings file : " & FileName & vbCrLf)
        Else
            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)

                'Run the Settings file:
                XSeq.RunXSequence(XSettings, XStatus)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        'Dim SettingsFileName As String = txtNodeKey.Text & "Settings"
        Dim SettingsFileName As String = StartPageFileName & "Settings"

        Dim XDocSettings As New System.Xml.Linq.XDocument
        Project.ReadXmlData(SettingsFileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'Message.Add("No HTML Settings file : " & SettingsFileName & vbCrLf)
        Else
            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)
                'Run the Settings file:
                XSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub


    Private Sub XSeq_Instruction(Info As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn
            Case "Settings:Form:Name"
                FormName = Info

            Case "Settings:Form:Item:Name"
                ItemName = Info

            Case "Settings:Form:Item:Value"
                RestoreSetting(FormName, ItemName, Info)

            Case "Settings:Form:SelectId"
                SelectId = Info

            Case "Settings:Form:OptionText"
                RestoreOption(SelectId, Info)


            Case "Settings"

            Case "EndOfSequence"
                'Main.Message.Add("End of processing sequence" & Info & vbCrLf)

            Case Else
                'Main.Message.AddWarning("Unknown location: " & Locn & "  Info: " & Info & vbCrLf)
                Message.AddWarning("Unknown location: " & Locn & "  Info: " & Info & vbCrLf)

        End Select
    End Sub


    Public Sub RestoreSetting(ByVal FormName As String, ByVal ItemName As String, ByVal ItemValue As String)
        'Restore the setting value with the specified Form Name and Item Name.

        Me.WebBrowser1.Document.InvokeScript("RestoreSetting", New String() {FormName, ItemName, ItemValue})

    End Sub

    Public Sub RestoreOption(ByVal SelectId As String, ByVal OptionText As String)
        'Restore the Option text in the Select control with the Id SelectId.

        Me.WebBrowser1.Document.InvokeScript("RestoreOption", New String() {SelectId, OptionText})
    End Sub

    Private Sub SaveWebPageSettings()
        'Call the SaveSettings JavaScript function:
        Try
            Me.WebBrowser1.Document.InvokeScript("SaveSettings")
        Catch ex As Exception
            Message.AddWarning("Web page settings not saved: " & ex.Message & vbCrLf)
        End Try

    End Sub

    Public Function GetFormNo() As String
        'Return FormNo.ToString
        Return "-1"
    End Function

    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        Message.AddWarning(Msg)
    End Sub


    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.

    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)

    End Sub

    Public Sub ParentProjectName(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project name:
        RestoreSetting(FormName, ItemName, Project.ParentProjectName)
    End Sub

    Public Sub ParentProjectPath(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project path:
        RestoreSetting(FormName, ItemName, Project.ParentProjectPath)
    End Sub

    Public Sub ParentProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Parent Project parameter value:
        RestoreSetting(FormName, ItemName, Project.ParentParameter(ParameterName).Value)
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        RestoreSetting(FormName, ItemName, Project.Parameter(ParameterName).Value)
    End Sub

    Public Sub ApplicationNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Application Network:
        RestoreSetting(FormName, ItemName, Project.Parameter("AppNetName").Value)
    End Sub


#End Region 'Methods Called by JavaScript -------------------------------------------------------------------------------------------------------------------------------


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
        Import.RestoreSettings()

        RestoreFormSettings() 'This restores the Main form settings for the selected project.
        Project.ReadProjectInfoFile()

        Project.ReadParameters()
        Project.ReadParentParameters()
        If Project.ParentParameterExists("AppNetName") Then
            'Project.Parameter("AppNetName") = Project.ParentParameter("AppNetName")
            Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
            AppNetName = Project.Parameter("AppNetName").Value
        Else
            AppNetName = Project.GetParameter("AppNetName")
        End If

        Project.LockProject() 'Lock the project while it is open in this application.

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
        txtSettingsPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataPath.Text = Project.DataLocn.Path

    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Project.ShowParameters()
    End Sub

    Private Sub btnAddProject_Click(sender As Object, e As EventArgs) Handles btnAddProject.Click
        'Add the current project to the Message Service list.

        If Project.ParentProjectName <> "" Then
            Message.AddWarning("This project has a parent: " & Project.ParentProjectName & vbCrLf)
            Message.AddWarning("Child projects can not be added to the list." & vbCrLf)
            Exit Sub
        End If

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to AppNet:
                    Message.XAddText("Message sent to " & "MessageService" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)

                End If
            End If
        End If
    End Sub

    Private Sub btnOpenProject_Click(sender As Object, e As EventArgs) Handles btnOpenProject.Click
        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then

        Else
            Process.Start(Project.Path)
        End If
    End Sub

    Private Sub btnOpenSettings_Click(sender As Object, e As EventArgs) Handles btnOpenSettings.Click
        If Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SettingsLocn.Path)
        End If
    End Sub

    Private Sub btnOpenData_Click(sender As Object, e As EventArgs) Handles btnOpenData.Click
        If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.DataLocn.Path)
        End If
    End Sub

    Private Sub btnOpenSystem_Click(sender As Object, e As EventArgs) Handles btnOpenSystem.Click
        If Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SystemLocn.Path)
        End If
    End Sub

    Private Sub btnOpenAppDir_Click(sender As Object, e As EventArgs) Handles btnOpenAppDir.Click
        Process.Start(ApplicationInfo.ApplicationDir)
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        Timer2.Interval = 5000 '5 seconds
        Timer2.Enabled = True
        Timer2.Start()
    End Sub

    Private Sub TabPage1_Leave(sender As Object, e As EventArgs) Handles TabPage1.Leave
        Timer2.Enabled = False
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
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

        If System.IO.Directory.Exists(txtInputFileDir.Text) Then
            Dim FileList() As String = System.IO.Directory.GetFiles(txtInputFileDir.Text, "*.*")
            Dim I As Integer 'Loop index

            lstTextFiles.Items.Clear()

            For I = 0 To FileList.Count - 1
                lstTextFiles.Items.Add(FileList(I))
            Next
        Else
            Message.AddWarning("The directory: " & txtInputFileDir.Text & " doesnt exist!" & vbCrLf)
        End If
    End Sub

    Private Sub btnAddInputFilesToSeq_Click(sender As Object, e As EventArgs) Handles btnAddInputFilesToSeq.Click
        'Save the Input Files settings in the current Processing Sequence

        Dim I As Integer

        If IsNothing(Sequence) Then
            Message.AddWarning("The Sequencing form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Input Data: The text file directory and selected files:-->" & vbCrLf

            Sequence.rtbSequence.SelectedText = "<InputData>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <TextFileDirectory>" & Trim(txtInputFileDir.Text) & "</TextFileDirectory>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <TextFilesToProcess>" & vbCrLf

            If rbManual.Checked = True Then
                Sequence.rtbSequence.SelectedText = "  <SelectFileMode>" & "Manual" & "</SelectFileMode>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "    <Command>ClearSelectedFileList</Command>" & vbCrLf
                For I = 1 To lstTextFiles.SelectedItems.Count
                    Import.SelectedFiles(I - 1) = lstTextFiles.SelectedItems(I - 1)
                    Sequence.rtbSequence.SelectedText = "    <TextFile>" & lstTextFiles.SelectedItems(I - 1) & "</TextFile>" & vbCrLf
                Next
            Else
                Sequence.rtbSequence.SelectedText = "  <SelectFileMode>" & "SelectionFile" & "</SelectFileMode>" & vbCrLf
            End If
            Sequence.rtbSequence.SelectedText = "  <SelectionFilePath>" & Trim(txtSelectionFile.Text) & "</SelectionFilePath>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </TextFilesToProcess>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "</InputData>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()

        End If
    End Sub

    Private Sub btnSelectAllInputFiles_Click(sender As Object, e As EventArgs) Handles btnSelectAllInputFiles.Click
        'Select all input files in lstTextFiles

        Dim I As Integer
        For I = 0 To lstTextFiles.Items.Count - 1
            lstTextFiles.SetSelected(I, True)
        Next
        UpdateSelTextFileList()

    End Sub

#End Region 'Input Files Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Output Files Tab" '----------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnDatabase_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
        'Select the database file:

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
        Dim TempRegEx As Import.strucRegEx
        If Project.DataFileExists(txtRegExList.Text) Then
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
            Message.AddWarning("No RegEx List name has been specified." & vbCrLf)
            Exit Sub
        End If

        'Exit if there are no RegEx records in the _RegEx array:
        If IsNothing(Import.mRegEx) Then
            Message.AddWarning("No Regular Expressions have been specified." & vbCrLf)
            Exit Sub
        End If

        'Exit if there is an existing file that we don't want to overwrite:
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

        Project.SaveXmlData(RegExListFileName, RegExList)

    End Sub

    Private Sub btnAddToSeq_Click(sender As Object, e As EventArgs) Handles btnAddToSeq.Click
        'Add the MatchText settings to the current Import Sequence

        If IsNothing(Sequence) Then
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

        Dim TempRegEx As Import.strucRegEx

        If RegExIndex > 0 Then
            RegExIndex = RegExIndex - 1
            TempRegEx = Import.RegEx(RegExIndex)
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
            Import.RegExInsert(RegExIndex, TempRegEx)

            'Update DataGridView1:
            TempRow = DataGridView1.FirstDisplayedScrollingRowIndex 'Save the number of the first displayed row
            Dim MaxRow As Integer
            MaxRow = Import.RegExCount
            DataGridView1.Rows.Clear()
            DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically
            For I = 1 To MaxRow
                DataGridView1.Rows.Add(1)
                DataGridView1.Rows(I - 1).Cells(0).Value = Import.RegEx(I - 1).Name
                DataGridView1.Rows(I - 1).Cells(1).Value = Import.RegEx(I - 1).Descr
            Next

            'Highlight the current row:
            DataGridView1.ClearSelection()
            DataGridView1.Rows(RegExIndex).Selected = True
            DataGridView1.FirstDisplayedScrollingRowIndex = TempRow 'Restore the number of the first displayed row

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


            Import.RegExInsert(RegExIndex, TempRegEx)
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
            Import.RegExModify(RegExIndex, TempRegEx)

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
        UpdateMatches()

        Dim NValues As Integer
        NValues = Import.DbDestValues.GetUpperBound(1) + 1 'This is the number of columns in the DbDestValues array
        Message.Add("Number of columns in DbDestValues is: " & Str(NValues) & vbCrLf)

        Dim NRows As Integer
        NRows = Import.DbDestValues.GetUpperBound(0) + 1 'This is the number of rows in the DbDestValues array
        Message.Add("Number of rows in DbDestValues is: " & Str(NRows) & vbCrLf)

        Dim I As Integer
        For I = 0 To NRows - 1
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
        Dim TempDbDest As Import.strucDbDest
        Dim TempDbMult As Import.strucMultiplier
        If Project.DataFileExists(txtDbDestList.Text) Then
            Dim DbDestList As System.Xml.Linq.XDocument
            Project.ReadXmlData(txtDbDestList.Text, DbDestList)
            Import.DbDestListDescription = DbDestList.<DatabaseDestinations>.<Description>.Value
            txtListDescription.Text = Import.DbDestListDescription

            If DbDestList.<DatabaseDestinations>.<CreationDate>.Value = "" Then

            Else
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

        RefreshLocations()

    End Sub

    'Public Sub RefreshForm()
    Public Sub RefreshLocations()
        'Refresh the form:

        txtDbDestList.Text = Import.DbDestListName
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

        NCols = Import.DbDestValues.GetUpperBound(1) + 1
        NRows = Import.DbDestValues.GetUpperBound(0) + 1

        DataGridView3.ColumnCount = NCols
        DataGridView3.RowCount = NRows
        Dim RowNo As Integer
        Dim ColNo As Integer

        For RowNo = 0 To NRows - 1
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
            Message.AddWarning("No Database Destinations List name has been specified." & vbCrLf)
            Exit Sub
        End If

        'Exit if there are no RegEx records in the _RegEx array:
        If IsNothing(Import.mDbDest) Then
            Message.AddWarning("No Database Destinations have been specified." & vbCrLf)
            Exit Sub
        End If

        If IsNothing(Import.mMultiplierCodes) Then
            'Create a blank multiplier to prevent error creating XML file of database destinations:
            Dim TempMult As Import.strucMultiplier
            TempMult.RegExMultiplierVariable = ""
            TempMult.MultiplierCode = ""
            TempMult.MultiplierValue = 0
            Import.MultipliersAppend(TempMult)
        End If

        'Exit if there is an existing file that we don't want to overwrite:
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

        Project.SaveXmlData(DbDestListFileName, DbDestList)
        ListChanged = False
    End Sub

    Private Sub btnNewLocationList_Click(sender As Object, e As EventArgs) Handles btnNewLocationList.Click
        'Create a new DbDestinations list

        If ListChanged = True Then
            Dim dr As DialogResult
            dr = MessageBox.Show("Save the changes to the Database Destinations list?", "Notice", MessageBoxButtons.YesNoCancel)
            If dr = DialogResult.Yes Then
                'Exit if no RegEx List name has been specified:
                If Trim(txtDbDestList.Text) = "" Then
                    Message.AddWarning("No Databse Destinations List name has been specified." & vbCrLf)
                    Exit Sub
                End If

                'Exit if there are no RegEx records in the _RegEx array:
                If IsNothing(Import.mDbDest) Then
                    Message.AddWarning("No Database Destinations have been specified." & vbCrLf)
                    Exit Sub
                End If
                SaveDbDestList()
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
                Message.AddWarning("A row must be selected on the Databse Destination grid!" & vbCrLf)
            Else
                SelRow = DataGridView2.SelectedCells.Item(0).RowIndex
                If DataGridView2.Rows(SelRow).IsNewRow = True Then 'Uncommited row - new row cannot be appended
                    Message.AddWarning("Uncommitted row - new row cannot be appended. You may need to enter a RegEx Variable name!" & vbCrLf)
                Else
                    DataGridView2.Rows.Insert(SelRow + 1)
                    DataGridView3.Rows.Insert(SelRow + 1)
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
            Message.AddWarning("Uncommitted row - cannot be deleted" & vbCrLf)
        Else
            DataGridView2.Rows.RemoveAt(SelRow)
            DataGridView3.Rows.RemoveAt(SelRow)
            'Now update the Database Destination data in Import:
            Import.DbDestDelete(SelRow)
        End If

        If DbDestIndex > DataGridView2.RowCount - 1 Then 'The current Index is pointing past the last row in DataGridView2
            DbDestIndex = DataGridView2.RowCount - 1 '
        End If

    End Sub

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

        Dim I As Integer
        Dim strVarName As String 'This is used to hold the RegEx variable name in the current row
        Dim CapCount As Integer
        Dim J As Integer
        'Find matching RegEx Variables in the Database Destinations grid:
        For I = 1 To DataGridView2.RowCount 'Processes each row in the Data Destinations grid.
            If DataGridView2.Rows(I - 1).Cells(0).Value = Nothing Then 'No RegEx variable name is specified.
            Else
                strVarName = DataGridView2.Rows(I - 1).Cells(0).Value.ToString 'The RegEx variable name in the current grid row.
                If strVarName = txtRegExVariable.Text Then 'The RegExVariable at the current row matches the required variable to modify
                    If DataGridView2.Rows(I - 1).Cells(1).Value.ToString = "Single Value" Then

                    ElseIf DataGridView2.Rows(I - 1).Cells(1).Value.ToString = "Multiple Value" Then

                    End If
                    If IsNothing(DataGridView3.Rows(I - 1).Cells(0).Value) Then
                        'There is no text to modify
                    Else
                        Dim OutputString As String
                        Dim InputString As String
                        InputString = DataGridView3.Rows(I - 1).Cells(0).Value.ToString
                        ConvertDate(txtInputDateFormat.Text, txtOutputDateFormat.Text, InputString, OutputString)
                        DataGridView3.Rows(I - 1).Cells(0).Value = OutputString
                    End If
                End If
            End If
        Next
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

        Dim I As Integer
        Dim strVarName As String 'This is used to hold the RegEx variable name in the current row
        Dim CapCount As Integer
        Dim J As Integer
        'Find matching RegEx Variables in the Database Destinations grid:
        For I = 1 To DataGridView2.RowCount
            If DataGridView2.Rows(I - 1).Cells(0).Value = Nothing Then 'No RegEx variable name is specified.
            Else
                strVarName = DataGridView2.Rows(I - 1).Cells(0).Value.ToString 'The RegEx variable name in the current grid row.
                If strVarName = txtRegExVariable.Text Then 'The RegExVariable at the current row matches the required variable to modify
                    If IsNothing(DataGridView3.Rows(I - 1).Cells(0).Value) Then
                        txtTestInputString.Text = ""
                    Else
                        txtTestInputString.Text = DataGridView3.Rows(I - 1).Cells(0).Value.ToString
                    End If

                End If
            End If
        Next
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
            'Debug.Print(ex.Message)
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
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFixedValue
        End If
    End Sub

    Private Sub rbTextFileName_Click(sender As Object, e As EventArgs)
        If rbAppendTextFileName.Checked Then
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFileName
        End If
    End Sub

    Private Sub rbTextFileDirectory_Click(sender As Object, e As EventArgs)
        If rbAppendTextFileDirectory.Checked Then
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFileDir
        End If
    End Sub

    Private Sub rbTextFilePath_Click(sender As Object, e As EventArgs)
        If rbAppendTextFilePath.Checked Then
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendFilePath
        End If
    End Sub

    Private Sub rbCurrentDate_Click(sender As Object, e As EventArgs)
        If rbAppendCurrentDate.Checked Then
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentDate
        End If
    End Sub

    Private Sub rbCurrentTime_Click(sender As Object, e As EventArgs)
        If rbAppendCurrentTime.Checked Then
            Import.ModifyValuesType = Import.ModifyValuesTypes.AppendCurrentTime
        End If
    End Sub

    Private Sub rbCurrentDateTime_Click(sender As Object, e As EventArgs)
        If rbAppendCurrentDateTime.Checked Then
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
            ElseIf rbAppendFixedValue.Checked Then
                Sequence.rtbSequence.SelectedText = "  <FixedValue>" & txtFixedValue.Text & "</FixedValue>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_fixed_value</ModifyType>" & vbCrLf
            ElseIf rbAppendRegExVar.Checked Then
                Sequence.rtbSequence.SelectedText = "  <RegExVariableValueFrom>" & Trim(txtAppendRegExVar.Text) & "</RegExVariableValueFrom>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_RegEx_variable_value</ModifyType>" & vbCrLf
            ElseIf rbAppendTextFileName.Checked Then
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_file_name</ModifyType>" & vbCrLf
            ElseIf rbAppendTextFileDirectory.Checked Then
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_file_directory</ModifyType>" & vbCrLf
            ElseIf rbAppendTextFilePath.Checked Then
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_file_path</ModifyType>" & vbCrLf
            ElseIf rbAppendCurrentDate.Checked Then
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_current_date</ModifyType>" & vbCrLf
            ElseIf rbAppendCurrentTime.Checked Then
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_current_time</ModifyType>" & vbCrLf
            ElseIf rbAppendCurrentDateTime.Checked Then
                Sequence.rtbSequence.SelectedText = "  <ModifyType>Append_current_date_time</ModifyType>" & vbCrLf
            End If

            Sequence.rtbSequence.SelectedText = "</ModifyValues>" & vbCrLf

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
        'Connect to or disconnect from the Message System (ComNet).
        If ConnectedToComNet = False Then
            ConnectToComNet()
        Else
            DisconnectFromComNet()
        End If

    End Sub

    Private Sub ConnectToComNet()
        'Connect to the Message System. (ComNet)

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If MsgServiceAppPath = "" Then
            Message.AddWarning("Message Service application path is unknown." & vbCrLf)
        Else
            If ComNetRunning() Then
                'The Message Service is Running.
            Else 'The Message Service is NOT running'
                'Start the Message Service:
                If System.IO.File.Exists(MsgServiceExePath) Then 'OK to start the Message Service application:
                    Shell(Chr(34) & MsgServiceExePath & Chr(34), AppWinStyle.NormalFocus) 'Start Message Service application with no argument
                Else
                    'Incorrect Message Service Executable path.
                    Message.AddWarning("Message Service exe file not found. Service not started." & vbCrLf)
                End If
            End If
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.AddWarning("Client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds

                ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
                ConnectionName = client.Connect(AppNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False) 'UPDATED 2Feb19

                If ConnectionName <> "" Then
                    Message.Add("Connected to the Communication Network as " & ConnectionName & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToComNet = True
                    SendApplicationInfo()
                    client.GetMessageServiceAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).
                Else
                    Message.Add("Connection to the Communication Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Communication Network (Message Service) is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If

    End Sub


    Private Sub ConnectToComNet(ByVal ConnName As String)
        'Connect to the Communication Network with the connection name ConnName.

        If ConnectedToComNet = False Then
            Dim Result As Boolean

            If IsNothing(client) Then
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.AddWarning("client state is faulted. Connection not made!" & vbCrLf)
            Else
                Try
                    'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds
                    ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
                    ConnectionName = client.Connect(AppNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False) 'UPDATED 2Feb19

                    If ConnectionName <> "" Then
                        Message.Add("Connected to the Communication Network as " & ConnectionName & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                        btnOnline.Text = "Online"
                        btnOnline.ForeColor = Color.ForestGreen
                        ConnectedToComNet = True
                        SendApplicationInfo()
                        client.GetMessageServiceAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).
                    Else
                        Message.Add("Connection to the Communication Network failed!" & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    End If
                Catch ex As System.TimeoutException
                    Message.Add("Timeout error. Check if the Communication Network is running." & vbCrLf)
                Catch ex As Exception
                    Message.Add("Error message: " & ex.Message & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End Try
            End If
        Else
            Message.AddWarning("Already connected to the Communication Network." & vbCrLf)
        End If

    End Sub


    Private Sub DisconnectFromComNet()
        'Disconnect from the Communication Network.

        If ConnectedToComNet = True Then
            If IsNothing(client) Then
                Message.Add("Already disconnected from the Communication Network." & vbCrLf)
                btnOnline.Text = "Offline"
                btnOnline.ForeColor = Color.Red
                ConnectedToComNet = False
                ConnectionName = ""
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted." & vbCrLf)
                    ConnectionName = ""
                Else
                    Try
                        client.Disconnect(AppNetName, ConnectionName)

                        btnOnline.Text = "Offline"
                        btnOnline.ForeColor = Color.Red
                        ConnectedToComNet = False
                        ConnectionName = ""
                        Message.Add("Disconnected from the Communication Network." & vbCrLf)
                    Catch ex As Exception
                        Message.AddWarning("Error disconnecting from Communication Network: " & ex.Message & vbCrLf)
                    End Try
                End If
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

                Dim text As New XElement("Text", "Import")
                applicationInfo.Add(text)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)

                'Show the message sent to AppNet:
                Message.XAddText("Message sent to " & "MessageService" & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage("", "MessageService", doc.ToString) 'UPDATED 2Feb19

            End If
        End If
    End Sub

    Private Function ComNetRunning() As Boolean
        'Return True if ComNet (Message Service) is running.
        If System.IO.File.Exists(MsgServiceAppPath & "\Application.Lock") Then
            Return True
        Else
            Return False
        End If
    End Function

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

        If IsDBNull(Info) Then
            Info = ""
        End If

        Select Case Locn

            Case "ClientAppNetName"
                ClientAppNetName = Info 'The name of the Client Application Network requesting service. ADDED 2Feb19.

            Case "ClientName"
                ClientAppName = Info 'The name of the Client requesting service.

            Case "ClientConnectionName"
                ClientConnName = Info 'The name of the client requesting service.

            Case "ClientLocn" 'The Location within the Client requesting service.
                Dim statusOK As New XElement("Status", "OK") 'Add Status OK element when the Client Location is changed
                xlocns(xlocns.Count - 1).Add(statusOK)

                xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the instructions for the last location to the reply xmessage
                xlocns.Add(New XElement(Info)) 'Start the new location instructions

            Case "Main"
                 'Blank message - do nothing.

            Case "Main:Status"
                Select Case Info
                    Case "OK"
                        'Main instructions completed OK
                End Select

            Case "Command"
                Select Case Info
                    Case "ConnectToComNet" 'Startup Command
                        If ConnectedToComNet = False Then
                            ConnectToComNet()
                        End If
                End Select

                   'Startup Command Arguments ================================================
            Case "ProjectName"
                If Project.OpenProject(Info) = True Then
                    ProjectSelected = True 'Project has been opened OK.
                Else
                    ProjectSelected = False 'Project could not be opened.
                End If

            Case "ProjectID"
                Message.AddWarning("Add code to handle ProjectID parameter at StartUp!" & vbCrLf)
                'Note the AppNet will usually select a project using ProjectPath.

            Case "ProjectPath"
                If Project.OpenProjectPath(Info) = True Then
                    ProjectSelected = True 'Project has been opened OK.
                Else
                    ProjectSelected = False 'Project could not be opened.
                End If

            Case "ConnectionName"
                StartupConnectionName = Info
            '--------------------------------------------------------------------------

            'Application Information  =================================================
            'returned by client.GetMessageServiceAppInfoAsync()
            Case "MessageServiceAppInfo:Name"
                'The name of the Message Service Application. (Not used.)

            Case "MessageServiceAppInfo:ExePath"
                'The executable file path of the Message Service Application.
                MsgServiceExePath = Info

            Case "MessageServiceAppInfo:Path"
                'The path of the Message Service Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                MsgServiceAppPath = Info
           '---------------------------------------------------------------------------


            Case "EndOfSequence"
                'End of Information Vector Sequence reached.
                'Add Status OK element at the end of the sequence:
                Dim statusOK As New XElement("Status", "OK")
                xlocns(xlocns.Count - 1).Add(statusOK)

            Case Else
                Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                Message.AddWarning("            info: " & Info & vbCrLf)

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
            Message.AddWarning("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.AddWarning("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Try
                    client.SendMessage(ClientAppNetName, ClientConnName, MessageText) 'Added 2Feb19
                    MessageText = "" 'Clear the message after it has been sent.
                    ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
                    ClientConnName = "" 'Clear the Client Application Name after the message has been sent.
                    xlocns.Clear()
                Catch ex As Exception
                    Message.AddWarning("Error sending message: " & ex.Message & vbCrLf)
                End Try
            End If
        End If

        'Stop timer:
        Timer1.Enabled = False
    End Sub

    Private Sub btnGridProp_Click(sender As Object, e As EventArgs) Handles btnGridProp.Click

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

    Private Sub lstTextFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTextFiles.SelectedIndexChanged

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        'Keet the connection awake with each tick:

        If ConnectedToComNet = True Then
            Try
                If client.IsAlive() Then
                    Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
                    Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
                Else
                    Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf)
                    Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
                'Set interval to five minutes - try again in five minutes:
                Timer3.Interval = TimeSpan.FromMinutes(5).TotalMilliseconds '5 minute interval
            End Try
        Else
            Message.Add(Format(Now, "HH:mm:ss") & " Not connected." & vbCrLf)
        End If

    End Sub

    Private Sub chkConnect_LostFocus(sender As Object, e As EventArgs) Handles chkConnect.LostFocus
        If chkConnect.Checked Then
            Project.ConnectOnOpen = True
        Else
            Project.ConnectOnOpen = False
        End If
        Project.SaveProjectInfoFile()

    End Sub

#End Region 'Process XMessages --------------------------------------------------------------------------------------------------------------------------------------------------------

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class
