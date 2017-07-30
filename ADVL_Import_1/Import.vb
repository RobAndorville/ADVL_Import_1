Public Class Import
    'Import data system.

    'To Do:
    '4th May 2014: Add method to select a set of files using the SelectionFilePath.

#Region " Variable Declarations"

    'Declare Forms used by the application:
    'Public WithEvents DebugMessages As frmDebugMessages

    Structure strucRegEx 'Regular Expression entry structure
        Dim Name As String           'The name of the RegEx
        Dim Descr As String          'A description of the RegEx
        Dim RegEx As String          'The Regular Expression (RegEx) string
        Dim ExitOnMatch As Boolean   'If True, exit processing of the current RegEx list
        Dim ExitOnNoMatch As Boolean 'If True, exit processing of the current RegEx list
        Dim MatchStatus As String    'Set the Status string to this value if RegEx match
        Dim NoMatchStatus As String  'Set the Status string to this value if no RegEx match
    End Structure

    Structure strucDbDest 'Database destination parameters
        Dim RegExVariable As String
        Dim Type As String 'Single Value or Multiple Value
        Dim TableName As String
        Dim FieldName As String
        Dim StatusField As String '(Optional) The field where the status of the variable value is coded (OK, N/A)
    End Structure

    Structure strucMultiplier 'Multiplier structure
        Dim RegExMultiplierVariable As String
        Dim MultiplierCode As String
        Dim MultiplierValue As Single
    End Structure

    Structure strucFileInfo 'File information structure
        Dim Name As String
        Dim Description As String
        Dim CreationDate As DateTime
        Dim LastEditDate As DateTime
    End Structure

    Structure strucGridProp 'Regular Expression entry structure
        Dim Mult As Single
        Dim Status As String
    End Structure

    Structure strucTableList
        Dim TableName As String 'Table name (Unique list of table names)
        Dim MaxNValues As Integer 'Maximum number of values in a Multiple Value variable to be written to the table
        Dim MinNValues As Integer 'Minimum number of values in a Multiple Value variable
        Dim HasGaps As Boolean 'True if there are gaps with no values within a set of Multiple Value variable values
    End Structure

    Structure strucFieldList
        Dim TableName As String
        Dim FieldName As String
        Dim StatusFieldName As String
        Dim FieldType As String 'Text-Memo-Number-Date/Time-Currency-AutoNumber-Yes/No-Hyperlink Number:Byte-Integer-LongInteger-Single-Double
        Dim FieldLen As Integer
        Dim MultiValued As Boolean
        Dim Status As String 'Gap-Mis
        Dim NValues As Integer 'The number of values for this Field to write to the database
        'UPDATE: GridProp() contains the Multipliers corresponding to values in DataGridView1. (Mult field not required).
        'Dim Mult As Single 'The multiplier used to scale the Field values in the source document
        Dim RowNo As Integer 'The Row Number of the Field in DataGridView1
    End Structure

    Structure strucFieldValues
        Dim Value As String
        Dim Status As String
    End Structure

    Structure strucSelection 'Regular Expression entry structure
        Dim Type As String 'ClearAll, SelectAll, AddMatches, RemoveMatches
        Dim ReadMode As String 'ReadAll, ReadLines, ReadChars
        Dim ReadNo As Integer 'The number of lines or characters to read from the files
        Dim FileType As String 'txt, asc or used defined
        Dim RegEx As String 'The RegEx string used to match files.
    End Structure

    'Property variables:

    'Database destination parameters:
    'mDbDest is declared below with the associated Property code.
    'Public mDbDest() As strucDbDest 'List of Database Destinations for matched text
    'Array containing Database Destination information
    'This array has 5 columns: 
    'RegEx_Variable     The name of the variable to be matched in a Regular Expression.
    'Variable_Type      (Single Value, ...) The type of variable.
    'Destination_Table  The name of the table in the selected database that the values are to be written to.
    'Destination_Field  The name of the field in the table that the values are to be written to.
    'Status_Field       The name of the field to hold status codes for the corresponding values.
    'The array is dimensioned using column, row order: DbDest(0 to 4,NRows - 1) 
    'This order allows the number of rows to be changed using the Redim Preserve statement.

    Public TextStore As String = "" 'Holds the text for processing
    Dim textIn As System.IO.StreamReader

    Public DbDestValues(,) As String
    'This array stores text matches corresponding the the rows in the DbDest() array.
    'The array is dimensioned using row, column order: DbDestValues(0 to NRows - 1, 0 to NCols - 1)
    'This order allows the number of columns to be changed using the Redim Preserve statement.

    'Variables used for writing data to the database:
    Dim GridProp(,) As strucGridProp 'Array used to store properties for corresponding cells in DbDestinations.DataGridView1
    Dim TableList() As strucTableList 'Holds a list of destination table names and Min and Max number of values to be written
    Dim FieldList() As strucFieldList 'Holds a list of destination fields and other properties
    Dim FieldValues(,) As strucFieldValues 'Holds the destination field values

    Dim InsertCount As Integer = 0 'Keeps count of the number of items written or attempted to write to the database.

    'Variables used to connect to the database:
    'THe Property ConnectionString contains the database connection string
    Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
    Dim da As New System.Data.OleDb.OleDbDataAdapter
    Dim cmd As New System.Data.OleDb.OleDbCommand

    Private ImportStatus As New System.Collections.Specialized.StringCollection

    'Private WithEvents XSeq As New TDS_Utilities.RunXSequence 'This is used to run a set of XML Sequence statements. These are used to control the data importing.
    'Private WithEvents XSeq As New TDS_Utilities.XSequence 'This is used to run a set of XML Sequence statements. These are used to control the data importing.
    Private WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence 'This is used to run a set of XML Sequence statements. These are used to control the data importing.

    'Variables used to monitor the import process:
    Private ReadLineCount As Integer 'Variable used to keep a count of the lines read in a file.

    'List of the database types for which this class has predefined connection strings.
    Public Enum DatabaseTypeEnum
        User_defined_connection_string 'In this case, the connection string is defined manually
        Access2007To2013               'The connection string is defined when this database type is selected.
    End Enum

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    'Public ApplicationName As String 'The name of the application using the Import class.

    Public DataLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store data.

#End Region 'Variable Declarations

#Region " Properties"

    'LIST OF PROPERTIES:
    'ProjectPath                The path of the project directory containing import parameter files.
    'NOTE: ProjectPath is no longer used. Use SettingsLocn and DataLocn instead.

    'ImportSequenceName         The name of a data import processing sequence.
    'ImportSequenceDescription  A description of the import processing sequence.

    'TextFileDir                The directory containing text files to be imported.
    'SelTextFileCount           The number of text files selected for importing.
    'SelTextFiles()             An array contining the names of the text files selected for importing.
    'SelTextFileNumber          The index number of the current selected text file.
    'SelectTextFileMode         The method being used to select the text files for importing. Manual or SelectionFile.
    'SelectionFileName          The name of the selection file used to select the text files. This will be used if the SelectionFile mode is used.
    'TextFilePath               The path of a text file being imported.
    'TextFileOpen               True or False. Indicates if the text file has been opened for importing.

    'DatabasePath               The path of the database into which the text file data is being imported.
    'DatabaseType               The type of database (Access 2013 2010 2007). The same connection string is used to openAccess 2007 - 2013 databases. 
    'ConnectionString           The connection string used to open the database. 'http://www.connectionstrings.com/access/

    'RegExCount                 The number of entries in the Regular Expression list.
    'RegExListName              The name of the Regular Expression list.
    'RegExListDescr             A description of the Regular Expression list.
    'RegExListCreationDate      The creation date of the Regular Expression list.
    'RegExListLastEditDate      The last edit date of the Regular Expression list.
    'RegEx()                    An array of regular expressions used to match text.

    'DbDestListName             The name of a list of database destinations.
    'DbDestListDescription      A description of the list of database destinations.
    'DbDestListCreationDate     The date the destination list was created.
    'DbDestListLastEditDate     The date the destination list was last edited.
    'DbDestCount                The number of entries in the list of destinations.
    'DbDest()                   Array of Database Destinations.

    'MultiplierCode()           A list of multiplier codes and the corresponding multiplier values. 
    'MultiplierCodeCount        The number of entries in the list of multiplier codes.

    'ReadTextNLines             The number of lines parameter used in the Read Text methods.
    'ReadTextString             The string parameter used in the Read Text methods.



    'Private mProjectPath As String = ""
    'Property ProjectPath() As String
    '    Get
    '        Return mProjectPath
    '    End Get
    '    Set(ByVal value As String)
    '        mProjectPath = Trim(value)
    '    End Set
    'End Property

    'IMPORT SEQUENCE
    Private mImportSequenceName As String = ""
    Property ImportSequenceName() As String
        Get
            Return mImportSequenceName
        End Get
        Set(ByVal value As String)
            mImportSequenceName = value
        End Set
    End Property

    Private mImportSequenceDescription As String = ""
    Property ImportSequenceDescription() As String
        Get
            Return mImportSequenceDescription
        End Get
        Set(ByVal value As String)
            mImportSequenceDescription = value
        End Set
    End Property

    'INPUT DATA:
    Private mTextFileDir As String = ""
    Property TextFileDir() As String 'The selected text file directory.
        Get
            TextFileDir = mTextFileDir
        End Get
        Set(ByVal value As String)
            mTextFileDir = value
        End Set
    End Property

    ReadOnly Property SelectedFileCount() As Integer 'The number of selected text files.
        Get
            If IsNothing(mSelectedFiles) Then
                Return 0
            Else
                Return mSelectedFiles.Count
            End If
        End Get
    End Property

    Private mSelectedFiles() As String
    Property SelectedFiles(ByVal Index As Integer) As String 'The paths of the selected text files.
        Get
            If Index <= mSelectedFiles.Count - 1 Then
                Return mSelectedFiles(Index)
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If IsNothing(mSelectedFiles) Then
                If Index = 0 Then
                    ReDim mSelectedFiles(0 To 0)
                    mSelectedFiles(0) = value
                Else
                    MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
                End If
            Else
                If Index > mSelectedFiles.Count - 1 Then
                    ReDim Preserve mSelectedFiles(0 To Index)
                End If
                mSelectedFiles(Index) = value
            End If

        End Set
    End Property

    Private mSelectedFileNumber = 0 'The index number of the current selected text file (The first selected text file has _selTextFileNumber = 1
    ReadOnly Property SelectedFileNumber
        Get
            Return mSelectedFileNumber
        End Get
    End Property

    Private mSelectFileMode As String = "Manual" 'Manual or SelectionFile
    Property SelectFileMode As String
        Get
            Return mSelectFileMode
        End Get
        Set(ByVal value As String)
            If value = "Manual" Or value = "SelectionFile" Then
                mSelectFileMode = value
            Else 'Not a valid value '_selectionFilePath
            End If
        End Set
    End Property

    Private mSelectionFileName As String = "" 'Selection file name
    Property SelectionFileName As String
        Get
            Return mSelectionFileName
        End Get
        Set(ByVal value As String)
            mSelectionFileName = value
        End Set
    End Property

    Private mCurrentFilePath = "" 'The current text file being processed
    Property CurrentFilePath() As String
        Get
            Return mCurrentFilePath
        End Get
        Set(ByVal value As String)
            mCurrentFilePath = value
        End Set
    End Property

    Dim mFileOpen As Boolean = False 'True if the Test Text File is open
    Property FileOpen As Boolean
        Get
            Return mFileOpen
        End Get
        Set(ByVal value As Boolean)
            mFileOpen = value
        End Set
    End Property

    'DATABASE:
    Private mDatabasePath As String = ""
    Property DatabasePath() As String
        Get
            DatabasePath = mDatabasePath
        End Get
        Set(ByVal value As String)
            mDatabasePath = value

            ''Set up the database connection:
            ''Specify the connection string (Access 2007):
            'connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + _
            '"data source = " + mDatabasePath
            ''Connect to the Access database:
            'conn = New System.Data.OleDb.OleDbConnection(connectionString)
            'Try
            '    conn.Open()
            'Catch ex As Exception
            '    'NOTE: Need to return this error message in an event!!!!!
            '    'ShowMessage("Error opening database: " & DatabasePath & vbCrLf, Color.Red)
            '    'ShowMessage(ex.Message & vbCrLf & vbCrLf, Color.Blue)
            '    Exit Property
            'End Try

            'Update the connection string:
            If mDatabaseType = DatabaseTypeEnum.Access2007To2013 Then
                ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + mDatabasePath
            End If
        End Set

    End Property

    Private mDatabaseType As DatabaseTypeEnum
    Property DatabaseType As DatabaseTypeEnum
        Get
            DatabaseType = mDatabaseType
        End Get
        Set(value As DatabaseTypeEnum)
            mDatabaseType = value
            If value = DatabaseTypeEnum.Access2007To2013 Then
                'Set the appropriate connection string:
                ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + DatabasePath
            End If
        End Set
    End Property

    Private mConnectionString As String = ""
    Property ConnectionString As String
        Get
            ConnectionString = mConnectionString
        End Get
        Set(value As String)
            mConnectionString = value
        End Set
    End Property

    'MATCH TEXT:
    Private mRegExListInfo As strucFileInfo
    Property RegExListName As String
        Get
            'Return _regExListName
            Return mRegExListInfo.Name
        End Get
        Set(ByVal value As String)
            '_regExListName = value
            mRegExListInfo.Name = value
        End Set
    End Property

    Property RegExListDescr As String
        Get
            'Return _regExListDescr
            Return mRegExListInfo.Description
        End Get
        Set(ByVal value As String)
            '_regExListDescr = value
            mRegExListInfo.Description = value
        End Set
    End Property

    Property RegExListCreationDate As DateTime
        Get
            'Return _regExListCreationDate
            Return mRegExListInfo.CreationDate
        End Get
        Set(ByVal value As DateTime)
            '_regExListCreationDate = value
            mRegExListInfo.CreationDate = value
        End Set
    End Property

    Property RegExListLastEditDate As DateTime
        Get
            'Return _regExListLastEditDate
            Return mRegExListInfo.LastEditDate
        End Get
        Set(ByVal value As DateTime)
            '_regExListLastEditDate = value
            mRegExListInfo.LastEditDate = value
        End Set
    End Property

    Public mRegEx() As strucRegEx 'List of regular expressions
    Property RegEx(ByVal Index As Integer) As strucRegEx
        Get
            If IsNothing(mRegEx) Then
                RegEx.Name = ""
                RegEx.Descr = ""
                RegEx.RegEx = ""
                RegEx.ExitOnMatch = False
                RegEx.ExitOnNoMatch = False
                RegEx.MatchStatus = ""
                RegEx.NoMatchStatus = ""
            Else
                If Index <= mRegEx.Count - 1 Then
                    Return mRegEx(Index)
                Else
                    RegEx.Name = ""
                    RegEx.Descr = ""
                    RegEx.RegEx = ""
                    RegEx.ExitOnMatch = False
                    RegEx.ExitOnNoMatch = False
                    RegEx.MatchStatus = ""
                    RegEx.NoMatchStatus = ""
                End If
            End If

        End Get
        Set(ByVal value As strucRegEx)
            If Index >= mRegEx.Count Then
                ReDim Preserve mRegEx(0 To Index)
            End If
            mRegEx(Index) = value
        End Set
    End Property

    ReadOnly Property RegExCount() As Integer
        Get
            If IsNothing(mRegEx) Then
                Return 0
            Else
                Return mRegEx.Count
            End If
        End Get
    End Property

    'DATABSE DESTINATIONS:
    Private mDbDestListInfo As strucFileInfo
    Property DbDestListName As String
        Get
            'Return _dbDestListName
            Return mDbDestListInfo.Name
        End Get
        Set(ByVal value As String)
            '_dbDestListName = value
            mDbDestListInfo.Name = value
            'OpenDbDestListFile()
        End Set
    End Property

    Property DbDestListDescription As String
        Get
            'Return _dbDestListDescr
            Return mDbDestListInfo.Description
        End Get
        Set(ByVal value As String)
            '_dbDestListDescr = value
            mDbDestListInfo.Description = value
        End Set
    End Property

    Property DbDestListCreationDate As DateTime
        Get
            'Return _dbDestListCreationDate
            Return mDbDestListInfo.CreationDate
        End Get
        Set(ByVal value As DateTime)
            '_dbDestListCreationDate = value
            mDbDestListInfo.CreationDate = value
        End Set
    End Property

    Property DbDestListLastEditDate As DateTime
        Get
            'Return _dbDestListLastEditDate
            Return mDbDestListInfo.LastEditDate
        End Get
        Set(ByVal value As DateTime)
            '_dbDestListLastEditDate = value
            mDbDestListInfo.LastEditDate = value
        End Set
    End Property

    Public mDbDest() As strucDbDest 'List of Database Destinations for matched text
    Property DbDest(ByVal Index As Integer) As strucDbDest
        'Specifies the databse destinations of RegEx results variables.
        Get
            If Index <= mDbDest.Count - 1 Then
                Return mDbDest(Index)
            Else
                DbDest.RegExVariable = ""
                DbDest.Type = ""
                DbDest.TableName = ""
                DbDest.FieldName = ""
                DbDest.StatusField = ""
            End If
        End Get
        Set(ByVal value As strucDbDest)
            If Index <= mDbDest.Count - 1 Then
                mDbDest(Index) = value
            Else
                ReDim Preserve mDbDest(0 To Index)
                mDbDest(Index) = value
            End If
        End Set

    End Property

    ReadOnly Property DbDestCount() As Integer
        'Keeps a count of the number of database destinations specified
        Get
            If IsNothing(mDbDest) Then
                Return 0
            Else
                Return mDbDest.Count
            End If
        End Get

    End Property

    Private _useNullValueString As Boolean = False 'If True, the string stored in NullValueString, is used to represent a null value in the FieldValues array.
    'Data sources may use "N/A" or "--" strings to indicate that no data is available.
    'If UseNullValueString is True, a FieldValue containing the NullValueString will entered into the database as a null value.
    Property UseNullValueString As Boolean
        Get
            Return _useNullValueString
        End Get
        Set(value As Boolean)
            _useNullValueString = value
        End Set
    End Property

    Private _nullValueString As String = "" 'A string, such as "N/A" or "--", used to indicate a null value.
    Property NullValueString As String
        Get
            Return _nullValueString
        End Get
        Set(value As String)
            _nullValueString = value
        End Set
    End Property

    'MULTIPLIERS:
    Public mMultiplierCodes() As strucMultiplier
    Property MultiplierCode(ByVal Index As Integer) As strucMultiplier
        Get
            If Index <= mMultiplierCodes.Count - 1 Then
                Return mMultiplierCodes(Index)
            Else
                MultiplierCode.RegExMultiplierVariable = ""
                MultiplierCode.MultiplierCode = ""
                MultiplierCode.MultiplierValue = 0
            End If
        End Get
        Set(ByVal value As strucMultiplier)
            If Index >= mMultiplierCodes.Count Then
                ReDim Preserve mMultiplierCodes(0 To Index)
            End If
            mMultiplierCodes(Index) = value
        End Set

    End Property

    ReadOnly Property MultiplierCodeCount()
        Get
            If IsNothing(mMultiplierCodes) Then
                Return 0
            Else
                Return mMultiplierCodes.Count
            End If
        End Get

    End Property

    'READ TEXT:
    Dim mReadTextNLines As Integer = 1
    Property ReadTextNLines As Integer
        Get
            Return mReadTextNLines
        End Get
        Set(value As Integer)
            mReadTextNLines = value
        End Set
    End Property

    Dim mReadTextString As String = ""
    Property ReadTextString As String
        Get
            Return mReadTextString
        End Get
        Set(value As String)
            mReadTextString = value
        End Set
    End Property

    'MODIFY VALUES:
    Dim mModifyValuesRegExVariable As String = ""
    Property ModifyValuesRegExVariable As String 'The name of the RegEx Variable affected by the Modifiction
        Get
            Return mModifyValuesRegExVariable
        End Get
        Set(value As String)
            mModifyValuesRegExVariable = value
        End Set
    End Property

    'Public Enum ModifyValuesTypes
    '    ConvertDate
    '    ReplaceChars
    '    FixedValue
    '    FileName
    '    FileDir
    '    FilePath
    '    CurrentDate
    '    CurrentTime
    '    CurrentDateTime
    'End Enum
    Public Enum ModifyValuesTypes
        ClearValue
        ConvertDate
        ReplaceChars
        AppendFixedValue
        AppendRegExVarValue
        AppendFileName
        AppendFileDir
        AppendFilePath
        AppendCurrentDate
        AppendCurrentTime
        AppendCurrentDateTime
    End Enum

    'Dim mModifyValuesType As String = ""
    Dim mModifyValuesType As ModifyValuesTypes 'The type of modification.
    'Property ModifyValuesType As String
    Property ModifyValuesType As ModifyValuesTypes
        Get
            Return mModifyValuesType
        End Get
        'Set(value As String)
        Set(value As ModifyValuesTypes)
            mModifyValuesType = value
        End Set
    End Property

    Dim mModifyValuesInputDateFormat As String = "" 'The Input Date Format used by the Convert Date modification.
    Property ModifyValuesInputDateFormat As String
        Get
            Return mModifyValuesInputDateFormat
        End Get
        Set(value As String)
            mModifyValuesInputDateFormat = value
        End Set
    End Property

    Dim mModifyValuesOutputDateFormat As String = "" 'The Output Date Format used by the Convert Date modification
    Property ModifyValuesOutputDateFormat As String
        Get
            Return mModifyValuesOutputDateFormat
        End Get
        Set(value As String)
            mModifyValuesOutputDateFormat = value
        End Set
    End Property

    Dim mModifyValuesCharsToReplace As String = "" 'The Characters to Replace used by the Replace Characters modification.
    Property ModifyValuesCharsToReplace As String
        Get
            Return mModifyValuesCharsToReplace
        End Get
        Set(value As String)
            mModifyValuesCharsToReplace = value
        End Set
    End Property

    Dim mModifyValuesReplacementChars As String = "" 'The replacement characters used by the Replace Characters modification.
    Property ModifyValuesReplacementChars As String
        Get
            Return mModifyValuesReplacementChars
        End Get
        Set(value As String)
            mModifyValuesReplacementChars = value
        End Set
    End Property

    Dim mModifyValuesFixedValue As String = "" 'The value appended by the Append Fixed Value modification.
    Property ModifyValuesFixedValue As String
        Get
            Return mModifyValuesFixedValue
        End Get
        Set(value As String)
            mModifyValuesFixedValue = value
        End Set
    End Property

    'Dim _modifyValuesRegExVarToAppend As String = "" 'The name of the RegEx variable containing the value to append.
    Dim _modifyValuesRegExVarValFrom As String = "" 'The name of the RegEx variable containing the value to append.
    'Property ModifyValuesRegExVarToAppend As String
    Property ModifyValuesRegExVarValFrom As String
        Get
            'Return _ModifyValuesRegExVarToAppend
            Return _modifyValuesRegExVarValFrom
        End Get
        Set(value As String)
            '_ModifyValuesRegExVarToAppend = value
            _modifyValuesRegExVarValFrom = value
        End Set
    End Property

    'Private _recordSequence As Boolean  'If True then processing sequences manually applied will be recorded in the processing sequence.
    'Property RecordSequence As Boolean
    '    Get
    '        Return _recordSequence
    '    End Get
    '    Set(value As Boolean)
    '        _recordSequence = value
    '    End Set
    'End Property

#End Region 'Properties

#Region " General Methods"

    Public Sub SaveSettings()
        'Saves the import settings of the current project in the ImportSettings.xml file (in the Project Directory)
        'The contents of the project variables are saved.
        'The import settings are saved in the following categories: IMPORT SEQUENCE, INPUT DATA, DATABASE, MATCH TEXT and DATABASE DESTINATIONS.

        If IsNothing(mSelectedFiles) Then
            ReDim mSelectedFiles(0 To 0)
            mSelectedFiles(0) = ""
        End If

        'If mRegEx() is Nothing, there is an error writing the XML file!!!!!
        If IsNothing(mRegEx) Then
            ReDim mRegEx(0 To 0)
            mRegEx(0).Name = ""
            mRegEx(0).Descr = ""
            mRegEx(0).RegEx = ""
            mRegEx(0).ExitOnMatch = False
            mRegEx(0).ExitOnNoMatch = False
            mRegEx(0).MatchStatus = ""
            mRegEx(0).NoMatchStatus = ""
        End If

        If IsNothing(mDbDest) Then
            ReDim mDbDest(0 To 0)
            mDbDest(0).RegExVariable = ""
            mDbDest(0).Type = ""
            mDbDest(0).TableName = ""
            mDbDest(0).FieldName = ""
            mDbDest(0).StatusField = ""
        End If

        If IsNothing(mMultiplierCodes) Then
            ReDim mMultiplierCodes(0 To 0)
            mMultiplierCodes(0).RegExMultiplierVariable = ""
            mMultiplierCodes(0).MultiplierCode = ""
            mMultiplierCodes(0).MultiplierValue = 1
        End If

        Dim importSettings = <?xml version="1.0" encoding="utf-8"?>
                             <!---->
                             <!--Import Settings for the Import Text Into Database application.-->
                             <Settings>
                                 <!---->
                                 <!--IMPORT SEQUENCE:-->
                                 <ImportSequence>
                                     <ImportSequenceName><%= ImportSequenceName %></ImportSequenceName>
                                     <ImportSequenceDescription><%= ImportSequenceDescription %></ImportSequenceDescription>
                                 </ImportSequence>
                                 <!---->
                                 <!--INPUT DATA:-->
                                 <InputData>
                                     <!--Text File Directory:-->
                                     <TextFileDirectory><%= TextFileDir %></TextFileDirectory>
                                     <!---->
                                     <!--List of text files selected for processing:-->
                                     <SelectedFileList>
                                         <%= From item In mSelectedFiles
                                             Select
                                       <SelectedFile><%= item %></SelectedFile>
                                         %>
                                     </SelectedFileList>
                                     <!---->
                                     <!--Select file mode: -->
                                     <SelectFileMode><%= SelectFileMode %></SelectFileMode>
                                     <!---->
                                     <!--Selection file path: -->
                                     <SelectionFileName><%= SelectionFileName %></SelectionFileName>
                                 </InputData>
                                 <!---->
                                 <!--DATABASE:-->
                                 <Database>
                                     <DatabasePath><%= DatabasePath %></DatabasePath>
                                 </Database>
                                 <!---->
                                 <!--MATCH TEXT:-->
                                 <MatchText>
                                     <RegularExpressionListName><%= RegExListName %></RegularExpressionListName>
                                     <RegularExpressionListDescription><%= RegExListDescr %></RegularExpressionListDescription>
                                     <RegularExpressionListCreationDate><%= RegExListCreationDate %></RegularExpressionListCreationDate>
                                     <RegularExpressionListLastEditDate><%= RegExListLastEditDate %></RegularExpressionListLastEditDate>
                                     <!--RegEx parameters:-->
                                     <RegularExpressionList>
                                         <%= From item In mRegEx
                                             Select
                                       <RegEx>
                                           <Name><%= item.Name %></Name>
                                           <Descr><%= item.Descr %></Descr>
                                           <Text><%= item.RegEx %></Text>
                                           <ExitOnMatch><%= item.ExitOnMatch %></ExitOnMatch>
                                           <ExitOnNoMatch><%= item.ExitOnNoMatch %></ExitOnNoMatch>
                                           <MatchStatus><%= item.MatchStatus %></MatchStatus>
                                           <NoMatchStatus><%= item.NoMatchStatus %></NoMatchStatus>
                                       </RegEx>
                                         %>
                                     </RegularExpressionList>
                                 </MatchText>
                                 <!---->
                                 <!--DATABASE DESTINATIONS:-->
                                 <DatabaseDestinations>
                                     <DatabaseDestinationListName><%= DbDestListName %></DatabaseDestinationListName>
                                     <DatabaseDestinationListDescription><%= DbDestListDescription %></DatabaseDestinationListDescription>
                                     <DatabaseDestinationListCreationDate><%= DbDestListCreationDate %></DatabaseDestinationListCreationDate>
                                     <DatabaseDestinationListLastEditDate><%= DbDestListLastEditDate %></DatabaseDestinationListLastEditDate>
                                     <!--Db Dest Parameters:-->
                                     <DatabaseDestinationList>
                                         <%= From item In mDbDest
                                             Select
                                         <DbDest>
                                             <RegExVariable><%= item.RegExVariable %></RegExVariable>
                                             <Type><%= item.Type %></Type>
                                             <TableName><%= item.TableName %></TableName>
                                             <FieldName><%= item.FieldName %></FieldName>
                                             <StatusField><%= item.StatusField %></StatusField>
                                         </DbDest>
                                         %>
                                     </DatabaseDestinationList>
                                     <UseNullValueString><%= UseNullValueString %></UseNullValueString>
                                     <NullValueString><%= NullValueString %></NullValueString>
                                 </DatabaseDestinations>
                                 <!---->
                                 <!--MULTIPLIERS:-->
                                 <Multipliers>
                                     <!--Multiplier Parameters:-->
                                     <MultiplierList>
                                         <%= From item In mMultiplierCodes
                                             Select
                                         <Multiplier>
                                             <RegExVariable><%= item.RegExMultiplierVariable %></RegExVariable>
                                             <Code><%= item.MultiplierCode %></Code>
                                             <Value><%= item.MultiplierValue %></Value>
                                         </Multiplier>
                                         %>
                                     </MultiplierList>
                                 </Multipliers>
                                 <!---->
                                 <!--READ TEXT:-->
                                 <ReadText>
                                     <ReadTextNLines><%= ReadTextNLines %></ReadTextNLines>
                                     <ReadTextString><%= ReadTextString %></ReadTextString>
                                 </ReadText>
                                 <!---->
                                 <!--MODIFY VALUES:-->
                                 <ModifyValues>
                                     <ModifyValuesRegExVariable><%= ModifyValuesRegExVariable %></ModifyValuesRegExVariable>
                                     <ModifyValuesType><%= ModifyValuesType %></ModifyValuesType>
                                     <ModifyValuesInputDateFormat><%= ModifyValuesInputDateFormat %></ModifyValuesInputDateFormat>
                                     <ModifyValuesOutputDateFormat><%= ModifyValuesOutputDateFormat %></ModifyValuesOutputDateFormat>
                                     <ModifyValuesCharsToReplace><%= ModifyValuesCharsToReplace %></ModifyValuesCharsToReplace>
                                     <ModifyValuesReplacementChars><%= ModifyValuesReplacementChars %></ModifyValuesReplacementChars>
                                     <ModifyValuesFixedValue><%= ModifyValuesFixedValue %></ModifyValuesFixedValue>
                                 </ModifyValues>
                             </Settings>

        'importSettings.Save(ProjectPath & "\" & "ImportSettings.xml")

        SettingsLocn.SaveXmlData("ImportSettings.xml", importSettings)
    End Sub


    Public Sub RestoreSettings()
        'Restores the state of the project saved in an ImportSettings.xml file.

        Dim TempRegEx As strucRegEx
        Dim TempDbDest As strucDbDest


        'If System.IO.File.Exists(ProjectPath & "\" & "ImportSettings.xml") Then
        'Debug.Print("ProjectPath = " & ProjectPath & vbCrLf)
        'Dim importSettings As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\" & "ImportSettings.xml")
        Dim importSettings As System.Xml.Linq.XDocument

        SettingsLocn.ReadXmlData("ImportSettings.xml", importSettings)

        If importSettings Is Nothing Then
            ClearSettings()
            Exit Sub
        End If

        'Read Import Sequence parameters: ------------------------------------------------------------------------
        ImportSequenceName = importSettings.<Settings>.<ImportSequence>.<ImportSequenceName>.Value
        If ImportSequenceName = Nothing Then ImportSequenceName = "" 'Debug.Print("ImportSequenceName is Nothing")
        ImportSequenceDescription = importSettings.<Settings>.<ImportSequence>.<ImportSequenceDescription>.Value
        If ImportSequenceDescription = Nothing Then ImportSequenceDescription = ""

        'Read Input Data parameters: -----------------------------------------------------------------------------
        TextFileDir = importSettings.<Settings>.<InputData>.<TextFileDirectory>.Value
        If TextFileDir = Nothing Then TextFileDir = ""

        'Read list of files selected for processing:
        Dim selectedFileList = From item In importSettings.<Settings>.<InputData>.<SelectedFileList>.<SelectedFile>
        SelTextFilesClear()
        For Each item In selectedFileList
            If item <> "" Then
                SelTextFileAppend(item)
            End If
        Next

        'Read the Select File Mode:
        SelectFileMode = importSettings.<Settings>.<InputData>.<SelectFileMode>.Value
        If SelectFileMode = Nothing Then SelectFileMode = ""

        'Read the Selection File Name:
        SelectionFileName = importSettings.<Settings>.<InputData>.<SelectionFileName>.Value
        If SelectionFileName = Nothing Then SelectionFileName = ""

        'Read database parameters: --------------------------------------------------------------------------------
        DatabasePath = importSettings.<Settings>.<Database>.<DatabasePath>.Value
        If DatabasePath = Nothing Then DatabasePath = ""

        'Read the Match Text parameters: --------------------------------------------------------------------------
        RegExListName = importSettings.<Settings>.<MatchText>.<RegularExpressionListName>.Value
        If RegExListName = Nothing Then RegExListName = ""
        RegExListDescr = importSettings.<Settings>.<MatchText>.<RegularExpressionListDescription>.Value
        If RegExListDescr = Nothing Then RegExListDescr = ""
        'If importSettings.<Settings>.<MatchText>.<RegularExpressionListCreationDate>.Value = "" Then
        If IsNothing(importSettings.<Settings>.<MatchText>.<RegularExpressionListCreationDate>.Value) Then
            'Leave RegExListCreateDate unchanged.
        Else
            RegExListCreationDate = importSettings.<Settings>.<MatchText>.<RegularExpressionListCreationDate>.Value
            'RegExListCreationDate = CDate(importSettings.<Settings>.<MatchText>.<RegularExpressionListCreationDate>.Value)
        End If
        'If importSettings.<Settings>.<MatchText>.<RegularExpressionListLastEditDate>.Value = "" Then
        If IsNothing(importSettings.<Settings>.<MatchText>.<RegularExpressionListLastEditDate>.Value) Then
            'Leave RegExListLastEditDate unchanged.
        Else
            RegExListLastEditDate = importSettings.<Settings>.<MatchText>.<RegularExpressionListLastEditDate>.Value
            'RegExListLastEditDate = CDate(importSettings.<Settings>.<MatchText>.<RegularExpressionListLastEditDate>.Value)
        End If

        'Read the Regular Expression list:
        Dim RegExList = From item In importSettings.<Settings>.<MatchText>.<RegularExpressionList>.<RegEx>
        RegExClear()
        For Each item In RegExList
            TempRegEx.Name = item.<Name>.Value
            TempRegEx.Descr = item.<Descr>.Value
            TempRegEx.RegEx = item.<Text>.Value
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

            If TempRegEx.Name <> "" Then
                RegExAppend(TempRegEx)
            End If
        Next

        'Read the Database Destinations parameters: -------------------------------------------------------------------------
        DbDestListName = importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListName>.Value
        If DbDestListName = Nothing Then DbDestListName = ""
        DbDestListDescription = importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListDescription>.Value
        If DbDestListDescription = Nothing Then DbDestListDescription = ""
        'If importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListCreationDate>.Value = "" Then
        If IsNothing(importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListCreationDate>.Value) Then
            'Leave DbDestListCreationDate unchanged.
        Else
            DbDestListCreationDate = importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListCreationDate>.Value
        End If
        'If importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListLastEditDate>.Value = "" Then
        If IsNothing(importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListLastEditDate>.Value) Then
            'Leave DbDestListLastEditDate unchanged.
        Else
            DbDestListLastEditDate = importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationListLastEditDate>.Value
        End If

        Dim DbDestList = From item In importSettings.<Settings>.<DatabaseDestinations>.<DatabaseDestinationList>.<DbDest>
        DbDestClear()
        For Each item In DbDestList
            TempDbDest.RegExVariable = item.<RegExVariable>.Value
            TempDbDest.Type = item.<Type>.Value
            TempDbDest.TableName = item.<TableName>.Value
            TempDbDest.FieldName = item.<FieldName>.Value
            TempDbDest.StatusField = item.<StatusField>.Value
            DbDestAppend(TempDbDest)
        Next

        If IsNothing(importSettings.<Settings>.<DatabaseDestinations>.<UseNullValueString>.Value) Then
            UseNullValueString = False
        Else
            If importSettings.<Settings>.<DatabaseDestinations>.<UseNullValueString>.Value = "true" Then
                UseNullValueString = True
            Else
                UseNullValueString = False
            End If
        End If
        If IsNothing(importSettings.<Settings>.<DatabaseDestinations>.<NullValueString>.Value) Then
            NullValueString = ""
        Else
            NullValueString = importSettings.<Settings>.<DatabaseDestinations>.<NullValueString>.Value
        End If

        'Read Multiplier Parameters: --------------------------------------------------------------------------------------
        Dim TempMult As strucMultiplier
        Dim multiplierParams = From item In importSettings.<Settings>.<Multipliers>.<MultiplierList>.<Multiplier>
        MultipliersClear()
        For Each item In multiplierParams
            TempMult.RegExMultiplierVariable = item.<RegExVariable>.Value
            TempMult.MultiplierCode = item.<Code>.Value
            TempMult.MultiplierValue = item.<Value>.Value
            MultipliersAppend(TempMult)
        Next

        'Read Read Text parameters: ----------------------------------------------------------------------------------------
        ReadTextNLines = importSettings.<Settings>.<ReadText>.<ReadTextNLines>.Value
        If ReadTextNLines = Nothing Then ReadTextNLines = 0
        ReadTextString = importSettings.<Settings>.<ReadText>.<ReadTextString>.Value
        If ReadTextString = Nothing Then ReadTextString = ""

        'Read Modify Values parameters: ------------------------------------------------------------------------------------
        ModifyValuesRegExVariable = importSettings.<Settings>.<ModifyValues>.<ModifyValuesRegExVariable>.Value
        If ModifyValuesRegExVariable = Nothing Then ModifyValuesRegExVariable = ""
        If IsNothing(importSettings.<Settings>.<ModifyValues>.<ModifyValuesType>.Value) Then
            ModifyValuesType = ModifyValuesTypes.ConvertDate
        Else
            Select Case importSettings.<Settings>.<ModifyValues>.<ModifyValuesType>.Value
                Case "ConvertDate"
                    ModifyValuesType = ModifyValuesTypes.ConvertDate
                'Case "CurrentDate"
                Case "AppendCurrentDate"
                    ModifyValuesType = ModifyValuesTypes.AppendCurrentDate
                'Case "CurrentDateTime"
                Case "AppendCurrentDateTime"
                    ModifyValuesType = ModifyValuesTypes.AppendCurrentDateTime
                'Case "CurrentTime"
                Case "AppendCurrentTime"
                    ModifyValuesType = ModifyValuesTypes.AppendCurrentTime
                'Case "FileDir"
                Case "AppendFileDir"
                    ModifyValuesType = ModifyValuesTypes.AppendFileDir
                'Case "FileName"
                Case "AppendFileName"
                    ModifyValuesType = ModifyValuesTypes.AppendFileName
                'Case "FilePath"
                Case "AppendFilePath"
                    ModifyValuesType = ModifyValuesTypes.AppendFilePath
                'Case "FixedValue"
                Case "AppendFixedValue"
                    ModifyValuesType = ModifyValuesTypes.AppendFixedValue
                Case "ReplaceChars"
                    ModifyValuesType = ModifyValuesTypes.ReplaceChars
                Case "AppendRegExVarValue"
                    ModifyValuesType = ModifyValuesTypes.AppendRegExVarValue
                Case "ClearValue"
                    ModifyValuesType = ModifyValuesTypes.ClearValue
                Case Else
                    ModifyValuesType = ModifyValuesTypes.ConvertDate
            End Select
        End If
        'ModifyValuesType = importSettings.<Settings>.<ModifyValues>.<ModifyValuesType>.Value
        'If ModifyValuesType = Nothing Then ModifyValuesType = ""
        ModifyValuesInputDateFormat = importSettings.<Settings>.<ModifyValues>.<ModifyValuesInputDateFormat>.Value
        If ModifyValuesInputDateFormat = Nothing Then ModifyValuesInputDateFormat = ""
        ModifyValuesOutputDateFormat = importSettings.<Settings>.<ModifyValues>.<ModifyValuesOutputDateFormat>.Value
        If ModifyValuesOutputDateFormat = Nothing Then ModifyValuesOutputDateFormat = ""
        ModifyValuesCharsToReplace = importSettings.<Settings>.<ModifyValues>.<ModifyValuesCharsToReplace>.Value
        If ModifyValuesCharsToReplace = Nothing Then ModifyValuesCharsToReplace = ""
        ModifyValuesReplacementChars = importSettings.<Settings>.<ModifyValues>.<ModifyValuesReplacementChars>.Value
        If ModifyValuesReplacementChars = Nothing Then ModifyValuesReplacementChars = ""
        ModifyValuesFixedValue = importSettings.<Settings>.<ModifyValues>.<ModifyValuesFixedValue>.Value
        If ModifyValuesFixedValue = Nothing Then ModifyValuesFixedValue = ""

        'Else 'Import Settings xml file not found
        '    Debug.Print("No valid ProjectPath: " & ProjectPath & vbCrLf)
        'End If
    End Sub

    Public Sub ClearSettings()
        'Clear the import settings.

        ImportSequenceName = ""
        ImportSequenceDescription = ""
        TextFileDir = ""
        SelTextFilesClear()
        SelectFileMode = ""
        SelectionFileName = ""
        DatabasePath = ""
        RegExListName = ""
        RegExListDescr = ""
        RegExClear()
        DbDestListName = ""
        DbDestListDescription = ""
        DbDestClear()
        UseNullValueString = False
        NullValueString = ""
        MultipliersClear()
        ReadTextNLines = 0
        ReadTextString = ""
        ModifyValuesRegExVariable = ""
        ModifyValuesType = ModifyValuesTypes.ConvertDate
        ModifyValuesInputDateFormat = ""
        ModifyValuesOutputDateFormat = ""
        ModifyValuesCharsToReplace = ""
        ModifyValuesReplacementChars = ""
        ModifyValuesFixedValue = ""

    End Sub

#End Region 'General Methods

#Region " RegEx Methods"

    Public Sub RegExClear()
        'Clears the contents of the mRegEx array.
        mRegEx = Nothing
    End Sub

    Public Sub RegExAppend(ByVal RegEx As strucRegEx)
        'Appends the RegEx to the end of the mRegEx() array:

        Dim Count As Integer

        If IsNothing(mRegEx) Then 'RegEx() contains mo elements
            ReDim mRegEx(0 To 0)
            mRegEx(0) = RegEx
        Else
            Count = mRegEx.Count
            ReDim Preserve mRegEx(0 To Count)
            mRegEx(Count) = RegEx
        End If
    End Sub

    Public Sub RegExInsert(ByVal Index As Integer, ByVal RegEx As strucRegEx)
        'Inserts the RegEx at the Index loaction in the mRegEx() array:

        Dim Count As Integer 'The number of elements in the _RegEx() array
        Dim I As Integer 'Loop index

        If IsNothing(mRegEx) Then '_RegEx() contains no elements
            If Index = 0 Then
                ReDim mRegEx(0 To 0)
                mRegEx(0) = RegEx
            Else
                'NOTE: Convert this code to a warning message event!!!
                MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
            End If
        Else
            Count = mRegEx.Count
            If Index < Count Then 'The new elements will be inserted within the current array
                ReDim Preserve mRegEx(0 To Count)
                For I = Count To Index Step -1
                    mRegEx(I) = mRegEx(I - 1)
                Next
                mRegEx(Count) = RegEx
            Else
                If Index = Count Then 'The new element will be added to the end of the array
                    ReDim Preserve mRegEx(0 To Count)
                    mRegEx(Count) = RegEx
                Else 'Index value is too large - beyond the end of the current array
                    'NOTE: Convert this code to a warning message event!!!
                    MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
                End If
            End If
        End If
    End Sub

    Public Sub RegExModify(ByVal Index As Integer, ByVal RegEx As strucRegEx)
        'Modify the RegEx at the specified Index location

        Dim Count As Integer

        If IsNothing(mRegEx) Then '_RegEx() contains no elements
            If Index = 0 Then
                ReDim mRegEx(0 To 0)
                mRegEx(0) = RegEx
            Else
                'NOTE: Convert this code to a warning message event!!!
                MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
            End If
        Else
            Count = mRegEx.Count
            If Index < Count Then
                mRegEx(Index) = RegEx
            Else
                If Index = Count Then 'Append the RegEx to the end of the RegEx array
                    ReDim Preserve mRegEx(0 To Count)
                    mRegEx(Count) = RegEx
                Else
                    'Index is larger then the size of the _simpleRegEx() array.
                    'NOTE: Convert this code to a warning message event!!!
                    MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
                End If

            End If
        End If
    End Sub

    Public Sub RegExDelete(ByVal Index As Integer)
        'Deletes the RegEx located at the specified index location in the mRegEx() array

        Dim Count As Integer
        Dim I As Integer

        If IsNothing(mRegEx) Then '_RegEx() contains no elements

        Else
            Count = mRegEx.Count
            If Index < Count - 1 Then 'Move higher elements down
                For I = Index To Count - 2
                    mRegEx(I) = mRegEx(I + 1)
                Next
            End If
            ReDim Preserve mRegEx(0 To Count - 2)
        End If
    End Sub

    'Public Sub OpenRegExListFileOld()
    '    'Opens the Regular Expression List File with name RegExListName.

    '    'Read the XML file:
    '    Dim Index As Integer
    '    Dim TempRegEx As strucRegEx
    '    'If System.IO.File.Exists(TextToDatabase.ProjectPath & "\" & txtRegExList.Text & ".Regexlist") Then
    '    If System.IO.File.Exists(ProjectPath & "\" & RegExListName & ".Regexlist") Then
    '        'Dim RegExList As System.Xml.Linq.XDocument = XDocument.Load(TextToDatabase.ProjectPath & "\" & txtRegExList.Text & ".Regexlist")
    '        Dim RegExList As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\" & RegExListName & ".Regexlist")
    '        'txtListDescription.Text = RegExList.<RegularExpressionList>.<Descr>.Value
    '        'txtListDescription.Text = RegExList.<RegularExpressionList>.<Description>.Value
    '        RegExListDescr = RegExList.<RegularExpressionList>.<Description>.Value
    '        'TextToDatabase.RegExListDescr = txtListDescription.Text
    '        'TextToDatabase.RegExListName = txtRegExList.Text
    '        'TextToDatabase.RegExListCreationDate = CStr(RegExList.<RegularExpressionList>.<CreationDate>.Value)
    '        RegExListCreationDate = CStr(RegExList.<RegularExpressionList>.<CreationDate>.Value)
    '        'TextToDatabase.
    '        Dim RegExs = From item In RegExList.<RegularExpressionList>.<RegularExpression>
    '        Index = 0
    '        'TextToDatabase.RegExClear()
    '        RegExClear()
    '        For Each item In RegExs
    '            TempRegEx.Name = item.<Name>.Value
    '            TempRegEx.Descr = item.<Descr>.Value
    '            TempRegEx.RegEx = item.<RegEx>.Value
    '            If item.<ExitOnMatch>.Value = "true" Then
    '                TempRegEx.ExitOnMatch = True
    '            Else
    '                TempRegEx.ExitOnMatch = False
    '            End If
    '            If item.<ExitOnNoMatch>.Value = "true" Then
    '                TempRegEx.ExitOnNoMatch = True
    '            Else
    '                TempRegEx.ExitOnNoMatch = False
    '            End If
    '            TempRegEx.MatchStatus = item.<MatchStatus>.Value
    '            TempRegEx.NoMatchStatus = item.<NoMatchStatus>.Value
    '            RegExAppend(TempRegEx)
    '        Next
    '    Else
    '        'RaiseEvent Warning("RegEx List file not found: " & RegExListName)
    '        RaiseEvent ErrorMessage("RegEx List file not found: " & RegExListName & vbCrLf)
    '    End If

    'End Sub

    Public Sub OpenRegExListFile()
        'Opens the Regular Expression List File with name RegExListName.

        'Read the XML file:
        Dim Index As Integer
        Dim TempRegEx As strucRegEx

        'If DataLocn.FileExists(RegExListName & ".Regexlist") Then
        If DataLocn.FileExists(RegExListName) Then
            Dim RegExList As System.Xml.Linq.XDocument
            'DataLocn.ReadXmlData(RegExListName & ".Regexlist", RegExList)
            DataLocn.ReadXmlData(RegExListName, RegExList)
            RegExListDescr = RegExList.<RegularExpressionList>.<Description>.Value
            RegExListCreationDate = CStr(RegExList.<RegularExpressionList>.<CreationDate>.Value)
            Dim RegExs = From item In RegExList.<RegularExpressionList>.<RegularExpression>
            Index = 0
            RegExClear()
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
                RegExAppend(TempRegEx)
            Next
        Else
            RaiseEvent ErrorMessage("RegEx List file not found: " & RegExListName & vbCrLf)
        End If

    End Sub

    Public Sub RunRegEx(ByVal IndexNo As Integer)
        'Runs the Regular Expression in the RegEx() list corresponding to the specified index.

        'RegEx() contains the list of Regular Expressions.
        'RegExCount is the number of RegExs in the list.

        'Check if the RegExIndex number is valid:
        If IndexNo + 1 > RegExCount Then   'There is no RegEx at index position RegExIndex.
            'RaiseEvent Warning("There is no RegEx at index position " & Str(IndexNo))
            RaiseEvent ErrorMessage("There is no RegEx at index position " & Str(IndexNo) & vbCrLf)
            Exit Sub
        End If

        'Check if there is text in the TextStore for the RegEx:
        If TextStore = "" Then 'TextStore is blank.
            'RaiseEvent Notice("There is no text in the TextStore")
            RaiseEvent Message("There is no text in the TextStore" & vbCrLf)
            Exit Sub
        End If

        Try
            Dim myRegEx As New System.Text.RegularExpressions.Regex(RegEx(IndexNo).RegEx)
            Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(TextStore)
            If myMatch.Success Then 'The Regular Expression found a match.
                'RaiseEvent Notice("Match")
                SaveTextMatch(myMatch)
            Else
                'RaiseEvent Notice("No Match")
                RaiseEvent Message("No Match" & vbCrLf)
            End If
        Catch ex As Exception
            'RaiseEvent Warning(ex.GetType.ToString)
            RaiseEvent ErrorMessage(ex.GetType.ToString & vbCrLf)
            'RaiseEvent Warning(ex.Message.ToString)
            RaiseEvent ErrorMessage(ex.Message.ToString & vbCrLf)
        End Try

    End Sub

    Public Sub RunRegExList()
        'Runs all the Regular Expressions in the RegEx() list.

        'RegEx() contains the list of Regular Expressions.
        'RegExCount is the number of RegExs in the list.

        Dim I As Integer

        'Check if there is text in the TextStore for the RegEx:
        If TextStore = "" Then 'TextStore is blank.
            'RaiseEvent Message("There is no text in the TextStore" & vbCrLf)
            RaiseEvent ErrorMessage("There is no text in the TextStore" & vbCrLf)
            Exit Sub
        End If

        For I = 0 To RegExCount - 1
            'Process each RegEx:
            Try
                Dim myRegEx As New System.Text.RegularExpressions.Regex(RegEx(I).RegEx)
                Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(TextStore)
                If myMatch.Success Then 'The Regular Expression found a match.
                    'RaiseEvent Message(myMatch.Captures.Count & " matches for " & RegEx(I).Name)
                    'RaiseEvent Message(myMatch.Groups(1).Captures.Count & " matches for " & RegEx(I).Name)
                    SaveTextMatch(myMatch)
                Else
                    RaiseEvent Message("No Match: " & RegEx(I).Name)
                End If
            Catch ex As Exception
                RaiseEvent ErrorMessage(ex.GetType.ToString & vbCrLf)
                RaiseEvent ErrorMessage(ex.Message.ToString & vbCrLf)
            End Try
        Next

    End Sub

#End Region 'RegEx Methods

#Region " Match Text Methods"

    Public Sub SaveTextMatch(ByRef myMatch As System.Text.RegularExpressions.Match)
        'Write the RegEx text match into the Database Destinations table (DbDestValues):

        'Debug.Print("Running SaveTextMatch" & vbCrLf)

        'If IsNothing(DbDestValues) Then
        If IsNothing(mDbDest) Then 'The Database Destinations have not been specified
            RaiseEvent ErrorMessage("The Database Destinations have not been specified." & vbCrLf)
            'Debug.Print("DbDestValues() = Nothing" & vbCrLf)
            'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
            'ShowMessage("Running SaveTextMatch() subroutine. DbDestValues = Nothing!" & vbCrLf, Color.Red)
        Else
            Dim I As Integer
            Dim strVarName As String
            Dim CapCount As Integer
            Dim J As Integer
            'Find matching RegEx Variables in the Database Destinations grid:
            'For I = 1 To TextToDatabase.DbDestinations.DataGridView1.RowCount
            Dim NRows As Integer = 0
            Dim NCols As Integer = 0
            NRows = mDbDest.Count
            'UPDATE 27Nov16 - Bug found in subroutine:
            '                 10 items (dates) matched so DbDestValues redimensioned to fit 10 items.
            '                 later 4 items matched and DbDestValues reduced to fit 4 items resulting in loss of some of the earlier matches.
            'To fix, NCols is set to the current number of columns in DbDestValues
            'NCols = 1
            NCols = DbDestValues.GetLength(1)
            'RaiseEvent Message("DbDestValues.GetLength(1) = " & NCols)
            For I = 1 To NRows
                If DbDest(I - 1).RegExVariable = "" Then 'No RegEx variable name is specified.

                Else
                    strVarName = DbDest(I - 1).RegExVariable 'The RegEx variable name.
                    CapCount = myMatch.Groups(strVarName).Captures.Count 'Get the number of captures
                    If CapCount = 1 Then
                        DbDestValues(I - 1, 0) = myMatch.Groups(strVarName).ToString
                    ElseIf CapCount > 1 Then 'Multiple matches found.
                        If NCols < CapCount Then
                            ReDim Preserve DbDestValues(0 To NRows - 1, 0 To CapCount - 1)
                        End If
                        For J = 1 To CapCount
                            DbDestValues(I - 1, J - 1) = myMatch.Groups(strVarName).Captures.Item(J - 1).ToString
                            'Debug.Print("DbDestValues(" & Str(I - 1) & "," & Str(J - 1) & ") = " & myMatch.Groups(strVarName).Captures.Item(J - 1).ToString)
                        Next
                    Else 'No matches found.

                    End If
                End If
            Next
        End If
    End Sub

#End Region 'Match Text Methods

#Region " Select Files Methods"

    Public Sub SelTextFilesClear()
        'Clears the contents of the Selected Text Files array
        mSelectedFiles = Nothing
    End Sub

    Public Sub SelTextFileAppend(ByVal TextFileName As String)
        'Append a Selected Text File to the end of the Selected text Files array:

        Dim Count As Integer

        If IsNothing(mSelectedFiles) Then '_selTextFiles() contains no elements
            ReDim mSelectedFiles(0 To 0)
            mSelectedFiles(0) = TextFileName
        Else
            Count = mSelectedFiles.Count
            'First check if TextFileName is already in the list of selected files:
            Dim I As Integer
            Dim FileFound As Boolean = False
            For I = 0 To Count - 1
                If mSelectedFiles(I) = TextFileName Then
                    FileFound = True
                    Exit For
                End If
            Next
            If FileFound = False Then
                ReDim Preserve mSelectedFiles(0 To Count)
                mSelectedFiles(Count) = TextFileName
            End If
        End If
    End Sub

    Public Sub SelTextFileRemove(ByVal TextFileName As String)
        'Remove a text file if it is present in the list of selected files:

        Dim Count As Integer
        Dim FileFound As Boolean = False
        Dim FilePosn As Integer 'The position of the file if it is found in the file list

        If IsNothing(mSelectedFiles) Then

        Else
            Count = mSelectedFiles.Count
            If Count > 0 Then
                Dim I As Integer
                For I = 0 To Count - 1
                    If mSelectedFiles(I) = TextFileName Then
                        FileFound = True
                        FilePosn = I
                        Exit For
                    End If
                Next
                If FileFound = True Then 'Remove the file from the list:
                    For I = FilePosn To Count - 2
                        mSelectedFiles(I) = mSelectedFiles(I + 1)
                    Next
                    ReDim Preserve mSelectedFiles(0 To Count - 2)
                End If
            End If
        End If
    End Sub

#End Region 'Select Files Methods

#Region "Database Destinations Methods"

    Public Sub DbDestClear()
        'Clears the contents of the SimpleDbDest array.
        mDbDest = Nothing
        'Also clear the RegEx matched values array:
        DbDestValues = Nothing
    End Sub

    Public Sub DbDestAppend(ByVal DbDest As strucDbDest)
        'Appends the DbDest strings to the end of the mDbDest array.

        Dim Count As Integer

        If IsNothing(mDbDest) Then 'mDbDest() contains no elements
            ReDim mDbDest(0 To 0)
            ReDim DbDestValues(0 To 0, 0 To 0) 'Also redimension the corrsponding RegEx matched values array.
            mDbDest(0) = DbDest
        Else
            Count = mDbDest.Count
            ReDim Preserve mDbDest(0 To Count)
            ReDim DbDestValues(0 To Count, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
            mDbDest(Count) = DbDest
        End If
    End Sub

    Public Sub DbDestInsertBlank(ByVal Index As Integer)
        'Inserts a blank entry at the specified Index location

        Dim Count As Integer 'The number of elements in the mDbDest array
        Dim I As Integer 'loop index

        If IsNothing(mDbDest) Then 'mDbDest() contains no elements
            If Index = 0 Then
                ReDim mDbDest(0 To 0)
                ReDim DbDestValues(0 To 0, 0 To 0) 'Also redimension the corrsponding RegEx matched values array.
                mDbDest(0).FieldName = ""
                mDbDest(0).RegExVariable = ""
                mDbDest(0).StatusField = ""
                mDbDest(0).TableName = ""
                mDbDest(0).Type = ""
            Else
                'MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
                RaiseEvent ErrorMessage("DbDestInsertBlank(Index) error: Index number is too large! DbDest = nothing and Index <> 0")
            End If
        Else
            Count = mDbDest.Count
            If Index < Count Then 'The new element will be inserted within the current array.
                'ReDim Preserve mDbDest(Count + 1)
                'ReDim DbDestValues(0 To Count + 1, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
                'Make room for the new element in the array:
                ReDim Preserve mDbDest(0 To Count)
                ReDim DbDestValues(0 To Count, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
                'For I = Count To Index Step -1
                For I = Count - 1 To Index Step -1 'To fix out of range error - 16Nov16
                    mDbDest(I + 1) = mDbDest(I)
                Next
                mDbDest(Index).FieldName = ""
                mDbDest(Index).RegExVariable = ""
                mDbDest(Index).StatusField = ""
                mDbDest(Index).TableName = ""
                mDbDest(Index).Type = ""
            Else
                If Index = Count Then 'The new element will be added to the end of the array.
                    'ReDim Preserve mDbDest(Count + 1)
                    'ReDim DbDestValues(0 To Count + 1, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
                    ReDim Preserve mDbDest(0 To Index)
                    ReDim DbDestValues(0 To Index, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
                    'mDbDest(Count + 1).FieldName = ""
                    'mDbDest(Count + 1).RegExVariable = ""
                    'mDbDest(Count + 1).StatusField = ""
                    'mDbDest(Count + 1).TableName = ""
                    'mDbDest(Count + 1).Type = ""
                    mDbDest(Index).FieldName = ""
                    mDbDest(Index).RegExVariable = ""
                    mDbDest(Index).StatusField = ""
                    mDbDest(Index).TableName = ""
                    mDbDest(Index).Type = ""
                Else 'The new element has an Index value beyond the end of the current array.
                    'Leave the array unchanged.
                    'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                    MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
                    RaiseEvent ErrorMessage("DbDestInsertBlank(Index) error: Index number is too large! Index > Count")
                End If
            End If
        End If

    End Sub

    Public Sub DbDestInsert(ByVal Index As Integer, ByVal DbDest As strucDbDest)
        'Inserts the DbDest parameters at the specified Index location.
        'The existing parameters at the Index location are moved up, together with the later parameters.
        'Note Index is zero based: the first element has Index = zero.

        Dim Count As Integer 'The number of elements in the _simpleRegEx() array
        Dim I As Integer 'loop index

        If IsNothing(mDbDest) Then '_simpleDbDest() contains no elements
            If Index = 0 Then
                ReDim mDbDest(0 To 0)
                ReDim DbDestValues(0 To 0, 0 To 0) 'Also redimension the corrsponding RegEx matched values array.
                mDbDest(0) = DbDest
            Else
                'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
            End If
        Else
            Count = mDbDest.Count
            If Index < Count Then 'The new element will be inserted within the current array.
                ReDim Preserve mDbDest(Count + 1)
                ReDim DbDestValues(0 To Count + 1, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
                For I = Count To Index Step -1
                    mDbDest(I + 1) = mDbDest(I)
                Next
                mDbDest(Index) = DbDest
            Else
                If Index = Count Then 'The new element will be added to the end of the array.
                    ReDim Preserve mDbDest(Count + 1)
                    ReDim DbDestValues(0 To Count + 1, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
                    mDbDest(Count + 2) = DbDest
                Else 'The new element has an Index value beyond the end of the current array.
                    'Leave the array unchanged.
                    'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                    MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
                End If
            End If
        End If
    End Sub

    Public Sub DbDestModify(ByVal Index As Integer, ByVal DbDest As strucDbDest)
        'Modify the DbDest strings at the specified Index location

        Dim Count As Integer

        If IsNothing(mDbDest) Then 'mDbDest() contains no elements
            If Index = 0 Then
                ReDim mDbDest(0 To 0)
                mDbDest(0) = DbDest
                'Debug.Print("Creating DbDest(0) entry.")
            Else
                'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                MsgBox("Index number is too large!", MsgBoxStyle.Information, "Notice")
            End If
        Else
            Count = mDbDest.Count
            'If Index < Count Then
            If Index < Count Then
                mDbDest(Index) = DbDest
                'Debug.Print("Creating DbDest(" & Str(Index) & ") entry.")
            Else
                If Index = Count Then
                    ReDim Preserve mDbDest(0 To Count)
                    mDbDest(Index) = DbDest
                    'Debug.Print("Creating DbDest(" & Str(Index) & ") entry.")
                Else
                    'Index is larger then the size of the _simpleRegEx() array.
                End If

            End If
        End If
    End Sub

    Public Sub DbDestDelete(ByVal Index As Integer)
        'Deletes the mDbDest (database destination) parameters located at the specified index location

        Dim Count As Integer
        Dim I As Integer

        If IsNothing(mDbDest) Then 'mDbDest() contains no elements

        Else
            Count = mDbDest.Count - 1
            For I = Index To Count - 1
                mDbDest(I) = mDbDest(I + 1)
            Next
            ReDim Preserve mDbDest(Count - 1)
            ReDim DbDestValues(0 To Count - 1, 0 To 0) 'Also redimension the corresponding RegEx matched values array.
        End If

        'For debugging:
        'Display the updated contents of mDbDest
        For I = 1 To mDbDest.Count
            Debug.Print("Row: " & I & "  Field name: " & mDbDest(I - 1).FieldName)
        Next


    End Sub

    'Public Sub OpenDbDestListFileOld()
    '    'Opens the Database Destinations List File (DbDestListName).
    '    'This file should be located in the project directory.

    '    'Read the XML file:
    '    Dim Index As Integer
    '    Dim TempDbDest As strucDbDest
    '    Dim TempDbMult As strucMultiplier
    '    If System.IO.File.Exists(ProjectPath & "\" & DbDestListName & ".Dbdestlist") Then
    '        Dim DbDestList As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\" & DbDestListName & ".Dbdestlist")
    '        DbDestListDescription = DbDestList.<DatabaseDestinations>.<Description>.Value

    '        If DbDestList.<DatabaseDestinations>.<CreationDate>.Value = "" Then

    '        Else
    '            DbDestListCreationDate = DbDestList.<DatabaseDestinations>.<CreationDate>.Value
    '        End If

    '        If DbDestList.<DatabaseDestinations>.<LastEditDate>.Value = "" Then

    '        Else
    '            DbDestListCreationDate = DbDestList.<DatabaseDestinations>.<CreationDate>.Value
    '        End If

    '        Dim DbDests = From item In DbDestList.<DatabaseDestinations>.<DestinationList>.<DatabaseDestination>
    '        Index = 0
    '        DbDestClear()
    '        'Read each Database Destination entry:
    '        For Each item In DbDests
    '            TempDbDest.RegExVariable = item.<RegExVariable>.Value
    '            TempDbDest.Type = item.<Type>.Value
    '            TempDbDest.TableName = item.<TableName>.Value
    '            TempDbDest.FieldName = item.<FieldName>.Value
    '            TempDbDest.StatusField = item.<StatusField>.Value
    '            DbDestAppend(TempDbDest)
    '        Next
    '        'Dimension the DbDestValues() array:
    '        'The array is dimensioned using row, column order: DbDestValues(0 to NRows - 1, 0 to NCols - 1)
    '        ReDim DbDestValues(0 To DbDestCount - 1, 0 To 0)

    '        'Read the Multipliers:
    '        Dim DbMult = From item In DbDestList.<DatabaseDestinations>.<MultiplierList>.<Multiplier>
    '        Index = 0
    '        MultipliersClear()
    '        'Read each Multiplier entry:
    '        For Each item In DbMult
    '            TempDbMult.RegExMultiplierVariable = item.<RegExMultiplier>.Value
    '            TempDbMult.MultiplierCode = item.<MultiplierCode>.Value
    '            TempDbMult.MultiplierValue = item.<MultiplierValue>.Value
    '            MultipliersAppend(TempDbMult)
    '        Next
    '    Else
    '        'RaiseEvent Warning("Database Destinations List file not found: " & DbDestListName)
    '        RaiseEvent ErrorMessage("Database Destinations List file not found: " & DbDestListName & vbCrLf)
    '    End If
    'End Sub

    Public Sub OpenDbDestListFile()
        'Opens the Database Destinations List File (DbDestListName).

        'Read the XML file:
        Dim Index As Integer
        Dim TempDbDest As strucDbDest
        Dim TempDbMult As strucMultiplier

        'If DataLocn.FileExists(DbDestListName & ".Dbdestlist") Then
        If DataLocn.FileExists(DbDestListName) Then
            Dim DbDestList As System.Xml.Linq.XDocument
            'DataLocn.ReadXmlData(DbDestListName & ".Dbdestlist", DbDestList)
            DataLocn.ReadXmlData(DbDestListName, DbDestList)
            DbDestListDescription = DbDestList.<DatabaseDestinations>.<Description>.Value

            If DbDestList.<DatabaseDestinations>.<CreationDate>.Value = "" Then

            Else
                DbDestListCreationDate = DbDestList.<DatabaseDestinations>.<CreationDate>.Value
            End If

            If DbDestList.<DatabaseDestinations>.<LastEditDate>.Value = "" Then

            Else
                'DbDestListCreationDate = DbDestList.<DatabaseDestinations>.<CreationDate>.Value
                DbDestListLastEditDate = DbDestList.<DatabaseDestinations>.<LastEditDate>.Value
            End If

            Dim DbDests = From item In DbDestList.<DatabaseDestinations>.<DestinationList>.<DatabaseDestination>
            Index = 0
            DbDestClear()
            'Read each Database Destination entry:
            For Each item In DbDests
                TempDbDest.RegExVariable = item.<RegExVariable>.Value
                TempDbDest.Type = item.<Type>.Value
                TempDbDest.TableName = item.<TableName>.Value
                TempDbDest.FieldName = item.<FieldName>.Value
                TempDbDest.StatusField = item.<StatusField>.Value
                DbDestAppend(TempDbDest)
            Next
            'Dimension the DbDestValues() array:
            'The array is dimensioned using row, column order: DbDestValues(0 to NRows - 1, 0 to NCols - 1)
            ReDim DbDestValues(0 To DbDestCount - 1, 0 To 0)

            'Read the Multipliers:
            Dim DbMult = From item In DbDestList.<DatabaseDestinations>.<MultiplierList>.<Multiplier>
            Index = 0
            MultipliersClear()
            'Read each Multiplier entry:
            For Each item In DbMult
                TempDbMult.RegExMultiplierVariable = item.<RegExMultiplier>.Value
                TempDbMult.MultiplierCode = item.<MultiplierCode>.Value
                TempDbMult.MultiplierValue = item.<MultiplierValue>.Value
                MultipliersAppend(TempDbMult)
            Next
        Else
            RaiseEvent ErrorMessage("Database Destinations List file not found: " & DbDestListName & vbCrLf)
        End If

    End Sub

    Public Sub RemoveBlankRows()
        'Removes blank rows from Database Destinations
        'Remove blank entries from DbDest

        Dim I As Integer 'Loop index
        Dim J As Integer 'Loop index
        Dim MaxRow As Integer
        Dim NRowsRemoved As Integer = 0
        MaxRow = mDbDest.Count

        For I = 0 To MaxRow - 1
            If mDbDest(I).FieldName = "" Then
                For J = I To MaxRow - 2
                    mDbDest(J) = mDbDest(J + 1)
                Next
                NRowsRemoved = NRowsRemoved + 1
            End If
        Next

        ReDim Preserve mDbDest(0 To MaxRow - 1 - NRowsRemoved)

    End Sub

    Public Sub ClearDbDestValues()
        'Clear the values in the DbDestValues() array

        Dim NRows As Integer = 1
        Dim NCols As Integer = 1

        NRows = mDbDest.Count

        ReDim DbDestValues(0 To NRows - 1, 0 To NCols - 1)

    End Sub

#End Region 'Database Destinations Methods

#Region "Modify Values Methods"

    Public Sub ModifyValuesApply()
        'Apply the modification to the Database Destinations table:
        'This method uses these properties:
        '    ModifyValuesType (Convert_date, Replace_characters, Fixed_value, Text_file_name, Text_file_directory, Text_file_path, Current_date, Current_time, Current_date_time)


        Dim I As Integer
        Dim strVarName As String 'This is used to hold the RegEx variable name in the current row
        Dim CapCount As Integer
        Dim J As Integer
        Dim K As Integer
        Dim RegExVarRow As Integer = -1 'The row number of the RegEx Variable used in a AppendRegExVarValue modification
        Dim RegExVarType As String = "" 'The type of the RegEx Variable used in a AppendRegExVarValue modification

        'If the ModifyValuesType is AppendRegExVarValue, find the row of the RegEx variable:
        If ModifyValuesType = ModifyValuesTypes.AppendRegExVarValue Then
            For I = 1 To DbDestCount
                'ModifyValuesRegExVar contains the name of the RegEx variable used in the modification
                'If DbDest(I - 1).RegExVariable = ModifyValuesRegExVarToAppend Then
                If DbDest(I - 1).RegExVariable = ModifyValuesRegExVarValFrom Then
                    RegExVarRow = I - 1
                    RegExVarType = DbDest(I - 1).Type
                    'RaiseEvent Message("RegEx variable to append: " & ModifyValuesRegExVarToAppend & " found at row number: " & RegExVarRow & " Type: " & RegExVarType)
                    RaiseEvent Message("RegEx variable to append: " & ModifyValuesRegExVarValFrom & " found at row number: " & RegExVarRow & " Type: " & RegExVarType)
                    Exit For
                End If
            Next
            If RegExVarRow = -1 Then
                'RaiseEvent ErrorMessage("Modify values error: Append RevEx variable value: The variable could not be found: " & ModifyValuesRegExVarToAppend)
                RaiseEvent ErrorMessage("Modify values error: Append RevEx variable value: The variable could not be found: " & ModifyValuesRegExVarValFrom)
                Exit Sub
            End If
        End If

        ''Find matching RegEx Variables in the Database Destinations grid:
        'Find matching RegEx Variables in the Data Destinations grid:
        For I = 1 To DbDestCount
            If DbDest(I - 1).RegExVariable = "" Then 'No RegEx variable name is specified.
            Else
                strVarName = DbDest(I - 1).RegExVariable 'The RegEx variable name in the current grid row.
                If strVarName = ModifyValuesRegExVariable Then 'The RegExVariable at the current row matches the required variable to modify
                    For K = 0 To DbDestValues.GetUpperBound(1)
                        Dim OutputString As String
                        Dim InputString As String
                        InputString = DbDestValues(I - 1, K)
                        'TO DO: UPDATE CODE TO INCLUDE OTHER DATA MODIFICATIONS !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        'ModifyValuesType = prop 'Convert_date or Replace_characters
                        'Debug.Print("ModifyValuesType = " & ModifyValuesType & vbCrLf)
                        If ModifyValuesType = ModifyValuesTypes.ClearValue Then
                            OutputString = ""
                            'If ModifyValuesType = "Convert_date" Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.ConvertDate Then
                            If IsNothing(DbDestValues(I - 1, K)) Then
                                'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
                                'ShowMessage("Convert_date error: Text to modify = Nothing: " & vbCrLf, Color.Red)
                                RaiseEvent ErrorMessage("Convert_date error: Text to modify = Nothing")
                            Else
                                ConvertDate(ModifyValuesInputDateFormat, ModifyValuesOutputDateFormat, InputString, OutputString)
                            End If

                            'ElseIf ModifyValuesType = "Replace_characters" Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.ReplaceChars Then
                            If IsNothing(DbDestValues(I - 1, K)) Then
                                'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
                                'ShowMessage("Replace_characters error: Text to modify = Nothing: " & vbCrLf, Color.Red)
                                RaiseEvent ErrorMessage("Replace_characters error: Text to modify = Nothing")
                            Else
                                OutputString = InputString
                                RaiseEvent Message("Replacing characters: " & ModifyValuesCharsToReplace & " with: " & ModifyValuesReplacementChars)
                                RaiseEvent Message("Input string: " & InputString)
                                'OutputString.Replace(ModifyValuesCharsToReplace, ModifyValuesReplacementChars)
                                OutputString = OutputString.Replace(ModifyValuesCharsToReplace, ModifyValuesReplacementChars)

                                RaiseEvent Message("Output string: " & OutputString)
                            End If
                            'ElseIf ModifyValuesType = "Fixed_value" Then
                            'ElseIf ModifyValuesType = ModifyValuesTypes.FixedValue Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendFixedValue Then
                            'OutputString = ModifyValuesFixedValue
                            OutputString = InputString & ModifyValuesFixedValue
                            'ElseIf ModifyValuesType = "Text_file_name" Then
                            'ElseIf ModifyValuesType = ModifyValuesTypes.FileName Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendRegExVarValue Then
                            'RegExVarRow is the row corresponding to the specified RegEx variable.
                            'OutputString = InputString & DbDestValues(I - 1, K)
                            If RegExVarType = "Single value" Then
                                OutputString = InputString & DbDestValues(RegExVarRow, 0)
                            Else
                                OutputString = InputString & DbDestValues(RegExVarRow, K)
                            End If
                            RaiseEvent Message("RegEx variable value appended: " & OutputString)
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendFileName Then
                            'OutputString = System.IO.Path.GetFileName(CurrentFilePath)
                            OutputString = InputString & System.IO.Path.GetFileName(CurrentFilePath)
                            'ElseIf ModifyValuesType = "Text_file_directory" Then
                            'ElseIf ModifyValuesType = ModifyValuesTypes.FileDir Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendFileDir Then
                            'OutputString = System.IO.Path.GetDirectoryName(CurrentFilePath)
                            OutputString = InputString & System.IO.Path.GetDirectoryName(CurrentFilePath)
                            'ElseIf ModifyValuesType = "Text_file_path" Then
                            'ElseIf ModifyValuesType = ModifyValuesTypes.FilePath Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendFilePath Then
                            'OutputString = CurrentFilePath
                            OutputString = InputString & CurrentFilePath
                            'ElseIf ModifyValuesType = "Current_date" Then
                            'ElseIf ModifyValuesType = ModifyValuesTypes.CurrentDate Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendCurrentDate Then
                            'OutputString = Format(Now, "d-MMM-yyyy")
                            OutputString = InputString & Format(Now, "d-MMM-yyyy")
                            'ElseIf ModifyValuesType = "Current_time" Then
                            'ElseIf ModifyValuesType = ModifyValuesTypes.CurrentTime Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendCurrentTime Then
                            'OutputString = Format(Now, "H:mm:ss")
                            OutputString = InputString & Format(Now, "H:mm:ss")
                            'ElseIf ModifyValuesType = "Current_date_time" Then
                            'ElseIf ModifyValuesType = ModifyValuesTypes.CurrentDateTime Then
                        ElseIf ModifyValuesType = ModifyValuesTypes.AppendCurrentDateTime Then
                            'OutputString = Format(Now, "d-MMM-yyyy H:mm:ss")
                            OutputString = InputString & Format(Now, "d-MMM-yyyy H:mm:ss")
                        Else
                            'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
                            'ShowMessage("Unrecognised Modify Values type: " & ModifyValuesType & vbCrLf, Color.Red)
                            RaiseEvent ErrorMessage("Unrecognised Modify Values type: " & ModifyValuesType)
                        End If
                        DbDestValues(I - 1, K) = OutputString
                    Next
                End If
            End If
        Next
    End Sub

    Public Sub TestModifyValuesApply(ByVal InputString As String, ByRef OutputString As String)
        'Test the modification on the specified InputString. The result is put in OutputString.
        'This method uses these properties:
        '    ModifyValuesType (Convert_date, Replace_characters, Fixed_value, Text_file_name, Text_file_directory, Text_file_path, Current_date, Current_time, Current_date_time)

        'Debug.Print("ModifyValuesType = " & ModifyValuesType & vbCrLf)

        If ModifyValuesType = "Convert_date" Then
            If IsNothing(InputString) Then
                'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
                'ShowMessage("Convert_date error: Text to modify = Nothing: " & vbCrLf, Color.Red)
            Else
                ConvertDate(ModifyValuesInputDateFormat, ModifyValuesOutputDateFormat, InputString, OutputString)
            End If

        ElseIf ModifyValuesType = "Replace_characters" Then
            Debug.Print("InputString = " & InputString & vbCrLf)
            If IsNothing(InputString) Then
                'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
                'ShowMessage("Replace_characters error: Text to modify = Nothing: " & vbCrLf, Color.Red)
                OutputString = ""
            Else
                OutputString = InputString.Replace(ModifyValuesCharsToReplace, ModifyValuesReplacementChars)
            End If
            'Debug.Print("ModifyValuesCharsToReplace = " & ModifyValuesCharsToReplace & vbCrLf)
            'Debug.Print("ModifyValuesReplacementChars = " & ModifyValuesReplacementChars & vbCrLf)
            'Debug.Print("OutputString = " & OutputString & vbCrLf)
        ElseIf ModifyValuesType = "Fixed_value" Then
            OutputString = ModifyValuesFixedValue
        ElseIf ModifyValuesType = "Text_file_name" Then
            If CurrentFilePath = "" Then
                OutputString = ""
            Else
                OutputString = System.IO.Path.GetFileName(CurrentFilePath)
            End If
        ElseIf ModifyValuesType = "Text_file_directory" Then
            If CurrentFilePath = "" Then
                OutputString = ""
            Else
                OutputString = System.IO.Path.GetDirectoryName(CurrentFilePath)
            End If
        ElseIf ModifyValuesType = "Text_file_path" Then
            OutputString = CurrentFilePath
        ElseIf ModifyValuesType = "Current_date" Then
            OutputString = Format(Now, "d-MMM-yyyy")
        ElseIf ModifyValuesType = "Current_time" Then
            OutputString = Format(Now, "H:mm:ss")
        ElseIf ModifyValuesType = "Current_date_time" Then
            OutputString = Format(Now, "d-MMM-yyyy H:mm:ss")
        Else
            'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
            'ShowMessage("Unrecognised Modify Values type: " & ModifyValuesType & vbCrLf, Color.Red)
        End If

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
            'NOTE: Convert this code to a warning message event!!! -----------------------------------------------------------------------------------
            Debug.Print("Convert Date Error:" & ex.Message)
        End Try
    End Sub

#End Region 'Modify Values Methods

#Region " Multipliers Methods"

    Public Sub MultipliersAppend(ByVal DbMult As strucMultiplier)
        'Append a Multiplier record to the end of the Multipliers array:

        Dim Count As Integer

        If IsNothing(mMultiplierCodes) Then 'mMultiplierCodes() contains no elements
            ReDim mMultiplierCodes(0 To 0)
            mMultiplierCodes(0) = DbMult
        Else
            Count = mMultiplierCodes.Count
            ReDim Preserve mMultiplierCodes(0 To Count)
            mMultiplierCodes(Count) = DbMult 'zero based array index
        End If
    End Sub

    Public Sub MultipliersClear()
        'Clears the contents of the Multipliers array
        mMultiplierCodes = Nothing
    End Sub

    Public Sub MultiplierValue(ByVal RegExVariable As String, ByVal MultiplierCode As String, ByRef MultiplierValue As Single, ByRef NoError As Boolean)
        'Returns the Multiplier Value corresponding to the RegExVariable and Multiplier Code.
        'If the Value is found without any errors then NoError = True.

        Dim I As Integer

        NoError = False

        If IsNothing(mMultiplierCodes) Then

        Else
            For I = 0 To mMultiplierCodes.Count - 1
                If (mMultiplierCodes(I).RegExMultiplierVariable = RegExVariable) And (mMultiplierCodes(I).MultiplierCode = MultiplierCode) Then
                    MultiplierValue = mMultiplierCodes(I).MultiplierValue
                    NoError = True
                    Exit For
                End If
            Next
        End If

    End Sub

#End Region 'Multipliers Methods

#Region " Read Text Methods"

    Public Sub SelectFirstFile()
        'Selects the first file in the list of input files.
        'The first file path is placed in the TextFilePath property and the file is opened.

        If SelectedFileCount > 0 Then 'There is at least one file in the input file list.
            mSelectedFileNumber = 0
            CurrentFilePath = SelectedFiles(0)
            OpenTextFile()
            'RaiseEvent Notice(vbCrLf & "First file opened: " & CurrentFilePath & vbCrLf & "Lines read: ")
            'RaiseEvent Message(vbCrLf & "First file opened: " & CurrentFilePath & vbCrLf & "Lines read: " & vbCrLf)
            RaiseEvent Message(vbCrLf & "First file opened: " & CurrentFilePath & vbCrLf & "Lines read: ")
            GoToStartOfText()

            ReadLineCount = 0 'Reset ReadLineCount.

            'ImportStatus.Remove("No_more_input_files")
            'ImportStatus.Remove("At_end_of_file")
            ImportStatusRemove("No_more_input_files")
            ImportStatusRemove("At_end_of_file")
        Else 'There are no files in the input file list.
            ImportStatus.Add("No_more_input_files")
        End If
    End Sub

    Public Sub SelectNextFile()
        'Selects the next file in the list of input files.
        'The next file path is placed in the TextFilePath property and the file is opened.

        'If SelectedFileNumber < SelectedFileCount Then 'There is at least one more input file in the list.
        If SelectedFileNumber < SelectedFileCount - 1 Then 'There is at least one more input file in the list.
            mSelectedFileNumber = mSelectedFileNumber + 1
            CurrentFilePath = SelectedFiles(mSelectedFileNumber)

            OpenTextFile()
            'RaiseEvent Notice(vbCrLf & vbCrLf & "New file opened: " & CurrentFilePath & vbCrLf & "Lines read: ")
            'RaiseEvent Message(vbCrLf & vbCrLf & "New file opened: " & CurrentFilePath & vbCrLf & "Lines read: " & vbCrLf)
            RaiseEvent Message(vbCrLf & "New file opened: " & CurrentFilePath & vbCrLf & "Lines read: ")
            GoToStartOfText()

            ReadLineCount = 0 'Reset ReadLineCount.

            'ImportStatus.Remove("No_more_input_files")
            'ImportStatus.Remove("At_end_of_file")
            ImportStatusRemove("No_more_input_files")
            'ImportStatusRemove("No_more_text_files")
            ImportStatusRemove("At_end_of_file")
        Else
            ImportStatus.Add("No_more_input_files")
            'ImportStatus.Add("No_more_text_files")
        End If

    End Sub

    Public Sub OpenTextFile()
        'Opens the selected taxt file.
        'The selected text file path is found in the TextFilePath property.

        FileOpen = False
        If Trim(CurrentFilePath) = "" Then
            'No test file has been selected
        Else
            If System.IO.File.Exists(CurrentFilePath) Then
                If IsNothing(textIn) = False Then
                    textIn = Nothing
                End If
                Try
                    textIn = New System.IO.StreamReader(New System.IO.FileStream(CurrentFilePath, IO.FileMode.Open, System.IO.FileAccess.Read))
                    FileOpen = True
                    'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage("Opened text file: " & TextFilePath & vbCrLf, Color.Black)
                Catch ex As Exception
                    'ShowMessage(vbCrLf, Color.Red)
                    'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage("Error opening text file" & TextFilePath, Color.Red)
                    'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage(ex.Message & vbCrLf & vbCrLf, Color.Blue)
                End Try
            Else
                'Selected test file does not exist
                'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                'ShowMessage("The selected text file does not exist:" & TextFilePath & vbCrLf, Color.Black)
                'SequenceStatus = "At_end_of_file"
                ImportStatus.Add("At_end_of_file")
            End If
        End If
    End Sub

    Public Sub GoToStartOfText()
        If FileOpen = True Then
            TextStore = ""
            textIn.DiscardBufferedData()
            textIn.BaseStream.Position = 0
        End If
    End Sub

    Public Sub ReadAllText()
        'Read the entire test file:
        If Trim(CurrentFilePath) = "" Then
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("No text file has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If System.IO.File.Exists(CurrentFilePath) Then
            'rtbFile.LoadFile(TestFile, RichTextBoxStreamType.PlainText)
            'TextStore = My.Computer.FileSystem.ReadAllText(TextFilePath)
            'OpenTextFile()
            If IsNothing(textIn) Then
                'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                'ShowMessage("The text file could not be opened." & vbCrLf, Color.Red)
                'Exit Sub
            Else
                GoToStartOfText()
                TextStore = textIn.ReadToEnd
            End If
        Else
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("Text file path is not valid.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Public Sub ReadClipboard()
        'Read the text in the clipboard
        TextStore = My.Computer.Clipboard.GetText
    End Sub

    Public Sub ReadNextLineOfText()
        'Read the next line in the text file:

        If Trim(CurrentFilePath) = "" Then
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("No text file has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If System.IO.File.Exists(CurrentFilePath) Then
            If FileOpen = True Then
                If textIn.EndOfStream = False Then
                    If IsNothing(textIn) Then
                        'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                        'ShowMessage("The text file could not be opened." & vbCrLf, Color.Red)
                        'Exit Sub
                    Else
                        TextStore = textIn.ReadLine
                        ReadLineCount = ReadLineCount + 1 'Increment ReadLineCount
                        If textIn.EndOfStream = True Then
                            'SequenceStatus = "At_end_of_file"
                            ImportStatus.Add("At_end_of_file")
                            'Debug.Print("End of file reached")
                            'The last line has been read and the end of line has been flagged!
                        End If
                    End If
                Else
                    'SequenceStatus = "At_end_of_file"
                    ImportStatus.Add("At_end_of_file")
                    'Debug.Print("End of file reached")
                End If

            Else
                'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
                'MessageBox.Show("Selected text file is not open.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("Text file path is not valid.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Public Sub ReadNLinesOfText(ByVal NLines As Integer)
        'Read the next NLines in the text file:

        If Trim(CurrentFilePath) = "" Then
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("No text file has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If System.IO.File.Exists(CurrentFilePath) Then
            If FileOpen = True Then
                If NLines > 0 Then
                    Dim I As Integer
                    Dim StrRead As New System.Text.StringBuilder
                    StrRead.Append(textIn.ReadLine & vbCrLf)
                    For I = 2 To NLines
                        StrRead.Append(textIn.ReadLine & vbCrLf)
                    Next
                    TextStore = StrRead.ToString
                Else
                    TextStore = ""
                End If

            Else

            End If
        Else
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("Text file path is not valid.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    Public Sub SkipNLinesOfText(ByVal NLines As Integer)
        'Skip the next NLines in the text file:

        If Trim(CurrentFilePath) = "" Then
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("No text file has been specified.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If System.IO.File.Exists(CurrentFilePath) Then
            If FileOpen = True Then
                If NLines > 0 Then
                    Dim I As Integer
                    'Dim StrRead As New System.Text.StringBuilder
                    'StrRead.Append(textIn.ReadLine & vbCrLf)
                    textIn.ReadLine()
                    For I = 2 To NLines
                        'StrRead.Append(textIn.ReadLine & vbCrLf)
                        textIn.ReadLine()
                    Next
                    'TextStore = StrRead.ToString
                Else
                    'TextStore = ""
                End If

            Else

            End If
        Else
            'NOTE: Convert this code to a warning message event!!! ---------------------------------------------------------------------------------
            'MessageBox.Show("No test file path is not valid.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Public Sub ReadTextToString(ByVal StrMatch As String)
        'Read to the specified string 

        Dim StrRead As New System.Text.StringBuilder
        Dim MatchPosn As Integer
        Dim CharVal As Integer

        StrRead.Clear()
        MatchPosn = 0

        Do Until ((MatchPosn <> 0) Or (textIn.Peek = -1))
            'Debug.Print("<" & textIn.Peek & ">")
            CharVal = textIn.Read
            If (CharVal >= 0) And (CharVal <= 127) Then
                StrRead.Append(Chr(CharVal))
                MatchPosn = InStr(StrRead.ToString, StrMatch)
            End If
        Loop

        If MatchPosn = 0 Then
            TextStore = StrRead.ToString
        Else
            StrRead.Remove(MatchPosn - 1, Len(StrMatch)) 'Remove the matched string from the end of StrRead
            TextStore = StrRead.ToString
        End If

    End Sub

    Public Sub SkipTextPastString(ByVal StrMatch As String)
        'Skip past the specified string 

        Dim StrRead As New System.Text.StringBuilder
        Dim MatchPosn As Integer
        Dim CharVal As Integer

        StrRead.Clear()
        MatchPosn = 0

        Do Until ((MatchPosn <> 0) Or (textIn.Peek = -1))
            'Debug.Print("<" & textIn.Peek & ">")
            CharVal = textIn.Read
            If (CharVal >= 0) And (CharVal <= 127) Then
                StrRead.Append(Chr(CharVal))
                MatchPosn = InStr(StrRead.ToString, StrMatch)
            End If
        Loop
    End Sub

#End Region 'Read Text Methods

#Region " Write to Database Methods"

    Public Sub OpenDatabase()
        'Open the database.

        If ConnectionString = "" Then

        Else
            conn = New System.Data.OleDb.OleDbConnection(ConnectionString)
            conn.Open()
            da.InsertCommand = cmd
            da.InsertCommand.Connection = conn
        End If


    End Sub

    Public Sub CloseDatabase()
        'Close the database.

        da.InsertCommand.Connection.Close()

    End Sub

    Public Sub SetMultipliers()
        'Process the Multiplier rows in the RegEx variable destination grid 
        '  so that the corresponding multiplier cells in the GridProp(,) properties array contain the correct multiplier values.
        'Debug.Print("Running SetMultipliers" & vbCrLf)

        Dim NRows As Integer
        Dim NCols As Integer
        Dim RowNo As Integer
        Dim ValNo As Integer
        Dim DestTable As String
        Dim DestField As String
        Dim I As Integer
        Dim J As Integer
        Dim NextVarType As String
        Dim NextDestTable As String
        Dim NextDestField As String
        Dim MultCode As String
        Dim MultValue As Single
        Dim RegExVar As String
        Dim Valid As Boolean

        NRows = DbDestCount

        NCols = DbDestValues.GetUpperBound(1) + 1

        ReDim GridProp(0 To NRows - 1, 0 To NCols - 1) 'Resize the GridProp array to match the size of DbDestValues().

        'GridProp will contain Mult and Status cells to hold corresponding Multiplier and Status values.
        'Multiplier: Some tables include a multiplier code for financial values. The multiplier code is converted to a value that is multiplied with the values.
        'Status: Some financial values may be missing or have uncetain values. This information is stored in a corresponding Status field.

        'Initialise GridProp array:
        For I = 0 To NRows - 1
            'If Trim(DbDestinations.DataGridView1.Rows(I).Cells(1).Value.ToString) = "Single Multiplier" Then 'Set all multiplier values to zero
            If DbDest(I).Type = "Single Multiplier" Then 'Set all multiplier values to zero
                'Debug.Print("DbDest(" & I & ").Type = Single Multiplier" & vbCrLf)
                For J = 0 To NCols - 1
                    GridProp(I, J).Mult = 0
                Next
                'ElseIf Trim(DbDestinations.DataGridView1.Rows(I).Cells(1).Value.ToString) = "Multiple Multiplier" Then 'Set all multiplier values to zero
            ElseIf DbDest(I).Type = "Multiple Multiplier" Then 'Set all multiplier values to zero
                'Debug.Print("DbDest(" & I & ").Type = Multiple Multiplier" & vbCrLf)
                For J = 0 To NCols - 1
                    GridProp(I, J).Mult = 0
                Next
            Else 'Value row: Set all multiplier values to default value of one:
                'Debug.Print("DbDest(" & I & ").Type = Value" & vbCrLf)
                For J = 0 To NCols - 1
                    GridProp(I, J).Mult = 1
                    'Debug.Print("GridProp(" & I & "," & J & ").Mult = " & GridProp(I, J).Mult)
                Next
            End If
        Next

        'Check for duplicate definitions in DataGridView1:
        If CheckForDuplicates() = True Then
            'MsgBox("Duplicate entry in Database Destinations!", MsgBoxStyle.Exclamation, "Notice")
            'RaiseEvent ErrorMessage("Duplicate entry in Database Destinations!")
            'Display the contents of the GridProp array in the Debug window:
            ShowGridProp()
            PrintDbDestData()
            MsgBox("Duplicate entry in Database Destinations!", MsgBoxStyle.Exclamation, "Notice")
            Exit Sub
        Else 'No duplicates
            'OK to continue
        End If

        'Find each of the rows with variable type: Single Multiplier or Multiple Multiplier.
        For RowNo = 0 To NRows - 1
            'Scan each row in DataGridView1 for Multiplier codes.
            If DbDest(RowNo).Type = "Single Multiplier" Then 'The current row contains a multiplier code
                'Multiplier codes are usually placed in the first value cell. This code applies to all value cells in the corresponding Destination Field.
                'The corresponding field values can be determined from the Destination Table and Destination Field columns.

                RegExVar = DbDest(RowNo).RegExVariable
                MultCode = DbDestValues(RowNo, 0)
                'Find Multiplier Value:
                MultiplierValue(RegExVar, MultCode, MultValue, Valid)

                If Valid = False Then
                    'MsgBox("Multiplier Value not found for RegEx Variable: " & RegExVar & " and Code: " & MultCode, MsgBoxStyle.Exclamation, "Notice")
                    RaiseEvent ErrorMessage("Multiplier Value not found for RegEx Variable: " & RegExVar & " and Code: " & MultCode)
                    Exit Sub
                End If

                'Find the Row containing the corresponding Destination Field values:
                DestTable = DbDest(RowNo).TableName
                DestField = DbDest(RowNo).FieldName
                For I = 0 To NRows - 1 'Search all rows in DataGridView1
                    If I <> RowNo Then 'Not at original Multiplier row
                        NextVarType = DbDest(I).Type
                        If (NextVarType <> "Single Multiplier") And (NextVarType <> "Multiple Multiplier") Then 'This is a value row
                            NextDestTable = DbDest(I).TableName
                            NextDestField = DbDest(I).FieldName
                            If (NextDestTable = DestTable) And (NextDestField = DestField) Then 'This value row corresponds to the current Multiplier
                                'Place multiplier values in the corresponding cells in the GridProparray:
                                For J = 0 To NCols - 1
                                    GridProp(I, J).Mult = MultValue
                                Next
                            End If
                        End If
                    End If
                Next
            ElseIf DbDest(RowNo).Type = "Multiple Multiplier" Then 'The current row contains a multiplier code

                RegExVar = DbDest(RowNo).RegExVariable
                Dim MultVals(0 To NCols - 1) As Single

                'Process each Multiplier Code:
                For ValNo = 1 To NCols
                    MultCode = DbDestValues(RowNo, ValNo)
                    'Find Multiplier Value:
                    MultiplierValue(RegExVar, MultCode, MultValue, Valid)

                    If Valid = False Then
                        'MsgBox("Multiplier Value not found for RegEx Variable: " & RegExVar & " and Code: " & MultCode, MsgBoxStyle.Exclamation, "Notice")
                        RaiseEvent ErrorMessage("Multiplier Value not found for RegEx Variable: " & RegExVar & " and Code: " & MultCode)
                        Exit Sub
                    End If

                    MultVals(ValNo - 1) = MultValue
                Next

                'Find the Row (in DataGridView1) containing the corresponding Destination Field values:
                DestTable = DbDest(RowNo).TableName
                DestField = DbDest(RowNo).FieldName
                For I = 0 To NRows - 1 'Search all rows in DataGridView1
                    If I <> RowNo Then 'Not at original Multiplier row
                        NextVarType = DbDest(I).Type
                        If (NextVarType <> "Single Multiplier") And (NextVarType <> "Multiple Multiplier") Then 'This is a value row
                            NextDestTable = DbDest(I).TableName
                            NextDestField = DbDest(I).FieldName
                            If (NextDestTable = DestTable) And (NextDestField = DestField) Then 'This value row corresponds to the current Multiplier
                                'Place multiplier values in the corresponding cells in the GridProparray:
                                For J = 0 To NCols - 1
                                    GridProp(I, J).Mult = MultVals(J)
                                Next
                            End If
                        End If
                    End If
                Next

            End If
        Next


        'FOR DEBUGGING: --------------------------------------------------------
        'Display the contents of the GridProp array in the Debug window:
        'Debug.Print("GridProp(,)")
        'Debug.WriteLine(String.Format("{0,-8}{1,-13}{2,-10}{3,-10}", "Row No", "Multiplier", "Status", "..."))
        'For RowNo = 0 To GridProp.GetUpperBound(0)
        '    Debug.Write(String.Format("{0,-8}", RowNo))
        '    For ValueNo = 0 To GridProp.GetUpperBound(1)
        '        Debug.Write(String.Format("{0,-10}{1,-4}", GridProp(RowNo, ValueNo).Mult, GridProp(RowNo, ValueNo).Status))
        '    Next
        '    Debug.WriteLine("")
        'Next
        'Debug.Print("")
        ' -----------------------------------------------------------------------

    End Sub

    Private Sub ShowGridProp()
        'For debugging: Show the GridProp array:
        Debug.Print("GridProp(,)")
        Debug.WriteLine(String.Format("{0,-8}{1,-13}{2,-10}{3,-10}", "Row No", "Multiplier", "Status", "..."))
        For RowNo = 0 To GridProp.GetUpperBound(0)
            Debug.Write(String.Format("{0,-8}", RowNo))
            For ValueNo = 0 To GridProp.GetUpperBound(1)
                Debug.Write(String.Format("{0,-10}{1,-4}", GridProp(RowNo, ValueNo).Mult, GridProp(RowNo, ValueNo).Status))
            Next
            Debug.WriteLine("")
        Next
        Debug.Print("")
    End Sub

    Private Sub PrintDbDestData()
        'For debugging: Print DbDest data

        Dim Rows As Integer = DbDestCount
        'Dim Cols As Integer
        Dim I As Integer

        For I = 1 To Rows
            Debug.Print("Contents of DbDest() --------------------------------------")
            Debug.Print("Row: " & I)
            Debug.Print("RegExVariable: " & DbDest(I - 1).RegExVariable)
            Debug.Print("Type: " & DbDest(I - 1).Type)
            Debug.Print("Table name: " & DbDest(I - 1).TableName)
            Debug.Print("Field name: " & DbDest(I - 1).FieldName)
            Debug.Print("Status field: " & DbDest(I - 1).StatusField)
        Next
    End Sub

    Public Function ReturnDbDestData() As String
        'Return a string containing the contents of DbDest()
        Dim sb As New System.Text.StringBuilder

        sb.Append("------------ DbDest() ---------------" & vbCrLf)
        Dim Rows As Integer = DbDestCount
        Dim I As Integer

        For I = 1 To Rows
            sb.Append("Row: " & I & vbCrLf)
            sb.Append("RegExVariable: " & DbDest(I - 1).RegExVariable & vbCrLf)
            sb.Append("Type: " & DbDest(I - 1).Type & vbCrLf)
            sb.Append("Table name: " & DbDest(I - 1).TableName & vbCrLf)
            sb.Append("Field name: " & DbDest(I - 1).FieldName & vbCrLf)
            sb.Append("Status field: " & DbDest(I - 1).StatusField & vbCrLf)
        Next

        sb.Append("------------------------------------" & vbCrLf)

        Return sb.ToString

    End Function

    Private Function CheckForDuplicates() As Boolean
        'Checks for duplicate entries in DataGridView1

        Dim I As Integer 'Loop index
        Dim MaxRow As Integer
        Dim Result As Boolean

        'MaxRow = DbDestinations.DataGridView1.Rows.Count
        MaxRow = DbDestCount
        'If DbDestinations.DataGridView1.Rows(MaxRow - 1).IsNewRow = True Then
        '    MaxRow = MaxRow - 1
        'End If
        Result = False

        For I = 0 To MaxRow - 1
            If CheckForDuplicatesAt(I) = True Then
                Result = True
                Exit For
            End If
        Next

        Return Result

    End Function

    Private Function CheckForDuplicatesAt(ByVal Row As Integer) As Boolean
        'Checks for duplicate entries at location Row in DbDest()
        '    There should be only 1 Value row for a Table/Field destination
        '    There should be only 0 or 1 Multiplier for each Table/Field data destination
        '    Regular Expression (RegEx) variables can be repeated:
        '        A Value can be written to more than one table
        '        A Multiplier can be used for more than one Field

        Dim MaxRow As Integer
        Dim I As Integer 'Loop index
        Dim RegExVar As String
        Dim VarType As String
        Dim DestTable As String
        Dim DestField As String
        Dim ValueCount As Integer 'Count of value rows corresponding to the RegExVar
        Dim MultCount As Integer 'Count of the multiplier rows corresponding to the RegExVar
        Dim NextRegExVar As String
        Dim NextVarType As String
        Dim NextDestTable As String
        Dim NextDestField As String

        MaxRow = DbDestCount

        RegExVar = DbDest(Row).RegExVariable
        VarType = DbDest(Row).Type
        DestTable = DbDest(Row).TableName
        DestField = DbDest(Row).FieldName


        If VarType = "Single Multiplier" Then
            MultCount = 1
            ValueCount = 0
        Else
            MultCount = 0
            ValueCount = 1
        End If

        For I = Row + 1 To MaxRow - 1
            NextRegExVar = DbDest(I).RegExVariable
            NextVarType = DbDest(I).Type
            NextDestTable = DbDest(I).TableName
            NextDestField = DbDest(I).FieldName

            'Multiple RegEx variable names are possible, but not pointing to the same Table/Field.

            'Look for duplicate Table/Field destinations:
            'Look for duplicate Multipliers applied to the same Table/Field
            'Look for duplicate Variables applied to the same Table/Field
            If (DestTable = NextDestTable) And (DestField = NextDestField) Then
                If (NextVarType = "Single Multiplier") Or (NextVarType = "Multiple Multiplier") Then
                    MultCount = MultCount + 1
                    If MultCount > 1 Then
                        'MsgBox("Two or more Multipliers for Field: " & DestField, MsgBoxStyle.Exclamation, "Notice")
                        RaiseEvent ErrorMessage("Two or more Multipliers for Field: " & DestField)
                        PrintDbDestData()
                    End If
                Else
                    ValueCount = ValueCount + 1
                    If ValueCount > 1 Then
                        'MsgBox("Two or more Value Rows for Field: " & DestField, MsgBoxStyle.Exclamation, "Notice")
                        RaiseEvent ErrorMessage("Two or more Value Rows for Field: " & DestField)
                        PrintDbDestData()
                    End If
                End If
            End If
        Next

        If ValueCount > 1 Then 'More than one value type corresponds to the Table/Field destination at location Row
            Return True
        Else
            If MultCount > 1 Then 'More then one Multiplier type corresponds to the Table/Field destination at location Row
                Return True
            Else
                Return False
            End If
        End If

    End Function



    Public Sub ConvertYYYYMMDDToDate(ByVal RegExVar As String)
        'Converts variables matched with the specified RegEx (RegExVar) to a date string
        'The variables have the format YYYYMMDD
        'where YYYY is a four digit date eg 2010
        '      MM is a two digit month eg 09
        '      DD is a two digit day eg 30

        'To Do ......................

    End Sub

    Public Sub GetTableList()
        'Gets a list of destination Tables from DbDest() (Previously DataGridView1)
        '    The list is stored in TableList()

        'DataGridView1 contains the raw data extracted from the text file using the Regular Expressions
        'Data for each table is saved in the TableList array.
        'TableList() Contains the following fields:
        '    TableName
        '    MaxNValues - the maximum number of values in a Multiple Value field to be written to the table UPDATE REQUIRED: check for Multiple Value ************
        '    MinNValues - the minimum number of values in a Multiple Value fields to be written to the table UPDATE REQUIRED: check for Multiple Value ***********
        '    HasGaps    - True is a Multiple Value field has one or more gaps (spaces) between any values UPDATE REQUIRED: check for Multiple Value **************

        Dim RowNo As Integer
        Dim LastRow As Integer
        Dim LastCol As Integer
        Dim NewTableName As Boolean
        Dim NextTableName As String
        Dim I As Integer
        Dim NValues As Integer
        Dim BlankValue As Boolean
        Dim ValueGaps As Boolean
        Dim CurTableName As String

        'Get the first Table Name:
        ReDim TableList(0 To 0)
        'TableList(0).TableName = Trim(DbDestinations.DataGridView1.Rows(0).Cells(2).Value.ToString)
        TableList(0).TableName = DbDest(0).TableName

        'Get the remaining Table Names:
        'LastRow = DbDestinations.DataGridView1.Rows.Count
        LastRow = DbDestCount
        'If DbDestinations.DataGridView1.Rows(LastRow - 1).IsNewRow = True Then
        '    LastRow = LastRow - 1
        'End If
        For RowNo = 1 To LastRow - 1
            NewTableName = True
            'NextTableName = Trim(DbDestinations.DataGridView1.Rows(RowNo).Cells(2).Value.ToString)
            NextTableName = DbDest(RowNo).TableName
            For I = 0 To TableList.Count - 1
                If NextTableName = TableList(I).TableName Then 'NextTableName is already in the list
                    NewTableName = False
                    Exit For
                End If
            Next
            If NewTableName = True Then 'Add table name to the list
                ReDim Preserve TableList(0 To TableList.Count)
                TableList(TableList.Count - 1).TableName = NextTableName
            End If
        Next

        'Find the Min and Max number of values to be written to each table:

        'LastCol = DbDestinations.DataGridView1.ColumnCount
        LastCol = DbDestValues.GetUpperBound(1)

        'Initialise HasGaps MinNValues and MaxNValues in TableList()
        For I = 0 To TableList.Count - 1
            TableList(I).HasGaps = False
            'TableList(I).MinNValues = LastCol - 5
            TableList(I).MinNValues = LastCol
            TableList(I).MaxNValues = 0
        Next

        'Process each row in DataGridView1
        For RowNo = 0 To LastRow - 1
            'CurTableName = Trim(DbDestinations.DataGridView1.Rows(RowNo).Cells(2).Value.ToString) 'Current table name
            CurTableName = DbDest(RowNo).TableName 'Current table name
            'If Trim(DbDestinations.DataGridView1.Rows(RowNo).Cells(1).Value.ToString) = "Multiple Value" Then 'Field with multiple values to be written
            If DbDest(RowNo).Type = "Multiple Value" Then 'Field with multiple values to be written
                NValues = 0
                BlankValue = False 'Initialise value
                ValueGaps = False
                'For I = 5 To LastCol - 1
                'For I = 0 To LastCol - 1
                For I = 0 To LastCol  '*** UPDATED 22Dec16 ***
                    'If DbDestinations.DataGridView1.Rows(RowNo).Cells(I).Value = Nothing Then
                    If DbDestValues(RowNo, I) = Nothing Then
                        BlankValue = True
                        'ElseIf Trim(DbDestinations.DataGridView1.Rows(RowNo).Cells(I).Value.ToString) = "" Then 'This column contains a blank value
                    ElseIf DbDestValues(RowNo, I) = "" Then 'This column contains a blank value
                        BlankValue = True
                    Else 'This cell contains a value
                        NValues = NValues + 1
                        If BlankValue = True Then
                            ValueGaps = True
                        End If
                    End If
                Next

                'ElseIf Trim(DbDestinations.DataGridView1.Rows(RowNo).Cells(1).Value.ToString) = "Single Value" Then 'Field with single value to be written
            ElseIf DbDest(RowNo).Type = "Single Value" Then 'Field with single value to be written
                NValues = 1
                ValueGaps = False
            End If
            'Update TableList():
            For I = 0 To TableList.Count - 1
                If CurTableName = TableList(I).TableName Then 'Update matching Table Name record:
                    If NValues < TableList(I).MinNValues Then
                        TableList(I).MinNValues = NValues
                    End If
                    If NValues > TableList(I).MaxNValues Then
                        TableList(I).MaxNValues = NValues
                    End If
                    If ValueGaps = True Then
                        TableList(I).HasGaps = True
                    End If
                Else

                End If
            Next
        Next

        'FOR DEBUGGING: --------------------------------------------------------
        'Print(TableList())
        'Debug.Print("")
        'Debug.Print("TableList()")
        'Debug.WriteLine(String.Format("{0,-32}{1,-14}{2,-14}{3,-10}", "Table Name", "Min N Values", "Max N Values", "Has Gaps"))
        'For RowNo = 0 To TableList.Count - 1
        '    Debug.WriteLine(String.Format("{0,-32}{1,-14}{2,-14}{3,-10}", TableList(RowNo).TableName, TableList(RowNo).MinNValues, TableList(RowNo).MaxNValues, TableList(RowNo).HasGaps))
        'Next
        'Debug.Print("")
        DebugWriteTableList()
        ' -----------------------------------------------------------------------

    End Sub

    Private Sub DebugWriteTableList()
        'For debugging: Write contents of the TableList array to the Messages window:

        'SetMessageColor(Color.Black)
        'SetMessageFontName(FontList.Lucida_Sans_Typewriter)
        'AddMessage("TableList()" & vbCrLf)
        'AddMessage(String.Format("{0,-32}{1,-14}{2,-14}{3,-10}", "Table Name", "Min N Values", "Max N Values", "Has Gaps") & vbCrLf)
        'For RowNo = 0 To TableList.Count - 1
        '    SetMessageColor(Color.Blue)
        '    AddMessage(String.Format("{0,-32}{1,-14}{2,-14}{3,-10}", TableList(RowNo).TableName, TableList(RowNo).MinNValues, TableList(RowNo).MaxNValues, TableList(RowNo).HasGaps) & vbCrLf)
        'Next
        'SetMessageColor(Color.Black)
        'AddMessage(vbCrLf)

    End Sub

    Public Sub GetFieldList()
        'Gets a list of destination Fields from DbDest() (Previously DataGridView1)
        '    The list is stored in FieldList()

        'DataGridView1 contains the raw data extracted from the text file using the Regular Expressions
        'Data for each Field is saved in the FieldList array.
        'FieldList() Contains the following fields:
        '    TableName
        '    FieldName 
        '    FieldType  'Text-Memo-Number-Date/Time-Currency-AutoNumber-Yes/No-Hyperlink Number:Byte-Integer-LongInteger-Single-Double
        '    FieldLen
        '    Status     'Gap-Mis
        '    NValues    'The number of values for this Field to write to the database
        '    Mult       'The multiplier used to scale the Field values in the source document
        '    RowNo      'The Row Number of the Field in DataGridView1

        'NOTE: TableList() must be filled before this subroutine is run!

        Dim RowNo As Integer
        Dim LastRow As Integer
        Dim VarType As String
        Dim FieldNo As Integer
        Dim MultCode As String
        Dim MultVal As Single

        ReDim FieldList(0 To 0)

        LastRow = DbDestCount

        FieldNo = 0

        For RowNo = 0 To LastRow - 1 'Scan all rows in DataGridView1
            VarType = DbDest(RowNo).Type
            If (VarType <> "Single Multiplier") And (VarType <> "Multiple Multiplier") Then 'VarType is a Field to be written to the database
                ReDim Preserve FieldList(0 To FieldNo)
                FieldList(FieldNo).TableName = DbDest(RowNo).TableName
                FieldList(FieldNo).FieldName = DbDest(RowNo).FieldName
                If IsNothing(DbDest(RowNo).StatusField) Then
                    mDbDest(RowNo).StatusField = ""
                End If
                FieldList(FieldNo).StatusFieldName = DbDest(RowNo).StatusField
                FieldList(FieldNo).Status = "" 'Initial value. This value is modified when values are read from DataGridView1 and written to the database.
                If DbDest(RowNo).Type = "Single Value" Then
                    FieldList(FieldNo).MultiValued = False
                Else
                    FieldList(FieldNo).MultiValued = True
                End If
                FieldList(FieldNo).RowNo = RowNo 'The Row Number of the Field in DataGridView1
                FieldNo = FieldNo + 1
            End If
        Next


        'Get the Field Type and Length for each Field entry in FieldList(): --------------------------------------------------------------------------
        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Process through each Table in TableList()
        Dim NTables As Integer
        Dim TableNo As Integer
        Dim TableName As String
        Dim FieldName As String

        NTables = TableList.Count

        For TableNo = 0 To NTables - 1
            TableName = TableList(TableNo).TableName
            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + DatabasePath
            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            Dim restrictions As String() = New String() {Nothing, Nothing, TableName, Nothing} 'Get the Schema just for table TableName
            dt = conn.GetSchema("Columns", restrictions) 'Get Columns Schema

            'DATA_TYPE
            'Empty = 0
            'SmallInt = 2
            'Integer = 3
            'Single = 4
            'Double = 5
            'Currency = 6
            'Date = 7
            'BSTR = 8      A null-terminated character string of Unicode characters (DBTYPE_BSTR).
            'IDispatch = 9 A pointer to an IDispatch interface (DBTYPE_IDISPATCH). 
            'Error = 10    A 32-bit error code (DBTYPE_ERROR). 
            'Boolean = 11
            'Variant = 12
            'IUnknown = 13
            'Decimal = 14
            'TinyInt = 16
            'UnsignedTinyInt = 17
            'UnsignedSmallInt = 18
            'UnsignedInt = 19
            'BigInt = 20
            'UnsignedBigInt = 21
            'Filetime = 64
            'Guid = 72
            'Binary = 128
            'Char = 129    A character string (DBTYPE_STR). 
            'WChar = 130   A null-terminated stream of Unicode characters (DBTYPE_WSTR). 
            'Numeric = 131
            'DBDate = 133
            'DBTime = 134
            'DBTimeStamp = 135
            'PropVariant = 138
            'VarNumeric = 139
            'VarChar = 200
            'LongVarChar = 201
            'VarWChar = 202
            'LongVarWChar = 203
            'VarBinary = 204
            'LongVarBinary = 205

            For FieldNo = 0 To FieldList.Count - 1
                If FieldList(FieldNo).TableName = TableName Then
                    For RowNo = 0 To dt.Rows.Count - 1
                        If dt.Rows(RowNo).Item("COLUMN_NAME").ToString = FieldList(FieldNo).FieldName Then
                            FieldList(FieldNo).FieldType = dt.Rows(RowNo).Item("DATA_TYPE").ToString
                            If FieldList(FieldNo).FieldType = "7" Then '
                                FieldList(FieldNo).FieldLen = 0
                            ElseIf FieldList(FieldNo).FieldType = "3" Then 'Integer
                                FieldList(FieldNo).FieldLen = 0
                            ElseIf FieldList(FieldNo).FieldType = "6" Then 'Currency
                                FieldList(FieldNo).FieldLen = 0
                            ElseIf FieldList(FieldNo).FieldType = "4" Then 'Single
                                FieldList(FieldNo).FieldLen = 0
                            ElseIf FieldList(FieldNo).FieldType = "130" Then 'WChar
                                FieldList(FieldNo).FieldLen = dt.Rows(RowNo).Item("CHARACTER_MAXIMUM_LENGTH")
                            Else
                                If IsDBNull(dt.Rows(RowNo).Item("CHARACTER_MAXIMUM_LENGTH")) Then
                                    FieldList(FieldNo).FieldLen = 0
                                Else
                                    FieldList(FieldNo).FieldLen = dt.Rows(RowNo).Item("CHARACTER_MAXIMUM_LENGTH")
                                End If

                            End If
                        End If
                    Next 'RowNo
                End If
            Next 'FieldNo

            conn.Close()

        Next 'TableNo
        '------------------------------------------------------------------------------------------------------------------------

        'Process Multipliers:-------------------------------------------------------------------------------
        'Initialise multiplier values in FieldList() to 1:
        'For I = 0 To FieldList.Count - 1
        '    FieldList(I).Mult = 1
        'Next

        'Dim FieldName As String
        'Dim NoError As Boolean
        'For RowNo = 0 To LastRow - 1
        '    If Trim(DataGridView1.Rows(RowNo).Cells(1).Value.ToString) = "Single Multiplier" Then 'Single Multiplier code
        '        MultCode = Trim(DataGridView1.Rows(RowNo).Cells(5).Value.ToString)
        '        ImportTextIntoDatabase.MultiplierValue(Trim(DataGridView1.Rows(RowNo).Cells(0).Value.ToString), MultCode, MultVal, NoError)
        '        'Usage: ImportTextIntoDatabase.MultiplierValue(RegExVariable, MultiplierCode, MultiplierVal, NoError)
        '        '       The MultiplierValue is returned along with the NoError flag (= True if MultiplierVal found OK).
        '        If NoError = False Then
        '            MultVal = 0
        '        End If
        '        'Put the Multiplier Value in the FieldList() array
        '        TableName = Trim(DataGridView1.Rows(RowNo).Cells(2).Value.ToString)
        '        FieldName = Trim(DataGridView1.Rows(RowNo).Cells(3).Value.ToString)
        '        For I = 0 To FieldList.Count - 1
        '            If (FieldList(I).TableName = TableName) And (FieldList(I).FieldName = FieldName) Then
        '                FieldList(I).Mult = MultVal
        '                Exit For
        '            End If
        '        Next

        '    ElseIf Trim(DataGridView1.Rows(RowNo).Cells(1).Value.ToString) = "Multiple Multiplier" Then 'Multiple Multiplier code
        '        For I = 0 To FieldList.Count - 1
        '            MultCode = Trim(DataGridView1.Rows(RowNo).Cells(5 + I).Value.ToString)
        '            ImportTextIntoDatabase.MultiplierValue(Trim(DataGridView1.Rows(RowNo).Cells(0).Value.ToString), MultCode, MultVal, NoError)
        '            'Usage: ImportTextIntoDatabase.MultiplierValue(RegExVariable, MultiplierCode, MultiplierVal, NoError)
        '            '       The MultiplierValue is returned along with the NoError flag (= True if MultiplierVal found OK).
        '            If NoError = False Then
        '                MultVal = 0
        '            End If
        '            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '        Next
        '    Else 'Not a (Single or Multiple) Multiplier code

        '    End If
        'Next

        Dim HasGaps As Boolean
        Dim Gap As Boolean
        Dim LastValuePosn As Integer
        Dim ValNo As Integer
        Dim ValStr As String

        'Find the number of values in each Multple Value field:
        For RowNo = 0 To LastRow - 1
            If DbDest(RowNo).Type = "Multiple Value" Then 'Multiple values to be counted
                TableName = DbDest(RowNo).TableName
                FieldName = DbDest(RowNo).FieldName
                HasGaps = False
                Gap = False
                LastValuePosn = 0

                'For ValNo = 1 To DbDestValues.GetUpperBound(1) 'First 5 columns contain non-value data
                For ValNo = 1 To DbDestValues.GetUpperBound(1) + 1 'First 5 columns contain non-value data '*** UPDATED 22Dec16 ***
                    If IsNothing(DbDestValues(RowNo, ValNo - 1)) Then
                        'This code has not been tested: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        Gap = True 'Indicate the current value position is a gap
                    Else
                        ValStr = DbDestValues(RowNo, ValNo - 1)
                        If ValStr = "" Then
                            Gap = True 'Indicate the current value position is a gap
                        Else
                            LastValuePosn = ValNo 'Update the position of the last non-gap value.
                            If Gap = True Then
                                HasGaps = True 'Record that the set of values contains at least one gap
                            End If
                            Gap = False 'Indicates that the current value position is not a gap
                        End If
                    End If
                Next 'ValNo

                'Update the FieldList() entry:
                For I = 0 To FieldList.Count - 1
                    If (TableName = FieldList(I).TableName) And (FieldName = FieldList(I).FieldName) Then
                        FieldList(I).NValues = LastValuePosn
                        If HasGaps = True Then
                            FieldList(I).Status = "Gaps"
                        Else
                            FieldList(I).Status = "NoGaps"
                        End If
                    End If
                Next
            ElseIf DbDest(RowNo).Type = "Single Value" Then 'Single value
                TableName = DbDest(RowNo).TableName
                FieldName = DbDest(RowNo).FieldName
                If IsNothing(DbDestValues(RowNo, 0)) Then
                    ValStr = ""
                Else
                    ValStr = DbDestValues(RowNo, 0)
                End If

                'Update the FieldList() entry:
                For I = 0 To FieldList.Count - 1
                    If (TableName = FieldList(I).TableName) And (FieldName = FieldList(I).FieldName) Then
                        FieldList(I).NValues = 1
                        If ValStr = "" Then
                            FieldList(I).Status = "MIS"
                        Else
                            FieldList(I).Status = "OK"
                        End If
                    End If
                Next
            End If
        Next

        'FOR DEBUGGING: --------------------------------------------------------
        'Display the contents of the FieldList() array in the Debug window
        'Dim tabs As Single() = {64, 32, 8, 8, 8, 8, 8}
        'Dim stringFormat As New StringFormat()
        'stringFormat.SetTabStops(0, tabs)
        'Debug.Print("FieldList()")
        'Format codes are in the curly brackets: {0,-10} 
        'The first number (zero based) in the bracket corresponds to the element number in  a list of values
        'for formatting strings ,-10 the - means left aligned and the 10 means pad to ten chars
        'http://idunno.org/archive/2004/14/01/122.aspx
        'http://www.builderau.com.au/program/dotnet/soa/Easily-format-string-output-with-String-Format/0,339028399,339177160,00.htm

        'Debug.WriteLine(String.Format("{0,-8}{1,-32}{2,-38}{3,-44}{4,-12}{5,-14}{6,-14}{7,-8}{8,-10}{9,-8}", "Row No", "Table Name", "Field Name", "Status Field", "Field Type", "Field Length", "Multi-Valued", "Status", "NValues", "Row No"))
        'For RowNo = 0 To FieldList.Count - 1
        '    Debug.Write(String.Format("{0,-8}", RowNo))
        '    'Debug.WriteLine(String.Format("{0,-32}{1,-24}{2,-30}{3,-12}{4,-14}{5,-14}{6,-12}{7,-8}{8,-10}{9,-8}", FieldList(RowNo).TableName, FieldList(RowNo).FieldName, FieldList(RowNo).StatusFieldName, FieldList(RowNo).FieldType, FieldList(RowNo).FieldLen, FieldList(RowNo).MultiValued, FieldList(RowNo).Mult, FieldList(RowNo).Status, FieldList(RowNo).NValues, FieldList(RowNo).RowNo))
        '    Debug.WriteLine(String.Format("{0,-32}{1,-38}{2,-44}{3,-12}{4,-14}{5,-14}{6,-8}{7,-10}{8,-8}", FieldList(RowNo).TableName, FieldList(RowNo).FieldName, FieldList(RowNo).StatusFieldName, FieldList(RowNo).FieldType, FieldList(RowNo).FieldLen, FieldList(RowNo).MultiValued, FieldList(RowNo).Status, FieldList(RowNo).NValues, FieldList(RowNo).RowNo))
        'Next
        'Debug.Print("")
        DebugWriteFieldList()
        ' -----------------------------------------------------------------------

    End Sub

    Private Sub DebugWriteFieldList()
        'For debugging: Write contents of the FieldList array to the Messages window:

        'SetMessageColor(Color.Black)
        'SetMessageFontName(FontList.Lucida_Sans_Typewriter)
        'AddMessage("FieldList()" & vbCrLf)
        'AddMessage(String.Format("{0,-8}{1,-32}{2,-38}{3,-44}{4,-12}{5,-14}{6,-14}{7,-8}{8,-10}{9,-8}", "Row No", "Table Name", "Field Name", "Status Field", "Field Type", "Field Length", "Multi-Valued", "Status", "NValues", "Row No") & vbCrLf)
        'For RowNo = 0 To FieldList.Count - 1
        '    SetMessageColor(Color.Blue)
        '    AddMessage(String.Format("{0,-8}", RowNo))
        '    AddMessage(String.Format("{0,-32}{1,-38}{2,-44}{3,-12}{4,-14}{5,-14}{6,-8}{7,-10}{8,-8}", FieldList(RowNo).TableName, FieldList(RowNo).FieldName, FieldList(RowNo).StatusFieldName, FieldList(RowNo).FieldType, FieldList(RowNo).FieldLen, FieldList(RowNo).MultiValued, FieldList(RowNo).Status, FieldList(RowNo).NValues, FieldList(RowNo).RowNo) & vbCrLf)
        'Next
        'SetMessageColor(Color.Black)
        'AddMessage(vbCrLf)

    End Sub

    Public Sub GetFieldValues()
        'Gets a list of destination Field Values from DbDestValues(,) (previously DataGridView1)
        '    The list is stored in FieldValues(,)

        'Debug.Print("Running GetFieldValues" & vbCrLf)

        'DataGridView1 contains the raw data extracted from the text file using the Regular Expressions
        'Data for each Value is saved in the FieldValues array.

        'The GetFieldValues subroutine puts the data values in the FieldValues array:
        'FieldValues(,) Contains the following fields:
        '    Value  'If the value is a number, any corresponding multiplier is applied
        '    Status     'OK: the value is valid, N/A: the value is not available, MIS: missing value, ERR: the value has an error, W01: warning 01 (to be defined), W02 etc

        'FieldList() Contains the following fields:
        '    TableName
        '    FieldName 
        '    FieldType  'Text-Memo-Number-Date/Time-Currency-AutoNumber-Yes/No-Hyperlink Number:Byte-Integer-LongInteger-Single-Double
        '    FieldLen
        '    Status     'Gap-Mis
        '    NValues    'The number of values for this Field to write to the database
        '    Mult       'The multiplier used to scale the Field values in the source document
        '    RowNo      'The Row Number of the Field in DataGridView1

        Dim NFields As Integer
        Dim MaxNValues As Integer
        Dim FieldNo As Integer
        Dim RowNo As Integer
        Dim ValueNo As Integer
        Dim ValueStr As String
        Dim ValueDbl As Double

        NFields = FieldList.Count
        'MaxNValues = DbDestValues.GetUpperBound(1)
        MaxNValues = DbDestValues.GetUpperBound(1) + 1 '*** UPDATED 22Dec16 ***

        'For debugging:
        'RaiseEvent Message("GetFieldValues: MaxNValues = " & MaxNValues & vbCrLf)

        ReDim FieldValues(0 To NFields - 1, 0 To MaxNValues)

        For FieldNo = 0 To NFields - 1 'Process each field
            RowNo = FieldList(FieldNo).RowNo 'Get the Row Number in DataGridView1

            If FieldList(FieldNo).NValues = 0 Then 'No values to process
                FieldValues(FieldNo, 0).Value = "0"
                FieldValues(FieldNo, 0).Status = "MIS"
            Else
                For ValueNo = 0 To FieldList(FieldNo).NValues - 1 'Process each Value
                    If IsNothing(DbDestValues(RowNo, ValueNo)) Then
                        ValueStr = ""
                        'Debug.Print("DbDestValues(" & RowNo & "," & ValueNo & ") = Nothing" & vbCrLf)
                    Else
                        ValueStr = DbDestValues(RowNo, ValueNo)
                        'Debug.Print("DbDestValues(" & RowNo & "," & ValueNo & ") = " & ValueStr & vbCrLf)
                    End If

                    If ValueStr = "N/A" Then
                        FieldValues(FieldNo, ValueNo).Value = "0"
                        FieldValues(FieldNo, ValueNo).Status = "N/A"
                    ElseIf ValueStr = "n/a" Then
                        FieldValues(FieldNo, ValueNo).Value = "0"
                        FieldValues(FieldNo, ValueNo).Status = "N/A"
                    ElseIf Trim(ValueStr) = "" Then
                        FieldValues(FieldNo, ValueNo).Value = "0"
                        FieldValues(FieldNo, ValueNo).Status = "MIS"
                    Else 'Valid value
                        FieldValues(FieldNo, ValueNo).Status = "OK"
                        If GridProp(RowNo, ValueNo).Mult = 1 Then
                            FieldValues(FieldNo, ValueNo).Value = ValueStr
                            'Debug.Print("GridProp(" & RowNo & "," & ValueNo & ").Mult = 1" & vbCrLf)
                            'Debug.Print("FieldValues(" & FieldNo & "," & ValueNo & ").Value = " & ValueStr & vbCrLf)
                        Else
                            'Debug.Print("GridProp(" & RowNo & "," & ValueNo & ").Mult =" & Str(GridProp(RowNo, ValueNo).Mult) & vbCrLf)
                            ValueDbl = CDbl(Val(ValueStr.Replace(",", ""))) 'Commas
                            ValueDbl = ValueDbl * GridProp(RowNo, ValueNo).Mult
                            ValueStr = Str(ValueDbl)
                            'Debug.Print("Rescaled value = " & ValueStr & vbCrLf)
                            FieldValues(FieldNo, ValueNo).Value = ValueStr
                            'Debug.Print("FieldValues(" & FieldNo & "," & ValueNo & ").Value = " & ValueStr & vbCrLf)
                        End If
                    End If
                Next
            End If
        Next

        'FOR DEBUGGING: --------------------------------------------------------
        'Display the contents of the FieldValues array in the Debug window:
        'Debug.Print("FieldValues(,)")
        'Debug.WriteLine(String.Format("{0,-8}{1,-14}{2,-12}{3,-12}", "Row No", "Field Value", "Status", "..."))
        'For RowNo = 0 To FieldValues.GetUpperBound(0)
        '    Debug.Write(String.Format("{0,-8}", RowNo))
        '    For ValueNo = 0 To FieldValues.GetUpperBound(1)
        '        Debug.Write(String.Format("{0,-14}{1,-4}", FieldValues(RowNo, ValueNo).Value, FieldValues(RowNo, ValueNo).Status))
        '    Next
        '    Debug.WriteLine("")
        'Next
        'Debug.Print("")
        DebugWriteFieldValues()
        ' -----------------------------------------------------------------------
    End Sub

    Private Sub DebugWriteFieldValues()
        'For debugging: Write contents of the FieldValues array to the Messages window:

        'SetMessageColor(Color.Black)
        'SetMessageFontName(FontList.Lucida_Sans_Typewriter)
        'AddMessage("FieldValues(,)" & vbCrLf)

        'AddMessage(String.Format("{0,-8}{1,-14}{2,-12}{3,-12}", "Row No", "Field Value", "Status", "...") & vbCrLf)
        'For RowNo = 0 To FieldValues.GetUpperBound(0)
        '    AddMessage(String.Format("{0,-8}", RowNo))
        '    For ValueNo = 0 To FieldValues.GetUpperBound(1)
        '        SetMessageColor(Color.Blue)
        '        AddMessage(String.Format("{0,-14}{1,-4}", FieldValues(RowNo, ValueNo).Value, FieldValues(RowNo, ValueNo).Status) & vbCrLf)
        '    Next
        'Next
        'SetMessageColor(Color.Black)
        'AddMessage(vbCrLf)

    End Sub

    Public Sub ProcessMatches()
        'Process the matches.
        'This step is required before they can be written to the database.

        SetMultipliers()
        GetTableList()
        GetFieldList()
        GetFieldValues()

    End Sub

    'Public Sub WriteToDatabase(ByRef da As System.Data.OleDb.OleDbDataAdapter)
    'The data adaptor da is now declared within this class. - No need to pass this object as an argument.
    Public Sub WriteToDatabase()
        'Writes the data in DataGridView1 to the database

        Dim TabNo As Integer 'Table Number
        Dim TableName As String
        Dim ValNo As Integer 'Value Number (Used to write multiple value fields)
        Dim FldNo As Integer 'Field Number
        Dim FirstField As Boolean

        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim FieldNames As String 'Holds the list of Field Names used in the INSERT commandString
        Dim FieldVals As String 'Holds the list of Field Values used in the INSERT commandString

        'For debugging:
        'RaiseEvent Message("Starting WriteToDatabase()" & vbCrLf)
        'RaiseEvent Message("TableList.Count = " & TableList.Count & vbCrLf)
        'RaiseEvent Message("FieldList.Count = " & FieldList.Count & vbCrLf)

        For TabNo = 0 To TableList.Count - 1 'Process each Table
            TableName = TableList(TabNo).TableName

            'Get list of FieldNames:
            FirstField = True
            FieldNames = ""
            For FldNo = 0 To FieldList.Count - 1
                If FieldList(FldNo).TableName = TableName Then
                    If FirstField = True Then
                        FieldNames = "[" & FieldList(FldNo).FieldName & "]"
                        If FieldList(FldNo).StatusFieldName <> "" Then
                            FieldNames = FieldNames & ", [" & FieldList(FldNo).StatusFieldName & "]"
                        End If
                        FirstField = False
                    Else
                        FieldNames = FieldNames & ", [" & FieldList(FldNo).FieldName & "]"
                        If FieldList(FldNo).StatusFieldName <> "" Then
                            FieldNames = FieldNames & ", [" & FieldList(FldNo).StatusFieldName & "]"
                        End If
                    End If
                End If
            Next 'FldNo

            'For debugging:
            'RaiseEvent Message("TableList(TabNo).MaxNValues = " & TableList(TabNo).MaxNValues & vbCrLf)


            'Get list of Field Values:
            For ValNo = 0 To TableList(TabNo).MaxNValues - 1
                FirstField = True
                FieldVals = ""
                For FldNo = 0 To FieldList.Count - 1
                    If FieldList(FldNo).TableName = TableName Then
                        If FirstField = True Then
                            If FieldList(FldNo).MultiValued = False Then 'Single valued field
                                If (FieldList(FldNo).FieldType = 130) Then 'String - needs single quotes
                                    FieldVals = "'" & FieldValues(FldNo, 0).Value.Replace("'", "").Replace(vbDouble, "").Replace(vbCrLf, "") & "'"
                                    FirstField = False
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Status & "'"
                                    End If
                                ElseIf FieldList(FldNo).FieldType = 7 Then 'Date - needs single quotes
                                    FieldVals = "'" & FieldValues(FldNo, 0).Value.Replace(".", "/") & "'" 'Use / to delimit day/month/year
                                    FirstField = False
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Status & "'"
                                    End If
                                Else 'Values - do not need single quotes
                                    'UPDATE 23Dec16 - Check for NullValueString:
                                    If UseNullValueString = True Then
                                        If FieldValues(FldNo, 0).Value = NullValueString Then
                                            'FieldVals = " "
                                            'FieldVals = ""
                                            'FieldVals = "'Null'"
                                            FieldVals = " null"
                                        Else
                                            FieldVals = FieldValues(FldNo, 0).Value.Replace(",", "").Replace("$", "")
                                        End If
                                    Else
                                        'Original code:
                                        FieldVals = FieldValues(FldNo, 0).Value.Replace(",", "").Replace("$", "")
                                    End If
                                    'FieldVals = FieldValues(FldNo, 0).Value.Replace(",", "").Replace("$", "")
                                    FirstField = False
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Status & "'"
                                    End If
                                End If
                            Else 'Multi-Valued field
                                If (FieldList(FldNo).FieldType = 130) Then 'String - needs single quotes
                                    FieldVals = "'" & FieldValues(FldNo, ValNo).Value.Replace("'", "").Replace(vbDouble, "").Replace(vbCrLf, "") & "'"
                                    FirstField = False
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Status & "'"
                                    End If
                                ElseIf FieldList(FldNo).FieldType = 7 Then 'Date - needs single quotes
                                    FieldVals = "'" & FieldValues(FldNo, ValNo).Value.Replace(".", "/") & "'" 'Use / to delimit day/month/year
                                    FirstField = False
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Status & "'"
                                    End If
                                Else 'Values - do not need single quotes
                                    If FieldValues(FldNo, ValNo).Value <> Nothing Then
                                        'UPDATE 23Dec16 - Check for NullValueString:
                                        If UseNullValueString = True Then
                                            If FieldValues(FldNo, ValNo).Value = NullValueString Then
                                                'FieldVals = " "
                                                'FieldVals = ""
                                                'FieldVals = "'Null'"
                                                FieldVals = " null"
                                            Else
                                                FieldVals = FieldValues(FldNo, ValNo).Value.Replace(",", "").Replace("$", "")
                                            End If
                                        Else
                                            'Original code:
                                            FieldVals = FieldValues(FldNo, ValNo).Value.Replace(",", "").Replace("$", "")
                                        End If
                                        'FieldVals = FieldValues(FldNo, ValNo).Value.Replace(",", "").Replace("$", "")
                                        If FieldList(FldNo).StatusFieldName <> "" Then
                                            FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Status & "'"
                                        End If
                                    Else
                                        FieldVals = "0"
                                        FieldVals = FieldVals & ", '" & "MIS" & "'"
                                    End If
                                    FirstField = False
                                End If

                            End If
                        Else 'FirstField = False
                            If FieldList(FldNo).MultiValued = False Then 'Single valued field
                                If (FieldList(FldNo).FieldType = 130) Then 'String - needs single quotes
                                    FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Value.Replace("'", "").Replace(vbDouble, "").Replace(vbCrLf, "") & "'"
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Status & "'"
                                    End If
                                ElseIf FieldList(FldNo).FieldType = 7 Then 'Date - needs single quotes
                                    FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Value.Replace(".", "/") & "'" 'Use / to delimit day/month/year
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Status & "'"
                                    End If
                                Else 'Values - do not need single quotes
                                    'UPDATE 23Dec16 - Check for NullValueString:
                                    If UseNullValueString = True Then
                                        If FieldValues(FldNo, 0).Value = NullValueString Then
                                            'FieldVals = FieldVals & ", "
                                            'FieldVals = FieldVals & ","
                                            'FieldVals = FieldVals & ", 'Null'"
                                            FieldVals = FieldVals & ", null"
                                        Else
                                            FieldVals = FieldVals & ", " & FieldValues(FldNo, 0).Value.Replace(",", "").Replace("$", "")
                                        End If
                                    Else
                                        'Original code:
                                        FieldVals = FieldVals & ", " & FieldValues(FldNo, 0).Value.Replace(",", "").Replace("$", "")
                                    End If
                                    '   FieldVals = FieldVals & ", " & FieldValues(FldNo, 0).Value.Replace(",", "").Replace("$", "")
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, 0).Status & "'"
                                    End If
                                End If
                            Else 'Multi-Valued field
                                If (FieldList(FldNo).FieldType = 130) Then 'String - needs single quotes
                                    FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Value.Replace("'", "").Replace(vbDouble, "").Replace(vbCrLf, "") & "'"
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Status & "'"
                                    End If
                                ElseIf FieldList(FldNo).FieldType = 7 Then 'Date - needs single quotes
                                    FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Value.Replace(".", "/") & "'" 'Use / to delimit day/month/year
                                    If FieldList(FldNo).StatusFieldName <> "" Then
                                        FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Status & "'"
                                    End If
                                Else 'Values - do not need single quotes
                                    If FieldValues(FldNo, ValNo).Value <> Nothing Then
                                        'UPDATE 23Dec16 - Check for NullValueString:
                                        If UseNullValueString = True Then
                                            If FieldValues(FldNo, ValNo).Value = NullValueString Then
                                                'FieldVals = FieldVals & ", "
                                                'FieldVals = FieldVals & ","
                                                'FieldVals = FieldVals & ", 'Null'"
                                                FieldVals = FieldVals & ", null"
                                            Else
                                                FieldVals = FieldVals & ", " & FieldValues(FldNo, ValNo).Value.Replace(",", "").Replace("$", "")
                                            End If
                                        Else
                                            'Original code:
                                            FieldVals = FieldVals & ", " & FieldValues(FldNo, ValNo).Value.Replace(",", "").Replace("$", "")
                                        End If

                                        If FieldList(FldNo).StatusFieldName <> "" Then
                                            FieldVals = FieldVals & ", '" & FieldValues(FldNo, ValNo).Status & "'"
                                        End If
                                    Else
                                        FieldVals = FieldVals & ", '0'"
                                        If FieldList(FldNo).StatusFieldName <> "" Then
                                            FieldVals = FieldVals & ", 'MIS'"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next 'FldNo

                'Specify the commandString to query the database:
                commandString = "INSERT INTO " & TableName & "(" & FieldNames & ") VALUES (" & FieldVals & ")"

                'For debugging:
                'RaiseEvent Message("commandString = " & commandString & vbCrLf)

                Try
                    InsertCount = InsertCount + 1
                    'THE FOLLOWING CODE CAUSES AND ERROR:
                    'If (InsertCount Mod 100) = 0 Then
                    '    'Debug.Print("Insert Count = " & InsertCount)
                    '    'NOTE: Convert this code to a  message event!!! ---------------------------------------------------------------------------------
                    '    'ShowMessage("Insert Count = " & InsertCount & vbCrLf, Color.Black)
                    '    DebugMessages.MessageColor = Drawing.Color.Black
                    '    DebugMessages.MessageFontName = frmDebugMessages.FontList.Arial
                    '    DebugMessages.MessageFontSize = 10
                    '    DebugMessages.MessageFontStyle = Drawing.FontStyle.Regular

                    'End If
                    'ERROR MESSAGE(S): Error with the following InsertCommand string:
                    'INSERT INTO Daily_Prices([ASX_Code], [Trade_Date], [Open], [High], [Low], [Close], [Volume]) VALUES ('CSLVZX', '02 January 2009', 2.66, 2.66, 2.66, 2.66, 0)
                    'Insert Command Error: Object reference not set to an instance of an object.
                    'I DONT KNOW WHY IT CAUSES THIS ERROR!!!!!!!
                    'COMMENTING OUT THE CODE STOPS THE ERROR!!!!!

                    da.InsertCommand.CommandText = commandString
                    da.InsertCommand.ExecuteNonQuery()
                    'Debug.WriteLine("OK: " & commandString)
                    'Debug.WriteLine("")
                Catch ex As Exception
                    'Debug.WriteLine("")
                    'NOTE: Convert this code to a  message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage(vbCrLf, Color.Red)
                    'Debug.WriteLine("Error with the following InsertCommand string:")
                    'NOTE: Convert this code to a  message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage("Error with the following InsertCommand string:", Color.Red)
                    'Debug.WriteLine(commandString)
                    'NOTE: Convert this code to a  message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage(commandString & vbCrLf, Color.Black)
                    'Debug.Write(ex.ToString)
                    'Debug.WriteLine("Insert Command Error: " & ex.Message)
                    'Debug.WriteLine("")
                    'NOTE: Convert this code to a  message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage(ex.Message & vbCrLf & vbCrLf, Color.Blue)
                    'NOTE: Convert this code to a  message event!!! ---------------------------------------------------------------------------------
                    'ShowMessage("Input text: " & TextStore & vbCrLf, Color.Black)
                    'RaiseEvent Warning("Error with the following InsertCommand string:" & vbCrLf)
                    RaiseEvent ErrorMessage("Error with the following InsertCommand string:" & vbCrLf)
                    'RaiseEvent Warning(commandString & vbCrLf)
                    RaiseEvent ErrorMessage(commandString & vbCrLf)
                    'RaiseEvent Warning("Insert Command Error: " & ex.Message & vbCrLf)
                    RaiseEvent ErrorMessage("Insert Command Error: " & ex.Message & vbCrLf)
                End Try
            Next 'ValNo
        Next 'TabNo

        'da.InsertCommand.Connection.Close()

    End Sub


#End Region 'Write to Database Methods

#Region " Modify Values Code"

    'Public Sub ApplyModifyValues()
    '    'Apply the modification to the Database Destinations table:

    '    Dim I As Integer
    '    Dim strVarName As String 'This is used to hold the RegEx variable name in the current row
    '    Dim CapCount As Integer
    '    Dim J As Integer
    '    'Find matching RegEx Variables in the Database Destinations array:
    '    'For I = 1 To TDS_Import.DbDestinations.DataGridView1.RowCount
    '    For I = 1 To mDbDest.Count
    '        'If TDS_Import.DbDestinations.DataGridView1.Rows(I - 1).Cells(0).Value = Nothing Then 'No RegEx variable name is specified.
    '        If mDbDest(I - 1).RegExVariable = Nothing Then
    '        Else
    '            'strVarName = TDS_Import.DbDestinations.DataGridView1.Rows(I - 1).Cells(0).Value.ToString 'The RegEx variable name in the current grid row.
    '            strVarName = mDbDest(I - 1).RegExVariable 'The RegEx variable name in the current destination row.
    '            'If strVarName = txtRegExVariable.Text Then 'The RegExVariable at the current row matches the required variable to modify
    '            'ModifyValuesRegExVariable holds the name of the RegEx variable to modify
    '            If strVarName = ModifyValuesRegExVariable Then
    '                'If IsNothing(TDS_Import.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value) Then
    '                'If IsNothing(TDS_Import.DbDestinations.DataGridView2.Rows(I - 1).Cells(0).Value) Then
    '                If IsNothing(DbDestValues(I - 1, 0)) Then
    '                    'There is no text to modify
    '                Else
    '                    'txtTestInputString.Text = TextToDatabase.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value.ToString
    '                    Dim OutputString As String
    '                    Dim InputString As String
    '                    'InputString = TDS_Import.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value.ToString
    '                    InputString = DbDestValues(I - 1, 0).ToString
    '                    'ConvertDate(txtInputDateFormat.Text, txtOutputDateFormat.Text, InputString, OutputString)
    '                    ConvertDate(ModifyValuesInputDateFormat, ModifyValuesOutputDateFormat, InputString, OutputString)

    '                    'TDS_Import.DbDestinations.DataGridView1.Rows(I - 1).Cells(5).Value = OutputString
    '                    DbDestValues(I - 1, 0) = OutputString

    '                End If

    '            End If

    '        End If
    '    Next

    'End Sub

#End Region 'Modify Values Code

#Region " Import Status Methods"

    Public Sub ImportStatusRemove(ByVal Status As String)
        'Remove the specified Status string from the ImportStatus collection if it is present:

        If IsNothing(ImportStatus) Then
            'ImportStatus contains no Status strings
        Else
            If ImportStatus.Contains(Status) Then
                ImportStatus.Remove(Status)
            End If

        End If
    End Sub

    Public Function ImportStatusContains(ByVal Status As String) As Boolean
        'Check if the ImportStatus collection contains the specified Status string:

        If IsNothing(ImportStatus) Then
            Return False
        Else
            If ImportStatus.Contains(Status) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    'Private Sub SequenceStatusTest(ByVal Status As String)
    '    'Code example using the ImportStatus collection:

    '    ImportStatus.Add("At_end_of_file")

    '    ImportStatus.Remove("At_end_of_file")

    '    If ImportStatus.Contains("At_end_of_file") Then

    '    End If

    'End Sub

#End Region 'Import Status Methods

#Region " Run Sequence Methods"

    Public Sub RunXSequence(ByRef xDoc As System.Xml.XmlDocument)
        Debug.Print("Running Import.RunXSequence(xDoc)")
        XSeq.RunXSequence(xDoc, ImportStatus)
    End Sub

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        'RaiseEvent Warning(ErrMsg)
        RaiseEvent ErrorMessage(ErrMsg)
    End Sub

    'Private Sub XSeq_Instruction_Old(Path As String, Prop As String) Handles XSeq.Instruction_Old
    '    'Execute each instruction produced by running the XSequence file.

    '    Debug.Print("Execute instruction: Path = " & Path)

    '    Select Case Path


    '        'Input Data:
    '        Case "InputData:TextFileDirectory"
    '            Debug.Print("Execute instruction: InputData:TextFileDirectory - TextFileDir = " & Prop)
    '            TextFileDir = Prop
    '        Case "InputData:TextFilesToProcess:SelectFileMode"
    '            SelectFileMode = Prop
    '        Case "InputData:TextFilesToProcess:Command"
    '            If Prop = "ClearSelectedFileList" Then
    '                SelTextFilesClear()
    '            End If
    '        Case "InputData:TextFilesToProcess:TextFile"
    '            SelTextFileAppend(Prop)
    '        Case "InputData:TextFilesToProcess:SelectionFilePath"
    '            SelectionFileName = Prop

    '            'Database:
    '        Case "Database:DatabasePath"
    '            Debug.Print("Execute instruction: Database:DatabasePath - DatabasePath = " & Prop)
    '            DatabasePath = Prop

    '        Case "Database:DatabaseType"
    '            If Prop = "Access2007To2013" Then
    '                DatabaseType = DatabaseTypeEnum.Access2007To2013
    '            ElseIf Prop = "User_defined_connection_string" Then
    '                DatabaseType = DatabaseTypeEnum.User_defined_connection_string
    '            End If


    '            'Database Destinations List:
    '        Case "DatabaseDestinationsList"
    '            DbDestListName = Prop
    '            OpenDbDestListFile()

    '            'Read Text Commands:
    '        Case "ReadTextCommand"
    '            Debug.Print("Execute instruction: ReadTextCommand")
    '            Debug.Print("Command = " & Prop)
    '            If Prop = "OpenFirstFile" Then
    '                Debug.Print("Running SelectFirstFile()")
    '                SelectFirstFile()
    '            ElseIf Prop = "OpenNextFile" Then
    '                Debug.Print("Running SelectNextFile()")
    '                SelectNextFile()
    '            ElseIf Prop = "ReadNextLine" Then
    '                Debug.Print("Running ReadNextLineOfText()")
    '                ReadNextLineOfText()
    '                If ReadLineCount Mod 200 = 0 Then
    '                    'RaiseEvent Notice(" " & ReadLineCount) 'Show the number of lines read
    '                    RaiseEvent Message(" " & ReadLineCount & vbCrLf) 'Show the number of lines read
    '                End If
    '            End If

    '            'Match Text RegEx List:
    '        Case "MatchTextRegExList"
    '            RegExListName = Prop
    '            OpenRegExListFile()

    '            'Processing Commands:
    '        Case "ProcessingCommand"
    '            If Prop = "RunRegExList" Then
    '                RunRegExList()
    '            ElseIf Prop = "OpenDatabase" Then
    '                OpenDatabase()
    '            ElseIf Prop = "ProcessMatches" Then
    '                ProcessMatches()
    '            ElseIf Prop = "WriteToDatabase" Then
    '                WriteToDatabase()
    '            ElseIf Prop = "CloseDatabase" Then
    '                CloseDatabase()
    '            End If

    '            'Modify Values:
    '        Case "ModifyValues:RegExVariable"
    '            ModifyValuesRegExVariable = Prop
    '        Case "ModifyValues:InputDateFormat"
    '            ModifyValuesInputDateFormat = Prop
    '        Case "ModifyValues:OutputDateFormat"
    '            ModifyValuesOutputDateFormat = Prop
    '        Case "ModifyValues:CharactersToReplace"
    '            ModifyValuesCharsToReplace = Prop
    '        Case "ModifyValues:ReplacementCharacters"
    '            ModifyValuesReplacementChars = Prop
    '        Case "ModifyValues:FixedValue"
    '            mModifyValuesFixedValue = Prop
    '        Case "ModifyValues:ModifyType"
    '            If Prop = "Convert_date" Then
    '                ModifyValuesType = "Convert_date"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Replace_characters" Then
    '                ModifyValuesType = "Replace_characters"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Fixed_value" Then
    '                ModifyValuesType = "Fixed_value"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Text_file_name" Then
    '                ModifyValuesType = "Text_file_name"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Text_file_directory" Then
    '                ModifyValuesType = "Text_file_directory"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Text_file_path" Then
    '                ModifyValuesType = "Text_file_path"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Current_date" Then
    '                ModifyValuesType = "Current_date"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Current_time" Then
    '                ModifyValuesType = "Current_time"
    '                ModifyValuesApply()
    '            ElseIf Prop = "Current_date_time" Then
    '                ModifyValuesType = "Current_date_time"
    '                ModifyValuesApply()
    '            End If

    '        Case "EndOfSequence"
    '            'RaiseEvent Notice(vbCrLf & "End of processing sequence" & vbCrLf) 'Show the number of lines read
    '            RaiseEvent Message(vbCrLf & "End of processing sequence" & vbCrLf) 'Show the number of lines read
    '    End Select

    '    'RaiseEvent Notice("Path: " & Path & "  Prop: " & Prop) 'For debugging only!

    'End Sub

    Private Sub XSeq_Instruction(Info As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSequence file.

        'Debug.Print("Execute noxel: Info = " & Info & "  Locn = " & Locn)

        Select Case Locn
            'Input Data: ------------------------------------------------------------
            Case "InputData:TextFileDirectory"
                TextFileDir = Info
            Case "InputData:TextFilesToProcess:SelectFileMode"
                SelectFileMode = Info
            Case "InputData:TextFilesToProcess:Command"
                If Info = "ClearSelectedFileList" Then
                    SelTextFilesClear()
                End If
            Case "InputData:TextFilesToProcess:TextFile"
                SelTextFileAppend(Info)
            Case "InputData:TextFilesToProcess:SelectionFilePath"
                SelectionFileName = Info

           'Database: ----------------------------------------------------------------
           ''Old version:
           ' Case "Database:DatabasePath"
           '     Debug.Print("Execute instruction: Database:DatabasePath - DatabasePath = " & Info)
           '     DatabasePath = Info
           '     'Old version:
           ' Case "Database:DatabaseType"
           '     If Info = "Access2007To2013" Then
           '         DatabaseType = DatabaseTypeEnum.Access2007To2013
           '     ElseIf Info = "User_defined_connection_string" Then
           '         DatabaseType = DatabaseTypeEnum.User_defined_connection_string
           '     End If
           '     'New version:
           ' Case "Database"
           '     Debug.Print("Execute instruction: Database: - DatabasePath = " & Info)
           '     DatabasePath = Info

                'Newest version:
            Case "Database:Path"
                'Debug.Print("Execute instruction: Database:Path - DatabasePath = " & Info)
                DatabasePath = Info
            Case "Database:Type"
                If Info = "Access2007To2013" Then
                    DatabaseType = DatabaseTypeEnum.Access2007To2013
                ElseIf Info = "User_defined_connection_string" Then
                    DatabaseType = DatabaseTypeEnum.User_defined_connection_string
                End If

            'Database Destinations List: ---------------------------------------------
            Case "DatabaseDestinationsList"
                DbDestListName = Info
                OpenDbDestListFile()

            'Read Text Commands: ------------------------------------------------------
            Case "ReadTextCommand"
                'Debug.Print("Execute instruction: ReadTextCommand")
                'Debug.Print("Command = " & Info)
                If Info = "OpenFirstFile" Then
                    'Debug.Print("Running SelectFirstFile()")
                    SelectFirstFile()
                ElseIf Info = "OpenNextFile" Then
                    'Debug.Print("Running SelectNextFile()")
                    SelectNextFile()
                ElseIf Info = "ReadNextLine" Then
                    'Debug.Print("Running ReadNextLineOfText()")
                    ReadNextLineOfText()
                    If ReadLineCount Mod 200 = 0 Then
                        'RaiseEvent Notice(" " & ReadLineCount) 'Show the number of lines read
                        'RaiseEvent Message(" " & ReadLineCount & vbCrLf) 'Show the number of lines read
                        RaiseEvent Message(" " & ReadLineCount) 'Show the number of lines read
                    End If
                End If

            'Match Text RegEx List: --------------------------------------------------------
            Case "MatchTextRegExList"
                RegExListName = Info
                OpenRegExListFile()

            'Processing Commands: ----------------------------------------------------------
            Case "ProcessingCommand"
                If Info = "RunRegExList" Then
                    RunRegExList()
                ElseIf Info = "OpenDatabase" Then
                    DatabaseType = DatabaseTypeEnum.Access2007To2013
                    OpenDatabase()
                ElseIf Info = "ProcessMatches" Then
                    ProcessMatches()
                ElseIf Info = "WriteToDatabase" Then
                    WriteToDatabase()
                ElseIf Info = "CloseDatabase" Then
                    CloseDatabase()
                Else
                    RaiseEvent ErrorMessage("Unknown Processing Command: " & Info & vbCrLf)
                End If

                'Modify Values: -----------------------------------------------------
            Case "ModifyValues:RegExVariable"
                ModifyValuesRegExVariable = Info
            Case "ModifyValues:InputDateFormat"
                ModifyValuesInputDateFormat = Info
            Case "ModifyValues:OutputDateFormat"
                ModifyValuesOutputDateFormat = Info
            Case "ModifyValues:CharactersToReplace"
                ModifyValuesCharsToReplace = Info
            Case "ModifyValues:ReplacementCharacters"
                ModifyValuesReplacementChars = Info
            Case "ModifyValues:FixedValue"
                mModifyValuesFixedValue = Info
            'Case "ModifyValues:RegExVariableToAppend"
            Case "ModifyValues:RegExVariableValueFrom"
                '_modifyValuesRegExVarToAppend = Info
                _modifyValuesRegExVarValFrom = Info
            Case "ModifyValues:ModifyType"
                If Info = "Convert_date" Then
                    'ModifyValuesType = "Convert_date"
                    ModifyValuesType = ModifyValuesTypes.ConvertDate
                    ModifyValuesApply()
                ElseIf Info = "Clear_value" Then
                    ModifyValuesType = ModifyValuesTypes.ClearValue
                    ModifyValuesApply()
                ElseIf Info = "Replace_characters" Then
                    'ModifyValuesType = "Replace_characters"
                    ModifyValuesType = ModifyValuesTypes.ReplaceChars
                    ModifyValuesApply()
                    'ElseIf Info = "Fixed_value" Then
                ElseIf Info = "Append_fixed_value" Then
                    'ModifyValuesType = "Fixed_value"
                    'ModifyValuesType = ModifyValuesTypes.FixedValue
                    ModifyValuesType = ModifyValuesTypes.AppendFixedValue
                    ModifyValuesApply()
                    'ElseIf Info = "Text_file_name" Then
                ElseIf Info = "Append_RegEx_variable_value" Then
                    ModifyValuesType = ModifyValuesTypes.AppendRegExVarValue
                    ModifyValuesApply()
                ElseIf Info = "Append_file_name" Then
                    'ModifyValuesType = "Text_file_name"
                    'ModifyValuesType = ModifyValuesTypes.FileName
                    ModifyValuesType = ModifyValuesTypes.AppendFileName
                    ModifyValuesApply()
                    'ElseIf Info = "Text_file_directory" Then
                ElseIf Info = "Append_file_directory" Then
                    'ModifyValuesType = "Text_file_directory"
                    'ModifyValuesType = ModifyValuesTypes.FileDir
                    ModifyValuesType = ModifyValuesTypes.AppendFileDir
                    ModifyValuesApply()
                    'ElseIf Info = "Text_file_path" Then
                ElseIf Info = "Append_file_path" Then
                    'ModifyValuesType = "Text_file_path"
                    'ModifyValuesType = ModifyValuesTypes.FilePath
                    ModifyValuesType = ModifyValuesTypes.AppendFilePath
                    ModifyValuesApply()
                    'ElseIf Info = "Current_date" Then
                ElseIf Info = "Append_current_date" Then
                    'ModifyValuesType = "Current_date"
                    'ModifyValuesType = ModifyValuesTypes.CurrentDate
                    ModifyValuesType = ModifyValuesTypes.AppendCurrentDate
                    ModifyValuesApply()
                    'ElseIf Info = "Current_time" Then
                ElseIf Info = "Append_current_time" Then
                    'ModifyValuesType = "Current_time"
                    'ModifyValuesType = ModifyValuesTypes.CurrentTime
                    ModifyValuesType = ModifyValuesTypes.AppendCurrentTime
                    ModifyValuesApply()
                    'ElseIf Info = "Current_date_time" Then
                ElseIf Info = "Append_current_date_time" Then
                    'ModifyValuesType = "Current_date_time"
                    'ModifyValuesType = ModifyValuesTypes.CurrentDateTime
                    ModifyValuesType = ModifyValuesTypes.AppendCurrentDateTime
                    ModifyValuesApply()
                End If

            Case "EndOfSequence"
                RaiseEvent Message(vbCrLf & "End of processing sequence" & vbCrLf) 'Show the number of lines read

            Case Else
                RaiseEvent ErrorMessage("Unknown Locn: " & Locn & vbCrLf)
        End Select

    End Sub

    'Private Sub RunXSequence_Instruction(Path As String, Prop As String) Handles RunXSequence.Instruction
    '    RaiseEvent Notice("Path: " & Path & "  Prop: " & Prop & vbCrLf)
    'End Sub

#End Region 'Run Sequence Methods


    'Public Sub AddMessage(ByVal strMsg As String)
    '    'Add the message to the DebugMessages Form

    '    'Check is the form is open:
    '    If IsNothing(DebugMessages) Then
    '        DebugMessages = New frmDebugMessages
    '        DebugMessages.Show()
    '    Else
    '        DebugMessages.Show()
    '    End If

    '    Dim StrLen As Integer
    '    Dim StrStart As Integer

    '    StrStart = DebugMessages.rtbMessages.TextLength
    '    StrLen = strMsg.Length

    '    DebugMessages.rtbMessages.AppendText(strMsg)
    '    DebugMessages.rtbMessages.Select(StrStart, StrLen)
    '    'DebugMessages.rtbMessages.SelectionColor = MessageColor
    '    DebugMessages.rtbMessages.SelectionColor = DebugMessages.MessageColor
    '    'DebugMessages.rtbMessages.SelectionFont = New Font(MessageFontName, MessageFontSize, MessageFontStyle)
    '    DebugMessages.rtbMessages.SelectionFont = New System.Drawing.Font(DebugMessages.MessageFontName, DebugMessages.MessageFontSize, DebugMessages.MessageFontStyle)

    '    DebugMessages.Refresh() '

    '    'Scroll to the end of the message window:
    '    DebugMessages.rtbMessages.ScrollToCaret()

    '    DebugMessages.BringToFront()
    'End Sub

#Region " Events"

    'Public Event Status(ByVal StatusText As String) 'Send a status message

    'Public Event Warning(ByVal WarningText As String) 'Send a warning message
    Public Event ErrorMessage(ByVal Message As String) 'Send an error message.

    'Public Event Notice(ByVal NoticeText As String) 'Send a notice message
    Public Event Message(ByVal Message As String) 'Send a message

#End Region 'Events

End Class 'Import
