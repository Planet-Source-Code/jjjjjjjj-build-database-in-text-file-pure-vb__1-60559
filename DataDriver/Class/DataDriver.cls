VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBaseDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^¶¶¶^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^¶¶^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^¶¶¶^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^¶¶^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^¶¶^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^¶¶^^^^¶¶¶¶¶^^^¶¶¶¶¶¶^^^¶¶¶¶¶^^^¶¶^^^^^¶¶^^^^¶¶¶¶¶^^^^¶¶¶¶¶^^^^¶¶¶¶¶^^^^^^^^¶¶^^^^^^¶¶^^¶¶^¶¶¶^^¶¶^^¶¶^^^^¶¶^^^¶¶¶¶¶^^^¶¶^¶¶¶^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^^¶¶^^¶¶^^^¶¶^^^¶¶^^^^^¶¶^^^¶¶^^¶¶^^^^¶¶^^^^¶¶^^^¶¶^^¶¶^^^^¶^^¶¶^^^¶¶^^^^^^^¶¶^^^^^^^¶¶^¶¶¶¶¶¶^^¶¶^^¶¶^^^^¶¶^^¶¶^^^¶¶^^¶¶¶¶¶¶^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^^¶¶^^^^^^^^¶¶^^¶¶^^^^^^^^^^^¶¶^¶¶¶¶¶¶¶¶^^^^^^^^^^¶¶^¶¶^^^^^^¶¶^^^^^¶¶^^^^^^¶¶^^^^^^^¶¶^¶¶^^^^^^¶¶^^¶¶^^^^¶¶^¶¶^^^^^¶¶^¶¶^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^^¶¶^^^^¶¶¶¶¶¶^^¶¶^^^^^^^¶¶¶¶¶¶^¶¶^^^^^¶¶^^^^^¶¶¶¶¶¶^¶¶¶¶^^^^¶¶^^^^^¶¶^^^^^^¶¶^^^^^^^¶¶^¶¶^^^^^^¶¶^^^¶¶^^¶¶^^¶¶^^^^^¶¶^¶¶^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^^¶¶^^¶¶^^^^¶¶^^¶¶^^^^^¶¶^^^^¶¶^¶¶^^^^^^¶¶^^¶¶^^^^¶¶^^¶¶¶¶¶^^¶¶¶¶¶¶¶¶¶^^^^^^¶¶^^^^^^^¶¶^¶¶^^^^^^¶¶^^^¶¶^^¶¶^^¶¶¶¶¶¶¶¶¶^¶¶^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^¶¶^^¶¶^^^^^¶¶^^¶¶^^^^¶¶^^^^^¶¶^¶¶^^^^^^¶¶^¶¶^^^^^¶¶^^^^¶¶¶¶^¶¶^^^^^^^^^^^^^¶¶^^^^^^¶¶^^¶¶^^^^^^¶¶^^^¶¶^^¶¶^^¶¶^^^^^^^^¶¶^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^^^¶¶^^¶¶^^^^^¶¶^^¶¶^^^^¶¶^^^^^¶¶^¶¶^^^^^^¶¶^¶¶^^^^^¶¶^^^^^^¶¶^¶¶^^^^^^^^^^^^^¶¶^^^^^^¶¶^^¶¶^^^^^^¶¶^^^^¶¶¶¶^^^¶¶^^^^^^^^¶¶^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶^^^^¶¶¶^^^¶¶^^^^¶¶¶^^¶¶^^^^¶¶^^^^¶¶¶^¶¶^^^^^¶¶^^¶¶^^^^¶¶¶^¶^^^^¶¶^^¶¶^^^^¶¶^^^^^^¶¶^^^^¶¶¶^^^¶¶^^^^^^¶¶^^^^¶¶¶¶^^^^¶¶^^^^¶¶^¶¶^^^^^^^^^^^^^^$
'$^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^¶¶¶¶¶^¶¶^^^¶¶¶¶^^¶¶¶¶¶^¶¶^¶¶¶¶¶¶¶¶^^^^¶¶¶¶¶^¶¶^^¶¶¶¶¶^^^^¶¶¶¶¶¶^^^^^^^¶¶¶¶¶¶¶^^^^^¶¶^^^^^^¶¶^^^^^¶¶^^^^^^¶¶¶¶¶¶^^¶¶^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶¶¶¶^^^^^^^^^^^^^^^^¶¶¶¶^¶¶^^^^^^^^^^^^^^^^¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^¶¶^¶¶^^¶¶^^^^^^^^^^¶¶^¶¶^¶¶¶¶¶^¶¶¶^^^^^^^^¶¶^^¶¶¶¶^^^¶¶¶¶^^¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶¶¶¶^^¶¶^^¶¶^^^^^^^^^^¶¶^¶¶^¶¶^^¶¶^^¶¶^^^^^^^¶¶^¶¶^^¶¶^¶¶^^^^¶¶^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^¶¶^^¶¶¶¶^^^^^^^^^^^¶¶^¶¶^¶¶^^¶¶^^¶¶^^^^^^^¶¶^¶¶^^¶¶^¶¶¶¶^^¶¶¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^¶¶^^¶¶¶¶^^^^^^^^^^^¶¶^¶¶^¶¶^^¶¶^^¶¶^^^^^^^¶¶^¶¶^^¶¶^^¶¶¶¶^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^¶¶^^^¶¶^^^^^^^¶¶^¶¶^¶¶^^¶¶^^¶¶^^^^^^^¶¶^¶¶^^¶¶^^^^¶¶^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶¶¶¶^^^^¶¶^^^¶¶^^^^¶¶¶¶^^¶¶^¶¶^^¶¶^^¶¶^^^^¶¶¶¶^^^¶¶¶¶^^¶¶¶¶^^^¶¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Source Code   : DataBase Driver
' Auther        : Jim Jose
' eMail         : jimjosev33@yahoo.com
' Purpose       : A Simple Text DataBase Driver
' Status        : Not yet completed!
' Copyright Jim Jose, Gtech Creations - 2005
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' How to Create a New text DataBase:
'
'   Creating a new database is realy simple. We are using a txt file to
' store data. All you want to do is save a *.txt file with the database
' parameters.
' ie: If you want a Table with two Fields then save the file with
'
'               40,20>
'
' Where, 40 - Length of first field
'        20 - Length of second field
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------

'The length used to store the DataBase properties
Private m_IDLength As Long

'The FilePath
Private m_FileName As String

'The Fields: The Values are loaded to this
Private m_Fields() As String

'The length of each field
Private m_FieldLen() As String

'The length of entire record
Private m_RecordLength As Long

'The AbsolutePosition parameter
Private m_RecordPosition As Long

'The Events
Public Event BeforeSave()
Public Event AfterDelete()
Public Event BeforeUpdate()
Public Event RecordChanged()
Public Event FileConnected()
Public Event FieldChanged(ByVal Index As Long)

'-------------------------------------------------------------------------
' Procedure  : OpenFile
' Auther     : Jim Jose
' Input      : DataBase FilePath
' OutPut     : None
' Purpose    : To Open the database file and prepare the reading Parameters
'-------------------------------------------------------------------------

Public Function OpenFile(FileName As String)
'The variables
Dim x           As Long
Dim Bit(1000)   As Byte
Dim m_DataID    As String
    
    'Close the previous connection if any
    If State = True Then CloseConnection
    
    'Opening the file
    'Note that no data is retrived now
    Open FileName For Binary As #1
    
    'Getting the DataBase Parameters
    Get #1, 1, Bit()
    
    'The Parameter string
    m_DataID = StrConv(Bit, vbUnicode)
    m_DataID = Split(m_DataID, ">")(0)
    
    'Getting the field lengths
    m_FieldLen = Split(m_DataID, ",")
    
    'Some preperations
    m_FileName = FileName
    m_RecordLength = 0
    ReDim m_Fields(FieldCount - 1)
    m_IDLength = Len(m_DataID) + 2
    
    'Calculating the total record length
    'Just the sum of field lengths
    For x = 0 To FieldCount - 1
        m_RecordLength = m_RecordLength + Val(m_FieldLen(x))
    Next x

    'Call Him
    RaiseEvent FileConnected
    
    'Inform Me
    Debug.Print vbCrLf & "File Opened"
    Debug.Print "Field Count " & FieldCount
    Debug.Print "Field Lengths " & m_DataID
    Debug.Print "Record Length " & m_RecordLength
    Debug.Print "Record Count " & RecordCount
End Function

'-------------------------------------------------------------------------
' Procedure  : FieldCount
' Auther     : Jim Jose
' Input      : None
' OutPut     : The number of fields
' Purpose    : To get the number of fields in the DataBase
'-------------------------------------------------------------------------

Public Function FieldCount() As Long
On Error GoTo Handle
    'Getting The UBound
    FieldCount = UBound(m_FieldLen) + 1
Exit Function
Handle:
'If error then no fields
    FieldCount = 0
End Function

'-------------------------------------------------------------------------
' Procedure  : Field
' Auther     : Jim Jose
' Input      : Field Index
' OutPut     : The field data
' Purpose    : To get the field data of current position
'-------------------------------------------------------------------------

Public Property Get Field(ByVal Index As Long) As String
    Field = m_Fields(Index)
End Property

Public Property Let Field(ByVal Index As Long, fData As String)
    m_Fields(Index) = fData
    
    'Call Him
    RaiseEvent FieldChanged(Index)
End Property

'-------------------------------------------------------------------------
' Procedure  : Save
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Save the current fields as a new Record
'-------------------------------------------------------------------------

Public Sub Save()
'The Variables
Dim x           As Long
Dim Bit()       As Byte
Dim tmpStr      As String
Dim StartPos    As Long
    
    'Call Him
    RaiseEvent BeforeSave
    
    'Check if the currently entered datasize is in the alloted limit
    If CheckSize = False Then Err.Raise 11, , "The field data is longer than the alloted space.": Exit Sub
    
    'Clearing the file space
    ReDim Bit(m_RecordLength - 1) As Byte
    StartPos = m_IDLength + RecordCount * m_RecordLength
    Put #1, StartPos, Bit

    'Putting the data One by One
    For x = 0 To FieldCount - 1
        tmpStr = m_Fields(x)
        Put #1, StartPos, tmpStr
        StartPos = StartPos + Val(m_FieldLen(x))
    Next x
    
    'Close the connection such that the file is saved
    Close #1
    'Open againe we need it
    Open m_FileName For Binary As #1
    
    'Inform Me
    'debug.Print "Data Saved!"
End Sub

'-------------------------------------------------------------------------
' Procedure  : Update
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To update the current fields
'-------------------------------------------------------------------------

Public Sub Update()
'The variables
Dim x           As Long
Dim Bit()       As Byte
Dim tmpStr      As String
Dim StartPos    As Long

    'Call Him
    RaiseEvent BeforeUpdate
    
    'Check if the currently entered datasize is in the alloted limit
    If CheckSize = False Then Err.Raise 11, , "The field data is longer than the alloted space.": Exit Sub
    
    'Clearing the datspace
    StartPos = m_IDLength + m_RecordPosition * m_RecordLength
    ReDim Bit(m_RecordLength - 1) As Byte
    Put #1, StartPos, Bit
    
    'Putting the data one by one
    For x = 0 To FieldCount - 1
        tmpStr = m_Fields(x)
        Put #1, StartPos, tmpStr
        StartPos = StartPos + Val(m_FieldLen(x))
    Next x
    
    'Again Close it and open it
    Close #1
    Open m_FileName For Binary As #1
    
    'Please inform Me
    'debug.Print "Data Updated!"
End Sub

'-------------------------------------------------------------------------
' Procedure  : Delete
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To delete the current record
'-------------------------------------------------------------------------

Public Sub Delete()
'The variables
Dim Bit1()      As Byte
Dim Bit2()      As Byte
Dim StartPos    As Long

    'This is realy a big task to delete a record
    
    'Get the records upto the current record
    StartPos = m_IDLength + m_RecordPosition * m_RecordLength
    ReDim Bit1(StartPos) As Byte
    Get #1, 1, Bit1
    
    'Get the records after
    ReDim Bit2((RecordCount - m_RecordPosition - 1) * m_RecordLength) As Byte
    Get #1, StartPos + m_RecordLength, Bit2
    Close #1
    
    'Kill the file
    Kill m_FileName
   
    'ReBuilt it without the current record
    Open m_FileName For Binary As #1
    Put #1, 1, Bit1
    Put #1, StartPos, Bit2
    
    'Again Close it, Open it
    Close #1
    Open m_FileName For Binary As #1
    
    'Call Him
    RaiseEvent AfterDelete
    
    'So inform Me Again
    'debug.Print "Data Updated!"
End Sub

'-------------------------------------------------------------------------
' Procedure  : Delete
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To delete the current record
'-------------------------------------------------------------------------

Public Sub DeleteAll()
    'Kill the file
    Close #1
    Kill m_FileName
   
    'ReBuilt it without the current record
    Open m_FileName For Binary As #1
    Put #1, 1, Join(m_FieldLen, ",") & ">"
    
    Close #1
    Open m_FileName For Binary As #1
    
    'Call Him
    RaiseEvent AfterDelete
End Sub

'-------------------------------------------------------------------------
' Procedure  : MoveNext
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To move to the next record
'-------------------------------------------------------------------------

Public Sub MoveNext()
    m_RecordPosition = m_RecordPosition + 1
    LoadValues
End Sub

'-------------------------------------------------------------------------
' Procedure  : MovePrevious
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To move to the previous record
'-------------------------------------------------------------------------

Public Sub MovePrevious()
    m_RecordPosition = m_RecordPosition - 1
    LoadValues
End Sub

'-------------------------------------------------------------------------
' Procedure  : MoveFirst
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To move to the first record
'-------------------------------------------------------------------------

Public Sub MoveFirst()
    m_RecordPosition = 0
    LoadValues
End Sub

'-------------------------------------------------------------------------
' Procedure  : MoveFirst
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To move to the Last record
'-------------------------------------------------------------------------

Public Sub MoveLast()
    m_RecordPosition = RecordCount - 1
    LoadValues
End Sub

'-------------------------------------------------------------------------
' Procedure  : LoadValues
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To move to the first record
'-------------------------------------------------------------------------

Private Sub LoadValues()
'The variables
Dim x           As Long
Dim Bit()       As Byte
Dim StartPos    As Long
Dim txtData     As String
    
    'Inform Me
    'debug.Print "Record Position " & m_RecordPosition
    If m_RecordPosition < 0 Then Err.Raise 12, , "Back Of File [BOF]"
    If m_RecordPosition > RecordCount - 1 Then Err.Raise 13, , "End Of File [EOF]"
    
    'Getting the whole record
    StartPos = m_IDLength + m_RecordPosition * m_RecordLength
    ReDim Bit(m_RecordLength) As Byte
    Get #1, StartPos, Bit()
    
    'Convert it to TEXT
    txtData = StrConv(Bit, vbUnicode)
    
    'Load the field Data One By One
    StartPos = 1
    For x = 0 To FieldCount - 1
        m_Fields(x) = Mid(txtData, StartPos, m_FieldLen(x))
        StartPos = StartPos + m_FieldLen(x)
    Next x
    
    RaiseEvent RecordChanged
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : FindText
' Auther     : Jim Jose
' Input      : The Text to find ,Satrt Position ,Methods
' OutPut     : The RecordPosition
' Purpose    : To find a text in the database
'-------------------------------------------------------------------------

Public Function FindText(FindWhat As String, Optional ByVal Start As Long = 0, _
                            Optional Field As Long = -1, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Long
'The Vatiables
Dim x           As Long
Dim Bit()       As Byte
Dim txtPos      As Long
Dim txtData     As String
Dim FieldPos    As Long
Dim StartPos    As Long
Dim TotLength   As Long
    
    'Inform Me
    'debug.Print "Record Position " & m_RecordPosition
    
    'Check
    If Field >= FieldCount Then Err.Raise 11, , "Invalid field index"
    TotLength = RecordCount * m_RecordLength
    
    'Get the field position
    If Field = -1 Then GoTo Skip
    For x = 0 To Field - 1
        FieldPos = FieldPos + m_FieldLen(x)
    Next x
    
Skip:
    
    'Determining the Bit Size
    StartPos = m_IDLength + Start * m_RecordLength + FieldPos
    If Field = -1 Then
        ReDim Bit(m_RecordLength) As Byte
    Else
        ReDim Bit(m_FieldLen(Field)) As Byte
    End If

FindNow:
    'If the search passout the File then stope
    If StartPos >= TotLength Then FindText = -1: Exit Function
    
    'Get the Text
    Get #1, StartPos, Bit()
    txtData = StrConv(Bit, vbUnicode)
    
    'Check if there is...
    txtPos = InStr(1, txtData, FindWhat, Compare)
    
    'If not do again
    If txtPos = 0 Then StartPos = StartPos + m_RecordLength: GoTo FindNow
    
    'Wow We got it!
    FindText = Int(StartPos / m_RecordLength)
    
End Function

'-------------------------------------------------------------------------
' Procedure  : RecordCount
' Auther     : Jim Jose
' Input      : None
' OutPut     : The total number of records
' Purpose    : to get the total number of records
'-------------------------------------------------------------------------

Public Function RecordCount() As Long
    'Get it from the file Byte length
    RecordCount = (FileLen(m_FileName) - m_IDLength) / m_RecordLength
    If RecordCount < 0 Then RecordCount = 0
End Function

'-------------------------------------------------------------------------
' Procedure  : AbsolutePosition
' Auther     : Jim Jose
' Input      : None
' OutPut     : New Position
' Purpose    : To move the record as you wish
'-------------------------------------------------------------------------

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = m_RecordPosition
End Property

Public Property Let AbsolutePosition(ByVal vNewValue As Long)
    m_RecordPosition = vNewValue
    LoadValues
End Property


'-------------------------------------------------------------------------
' Procedure  : CloseConnection
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To close the current connection,clearup
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    CloseConnection
End Sub

Public Sub CloseConnection()
    Erase m_FieldLen
    Erase m_Fields
    Close #1
End Sub


'-------------------------------------------------------------------------
' Procedure  : CheckSize
' Auther     : Jim Jose
' Input      : None
' OutPut     : Yes or No
' Purpose    : To check if the current request is in the alloted limit
'-------------------------------------------------------------------------

Private Function CheckSize() As Boolean
Dim x As Long
For x = 0 To FieldCount - 1
    If Len(m_Fields(x)) > m_FieldLen(x) Then CheckSize = False: Exit Function
Next x
CheckSize = True
End Function

'-------------------------------------------------------------------------
' Procedure  : State
' Auther     : Jim Jose
' Input      : None
' OutPut     : Yes or No
' Purpose    : To get the class state ,Opened?
'-------------------------------------------------------------------------

Public Function State() As Boolean
    If FieldCount = 0 Then State = False Else State = True
End Function
