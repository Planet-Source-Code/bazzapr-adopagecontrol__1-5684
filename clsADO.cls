VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Rem 'Getting started:
Rem '
Rem 'Compile this into a DLL.
Rem '
Rem 'You can use this object in ASP too.
Rem '
Rem 'VB Example (Add the reference ADOPageControl first):
Rem 'Dim ADOTest As ADOPageControl
Rem '
Rem 'ADOTest.ConnectString = "dsn=mydb; uid=myuid; pwd=mypwd"
Rem 'ADOTest.PageSize = 10
Rem 'ADOTest.ShowPage = 1
Rem 'ADOTest.Sql = "SELECT * FROM MyTable"
Rem '
Rem 'If ADOTest.Execute = False Then
Rem '    Debug.Print ADOTest.ErrorC
Rem '    Exit Sub
Rem 'End If
Rem '
Rem 'Debug.Print ADOTest.PageCount
Rem '
Rem 'Do Until ADOTest.Rs.EOF
Rem '    Debug.Print ADOTest.Rs("MyField1")
Rem '    ADOTest.Rs.MoveNext
Rem 'Loop
Rem '

Option Explicit

Private Conn            As New ADODB.Connection
Private Rs1             As ADODB.Recordset
Private Field           As ADODB.Field

Private objRDO          As New RDSServer.DataFactory

Private m_PageCount     As Long
Private m_ErrorC        As String
Private NewADORSet()

Public ConnectString    As String
Public Sql              As String
Public ShowPage         As Long
Public PageSize         As Long
Public Rs               As ADODB.Recordset



Rem '
Rem 'This function Connects to the database
Rem 'using the ConnectString passed to the Object.
Rem '
Private Function Connect() As Boolean

    Connect = False
    
    On Error GoTo ErrorHandler
    
    Rem '
    Rem 'Did we recieve a ConnectString?
    Rem '
    If ConnectString = vbNullString Then
        ErrorC = "A connectstring is required!"
        Exit Function
    End If

    Rem '
    Rem 'Did we recieve an Sql string?
    Rem '
    If Sql = vbNullString Then
        ErrorC = "An SQL statement is required!"
        Exit Function
    End If
    
    
    Rem '
    Rem 'Connect!
    Rem '
    Set Rs1 = New ADODB.Recordset
    Rs1.CursorLocation = adUseClient
    Rs1.Open Sql, ConnectString, adOpenStatic, adLockReadOnly, adCmdText

    Connect = True

Exit Function

ErrorHandler:
    ErrorC = "[Connect] " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Function



Rem '
Rem 'This function populates an array with
Rem 'the tables field names.
Rem '
Private Function CreateNewCol() As Boolean
Dim NewCol(3)
Dim FieldCount As Long
Dim RecCount As Long

    On Error GoTo ErrorHandler
    
    CreateNewCol = False

    FieldCount = Rs1.Fields.Count

    Rem '
    Rem 'Redim the array we are going to use
    Rem 'to the number of Columns in the table.
    Rem '
    ReDim NewADORSet(FieldCount)
    
    Rem '
    Rem 'Because an array starts at 0 we need
    Rem 'do a loop with one less than the actual
    Rem 'FieldCount. Notice we did a redim using
    Rem 'the full number found in FieldCount. This
    Rem 'is done because we will later in this
    Rem 'function add one more element to the
    Rem 'array thusly eliminating the need to redim
    Rem 'again. Have a look just after this "For Next"
    Rem '
    For RecCount = 0 To FieldCount - 1

        Set Field = Rs1.Fields(RecCount)
                
        NewCol(0) = Field.Name          'The Field Name
        NewCol(1) = CInt(Field.Type)    'Field Type
        
        
        Rem '
        Rem 'The field size must be determined, but because
        Rem 'the field size is only required with certain
        Rem 'data types, we use this select case to determine
        Rem 'the data type and then to act accordingly.
        Rem '
        Select Case Field.Type
        
            Case adBSTR, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBinary, adVarBinary, adLongVarBinary
                NewCol(2) = CInt(Field.DefinedSize)     'Field Size
            Case Else
                NewCol(2) = CInt(-1)                    'Field Size
        
        End Select
        
        NewCol(3) = True    'Allow nulls for this field - True / False
        
        NewADORSet(RecCount) = NewCol
        
    Next RecCount
    
       
    Rem '
    Rem 'I added this for multiple reasons when
    Rem 'any of those reasons come to me I will
    Rem 'ad them :-).
    Rem '
    NewCol(0) = "End"
    NewCol(1) = CInt(adInteger)
    NewCol(2) = CInt(-1)
    NewCol(3) = True
    NewADORSet(FieldCount) = NewCol
    
    CreateNewCol = True
    
Exit Function

ErrorHandler:
    ErrorC = "[CreateNewCol] " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Function



Rem '
Rem 'Now we need to copy the queried database into
Rem 'the Dummy database.
Rem '
Private Function DoCopy(DummyRS As ADODB.Recordset) As Boolean
Dim RecCount As Long
Dim FieldCount As Long

    DoCopy = False

    On Error GoTo ErrorHandler

    Rem '
    Rem 'Before we continue! Is there any records?
    Rem '
    If Rs1.RecordCount = 0 Then
        ErrorC = "No Records"
        Exit Function
    End If
    
    Rem '
    Rem 'Set the PaperSize etc.
    Rem '
    Rs1.PageSize = PageSize
    PageCount = Rs1.PageCount
    
    Rem '
    Rem 'Set the page to the page we need to see IF
    Rem 'it is not larger than the number of pages.
    Rem '
    If ShowPage > PageCount Then
        ErrorC = "The ShowPage is larger than the amount of pages! [PageCount=" & PageCount
        Exit Function
    End If
    Rs1.AbsolutePage = ShowPage


    Rem '
    Rem 'Loop through the current table
    Rem 'and place new records into the
    Rem 'dummy table.
    Rem '
    Do Until Rs1.EOF Or RecCount = Rs1.PageSize
    
        RecCount = RecCount + 1
        DummyRS.AddNew
        
        For FieldCount = 0 To Rs1.Fields.Count - 1
            DummyRS(FieldCount).Value = Rs1(FieldCount).Value
        Next
      
        Rs1.MoveNext
        
    Loop

    Rem '
    Rem 'Place the dummy recordsets data into the recordset
    Rem 'made available to the user.
    Rem '
    Set Rs = DummyRS

    Rem '
    Rem 'As not to annoy the poor user we
    Rem 'must not forget to move to the
    Rem 'First Record!!!! Else he wont find
    Rem 'any records because he is already
    Rem 'at the last one.
    Rem '
    Rs.MoveFirst
    
    DoCopy = True

Exit Function

ErrorHandler:
    ErrorC = "[DoCopy] " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Function



Rem '
Rem 'Once the user has supplied the data he executes
Rem 'this function. This function is the initial User Interface.
Rem '
Public Function Execute() As Boolean

    On Error GoTo ErrorHandler

    Execute = False
        
    Rem '
    Rem 'Call the Connect function. If it
    Rem 'succeeds then continue.
    Rem '
    If Not Connect Then
        Exit Function
    End If
    
    Rem '
    Rem 'If the user supplied us with a ShowPage then
    Rem 'do what needs to be done to return the data
    Rem 'in the page format requested
    Rem '
    Rem 'Else
    Rem '
    Rem 'Pass back the whole Query.
    Rem '
    If ShowPage > 0 Then
        
        Rem '
        Rem 'Is there a PageSize? No, then
        Rem 'supply one.
        Rem '
        If PageSize = 0 Then
            PageSize = 20
        End If
        
        Rem '
        Rem 'Execute the CreateNewCol function. If it
        Rem 'succeeds then continue.
        Rem '
        If Not CreateNewCol Then
            Exit Function
        End If
        
        Set objRDO = CreateObject("RDSServer.DataFactory")
        
        Rem '
        Rem 'Take the NewADORSet array and convert it to
        Rem 'a RecordSet. Then call the DoCopy function
        Rem 'passing the newly created Recordset to be
        Rem 'populated with data.
        Rem 'If DoCopy succeeds then continue.
        Rem '
        If Not DoCopy(objRDO.CreateRecordSet(NewADORSet)) Then
            Set objRDO = Nothing
            Exit Function
        End If
        
        Set objRDO = Nothing
        
    Else
        Rem '
        Rem 'Set the Rs1 to the Recordset the user
        Rem 'has access to.
        Rem '
        Set Rs = Rs1
    End If
  

    Execute = True

Exit Function
    
ErrorHandler:
    ErrorC = "[Execute] " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Function

Public Property Get ErrorC() As String
    ErrorC = m_ErrorC
End Property


Private Property Let ErrorC(ByVal vNewErrorC As String)
    m_ErrorC = vNewErrorC
End Property


Public Property Get PageCount() As Long
    PageCount = m_PageCount
End Property


Private Property Let PageCount(ByVal vNewPageCount As Long)
    m_PageCount = vNewPageCount
End Property

