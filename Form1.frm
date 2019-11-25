VERSION 5.00
Begin VB.Form FormBailBondsInventory 
   BorderStyle     =   0  'None
   Caption         =   "Power Inventory Application"
   ClientHeight    =   4500
   ClientLeft      =   4530
   ClientTop       =   1680
   ClientWidth     =   5430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   4212
      Left            =   240
      TabIndex        =   21
      Top             =   0
      Width           =   4932
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton Cmd_Reciept 
         Caption         =   "&Receipt"
         Height          =   250
         Left            =   3360
         TabIndex        =   11
         Top             =   2040
         Width           =   1000
      End
      Begin VB.TextBox TxtCreateDate 
         Enabled         =   0   'False
         Height          =   288
         Left            =   3360
         TabIndex        =   4
         Top             =   480
         Width           =   1212
      End
      Begin VB.CommandButton Cmd_Clear 
         Caption         =   "&Clear"
         Height          =   250
         Left            =   1560
         TabIndex        =   18
         Top             =   3720
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_Save 
         Caption         =   "&Save"
         Height          =   250
         Left            =   2640
         TabIndex        =   19
         Top             =   3720
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_Delete 
         Caption         =   "&Delete"
         Height          =   250
         Left            =   3720
         TabIndex        =   20
         Top             =   3720
         Width           =   1000
      End
      Begin VB.TextBox Txt_PowerAmount 
         Height          =   288
         Left            =   1680
         TabIndex        =   15
         Top             =   3120
         Width           =   1572
      End
      Begin VB.TextBox Txt_PowerNumber 
         Height          =   288
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox Txt_AssignmentDate 
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   3360
         TabIndex        =   10
         Top             =   1680
         Width           =   1212
      End
      Begin VB.TextBox Txt_ExpirationDate 
         Height          =   288
         Left            =   3480
         TabIndex        =   17
         Top             =   3120
         Width           =   1212
      End
      Begin VB.ComboBox Cmb_Surety 
         DataField       =   "Name"
         Height          =   288
         ItemData        =   "Form1.frx":0442
         Left            =   240
         List            =   "Form1.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox Cmb_Agents 
         DataField       =   "Agent_Lname"
         DataSource      =   "Adodc2"
         Height          =   288
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton Cmd_Find 
         Caption         =   "&Find"
         Height          =   250
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1000
      End
      Begin VB.ComboBox Cmb_PowerSymbol 
         Enabled         =   0   'False
         Height          =   288
         ItemData        =   "Form1.frx":0446
         Left            =   240
         List            =   "Form1.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Created Date"
         Height          =   252
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "&Surety"
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1932
      End
      Begin VB.Label Label5 
         Caption         =   "&Agent"
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1452
      End
      Begin VB.Label Label6 
         Caption         =   "P&ower Amount"
         Height          =   252
         Left            =   1680
         TabIndex        =   14
         Top             =   2880
         Width           =   1092
      End
      Begin VB.Label Label7 
         Caption         =   "&Power Symbol"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   1452
      End
      Begin VB.Label Label8 
         Caption         =   "&Power Number"
         Height          =   252
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label9 
         Caption         =   "Ass&igned Date"
         Height          =   252
         Left            =   3360
         TabIndex        =   9
         Top             =   1440
         Width           =   1452
      End
      Begin VB.Label Label10 
         Caption         =   "&Expiration Date"
         Height          =   252
         Left            =   3480
         TabIndex        =   16
         Top             =   2880
         Width           =   1212
      End
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu MnuAddAssign 
         Caption         =   "Add/Assign Powers"
         Shortcut        =   ^A
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FormBailBondsInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Global recorsets and connections
Dim sConnString As String
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As ADODB.Recordset
'--------------------------
'Global variables
Dim How_many, from_pow As Double
'For dynamic creation of error
Dim BuildError As String
Dim BigError As Boolean
'variable for search and find
Dim State As Boolean
Option Explicit

Private Sub Load_Agents()
On Error GoTo MyErrHandler
        'Connection string
        sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
        conn.Open sConnString
        
        'Set the connection point to DB
        Set cmd.ActiveConnection = conn
        
        'Make SQL statement to get fields
        cmd.CommandText = "SELECT Agent_Lname, Agent_Fname, Id  FROM Agents "
        
        'Set command type
        cmd.CommandType = adCmdText
        
        'Execute SQL
        Set rs = cmd.Execute
    
        'See if connection is closed if so error occured when connecting
        If (conn.State = adStateClosed) Then
                MsgBox "Unable to Connect Database", vbCritical, "Connection Error!!!"
                End
        End If
        
        'Move to first record
        rs.MoveFirst
        
        'Loop through and add records to the combobox
        Cmb_Agents.AddItem "<-- No Selection -->", 0
        Cmb_Agents.ItemData(Cmb_Agents.NewIndex) = 0
        Do While (Not rs.EOF)
                Cmb_Agents.AddItem rs.Fields("Agent_Lname").Value '& ", " & rs.Fields("Agent_Id").value
                'Relate the id from agent table to the comboboxes index
                Cmb_Agents.ItemData(Cmb_Agents.NewIndex) = rs.Fields("id").Value
                rs.MoveNext
        Loop
        'Clean up connection
        rs.Close
        conn.Close
        Set rs = Nothing
        Set cmd = Nothing
        Set conn = Nothing
        Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub
Private Sub Load_Surety()
On Error GoTo MyErrHandler
        'Same as above except different SQL statement
        sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
        conn.Open sConnString
    
        Set cmd.ActiveConnection = conn
    
        cmd.CommandText = "SELECT Name,Id FROM Surety"
        cmd.CommandType = adCmdText
        Set rs = cmd.Execute
        
        
        If (conn.State = adStateClosed) Then
                MsgBox "Unable to Connect Database", vbCritical, "Connection Error!!!"
                End
        End If
      
        rs.MoveFirst
            
            Do While Not (rs.EOF)
                    Cmb_Surety.AddItem rs.Fields("Name").Value '& ", " & rs.Fields("Surety_Id").value
                    Cmb_Surety.ItemData(Cmb_Surety.NewIndex) = rs.Fields("Id").Value
                    rs.MoveNext
            Loop
      
        Set rs = Nothing
        Set cmd = Nothing
        conn.Close
        Set conn = Nothing
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub
Private Sub Load_Symbols()
 On Error GoTo MyErrHandler
        'Same as above except different SQL statement
        sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
        conn.Open sConnString
    
        Set cmd.ActiveConnection = conn
    
        cmd.CommandText = "SELECT * FROM Surety"
        cmd.CommandType = adCmdText
        Set rs = cmd.Execute
    
        If (conn.State = adStateClosed) Then
            MsgBox "Unable to Connect Database", vbCritical, "Connection Error!!!"
            End
        End If
      
        
        
        rs.MoveFirst
        
        Do While Not (rs.EOF)
            Cmb_PowerSymbol.AddItem rs.Fields("Symbol").Value '& ", " & rs.Fields("Surety_Id").value
            Cmb_PowerSymbol.ItemData(Cmb_PowerSymbol.NewIndex) = rs.Fields("Id").Value
            rs.MoveNext
        Loop
        
        Set rs = Nothing
        Set cmd = Nothing
        conn.Close
        Set conn = Nothing
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Cmb_Agents_Click()
    
       'If Cmb_Agents.ItemData(Cmb_Agents.ListIndex) > 0 Then
       
          'Txt_AssignmentDate.Text = Format(Now, "MM/DD/YYYY")
       
       'End If

End Sub

Private Sub Cmb_Surety_Click()
On Error GoTo MyErrHandler
Dim i As Integer
        For i = 0 To Cmb_PowerSymbol.ListCount - 1
        
           If Cmb_PowerSymbol.ItemData(i) = Cmb_Surety.ItemData(Cmb_Surety.ListIndex) Then
           
              Cmb_PowerSymbol.ListIndex = i
              Exit For
           End If
        Next i
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

'Private Sub Cmd_Add_Click()
'   'Connection variables and counters
'   Dim cnn1 As ADODB.Connection
'   Dim rst As ADODB.Recordset
'   Dim strCnn As String
'   Dim count As Integer
'   Dim strSQL As String
'
'   Txt_AssignmentDate.Text = ""
'
'       'Set connection
'       Set cnn1 = New ADODB.Connection
'       strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
'       cnn1.Open strCnn
'
'       'Make recordset object and set params so we can add records to DB
'       Set rst = New ADODB.Recordset
'       rst.CursorType = adOpenKeyset
'       rst.LockType = adLockOptimistic
'       rst.Open "Power", cnn1, , , adCmdTable
'
'       'Check form to see if values filled in and they are correct
'        If (Check_Form_Fields("CmdAdd")) Then
'                from_pow = Txt_FromPower.Text
'
'                'Add generic powers to db
'                For count = 0 To How_many - 1
'                        'Check to see if power is already being used if and skip over
'                        Set rst = New ADODB.Recordset
'                        strSQL = "SELECT * FROM Power WHERE Power_Num = " & from_pow & " And surety_id = " & Cmb_Surety.ItemData(Cmb_Surety.ListIndex)
'                        rst.Open strSQL, cnn1, adOpenKeyset, adLockOptimistic, adCmdText
'                        'Check to see if power already in use
'                        If rst.RecordCount > 0 Then
'                           MsgBox "Power Number - " & from_pow & " has already been used.", vbInformation
'                           Exit For
'                        End If
'
'                        rst.AddNew
'                        rst!Power_num = from_pow
'                        rst.Fields("Symbol").Value = Cmb_PowerSymbol.Text
'                        rst!Execution_Date = Date
'                        rst!Create_Date = Date
'                        rst!Defendant_Name = "NA"
'                        rst!PowerAmount = Txt_PowerAmount.Text
'                        rst!Court_Date = Date
'
'                        rst!Court_Loc = "NA"
'                        rst!Expiration_Date = Txt_ExpirationDate.Text
'                        'Double check not necessary since the errorcheck already
'                        'Should have caught it
'                        If IsDate(Txt_AssignmentDate.Text) Then
'                           rst!Assign_Date = Txt_AssignmentDate.Text
'                        End If
'                        rst!Bond_Amount = 0
'                        rst!Surety_Id = Cmb_Surety.ItemData(Cmb_Surety.ListIndex)
'
'                        If (Cmb_Agents.ListIndex >= 0) Then
'                            rst!Agent_Id = Cmb_Agents.ItemData(Cmb_Agents.ListIndex)
'                        Else
'                            rst!Agent_Id = 0
'                        End If
'                        from_pow = from_pow + 1
'                    rst.Update
'                Next count
'            'Clear controls
'            Cmd_Clear_Click
'            'Clean up
'            rst.Close
'            cnn1.Close
'            Set rst = Nothing
'            Set cnn1 = Nothing
'            MsgBox count & " powers added successfully", vbInformation
'        Else
'            'Show the errors if they occured
'            MsgBox BuildError, vbInformation
'            'Reset buildError so it could be reused
'            BuildError = ""
'            Cmd_Clear.Enabled = True
'        End If
'End Sub

'Private Sub Cmd_Assign_Click()
'Dim FromPow, ToPow As Long
'Dim cnn As ADODB.Connection
'Dim rst As ADODB.Recordset
'Dim strCnn As String
'Dim count As Integer
'Dim strSQL As String
'Dim i As Integer
''To Hold the powers that can not
''be assigned because they don't exist
'Dim NotUpdateAble(1000) As Long
'Dim NotUpdated As String
'Dim NIndex As Integer
''To Hold the updated powers
'Dim Updated(1000) As Long
'Dim UIndex As Integer
'Dim UpdatedStr As String
'
''To hold true if all powers in range
' 'were Updated and there was no pows that
' 'were not updated
'Dim AllUpdated As Boolean
'Dim Informstr As String
'
'Dim WhatWay As Boolean
'
'Informstr = ""
'UpdatedStr = ""
'NotUpdated = ""
'BuildError = ""
'
'If (IsNumeric(Txt_FromPower.Text) And IsNumeric(Txt_ToPower.Text) And Cmb_Agents.ListIndex >= 0 And Cmb_Surety.ListIndex >= 0) Then
'            'Grab from and to power that will be
'            'updated
'            FromPow = CLng(Txt_FromPower.Text)
'            ToPow = CLng(Txt_ToPower.Text)
'
'            'Connection
'            Set cnn = New ADODB.Connection
'            strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
'            cnn.Open strCnn
''----------------------------------------------------------------------------------
'' Working But Modified to reflect powers that cant be changed because are not existent
'            'Open recordset
''            Set rst = New ADODB.Recordset
''            'Query each power
''            Do
''                strSQL = "SELECT Power_Num,Agent_Id,Surety_Id FROM Power WHERE Power_Num = " & FromPow
''                rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
''                'Check if query returned anything
''                If (rst.RecordCount > 0) Then
''                    'if so update
''                    Do While Not (rst.EOF)
''                        rst!Agent_Id = Cmb_Agents.ItemData(Cmb_Agents.ListIndex)
''                        rst!Surety_Id = Cmb_Surety.ItemData(Cmb_Surety.ListIndex)
''                        rst.MoveNext
''                    Loop
''                End If
''                'move to next power
''                FromPow = FromPow + 1
''                rst.Close
''            'keep goin until we have reached topower
''            Loop Until (FromPow = ToPow)
''--------------------------------------------------------------------------------
'            AllUpdated = True
'            NIndex = 0
'            UIndex = 0
'            Set rst = New ADODB.Recordset
'            'Query each power
'            Do
'                strSQL = "SELECT Power_Num,Agent_Id,Surety_Id FROM Power WHERE Power_Num = " & FromPow
'                rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
'                'Check if query returned anything
'                If (rst.RecordCount > 0) Then
'                    'if so update
'                    Do While Not (rst.EOF)
'                        rst!Agent_Id = Cmb_Agents.ItemData(Cmb_Agents.ListIndex)
'                        rst!Surety_Id = Cmb_Surety.ItemData(Cmb_Surety.ListIndex)
'                        rst.MoveNext
'                    Loop
'                     Updated(UIndex) = FromPow
'                     UIndex = UIndex + 1
'                    'move to next power
'                    FromPow = FromPow + 1
'                    rst.Close
'                Else
'                  'Set to false since we know that one of the powers does not even exist
'                  AllUpdated = False
'                  'Grab the power and stuff it into array and increment index for next
'                  'non existent power
'                  NotUpdateAble(NIndex) = FromPow
'                  'Increment From power
'                  FromPow = FromPow + 1
'                  NIndex = NIndex + 1
'                  rst.Close
'                End If
'            'keep goin until we have reached topower
'            Loop Until (FromPow = ToPow)
'
'            'Clean up
'            cnn.Close
'            Set rst = Nothing
'            Set cnn = Nothing
'
'            If (AllUpdated) Then
'
'                'Need to get the from power again because do loop incremented it and we lost original value
'                FromPow = Txt_FromPower.Text
'                'Show user it was successfull
'                MsgBox "Powers " & FromPow & " to " & ToPow & " were successfully assigned to agent " & _
'                Cmb_Agents.List(Cmb_Agents.ListIndex) & " and surety " & _
'                Cmb_Surety.List(Cmb_Surety.ListIndex), vbInformation
'                Cmd_Clear_Click
'
'            Else
'                'If above failed we know there is at least one power that was not existent
'                'So we generate string that will show the powers that have been
'                'updated as well as the powers that have not been updated
'
''                For i = 0 To NIndex
''                    If (NotUpdateAble(i) <> 0) Then
''                        WhatWay = False
''                    Else
''                        WhatWay = True
''                    End If
''                Next i
''
''                If (WhatWay) Then
'                    For i = 0 To NIndex - 1
'                        NotUpdated = NotUpdated & NotUpdateAble(i) & " "
'                    Next i
'                    Informstr = Informstr & "Powers " & NotUpdated & " did not exist "
'                    'MsgBox Informstr, vbInformation
''                Else
'                    For i = 0 To UIndex - 1
'                       UpdatedStr = UpdatedStr & Updated(i) & " "
'                    Next i
'
'                    Informstr = Informstr & "Powers " & UpdatedStr & " were updated successfully " & vbCrLf
'                    MsgBox Informstr, vbInformation
''                End If
'                Cmd_Clear_Click
'            End If
'Else
'            'Error check to build errors
'            BuildError = "Error -"
'            BuildError = BuildError & vbCrLf
'
'            If (Not IsNumeric(Txt_FromPower.Text) Or IsNull(Txt_FromPower.Text)) Then
'                    BuildError = BuildError & "Need to fill in valid number for from power field" & vbCrLf
'            End If
'            If (Not IsNumeric(Txt_ToPower.Text) Or IsNull(Txt_ToPower.Text)) Then
'                    BuildError = BuildError & "Need to fill in valid number for to power field" & vbCrLf
'            End If
'            If (Cmb_Agents.ListIndex < 0) Then
'                    BuildError = BuildError & "Need to pick an agent to assign to" & vbCrLf
'            End If
'            If (Cmb_Surety.ListIndex < 0) Then
'                    BuildError = BuildError & "Need to pick an surety to assign to" & vbCrLf
'            End If
'                MsgBox BuildError, vbInformation
'            Cmd_Clear_Click
'
'End If
'End Sub

'Clear all controls
Private Sub Cmd_Clear_Click()
On Error GoTo MyErrHandler
        Txt_PowerNumber.Text = ""
        Txt_AssignmentDate.Text = ""
        Txt_ExpirationDate.Text = ""
        Txt_PowerAmount.Text = ""
        TxtCreateDate.Text = ""
        Cmb_Agents.Clear
        Cmb_Surety.Clear
        Cmb_PowerSymbol.Clear
        Load_Agents
        Load_Surety
        Load_Symbols
        Cmd_Delete.Enabled = False
        Cmd_Save.Enabled = False
        Cmd_Clear.Enabled = True
        BuildError = ""
        State = False
        'Check state so we could set the
        'reciept button
        'If (State = False) Then
            Cmd_Reciept.Enabled = True
        'Else
            'Cmd_Reciept.Enabled = True
        'End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Cmd_Delete_Click()
On Error GoTo MyErrHandler
   Dim cnn1 As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim strCnn As String
   Dim last_plus_one As Integer
   Dim choice As Integer
   Dim DeletedPower As Integer
   choice = 0
   
   'Check to see if power number is filled
   If (Txt_PowerNumber.Text = "") Then
            MsgBox "In order to delete you need to specify Power number", vbInformation
   Else
        'Get power to delete
        DeletedPower = Txt_PowerNumber.Text
        
        'Open connection and recordset
        Set cnn1 = New ADODB.Connection
        strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
        cnn1.Open strCnn
   
        Set rst = New ADODB.Recordset
        rst.CursorLocation = adUseClient
        rst.CursorType = adOpenStatic
        rst.LockType = adLockBatchOptimistic
        'Query power
        rst.Open "SELECT * FROM Power WHERE Power_Num = " & Txt_PowerNumber.Text, strCnn, , , adCmdText
            
            'Check to see if the query returned anything
            If (rst.RecordCount < 1) Then
                   MsgBox "Power " & Trim(Txt_PowerNumber) & " does not exist", vbInformation
                   Cmd_Clear_Click
                   Load_Agents
                   Load_Surety
                   Load_Symbols
            Else
                    'We know that it returned something so ask to see if user wants to delete
                    choice = MsgBox("You are about to delete power " & Trim(Txt_PowerNumber) & " are you sure", vbOKCancel)
                    If (choice = vbOK) Then
                        'Delete, update, clear form and close connections
                        rst.Delete
                        rst.UpdateBatch
                        rst.Close
                        cnn1.Close
                        Cmd_Clear_Click
                        MsgBox "Power " & DeletedPower & " deleted successfully", vbInformation
                        State = False
                    Else
                        'Cmd_Clear_Click
                    End If
            End If
   End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Cmd_Find_Click()
On Error GoTo MyErrHandler
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim strCnn As String
    Dim strSQL As String
    Dim i As Integer
    Dim SureId As Integer
    Dim SureName As String
    
    'Check to see if the power number they entered is valid
    If (IsNumeric(Txt_PowerNumber.Text) And Txt_PowerNumber.Text <> "") Then
        
            'Set connection
            Set cnn = New ADODB.Connection
            strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
            cnn.Open strCnn
       
      
            Set rst = New ADODB.Recordset
            strSQL = "SELECT * FROM Power WHERE Power_Num = " & Txt_PowerNumber.Text
            rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText

            'Check what query returned
            If (rst.RecordCount < 1) Then
                MsgBox "Could not find power", vbInformation
                Cmd_Clear.Enabled = False
                Exit Sub
            Else
                'Fill in all the fields on the form
                Txt_PowerNumber.Text = rst!Power_num
                If Not IsNull(rst!Assign_Date) Then
                   Txt_AssignmentDate.Text = rst!Assign_Date
                End If
                Txt_PowerAmount.Text = rst!PowerAmount
                Txt_ExpirationDate.Text = rst!Expiration_Date
                TxtCreateDate.Text = rst!Create_Date
                'Scan combo box and see if the id returned matches itemdata field if so
                'set the selected index to the one returned
                For i = 0 To Cmb_Surety.ListCount - 1
                   If Cmb_Surety.ItemData(i) = rst.Fields("Surety_Id").Value Then
                      Cmb_Surety.ListIndex = i
                      Exit For
                   
                   End If
                Next i
                
                For i = 0 To Cmb_Agents.ListCount - 1
                   If Cmb_Agents.ItemData(i) = rst.Fields("Agent_Id").Value Then
                      Cmb_Agents.ListIndex = i
                      Exit For
                   End If
                Next i
                
                    
                For i = 0 To Cmb_PowerSymbol.ListCount - 1
                   If Cmb_PowerSymbol.List(i) = rst.Fields("Symbol").Value Then
                      Cmb_PowerSymbol.ListIndex = i
                      Exit For
                   End If
                Next i
            End If
            
            'Clean up
            rst.Close
            cnn.Close
            Set rst = Nothing
            Set cnn = Nothing
            
            'Set the state = true since Check_Form_Fields function will use it
            'Later when save button is clicked
            State = True
            
            
            Cmd_Delete.Enabled = True
            Cmd_Clear.Enabled = True
            Cmd_Save.Enabled = True
            If (State = True) Then
                Cmd_Reciept.Enabled = True
            End If
            
    Else
            'If the above test failed issue error message
            MsgBox "Error - Invalid Power Number", vbInformation
            'Set state to false meaning nothing was found so it cant be saved
            State = False
            Txt_PowerNumber.Text = ""
            
            Cmd_Delete.Enabled = False
            Cmd_Clear.Enabled = False
            Cmd_Save.Enabled = False
            Cmd_Clear_Click
            
    End If
Exit Sub
MyErrHandler:
    ' clean up
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    
    If Not cnn Is Nothing Then
        If cnn.State = adStateOpen Then cnn.Close
    End If
    Set cnn = Nothing
    
    If Err <> 0 Then
        MsgBox Err.Source & "-->" & Err.Description, , "Error"
    End If

End Sub

Public Sub Cmd_Reciept_Click()
On Error GoTo MyErrHandler
   Dim cnn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim strCnn As String
   Dim count As Integer
   Dim strSQL As String
   'For call to Check_Reciept
   'which returns data to be parsed
   Dim CheckString As String
   'Array to hold data that is parsed
   Dim ParsedData() As String
   Dim MyDate As String
   Dim MyAgentId As Integer
   
   
   
    'Call error check
    CheckString = Check_Reciept
    If (CheckString = "ERROR") Then
        MsgBox "Error - need to pick agent and a valid date to print reciept", vbInformation
    Else
            'Fill array with string
            ParsedData = Split(CheckString, ",")
            MyDate = ParsedData(0)
            MyDate = FormatDateTime(MyDate)
            MyAgentId = ParsedData(1)
            MyAgentId = CInt(MyAgentId)
        
            
            'MsgBox (MyDate)
            'MsgBox (MyAgentId)
            
            'MsgBox (MyAgentLname)
            
            'Createe connection
            Set cnn = New ADODB.Connection
            strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
            cnn.Open strCnn
       
            'Recorsset
            Set rst = New ADODB.Recordset
            strSQL = "SELECT Power.Assign_Date, Power.Agent_Id, Power.Power_Num From Power WHERE (((Power.Assign_Date)= #" & MyDate & "#) AND ((Power.Agent_Id)=" & MyAgentId & "))"
            rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
            
            
                If rst.RecordCount < 1 Then
                    MsgBox "Agent " & Cmb_Agents.List(Cmb_Agents.ListIndex) & " does not have any powers assigned on date = " & MyDate, vbInformation
                    rst.Close
                    cnn.Close
                    Set rst = Nothing
                    Set cnn = Nothing
                    Exit Sub
                End If
                       
            'Clear everything
            Cmd_Clear.Enabled = True
            
            'Get the agents Last name so we could pass it to the report
            'MyAgentLname = Cmb_Agents.List(Cmb_Agents.ListIndex)
            
            'Call Loadreport to load the report
            Call LoadReport(rst)
            '---->Need to figure out how to clean up connection here
            '---->since if I close connections here report will not be able to be displayed
            
            'Call CloseReport(rst)
    End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Cmd_Save_Click()
On Error GoTo MyErrHandler
   Dim cnn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim strCnn As String
   Dim count As Integer
   Dim strSQL As String
   Dim Ok As Boolean
   
   
   'Call Check_for_fields function with CmdSave so it knows save button issued it
   'if it returns true then we know we could save the data since the error check passed
   If (Check_Form_Fields("CmdSave")) Then
        
            Set cnn = New ADODB.Connection
            strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
            cnn.Open strCnn
       
            'Get all fields for the power that we want to save
            Set rst = New ADODB.Recordset
            strSQL = "SELECT * FROM Power WHERE Power_Num = " & Txt_PowerNumber.Text
            rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText

        
        
            'create select statement for record that i want check to see
            'if it returned anything
            'rst!Power_Num = Txt_PowerNumber.Text
        
            'Update with new values
            rst.Fields("Symbol").Value = Cmb_PowerSymbol.Text
                        
            rst!Execution_Date = Date
            rst!Create_Date = Date
            rst!Defendant_Name = "NA"
            rst!PowerAmount = Txt_PowerAmount.Text
            rst!Court_Date = Date
            rst!Court_Loc = "NA"
            rst!Expiration_Date = Txt_ExpirationDate.Text
            rst!Assign_Date = Txt_AssignmentDate.Text
            'rst!Bond_Amount = 0
                        
        
            rst!Surety_Id = Cmb_Surety.ItemData(Cmb_Surety.ListIndex) '+1
                            
       
            rst!Agent_Id = Cmb_Agents.ItemData(Cmb_Agents.ListIndex)  '+1
                        
            rst.Update
             
            MsgBox "Power # " & rst.Fields("Power_Num").Value & " saved successfully", vbInformation
            'Cmd_Clear_Click
            rst.Close
            cnn.Close
            'Cmd_Clear_Click
    Else
        MsgBox BuildError, vbInformation
        BuildError = ""
    End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo MyErrHandler
    'Load combos
    Load_Agents
    Load_Surety
    Load_Symbols
    Cmd_Save.Enabled = False
    Cmd_Clear.Enabled = True
    Cmd_Delete.Enabled = False
    Cmd_Reciept.Enabled = True
    Cmd_Find.Enabled = True
    'If (State = False) Then
        'Cmd_Reciept.Enabled = False
    'End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Mnu_Exit_Click()
        Unload Me
End Sub

'Private Sub Txt_Howmany_LostFocus()
'    'Check to see if fields are valid once user hits tab
'    If (Txt_FromPower.Text <> "" And IsNumeric(Txt_FromPower.Text) And Txt_Howmany.Text <> "" And IsNumeric(Txt_Howmany.Text)) Then
'            How_many = Txt_Howmany.Text
'            from_pow = Txt_FromPower.Text
'            Txt_ToPower.Text = from_pow + How_many '- 1
'    Else
'            'MsgBox ("Need to fill in valid fields for from power and how many powers to add")
'            Txt_FromPower.Text = ""
'            Txt_Howmany.Text = ""
'            Txt_ToPower.Text = ""
'    End If
'End Sub


Private Function Check_Form_Fields(ButtonIndex As String) As Boolean
On Error GoTo MyErrHandler
BuildError = "Error -"
BuildError = BuildError & vbCrLf
'Variable that will be used
'to send back to calling function on error
BigError = False

Select Case ButtonIndex
    Case "CmdAdd"
    
        Txt_PowerNumber.Text = ""
        
        If (Cmb_Surety.ListIndex < 0) Then
                BuildError = BuildError & "Need to pick a surety" & vbCrLf
                
                If Not BigError Then Cmb_Surety.SetFocus
                BigError = True
        End If
   
        
        If (Cmb_PowerSymbol.ListIndex < 0) Then
                BuildError = BuildError & "Need to pick a symbol" & vbCrLf
                
                If Not BigError Then Cmb_PowerSymbol.SetFocus
                BigError = True
        End If
                
   
        If (Txt_ExpirationDate.Text = "") Then
                BuildError = BuildError & "Need to fill in Expiration Date " & vbCrLf
                BigError = True
                If Not BigError Then Txt_ExpirationDate.SetFocus
        Else
                If (Not IsDate(Txt_ExpirationDate.Text)) Then
                    BuildError = BuildError & "Need to fill in valid Expiration date" & vbCrLf
                    Txt_ExpirationDate.Text = ""
                    If Not BigError Then Txt_ExpirationDate.SetFocus
                    BigError = True
                End If
        End If
   
   
        If (Txt_PowerAmount.Text = "") Then
                BuildError = BuildError & "Need to fill in Power Amount" & vbCrLf
                If Not BigError Then Txt_PowerAmount.SetFocus
                BigError = True
                
        Else
                If (Not IsNumeric(Txt_PowerAmount.Text)) Then
                    BuildError = BuildError & "Need to fill in valid Power Amount" & vbCrLf
                    Txt_PowerAmount.Text = ""
                    If Not BigError Then Txt_PowerAmount.SetFocus
                    BigError = True
                End If
        End If
          
        If (BigError) Then
            Check_Form_Fields = False
        Else
            Check_Form_Fields = True
        End If


Case "CmdSave"
        'Check to see if anything was found before checking values and returning
        'true so that the record could be updated
        If (State) Then
                'If anything was in variable set it to nothing as well as set big error
                'for final decision
                BuildError = ""
                BigError = False
                
                
                If Cmb_Surety.ListIndex < 0 Then 'tap (IsEmpty(Cmb_Surety.SelText)) Then
                    BuildError = BuildError & "Need to pick surety before saving" & vbCrLf
                    If Not BigError Then Cmb_Surety.SetFocus
                
                    BigError = True
                End If
                        
                If Cmb_Agents.ListIndex < 0 Then 'tap (IsEmpty(Cmb_Agents.SelText)) Then
                    BuildError = BuildError & "Need to pick an agent befor saving" & vbCrLf
                    If Not BigError Then Cmb_Agents.SetFocus
                
                    BigError = True
                End If
                            
                If Cmb_PowerSymbol.ListIndex < 0 Then 'tap (IsEmpty(Cmb_PowerSymbol.SelText)) Then
                    BuildError = BuildError & "Need to pick a symbol befor saving" & vbCrLf
                    If Not BigError Then Cmb_PowerSymbol.SetFocus
                
                    BigError = True
                End If
                
                
                
                If (Txt_AssignmentDate.Text = "") Then
                    BuildError = BuildError & "Need to fill Assignment Date" & vbCrLf
                    If Not BigError Then Txt_AssignmentDate.SetFocus
                    BigError = True
                Else
                        If (Not IsDate(Txt_AssignmentDate.Text)) Then
                            BuildError = BuildError & "Need to fill in valid Assignment Date" & vbCrLf
                            Txt_AssignmentDate.Text = ""
                            If Not BigError Then Txt_AssignmentDate.SetFocus
                            BigError = True
                           
                        End If
                End If
                
                
                If (Txt_ExpirationDate.Text = "") Then
                    BuildError = BuildError & "Need to fill Expiration Date" & vbCrLf
                    BigError = True
                    If Not BigError Then Txt_ExpirationDate.SetFocus
                    
                Else
                        If (Not IsDate(Txt_ExpirationDate.Text)) Then
                            BuildError = BuildError & "Need to fill in valid Expiration Date" & vbCrLf
                            Txt_ExpirationDate.Text = ""
                            If Not BigError Then Txt_ExpirationDate.SetFocus
                            BigError = True
                            
                        End If
                End If
                
                
                If (Txt_PowerAmount.Text = "") Then
                    BuildError = BuildError & "Need to fill in Power Amount" & vbCrLf
                    BigError = True
                    If Not BigError Then Txt_PowerAmount.SetFocus
                Else
                        If (Not IsNumeric(Txt_PowerAmount.Text)) Then
                            BuildError = BuildError & "Need to fill in valid Power Amount" & vbCrLf
                            Txt_PowerAmount.Text = ""
                            If Not BigError Then Txt_PowerAmount.SetFocus
                            BigError = True
                        End If
                End If
   
   
        Else
            BuildError = ""
            BuildError = "Need to find a power before updating or saving setting"
        End If
          
        If (BigError) Then
                Check_Form_Fields = False
        Else
                Check_Form_Fields = True
        End If
End Select
Exit Function
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Function
Private Function Check_Reciept() As String
On Error GoTo MyErrHandler
Dim Lookup As String
    'See if the fields are valid if so return data to be parsed else we know error occured
    If (IsDate(Txt_AssignmentDate.Text) And Cmb_Agents.List(Cmb_Agents.ListIndex) <> "") Then
        Lookup = Txt_AssignmentDate.Text & "," & Cmb_Agents.ItemData(Cmb_Agents.ListIndex)
        Check_Reciept = Lookup
    Else
        Check_Reciept = "ERROR"
    End If
Exit Function
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Function

Private Sub MnuAddAssign_Click()
On Error GoTo MyErrHandler
    Form1.Show
    FormBailBondsInventory.Hide
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Txt_PowerNumber_KeyPress(KeyAscii As Integer)
On Error GoTo MyErrHandler
'If user hits return do a find
    If KeyAscii = 13 Then
        Call Cmd_Find_Click
        KeyAscii = 0
    End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub
Public Sub LoadReport(MyRs As Object)
On Error GoTo MyErrHandler
MyRs.MoveFirst 'Move to first record

'Cycle through and build the report
While Not (MyRs.EOF)
    With DataReport1.Sections("Section1").Controls 'section1 mean that section you create in datareport
        .Item("txtPowerNumber").DataField = MyRs("Power_Num").Name
        .Item("txtAssignDate").DataField = MyRs("Assign_Date").Name
        .Item("txtAgentId").DataField = MyRs("Agent_Id").Name
        '.Item("txtAgentId").DataField = MyAgentLname  'MyRs("Agent_Id").Name
    End With
    MyRs.MoveNext
Wend

'To Set Label caption
With DataReport1.Sections("Section2").Controls
    .Item("lblPowerNum").Caption = "Power #"
    .Item("lblAssignDate").Caption = "Assigment Date"
    .Item("lblAgentId").Caption = "Agent Id"
    
End With

'set report title
With DataReport1.Sections("Section4").Controls
    .Item("lblTitle").Caption = "Reciept For Agent " & Cmb_Agents.List(Cmb_Agents.ListIndex)
End With

'to set datasource for datareport
Set DataReport1.DataSource = MyRs

'show datareport
DataReport1.Show
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Public Sub CloseReport(MyRecordset As Object)
    MyRecordset.Close
    Set MyRecordset = Nothing
End Sub

