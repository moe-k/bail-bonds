VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4212
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   4932
      Begin VB.CommandButton Cmd_Clear 
         Caption         =   "&Clear"
         Height          =   250
         Left            =   3720
         TabIndex        =   26
         Top             =   3720
         Width           =   1000
      End
      Begin VB.ComboBox Cmb_PowerSymbol 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AddAssignFrm.frx":0000
         Left            =   240
         List            =   "AddAssignFrm.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Agents 
         DataField       =   "Agent_Lname"
         DataSource      =   "Adodc2"
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox Cmb_Surety 
         DataField       =   "Name"
         Height          =   315
         ItemData        =   "AddAssignFrm.frx":0004
         Left            =   240
         List            =   "AddAssignFrm.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Txt_ExpirationDate 
         Height          =   288
         Left            =   3480
         TabIndex        =   6
         Top             =   3120
         Width           =   1212
      End
      Begin VB.TextBox Txt_AssignmentDate 
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   3360
         TabIndex        =   4
         Top             =   1680
         Width           =   1212
      End
      Begin VB.TextBox Txt_PowerNumber 
         Height          =   288
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox Txt_PowerAmount 
         Height          =   288
         Left            =   1680
         TabIndex        =   5
         Top             =   3120
         Width           =   1572
      End
      Begin VB.TextBox TxtCreateDate 
         Enabled         =   0   'False
         Height          =   288
         Left            =   3360
         TabIndex        =   17
         Top             =   480
         Width           =   1212
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "&Expiration Date"
         Height          =   252
         Left            =   3480
         TabIndex        =   25
         Top             =   2880
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "Ass&igned Date"
         Height          =   252
         Left            =   3360
         TabIndex        =   24
         Top             =   1440
         Width           =   1452
      End
      Begin VB.Label Label8 
         Caption         =   "&Power Number"
         Height          =   252
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label7 
         Caption         =   "&Power Symbol"
         Height          =   252
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   1452
      End
      Begin VB.Label Label6 
         Caption         =   "P&ower Amount"
         Height          =   252
         Left            =   1680
         TabIndex        =   21
         Top             =   2880
         Width           =   1092
      End
      Begin VB.Label Label5 
         Caption         =   "&Agent"
         Height          =   252
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "&Surety"
         Height          =   252
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1932
      End
      Begin VB.Label Label13 
         Caption         =   "Created Date"
         Height          =   252
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1932
      Left            =   240
      TabIndex        =   10
      Top             =   4440
      Width           =   4932
      Begin VB.CommandButton Cmd_Assign 
         Caption         =   "Assi&gn"
         Height          =   250
         Left            =   3720
         TabIndex        =   12
         Top             =   1440
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_Add 
         Caption         =   "A&dd"
         Height          =   250
         Left            =   2520
         TabIndex        =   11
         Top             =   1440
         Width           =   1000
      End
      Begin VB.TextBox Txt_Howmany 
         Height          =   288
         Left            =   2520
         TabIndex        =   8
         Top             =   600
         Width           =   732
      End
      Begin VB.TextBox Txt_ToPower 
         Height          =   288
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1932
      End
      Begin VB.TextBox Txt_FromPower 
         Height          =   288
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1932
      End
      Begin VB.Label Label4 
         Caption         =   "&Number of Powers"
         Height          =   252
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "&From Power"
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "&To Power"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1332
      End
   End
End
Attribute VB_Name = "Form1"
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
Private Sub Cmd_Add_Click()
On Error GoTo MyErrHandler
   'Connection variables and counters
   Dim cnn1 As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim strCnn As String
   Dim count As Integer
   Dim strSQL As String
   
   Txt_AssignmentDate.Text = ""
   
       'Set connection
       Set cnn1 = New ADODB.Connection
       strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
       cnn1.Open strCnn
          
       'Make recordset object and set params so we can add records to DB
       Set rst = New ADODB.Recordset
       rst.CursorType = adOpenKeyset
       rst.LockType = adLockOptimistic
       rst.Open "Power", cnn1, , , adCmdTable
    
       'Check form to see if values filled in and they are correct
        If (Check_Form_Fields("CmdAdd")) Then
                from_pow = Txt_FromPower.Text
                    
                'Add generic powers to db
                For count = 0 To How_many - 1
                        'Check to see if power is already being used if and skip over
                        Set rst = New ADODB.Recordset
                        strSQL = "SELECT * FROM Power WHERE Power_Num = " & from_pow & " And surety_id = " & Cmb_Surety.ItemData(Cmb_Surety.ListIndex)
                        rst.Open strSQL, cnn1, adOpenKeyset, adLockOptimistic, adCmdText
                        'Check to see if power already in use
                        If rst.RecordCount > 0 Then
                           MsgBox "Power Number - " & from_pow & " has already been used.", vbInformation
                           Exit For
                        End If
                        
                        rst.AddNew
                        rst!Power_num = from_pow
                        rst.Fields("Symbol").Value = Cmb_PowerSymbol.Text
                        rst!Execution_Date = Date
                        rst!Create_Date = Date
                        rst!Defendant_Name = "NA"
                        rst!PowerAmount = Txt_PowerAmount.Text
                        rst!Court_Date = Date
                        
                        rst!Court_Loc = "NA"
                        rst!Expiration_Date = Txt_ExpirationDate.Text
                        'Double check not necessary since the errorcheck already
                        'Should have caught it
                        If IsDate(Txt_AssignmentDate.Text) Then
                           rst!Assign_Date = Txt_AssignmentDate.Text
                        End If
                        rst!Bond_Amount = 0
                        rst!Surety_Id = Cmb_Surety.ItemData(Cmb_Surety.ListIndex)
                        
                        If (Cmb_Agents.ListIndex >= 0) Then
                            rst!Agent_Id = Cmb_Agents.ItemData(Cmb_Agents.ListIndex)
                        Else
                            rst!Agent_Id = 0
                        End If
                        from_pow = from_pow + 1
                    rst.Update
                Next count
            'Clear controls
            Cmd_Clear_Click
            'Clean up
            rst.Close
            cnn1.Close
            Set rst = Nothing
            Set cnn1 = Nothing
            MsgBox count & " powers added successfully", vbInformation
            Cmd_Clear.Enabled = True
        Else
            'Show the errors if they occured
            MsgBox BuildError, vbInformation
            'Reset buildError so it could be reused
            BuildError = ""
            Cmd_Clear.Enabled = True
        End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

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
   
   
    
        If (Txt_FromPower.Text = "") Then
                BuildError = BuildError & "Need to fill in From Power" & vbCrLf
                BigError = True
                If Not BigError Then Txt_FromPower.SetFocus
        Else
                If (Not IsNumeric(Txt_FromPower.Text)) Then
                    BuildError = BuildError & "Need to fill in valid From Power # " & vbCrLf
                    BigError = True
                    Txt_FromPower.Text = ""
                    If Not BigError Then Txt_FromPower.SetFocus
                End If
        End If
   
   
        If (Txt_Howmany.Text = "") Then
                BuildError = BuildError & "Need to fill in How many powers" & vbCrLf
                BigError = True
                If Not BigError Then Txt_Howmany.SetFocus
        Else
                If (Not IsNumeric(Txt_FromPower.Text)) Then
                    BuildError = BuildError & "Need to fill in valid value for # of powers " & vbCrLf
                    Txt_FromPower.Text = ""
                    If Not BigError Then Txt_Howmany.SetFocus
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

Private Sub Cmd_Assign_Click()
On Error GoTo MyErrHandler
Dim FromPow, ToPow As Long
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim strCnn As String
Dim count As Integer
Dim strSQL As String
Dim i As Integer
'To Hold the powers that can not
'be assigned because they don't exist
Dim NotUpdateAble(1000) As Long
Dim NotUpdated As String
Dim NIndex As Integer
'To Hold the updated powers
Dim Updated(1000) As Long
Dim UIndex As Integer
Dim UpdatedStr As String

'To hold true if all powers in range
 'were Updated and there was no pows that
 'were not updated
Dim AllUpdated As Boolean
Dim Informstr As String

Dim WhatWay As Boolean

Informstr = ""
UpdatedStr = ""
NotUpdated = ""
BuildError = ""

If (IsNumeric(Txt_FromPower.Text) And IsNumeric(Txt_ToPower.Text) And Cmb_Agents.ListIndex >= 0 And Cmb_Surety.ListIndex >= 0) Then
            'Grab from and to power that will be
            'updated
            FromPow = CLng(Txt_FromPower.Text)
            ToPow = CLng(Txt_ToPower.Text)
            
            'Connection
            Set cnn = New ADODB.Connection
            strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BailBondDB.mdb"
            cnn.Open strCnn
'----------------------------------------------------------------------------------
' Working But Modified to reflect powers that cant be changed because are not existent
            'Open recordset
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
'                End If
'                'move to next power
'                FromPow = FromPow + 1
'                rst.Close
'            'keep goin until we have reached topower
'            Loop Until (FromPow = ToPow)
'--------------------------------------------------------------------------------
            AllUpdated = True
            NIndex = 0
            UIndex = 0
            Set rst = New ADODB.Recordset
            'Query each power
            Do
                strSQL = "SELECT Power_Num,Agent_Id,Surety_Id FROM Power WHERE Power_Num = " & FromPow
                rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
                'Check if query returned anything
                If (rst.RecordCount > 0) Then
                    'if so update
                    Do While Not (rst.EOF)
                        rst!Agent_Id = Cmb_Agents.ItemData(Cmb_Agents.ListIndex)
                        rst!Surety_Id = Cmb_Surety.ItemData(Cmb_Surety.ListIndex)
                        rst.MoveNext
                    Loop
                     Updated(UIndex) = FromPow
                     UIndex = UIndex + 1
                    'move to next power
                    FromPow = FromPow + 1
                    rst.Close
                Else
                  'Set to false since we know that one of the powers does not even exist
                  AllUpdated = False
                  'Grab the power and stuff it into array and increment index for next
                  'non existent power
                  NotUpdateAble(NIndex) = FromPow
                  'Increment From power
                  FromPow = FromPow + 1
                  NIndex = NIndex + 1
                  rst.Close
                End If
            'keep goin until we have reached topower
            Loop Until (FromPow = ToPow)

            'Clean up
            cnn.Close
            Set rst = Nothing
            Set cnn = Nothing
            
            If (AllUpdated) Then
            
                'Need to get the from power again because do loop incremented it and we lost original value
                FromPow = Txt_FromPower.Text
                'Show user it was successfull
                MsgBox "Powers " & FromPow & " to " & ToPow & " were successfully assigned to agent " & _
                Cmb_Agents.List(Cmb_Agents.ListIndex) & " and surety " & _
                Cmb_Surety.List(Cmb_Surety.ListIndex), vbInformation
                Cmd_Clear_Click
                
            Else
                'If above failed we know there is at least one power that was not existent
                'So we generate string that will show the powers that have been
                'updated as well as the powers that have not been updated
                
'                For i = 0 To NIndex
'                    If (NotUpdateAble(i) <> 0) Then
'                        WhatWay = False
'                    Else
'                        WhatWay = True
'                    End If
'                Next i
'
'                If (WhatWay) Then
                    For i = 0 To NIndex - 1
                        NotUpdated = NotUpdated & NotUpdateAble(i) & " "
                    Next i
                    Informstr = Informstr & "Powers " & NotUpdated & " did not exist "
                    'MsgBox Informstr, vbInformation
'                Else
                    For i = 0 To UIndex - 1
                       UpdatedStr = UpdatedStr & Updated(i) & " "
                    Next i
                 
                    Informstr = Informstr & "Powers " & UpdatedStr & " were updated successfully " & vbCrLf
                    MsgBox Informstr, vbInformation
'                End If
                Cmd_Clear_Click
            End If
Else
            'Error check to build errors
            BuildError = "Error -"
            BuildError = BuildError & vbCrLf
            
            If (Not IsNumeric(Txt_FromPower.Text) Or IsNull(Txt_FromPower.Text)) Then
                    BuildError = BuildError & "Need to fill in valid number for from power field" & vbCrLf
            End If
            If (Not IsNumeric(Txt_ToPower.Text) Or IsNull(Txt_ToPower.Text)) Then
                    BuildError = BuildError & "Need to fill in valid number for to power field" & vbCrLf
            End If
            If (Cmb_Agents.ListIndex < 0) Then
                    BuildError = BuildError & "Need to pick an agent to assign to" & vbCrLf
            End If
            If (Cmb_Surety.ListIndex < 0) Then
                    BuildError = BuildError & "Need to pick an surety to assign to" & vbCrLf
            End If
                MsgBox BuildError, vbInformation
End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Cmd_Clear_Click()
On Error GoTo MyErrHandler
        Txt_PowerNumber.Text = ""
        Txt_AssignmentDate.Text = ""
        Txt_ExpirationDate.Text = ""
        Txt_PowerAmount.Text = ""
        Txt_FromPower.Text = ""
        Txt_Howmany.Text = ""
        Txt_ToPower.Text = ""
        TxtCreateDate.Text = ""
        Cmb_Agents.Clear
        Cmb_Surety.Clear
        Cmb_PowerSymbol.Clear
        Load_Agents
        Load_Surety
        Load_Symbols
        Cmd_Add.Enabled = True
        Cmd_Clear.Enabled = True
        Cmd_Assign.Enabled = True
        BuildError = ""
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
    Cmd_Clear.Enabled = True
    Cmd_Assign.Enabled = True
    Cmd_Add.Enabled = True
    Form1.Caption = "Add Assign Powers"
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Form_Terminate()
On Error GoTo MyErrHandler
    FormBailBondsInventory.Show
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo MyErrHandler
  FormBailBondsInventory.Show
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Txt_Howmany_LostFocus()
On Error GoTo MyErrHandler
    'Check to see if fields are valid once user hits tab
    If (Txt_FromPower.Text <> "" And IsNumeric(Txt_FromPower.Text) And Txt_Howmany.Text <> "" And IsNumeric(Txt_Howmany.Text)) Then
            How_many = Txt_Howmany.Text
            from_pow = Txt_FromPower.Text
            Txt_ToPower.Text = from_pow + How_many '- 1
    Else
            'MsgBox ("Need to fill in valid fields for from power and how many powers to add")
            Txt_FromPower.Text = ""
            Txt_Howmany.Text = ""
            Txt_ToPower.Text = ""
    End If
Exit Sub
MyErrHandler:
    MsgBox "Error Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
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

