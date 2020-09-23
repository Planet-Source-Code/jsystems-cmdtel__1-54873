VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cmdtel v 1.0 by JSystems (jsystems@home.ro)"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   660
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private i As Integer
Private n As Long

Private Sub Data1_Reposition()
DoEvents
Text1.SelText = "#"
Text1.SelStart = Len(Text1.Text)
Text1.Refresh

End Sub

Private Sub Form_Load()
On Error GoTo l1
DoEvents

Me.Show
Text1.SelText = "wait..."
Text1.SelStart = Len(Text1.Text)
Text1.Refresh

Data1.DatabaseName = App.Path & "\telef.dat"
Data1.RecordSource = "select * from T order by Nume ASC"
Data1.Refresh

Text1.Text = ""
Text1.Text = Text1.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
" CMDTEL V 1.0 " & vbCrLf & vbCrLf & _
"Jsystems Inc. All rights reserved." & vbCrLf & vbCrLf & _
"     Type help for help." & vbCrLf & vbCrLf & ">"

Text1.SelStart = Len(Text1.Text)

Exit Sub
l1:
Text1.SelText = "ERROR: " & Err.Description & " returned " & Err.Number
Text1.Refresh
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
Text1.Locked = True
Else
Text1.Locked = False
End If

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo l1
Dim cmdT As String
i = i + 1
If KeyCode = vbKeyReturn Then
KeyCode = 0
n = Len(Text1.Text)
'....parancs azonosito
cmdT = Left(Right(Text1.Text, i + 1), Len(Right(Text1.Text, i + 1)) - 2)
Text1.SelText = CmdStart(cmdT)
Text1.SelText = ">"
i = 0
End If
Exit Sub
l1:
Text1.SelText = "ERROR: " & Err.Description & " returned " & Err.Number
Text1.Refresh

End Sub

Private Function CmdStart(ByVal Txt As String) As String
On Error GoTo l1
DoEvents

If Txt = "help" Then
CmdStart = "Command usage " & vbCrLf & _
"(You can use the jocer caracter (*)):" & vbCrLf & _
"find [name/part of name]-[address]-[number] |example: find nagy-*ciocarliei*" & vbCrLf & _
"                                                      find -*florilor*-122*" & vbCrLf & _
"                                                      find --112987" & vbCrLf & _
"                                                      find nagy istvan" & vbCrLf & vbCrLf & _
"USE ONLY lcase CARACTERS" & vbCrLf

ElseIf Txt = "find" Then
CmdStart = "Command usage " & vbCrLf & _
"(You can use the jocer caracter (*)):" & vbCrLf & _
"find [name/part of name]-[address]-[number] |example: find nagy-*ciocarliei*" & vbCrLf & _
"                                                      find -*florilor*-122*" & vbCrLf & _
"                                                      find --112987" & vbCrLf & _
"                                                      find nagy istvan" & vbCrLf & vbCrLf & _
"USE ONLY lcase CARACTERS" & vbCrLf
ElseIf Txt = "" Then
CmdStart = "" & vbCrLf
ElseIf Txt = "exit" Then
Unload Me
End
ElseIf Left(Txt, 5) = "find " Then

CmdStart = IsCommand(Txt)
Else
CmdStart = "Invalid command." & vbCrLf
End If
Exit Function
l1:
Text1.SelText = "ERROR: " & Err.Description & " returned " & Err.Number
Text1.Refresh

End Function

Private Function IsCommand(ByVal Txt As String) As String
On Error GoTo l1
DoEvents
Text1.SelText = "Searching..."
Text1.Refresh
Dim NumeX As String
Dim AdresaX As String
Dim NumberX As String
Dim OptionTxt As String
Dim startTxt As String
startTxt = Right(Txt, Len(Txt) - 5)
If Left(startTxt, 1) = "-" Then
    If Left(startTxt, 2) = "--" Then
        NumeX = ""
        AdresaX = ""
        NumberX = Right(startTxt, Len(startTxt) - 2)
    Else
        NumeX = ""
        For a = 2 To Len(startTxt)
            If Mid(startTxt, a, 1) = "-" Then
                AdresaX = Mid(startTxt, 2, a - 2)
                NumberX = Mid(startTxt, Len(AdresaX) + 3, Len(startTxt) - Len(AdresaX))
            Exit For
            End If
        Next
    End If
Else
For a = 1 To Len(startTxt)
            If Mid(startTxt, a, 1) = "-" Then
                NumeX = Mid(startTxt, 1, a - 1)
            Exit For
            Else
            NumeX = startTxt
            End If
Next
Dim aa As Single
For a = Len(NumeX) + 2 To Len(startTxt)
            aa = aa + 1
            If Mid(startTxt, a, 1) = "-" Then
                AdresaX = Mid(startTxt, Len(NumeX) + 2, aa - 1)
                NumberX = Mid(startTxt, Len(AdresaX) + Len(NumeX) + 3, Len(startTxt) - Len(AdresaX))
            Exit For
            Else
            AdresaX = Mid(startTxt, Len(NumeX) + 2, Len(startTxt))
            NumberX = ""
            End If
        Next
End If
Dim crit1 As String
Dim crit2 As String
Dim crit3 As String
Dim critX As String
crit1 = NumeX
crit2 = AdresaX
crit3 = NumberX
If crit1 = "" Then crit1 = "*"
If crit2 = "" Then crit2 = "*"
If crit3 = "" Then crit3 = "*"
critX = "where Nume LIKE '" & crit1 & "' AND Adresa LIKE '" & crit2 & "' AND Tel LIKE '" & crit3 & "'"

Data1.RecordSource = "select * from T " & critX & " order by Nume ASC"
Data1.Refresh
Dim resultTxt As String
resultTxt = "SQL Called: " & "select * from T " & critX & " order by Nume ASC" & vbCrLf & vbCrLf
If Data1.Recordset.RecordCount <> 0 Then
Data1.Recordset.MoveLast
resultTxt = resultTxt & Data1.Recordset.RecordCount & " mach(es) found." & vbCrLf
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
resultTxt = resultTxt & Trim(Data1.Recordset(0).Value) & "| " & Trim(Data1.Recordset(1).Value) & "| " & Trim(Data1.Recordset(2).Value) & "| " & Trim(Data1.Recordset(3).Value) & vbCrLf
Data1.Recordset.MoveNext
Loop
IsCommand = resultTxt & vbCrLf
Else
IsCommand = resultTxt & "No mach found." & vbCrLf
End If
Exit Function
l1:
Text1.SelText = "ERROR: " & Err.Description & " returned " & Err.Number
Text1.Refresh

End Function

