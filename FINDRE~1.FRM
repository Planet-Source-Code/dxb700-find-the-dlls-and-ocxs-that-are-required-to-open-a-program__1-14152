VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Required DLLs and OCXs Finder - by DXB700 - ICQ:7293201"
   ClientHeight    =   4755
   ClientLeft      =   5760
   ClientTop       =   3930
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6510
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   9600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog com 
      Left            =   6960
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please select a file to search for it's required dlls and ocxs"
      Filter          =   "EXE FILES |*.exe"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Click here to select a file to start searching"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   6960
      Top             =   120
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2625
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "The Job Has Been Completed :)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   1800
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   840
      TabIndex        =   6
      Top             =   3960
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Found:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1800
      TabIndex        =   4
      Top             =   3480
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Searching in file:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by DXB700
'strange_life@hotmail.com
'ICQ 7293201
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim contents As String
Dim x As Long
Dim y As Long
Dim r As String
Dim g As String
Dim b As Long
Dim filename As String
Dim L As Long
Sub StartTheJob()
On Error Resume Next
y = 0
x = 0
r = ""
g = ""
b = 0

List1.Clear
Open filename For Binary As #1
contents = Space(LOF(1))
Get #1, , contents
Close #1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
y = InStr(1, contents, ".DLL")
b = y
5:
b = b - 1
g = Mid(contents, b)
Text1.Text = g
If Len(Text1.Text) = 0 Then
x = b
GoTo 4
Else
GoTo 5
End If
4 r = Mid(contents, b + 1)
List1.AddItem r
'ok,next
y = InStr(y + 4, contents, ".DLL")
If y = 0 Then
GoTo FINDOCX
Else
b = y
g = ""
Text1.Text = ""
GoTo 5
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FINDOCX:
y = InStr(1, contents, ".OCX")
b = y
6:
b = b - 1
g = Mid(contents, b)
Text1.Text = g
If Len(Text1.Text) = 0 Then
x = b
GoTo 7
Else
GoTo 6
End If
7 r = Mid(contents, b + 1)
List1.AddItem r
'ok,next
y = InStr(y + 4, contents, ".OCX")
If y = 0 Then
GoTo finddll
Else
b = y
g = ""
Text1.Text = ""
GoTo 6
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
finddll:
y = InStr(1, contents, ".dll")
b = y
8:
b = b - 1
g = Mid(contents, b)
Text1.Text = g
If Len(Text1.Text) = 0 Then
x = b
GoTo 11
Else
GoTo 8
End If
11 r = Mid(contents, b + 1)
List1.AddItem r
'ok,next
y = InStr(y + 4, contents, ".dll")
If y = 0 Then
GoTo findsmallocx
Else
b = y
g = ""
Text1.Text = ""
GoTo 8
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
findsmallocx:
y = InStr(1, contents, ".ocx")
b = y
25:
b = b - 1
g = Mid(contents, b)
Text1.Text = g
If Len(Text1.Text) = 0 Then
x = b
GoTo 13
Else
GoTo 25
End If
13 r = Mid(contents, b + 1)
List1.AddItem r
'ok,next
y = InStr(y + 4, contents, ".ocx")
If y = 0 Then
If L = 1 Then
L = 0
Label4.Caption = List1.ListCount & " DLLs and OCXs."
Label5.Visible = True
Exit Sub
Else
L = 1
StartTheJob
End If
Else
b = y
g = ""
Text1.Text = ""
GoTo 25
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub Command2_Click()
List1.Clear
Label5.Visible = False
com.ShowOpen
If Len(com.filename) = 0 Then
Exit Sub
Else
filename = com.filename
Label2.Caption = filename
End If
List1.Clear
MsgBox "Note: The program might seem stuck while searching (if the program you specify has many dlls and ocxs), but it isn't. Please give the program atmost 1 minute for searching. You will be informed when the program completes the job", vbExclamation, "Please Read Carefuly"
StartTheJob
End Sub


