VERSION 5.00
Begin VB.Form ENETERPASSWORD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTER YOUR ACCESS PASSWORD"
   ClientHeight    =   1170
   ClientLeft      =   1155
   ClientTop       =   1740
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2940
      TabIndex        =   3
      Top             =   690
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONTINUE"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4230
      TabIndex        =   2
      Top             =   690
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   60
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   330
      Width           =   5415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "YOUR PASSWORD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   1725
   End
End
Attribute VB_Name = "ENETERPASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'CHECKS PASSWORD, IF CORRECT THEN RUNS THE
'PROGRAM WXYZ.EXE
'IF NOT MSGBOX IS DISPLAYED PROGRAM EXITS

'OPEN PASSWORD FILE
Open PASSWORDPATH For Input As #1
PASSWORDSTR = Input$(LOF(1), #1)
Close #1
'PLACE ENTERED PASSWORD IN CHECK FILE
Open CURDIRR & "\PIXCC.PWD" For Output As #1
Print #1, Text1.Text
Close #1
'REOPEN THE CHECK FILE TO ENTEREDPASSWORD STRING
'VARIABLE
Open CURDIRR & "\PIXCC.PWD" For Input As #1
ENTEREDPASSWORD = Input$(LOF(1), #1)
Close #1
'CHECK THE PASSWORD
If PASSWORDSTR = ENTEREDPASSWORD Then
RETVAL = Shell(PROGRAMPATH, 1)
'END PROGRAM
End
Else
MsgBox "THIS IS THE WRONG PASSWORD", vbOKOnly, "ACCESS DENIED"
End
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
'CHECKS PASSWORD, IF CORRECT THEN PASSWORD
'IS REMOVED
'IF NOT MSGBOX IS DISPLAYED PROGRAM EXITS

'OPEN PASSWORD FILE
Open PASSWORDPATH For Input As #1
PASSWORDSTR = Input$(LOF(1), #1)
Close #1
'PLACE ENTERED PASSWORD IN CHECK FILE
Open CURDIRR & "\PIXCC.PWD" For Output As #1
Print #1, Text1.Text
Close #1
'REOPEN THE CHECK FILE TO ENTEREDPASSWORD STRING
'VARIABLE
Open CURDIRR & "\PIXCC.PWD" For Input As #1
ENTEREDPASSWORD = Input$(LOF(1), #1)
Close #1
'CHECK THE PASSWORD
If PASSWORDSTR = ENTEREDPASSWORD Then
'OPEN FILE WITH PROGRAM NAME
Open CURDIRR & "\PIXCCC.PWD" For Input As #1
PROGRAM = Input$(LOF(1), #1)
Close #1
'REMOVE PASSWORD
Kill CURDIRR & "\PIXC.PWD"
Kill CURDIRR & "\PIXCC.PWD"
FileCopy CURDIRR & "\WXYZ.EXE", CURDIRR & "\" & PROGRAM
Kill CURDIRR & "\WXYZ.EXE"
Kill CURDIRR & "\PIXCCC.PWD"
'END PROGRAM
End
Else
MsgBox "THIS IS THE WRONG PASSWORD", vbOKOnly, "ACCESS DENIED"
End
End If
End Sub

Private Sub Form_Activate()
'PLACE THE DIRECTORY OF THE PASSWORD PROGRAM
'RUNNING IN THE CURDIRR STRING CARIABLE
'SO THAT THE PROGRAM KNOWS WHERE TO LOOK
'FOR THE PASSWORD FILE
CURDIRR = CurDir("C")
'ENTER THE PATH TO PASSWORD FILE IN
'PASSWORDPATH STRING VARIABLE
'PIXC.PWD IS A TEXT FILE THAT HOLDS THE PASSWORD
PASSWORDPATH = CURDIRR & "\PIXC.PWD"
'ENTER PATH TO PROGRAM (WXYZ.EXE) INTO
'PROGRAMPATH STRING VARIABLE
PROGRAMPATH = CURDIRR & "\WXYZ.EXE"
End Sub

Private Sub Text1_Change()
'CHECKS IF ANYTHING IS TYPED IN TEXT BOX
'IF SO THEN COMMAND BUTTON IS MADE ENABLED
If Text1.Text = "" Then
Command1.Enabled = 0
Else
Command1.Enabled = -1
End If
End Sub


