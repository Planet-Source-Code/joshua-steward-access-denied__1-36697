VERSION 5.00
Begin VB.Form MAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD PASSWORD"
   ClientHeight    =   4575
   ClientLeft      =   1155
   ClientTop       =   1740
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5565
   Begin VB.Frame STEP3 
      Caption         =   "STEP 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   60
      TabIndex        =   11
      Top             =   420
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CommandButton Command4 
         Caption         =   "GO BACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2790
         TabIndex        =   18
         Top             =   1890
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CONTINUE"
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
         Height          =   400
         Left            =   4080
         TabIndex        =   17
         Top             =   1890
         Width           =   1245
      End
      Begin VB.TextBox Text2 
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1410
         Width           =   5205
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   540
         Width           =   5205
      End
      Begin VB.Label LINE 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Left            =   120
         TabIndex        =   16
         Top             =   990
         Width           =   5205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CONFIRM PASSWORD"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1110
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ENTER YOUR PASSWORD"
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
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   2355
      End
   End
   Begin VB.Frame STEP2 
      Caption         =   "STEP 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   60
      TabIndex        =   6
      Top             =   420
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CommandButton Command3 
         Caption         =   "NO"
         Height          =   400
         Left            =   2790
         TabIndex        =   9
         Top             =   870
         Width           =   1245
      End
      Begin VB.CommandButton Command2 
         Caption         =   "YES"
         Default         =   -1  'True
         Height          =   400
         Left            =   4080
         TabIndex        =   8
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "IS THIS THE FILE YOU WISH TO ADD A PASSWORD TO?"
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
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   5070
      End
      Begin VB.Label FILENAME 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   510
         Width           =   5205
      End
   End
   Begin VB.Frame STEP1 
      Caption         =   "STEP 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   5445
      Begin VB.DriveListBox Drive1 
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
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   2475
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3465
         Left            =   60
         TabIndex        =   4
         Top             =   570
         Width           =   2475
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3410
         Left            =   2580
         Pattern         =   "*.EXE*"
         TabIndex        =   3
         Top             =   570
         Width           =   2775
      End
   End
   Begin VB.Label STEPNUM 
      AutoSize        =   -1  'True
      Caption         =   "1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   285
   End
   Begin VB.Label STEPDESCRIPTION 
      AutoSize        =   -1  'True
      Caption         =   "SELECT THE .EXE YOU WANT TO RESTRICT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   0
      Top             =   90
      Width           =   4815
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'CHECK TO SEE THAT BOTH PASSWORDS MATCH
'IF NOT INFORM THE USER AND
'ERASE THE FIELDS, AND PLACE THE FOCUS BACK TO
'THE ENTER PASSWORD TEXT BOX (DISABLE OTHER
'CONTROLS)
'IF PASSWORDS MATCH THEN
'RENAME THE SELECTED EXE FILE TO "WXYZ.EXE"
'RENAME THE PASSWORD PROGRAM TO THE ORIGNAL
'NAME OF THE SELECTED EXE AND COPY THE PASSWORD
'PROGRAM TO THE FOLDER OF THE SELECTED EXE
'ALSO MAKE A FILE THAT HOLDS THE PASSWORD IN
'IT  AND PUT IT IN THE FOLDER OF THE
'SELECTED EXE ALONG WITH A BLANK TXT FILE TO
'CHECK THE PASSWORD, WITH OUT THIS THE PASSOWRD
'THAT THE USER PUTS INTO THE PASSWORD PROGRAM
'WOULD NEVER MATCH THE ONE IN THE FILE
'YOU MUST PLACE THAT ENTERED PASSWORD
'INTO THE BLANK FILE AND THEN REOPEN THAT FILE
'THEN CHECK THE TEXT FROM THE BLANK FILE TO
'THE PASSWORD FILE
'RENAME THE PASSOWRD PROGRAM ALLOWS
'THE SHORTCUTS THAT EXISTS FOR THE SELECTED
'PRGRAM TO STILL BE USABLE EVEN THOUGH
'THEY WILL JUST BRING UP THE PASSWORD PROGRAM
If Text2.Text = Text1.Text Then
'RENAME THE SELECTED EXE FILE TO "WXYZ.EXE"
'USING FILECOPY, THEN KILL THE ORIGINAL (DELETE)
FileCopy PROGRAMPATH, PROGRAMPATHwoFILENAME & "WXYZ.EXE"
Kill PROGRAMPATH
'RENAME THE PASSWORD PROGRAM TO THE ORIGNAL
'NAME OF THE SELECTED EXE AND COPY THE PASSWORD
'PROGRAM TO THE FOLDER OF THE SELECTED EXE
FileCopy CURDIRR & "\PASS.EXE", PROGRAMPATH
'MAKE FILE THAT HOLDS PASSWORD IN IT AND PUT IT
'IN THE FOLDER OF THE SELECTED EXE AND THE BLANK
'CHECK FILE
Open PROGRAMPATHwoFILENAME & "\PIXC.PWD" For Output As #1
Print #1, Text2.Text
Close #1
Open PROGRAMPATHwoFILENAME & "\PIXCC.PWD" For Output As #1
Print #1, ""
Close #1
MsgBox "PASSWORD CREATED", vbOKOnly, "ADD PASSWORD"
'END PROGRAM
End
'PIXC.PWD AND PIXCC.PWD AND WXYZ.EXE
'ARE JUST USED FOR ODD FILENAMES THE
'NAMES DON'T MATTER
Else
MsgBox "PASSWORDS DON'T MATCH, TRY AGAIN", vbOKOnly, "WRONG PASSWORD"
Text1.Text = ""
Text2.Text = ""
Text2.Enabled = 0
Command1.Enabled = 0
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
'PLACE THE ENTIRE PATH TO THE FILE
'IN THE PROGRAMPATH STRING VARIABLE
PROGRAMPATH = File1.Path & "\" & File1.FILENAME
'PLACE PROGRAM PATH WITHOUT FILENAME IN
'PROGRAMPATHwoFILENAME STRING VARIABLE
PROGRAMPATHwoFILENAME = File1.Path & "\"
'CHANGE STEPNUM AND STEPDESCRIPTION
'TO REFLECT CHANGE TO STEP 3
STEPNUM.Caption = "3)"
STEPDESCRIPTION = "ENTER PASSWORD TO USE"
STEP2.Visible = 0
STEP3.Visible = -1
'SET FOCUS TO TEXT1 (ENTER PASSWORD TEXT BOX)
Text1.SetFocus

End Sub

Private Sub Command3_Click()
'CHANGE STEPNUM AND STEPDESCRIPTION
'TO REFLECT CHANGE BACK TO STEP 1
STEPNUM.Caption = "1)"
STEPDESCRIPTION = "SELECT THE .EXE YOU WANT TO RESTRICT"
STEP2.Visible = 0
STEP1.Visible = -1
End Sub

Private Sub Command4_Click()
'CHANGE STEPNUM AND STEPDESCRIPTION
'TO REFLECT CHANGE BACK TO STEP 2
STEPNUM.Caption = "2)"
STEPDESCRIPTION = "CONFIRM FILE SELECTION"
STEP3.Visible = 0
STEP2.Visible = -1
'SET FOCUS TO YES BUTTON TO QUICKLY
'PASS BY STEP IF USER IS SURE OF THEIR
'SELECTION (THEY JUST PRESS ENTER)
Command2.SetFocus
End Sub


Private Sub Dir1_Change()
'CHANGES FILE1'S PATH TO DIR1'S PATH
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
'GOES TO DRIVEERR IF DRIVE IS UNREADABLE
On Error GoTo DRIVEERR
'IF DRIVE IS READABLE THEN THIS WILL
'SWITCH DIR1'S PATH TO THE SELECTED
'DRIVE
Dir1.Path = Drive1.Drive
Exit Sub
DRIVEERR:
MsgBox "DRIVE IS CAN NOT BE ACCESSED", vbOKOnly, "DRIVE PROBLEM"
Resume Next
End Sub

Private Sub File1_DblClick()
'PLACE THE SELECTED FILE NAME IN THE PROGRAM
'STRING VARIABLE LOCATED IN THE DECLARATIONS
'SECTION OF THE MODULE NAMED MAINBAS
PROGRAM = File1.FILENAME
'PLACE THE FILE NAME IN THE FILENAME BOX ON
'THE STEP 2 FRAME
FILENAME.Caption = File1.FILENAME
'CHANGE STEPNUM AND STEPDESCRIPTION
'TO REFLECT CHANGE TO STEP 2
STEPNUM.Caption = "2)"
STEPDESCRIPTION = "CONFIRM FILE SELECTION"
STEP1.Visible = 0
STEP2.Visible = -1
'SET FOCUS TO YES BUTTON TO QUICKLY
'PASS BY STEP IF USER IS SURE OF THEIR
'SELECTION (THEY JUST PRESS ENTER)
Command2.SetFocus
End Sub


Private Sub Form_Activate()
CURDIRR = CurDir("C")
End Sub

Private Sub Text1_Change()
'CHECK IF ANYTHING IS TYPED IN THE
'TEXT BOX AND IF SO THEN ALLOW THE CONFIRM
'PASSWORD TEXT BOX TO BECOME ENABLED
If Text1.Text = "" Then
Text2.Enabled = 0
Else
Text2.Enabled = -1
End If
End Sub


Private Sub Text2_Change()
'CHECK IF ANYTHING IS TYPED IN THE
'TEXT BOX AND IF SO THEN ALLOW THE CONTINUE
'BUTTON TO BECOME ENABLED
If Text1.Text = "" Then
Command1.Enabled = 0
Else
Command1.Enabled = -1
End If
End Sub


