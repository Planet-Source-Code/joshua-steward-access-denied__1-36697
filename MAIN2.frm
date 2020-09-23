VERSION 5.00
Begin VB.Form MAIN2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REMOVE PASSWORD"
   ClientHeight    =   4605
   ClientLeft      =   1155
   ClientTop       =   1740
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5655
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
      Left            =   30
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CommandButton Command2 
         Caption         =   "NO"
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
         Left            =   2760
         TabIndex        =   9
         Top             =   750
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "YES"
         Default         =   -1  'True
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
         Left            =   4050
         TabIndex        =   8
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5175
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
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   5445
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
         TabIndex        =   2
         Top             =   570
         Width           =   2475
      End
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
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Label STEPDESCRIPTION 
      AutoSize        =   -1  'True
      Caption         =   "PICK THE FILE THAT CONTAINS THE PASSWORD"
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
      Left            =   300
      TabIndex        =   5
      Top             =   90
      Width           =   5340
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "MAIN2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'REMOVE THE PASSWORD

Kill PROGRAMPATHwoFILENAME & "PIXC.PWD"
Kill PROGRAMPATHwoFILENAME & "PIXCC.PWD"
FileCopy PROGRAMPATHwoFILENAME & "WXYZ.EXE", PROGRAMPATHwoFILENAME & PROGRAM
Kill PROGRAMPATHwoFILENAME & "WXYZ.EXE"
MsgBox "PASSWORD REMOVED", vbOKOnly, "REMOVE PASSWORD"
End
End Sub

Private Sub Command2_Click()
'CHANGE STEPNUM AND STEPDESCRIPTION
'TO REFLECT CHANGE BACK TO STEP 1
STEPNUM.Caption = "1)"
STEPDESCRIPTION = "PICK THE FILE THAT CONTAINS THE PASSWORD"
STEP2.Visible = 0
STEP1.Visible = -1
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
Label1.Caption = File1.FILENAME
'CHANGE STEPNUM AND STEPDESCRIPTION
'TO REFLECT CHANGE TO STEP 2
STEPNUM.Caption = "2)"
STEPDESCRIPTION = "CONFIRM FILE SELECTION"
STEP1.Visible = 0
STEP2.Visible = -1
'ADD INFORMATION TO STRING VARIABLES
PROGRAM = File1.FILENAME
PROGRAMPATH = File1.Path & "\" & File1.FILENAME
PROGRAMPATHwoFILENAME = File1.Path & "\"
End Sub


