VERSION 5.00
Begin VB.Form OPTIONN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPTION"
   ClientHeight    =   1140
   ClientLeft      =   3210
   ClientTop       =   3225
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   2610
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
      Height          =   400
      Left            =   30
      TabIndex        =   3
      Top             =   690
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONTINUE"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   690
      Width           =   1245
   End
   Begin VB.OptionButton Option2 
      Caption         =   "REMOVE PASSWORD"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ADD PASSWORD"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1695
   End
End
Attribute VB_Name = "OPTIONN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = -1 Then
MAIN.Show
OPTIONN.Hide
Else
If Option2.Value = -1 Then
MAIN2.Show
OPTIONN.Hide
End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub
