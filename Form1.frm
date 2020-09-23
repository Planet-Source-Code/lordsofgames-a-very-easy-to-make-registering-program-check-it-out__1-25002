VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Registering Process 2/3"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   6855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Forget It"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Continue to Register"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Picture         =   "Form1.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      Picture         =   "Form1.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      Picture         =   "Form1.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   -360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1931
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser register 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Please Wait till page fully loads..."
      Top             =   1200
      Width           =   6840
      ExtentX         =   12065
      ExtentY         =   6588
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Line Line5 
      X1              =   6840
      X2              =   0
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line4 
      X1              =   6840
      X2              =   6840
      Y1              =   4920
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   6960
      X2              =   6960
      Y1              =   5040
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6960
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   120
      X2              =   6720
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Sub Command1_Click()
On Error Resume Next
register.GoBack
End Sub
Private Sub Command2_Click()
On Error Resume Next
register.Gofoward
End Sub
Private Sub Command3_Click()
On Error Resume Next
register.Refresh
End Sub
Private Sub Command4_Click()
On Error Resume Next
register.GoHome
End Sub

Private Sub Command5_Click()
On Error Resume Next
If MsgBox("Are you sure? It will take a few seconds to load.", vbYesNo) = vbYes Then
 StartingAddress = ("http://www.mail.yahoo.com") 'Make a webpage where the user can type in the registration password and continue. Password protected sites work best, if you know what I mean.
 register.Navigate StartingAddress
    Me.Caption = register.LocationName
 End If
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
MsgBox ("By clicking Continue with Registering, you will be taken to our registering site. You must be connected first and have an official ISP.")
End Sub

