VERSION 5.00
Begin VB.Form frmApp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Application"
   ClientHeight    =   2445
   ClientLeft      =   4920
   ClientTop       =   2670
   ClientWidth     =   6120
   Icon            =   "frmApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6120
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   1215
      Left            =   4620
      Picture         =   "frmApp.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   570
      Width           =   1380
   End
   Begin VB.Label lblText 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2310
      Left            =   105
      TabIndex        =   1
      Top             =   90
      Width           =   4260
   End
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

    ShellAbout Me.hWnd, App.Title, "Created by Deepak Kumar Shaw. " & vbCrLf & Chr(34) & _
               "OmSoft Corporation Ptv. Ltd." & Chr$(34) & " India.", ByVal 0&
    
End Sub

Private Sub Form_Load()
Const MyTxt = "This is very simple, easy, fastest way to make your About form which gives a really profession look to your application." & vbCrLf & _
              "Just a Single line of code, no over heads." & vbCrLf & _
              "If you like, vote plz. Thanks."

lblText.Caption = MyTxt
End Sub

