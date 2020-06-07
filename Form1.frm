VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengurutkan Abjad Secara Ascending"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
'Ketika Command1 diklik, maka kumpulan abjad di Text1
'akan diurutkan dari terkecil hingga abjad yang 'terbesar dan hasilnya ditampilkan di Text2 menjadi:
'ABCDEFGHIJKLMNOPQRSTUVWXYZ
  Dim sAbjad() As String, ar As Integer, _
  br As Integer, sAbjadTemp As String
  Text2.Text = ""
  ReDim sAbjad(Len(Text1.Text) - 1)
  For ar = 1 To Len(Text1.Text)
    sAbjad(ar - 1) = Mid(Text1.Text, ar, 1)
  Next ar
  For ar = LBound(sAbjad) To UBound(sAbjad)
    For br = LBound(sAbjad) To UBound(sAbjad) - 1
      If sAbjad(br) > sAbjad(br + 1) Then
        sAbjadTemp = sAbjad(br + 1)
        sAbjad(br + 1) = sAbjad(br)
        sAbjad(br) = sAbjadTemp
      End If
    Next br
  Next ar
  For ar = LBound(sAbjad) To UBound(sAbjad)
    Text2.Text = Text2.Text & sAbjad(ar)
  Next ar
End Sub

Private Sub Form_Load()
  'Mula-mula, kumpulan abjad di Text1 belum urut...
  Text1.Text = "QWERTYUIOPASDFGHJKLZXCVBNM"
  Text2.Text = ""
End Sub


