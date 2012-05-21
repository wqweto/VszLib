VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   516
      Left            =   924
      TabIndex        =   1
      Top             =   1176
      Width           =   1944
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Compress"
      Height          =   516
      Left            =   924
      TabIndex        =   0
      Top             =   420
      Width           =   1944
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompress_Click()
    With New cVszArchive
        .AddFile App.Path & "\Form1.frm"
        .AddFile App.Path & "\Project1.vbp"
        .Parameter("x") = 3 '-- CompressionLevel = Fast
        .CompressArchive App.Path & "\test.7z"
    End With
    MsgBox "test.7z created ok", vbExclamation
End Sub

Private Sub cmdExtract_Click()
    With New cVszArchive
        .OpenArchive App.Path & "\test.7z"
        .Extract App.Path & "\Unpacked"
    End With
    MsgBox "test.7z unpacked ok", vbExclamation
End Sub
