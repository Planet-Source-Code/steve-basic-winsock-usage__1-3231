VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connecting to Server, please wait......"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComCtl2.Animation Animation1 
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      _Version        =   327680
      FullWidth       =   105
      FullHeight      =   57
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Animation1.AutoPlay = True
Animation1.open (App.Path & "/graphics/Search.avi")
Animation1.Play 1000000, 1

End Sub


