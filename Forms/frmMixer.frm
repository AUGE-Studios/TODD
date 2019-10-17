VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMixer 
   Caption         =   "TODD Mixer"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "ToddMixer"
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   741
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5805
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   5160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox cmdSample 
      Height          =   495
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSample_Click(Index As Integer)
  Caption = CStr(Index)
  If cmdSample(Index).Value = 0 Then cmdSample(Index).BackColor = tdStandard Else cmdSample(Index).BackColor = tdLightGreen
End Sub

Private Sub Form_Load()
  Dim lngIndex As Long
  For lngIndex = 1 To 56
    Load cmdSample(lngIndex)
  Next lngIndex
  For lngIndex = 0 To cmdSample.UBound
    With cmdSample(lngIndex)
      .Caption = CStr(lngIndex)
      .Visible = True
    End With 'With cmdSample(lngIndex)
  Next lngIndex
End Sub

Private Sub Form_Resize()
  Dim lngIndex As Long
  Dim sngWidth As Single, sngHeight As Single
  Dim lngRow As Long, lngCol As Long
  sngWidth = ScaleWidth / 10
  sngHeight = ScaleHeight / 10
  For lngIndex = 0 To cmdSample.UBound
    lngRow = lngIndex \ 10
    lngCol = lngIndex Mod 10
    cmdSample(lngIndex).Move lngCol * sngWidth, lngRow * sngHeight, sngWidth, sngHeight
  Next lngIndex
End Sub

