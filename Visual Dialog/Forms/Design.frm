VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Design 
   AutoRedraw      =   -1  'True
   Caption         =   "Template Form"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   8895
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstEvent 
      Height          =   1620
      Left            =   90
      TabIndex        =   13
      Top             =   6000
      Width           =   7155
   End
   Begin VB.Timer tmrLoad 
      Interval        =   1000
      Left            =   4050
      Top             =   6990
   End
   Begin VB.PictureBox DesignForm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5730
      Left            =   150
      Picture         =   "Design.frx":0000
      ScaleHeight     =   5730
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   120
      Width           =   7185
      Begin ComctlLib.Toolbar Toolbar 
         Height          =   390
         Index           =   0
         Left            =   5400
         TabIndex        =   11
         Top             =   450
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         ButtonWidth     =   609
         AllowCustomize  =   0   'False
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Sample Toolbar 1"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Sample Toolbar 2"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Sample Toolbar 3"
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.Slider Slider 
         Height          =   585
         Index           =   0
         Left            =   5490
         TabIndex        =   12
         Top             =   3060
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1032
         _Version        =   327682
         LargeChange     =   1
         Max             =   1
         TickStyle       =   2
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   0
         Left            =   1500
         TabIndex        =   9
         Top             =   5370
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin ComctlLib.StatusBar Status 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   5340
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   556
         SimpleText      =   "Status Bar"
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   1
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "Status Bar"
               TextSave        =   "Status Bar"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.DirListBox Dir 
         Height          =   3240
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   450
         Width           =   5145
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "&Help"
         Height          =   315
         Index           =   2
         Left            =   5490
         MouseIcon       =   "Design.frx":1021
         TabIndex        =   6
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "&Cancel"
         Height          =   315
         Index           =   1
         Left            =   5490
         MouseIcon       =   "Design.frx":1173
         TabIndex        =   5
         Top             =   4170
         Width           =   1455
      End
      Begin VB.OptionButton Radio 
         Caption         =   "Open as &Copy"
         Height          =   225
         Index           =   0
         Left            =   3840
         TabIndex        =   4
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CheckBox CheckBox 
         Caption         =   "&Dont Display in future"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   0
         Left            =   150
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Text"
         Top             =   3810
         Width           =   5145
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "&Open"
         Height          =   315
         Index           =   0
         Left            =   5490
         MouseIcon       =   "Design.frx":12C5
         TabIndex        =   1
         Top             =   3780
         Width           =   1455
      End
      Begin ComctlLib.TabStrip TabStrip 
         Height          =   5265
         Index           =   0
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   9287
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "TabStrip"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Design"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub CheckBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(CheckBox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List CheckBox(Index), 0
End Sub

Private Sub CheckBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(CheckBox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List CheckBox(Index), 1
End Sub

Private Sub CommandButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(CommandButton(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List CommandButton(Index), 0
End Sub

Private Sub CommandButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(CommandButton(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List CommandButton(Index), 1
End Sub

Private Sub Dir_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Dir(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Dir(Index), 0
End Sub

Private Sub Dir_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Dir(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Dir(Index), 1
End Sub

Private Sub Form_Load()
Me.Icon = Nothing
IDE.Icon = Nothing
SendMessage CommandButton(Index).hWnd, &HF4&, &H0&, 0&
DesignForm.AutoSize = True
Form_Resize
Progress(Index).Value = Progress(Index).Max
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then: Exit Sub
DesignForm.Top = 0
DesignForm.Left = 0
End Sub

Private Sub Progress_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Progress(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Progress(Index), 0
End Sub

Private Sub Progress_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Progress(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Progress(Index), 1
End Sub

Private Sub Radio_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Radio(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Radio(Index), 0
End Sub

Private Sub Radio_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Radio(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Radio(Index), 1
End Sub

Private Sub Slider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Slider(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  Add2List Slider(Index), 0
End Sub

Private Sub Slider_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Slider(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  Add2List Slider(Index), 1
End Sub

Private Sub Status_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Status(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Status(Index), 0
End Sub

Private Sub Status_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Status(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  Add2List Status(Index), 1
End Sub

Private Sub Tab_Click(Index As Integer)

End Sub

Private Sub Tabstrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(TabStrip(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
      Add2List TabStrip(Index), 0
End Sub

Private Sub TabStrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(TabStrip(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List TabStrip(Index), 1
End Sub

Private Sub Text_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Text(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Text(Index), 0
End Sub

Private Sub Text_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Text(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Text(Index), 1
End Sub

Private Sub tmrLoad_Timer()
IDE.Visible = True
End Sub

Private Sub Toolbar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Toolbar(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
    Add2List Toolbar(Index), 0
End Sub

Private Sub Toolbar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Toolbar(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  Add2List Toolbar(Index), 1
End Sub

Private Function Add2List(ctlName As Control, Events As Integer)
If Events = 0 Then: lstEvent.AddItem ctlName.Name & ctlName.Index & " clicked": Exit Function
If Events = 1 Then: lstEvent.AddItem ctlName.Name & ctlName.Index & "Top: " & ctlName.Top & ", Left: " & ctlName.Left: Exit Function
End Function
