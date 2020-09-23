VERSION 5.00
Begin VB.MDIForm IDE 
   BackColor       =   &H8000000C&
   Caption         =   "Visual Dialog++"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8025
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuhelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "IDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Design.Show
End Sub

Private Sub mnuabout_Click()
Dim Aboutstr As String
Aboutstr = Aboutstr & "Visual Dialog++ for Windows" & vbCrLf
Aboutstr = Aboutstr & "Version 1.0.0 Build (0?)" & vbCrLf & vbCrLf
Aboutstr = Aboutstr & "Copyright (C) 2003, Shukri Zahari" & vbCrLf
Aboutstr = Aboutstr & "" & vbCrLf
Aboutstr = Aboutstr & "Any comments, suggestions or vote, please send it to me..." & vbCrLf
Aboutstr = Aboutstr & "Ah.. don't forget to vote me on PSC.com, OK?"
MsgBox Aboutstr, vbInformation, "About"
End Sub

Private Sub mnuhelp_Click()
Dim Helpstr As String
Helpstr = Helpstr & "Click on the controls & move them around the form." & vbCrLf
Helpstr = Helpstr & "They will move according to your mouse like VB does." & vbCrLf
Helpstr = Helpstr & "You can code custom event or their position if you like to..."
MsgBox Helpstr, vbInformation, "Help"
End Sub
