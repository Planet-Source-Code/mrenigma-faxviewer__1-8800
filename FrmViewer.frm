VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#1.0#0"; "IMGEDIT.OCX"
Begin VB.Form frmMain 
   Caption         =   " Main Window"
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "FrmViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10965
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Tag             =   " Main Window"
   Begin VB.PictureBox ControlContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   9795
      TabIndex        =   11
      Top             =   0
      Width           =   9825
      Begin VB.CommandButton cmdButton 
         Caption         =   "Print Fax"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   9
         Left            =   8970
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "PRINT"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Exit Viewer"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   8
         Left            =   7980
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":11DC
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "EXIT"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Rotate Right"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   7
         Left            =   6930
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":1AA6
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "ROTATERIGHT"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Rotate Left"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   6
         Left            =   5940
         Picture         =   "FrmViewer.frx":2970
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "ROTATELEFT"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Fit Screen"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   5
         Left            =   4950
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":383A
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "FITSCREEN"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Fit Page"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   4
         Left            =   3960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":4704
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "FITPAGE"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Zoom Out"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   3
         Left            =   2970
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":55CE
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "ZOOMOUT"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Zoom In"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   2
         Left            =   1980
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":6498
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "ZOOMIN"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Next Page"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   1
         Left            =   990
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":7362
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "NEXTPAGE"
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Prev Page"
         CausesValidation=   0   'False
         Height          =   1005
         Index           =   0
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmViewer.frx":822C
         Style           =   1  'Graphical
         TabIndex        =   0
         Tag             =   "PREVPAGE"
         Top             =   0
         Width           =   1000
      End
   End
   Begin ImgeditLibCtl.ImgEdit imgPage 
      CausesValidation=   0   'False
      Height          =   7275
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   6525
      _Version        =   65536
      _ExtentX        =   11509
      _ExtentY        =   12832
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      SelectionRectangleEnabled=   0   'False
      AutoRefresh     =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdButton_Click(Index As Integer)
      On Error GoTo ErrBeep
      Select Case UCase(cmdButton(Index).Tag)
         Case "NEXTPAGE"
            iCurrentPage = iCurrentPage + 1
            If iCurrentPage > iPages Then
               iCurrentPage = iPages
            Else
               Call LoadPage
            End If
         Case "PREVPAGE"
            iCurrentPage = iCurrentPage - 1
            If iCurrentPage < 1 Then
               iCurrentPage = 1
            Else
               Call LoadPage
            End If
         Case "ZOOMIN"
            ChangeZoom imgPage, 2
         Case "ZOOMOUT"
            ChangeZoom imgPage, -2
         Case "FITPAGE"
            imgPage.FitTo 0
         Case "FITSCREEN"
            imgPage.FitTo 1
         Case "ROTATELEFT"
            imgPage.RotateLeft
            imgPage.FitTo 0
         Case "ROTATERIGHT"
            imgPage.RotateRight
            imgPage.FitTo 0
         Case "EXIT"
            Unload Me
            End
         Case "PRINT"
            PrintFax.Show 1, Me
            Unload PrintFax
            Set PrintFax = Nothing
      End Select
      Call UpdateTitle
      Exit Sub
ErrBeep:
      On Error GoTo 0
      Beep
End Sub
Sub LoadPage()
      ZoomPercent = imgPage.Zoom
      If bMultipage Then
         imgPage.Page = iCurrentPage
      Else
         imgPage.Image = sFiles(iCurrentPage)
      End If
      imgPage.Display
      imgPage.Refresh
      Call UpdateTitle
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim sTemp As String
        
      Call StartupWindowpos(Me, , True, True, , True)
      ' Setup buttons
'      Me.ControlContainer.Width = cmdButton(0).Width
'      For i = 1 To cmdButton.UBound
'         cmdButton(i).Width = cmdButton(i - 1).Width
'         cmdButton(i).Left = cmdButton(i - 1).Left + cmdButton(i).Width
'         ControlContainer.Width = ControlContainer.Width + cmdButton(i).Width
'      Next
      ControlContainer.Left = 0
      If Me.Width < ControlContainer.Width Then
         Me.Width = ControlContainer.Width + 120
      End If
        
      ' Setup List of files (if .bmp Files)
      i = 0
      iCurrentPage = 1
      ZoomPercent = 100
      sTemp = Dir(App.Path & "\*.tif")
      If Len(sTemp) Then
         bMultipage = True
         ReDim Preserve sFiles(1) As String
         sFiles(1) = App.Path & "\" & sTemp
         imgPage.Image = sFiles(1)
         iPages = imgPage.PageCount
         Me.Show
         DoEvents
         Call LoadPage
         imgPage.FitTo 1
         imgPage.Refresh
      Else
         sTemp = Dir(App.Path & "\Page*.*")
         If Len(sTemp) Then
            ReDim Preserve sFiles(1) As String
            While Len(sTemp)
               i = i + 1
               ReDim Preserve sFiles(UBound(sFiles()) + 1) As String
               sFiles(i) = App.Path & "\" & sTemp
               sTemp = Dir
            Wend
            iPages = UBound(sFiles()) - 1
            Call LoadPage
            imgPage.FitTo 1
            imgPage.Display
            imgPage.Refresh
         Else
            MsgBox "No Faxes Found", vbOKOnly + vbExclamation, "Fax Viewer"
            Unload Me
            End
         End If
      End If
      DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      ' Save Display Settings then Exit
      Call StartupWindowpos(Me, True, True, True, , True)
End Sub
Private Sub Form_Resize()

      On Error Resume Next
      ' If Me.Width < ControlContainer.Width Then
      ' Me.Width = ControlContainer.Width
      ' End If
      With imgPage
         .Width = Me.ScaleWidth - 1
         .Height = Me.ScaleHeight - ControlContainer.Height - 3
      End With
      imgPage.FitTo 1
      On Error GoTo 0
End Sub

Private Sub Image1_Click()

End Sub
