VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PrintFax 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Fax"
   ClientHeight    =   2265
   ClientLeft      =   6885
   ClientTop       =   5730
   ClientWidth     =   3420
   Icon            =   "Printfax.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   228
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Print Fax"
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin MSComCtl2.UpDown udPage 
         Height          =   315
         Left            =   2490
         TabIndex        =   11
         Top             =   1800
         Width           =   195
         _ExtentX        =   318
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPage"
         BuddyDispid     =   196611
         OrigLeft        =   720
         OrigTop         =   2130
         OrigRight       =   915
         OrigBottom      =   2445
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtPage 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2070
         TabIndex        =   10
         Text            =   "1"
         Top             =   1800
         Width           =   400
      End
      Begin VB.OptionButton optIndividualPage 
         Caption         =   "Print Individual Page"
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton optAllPages 
         Caption         =   "Print All Pages"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   1470
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox txtCopies 
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Text            =   "1"
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   315
         Left            =   2430
         TabIndex        =   4
         Top             =   660
         Width           =   855
      End
      Begin VB.ComboBox coPrinter 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   315
         Left            =   2430
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1140
         TabIndex        =   12
         Top             =   720
         Width           =   195
         _ExtentX        =   318
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196615
         OrigLeft        =   720
         OrigTop         =   2130
         OrigRight       =   915
         OrigBottom      =   2445
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Copies"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   450
      End
   End
End
Attribute VB_Name = "PrintFax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Files() As String

Private Sub cmdExit_Click()
      If cmdExit.Caption = "Exit" Then
         Me.Hide
      Else
         Printer.KillDoc
         lblStatus.Caption = "Document Cancelled"
      End If
End Sub

Private Sub cmdPrint_Click()
Dim aprinter As Printer
Dim Returned As Integer
Dim iPageNo As Integer
Dim j As Integer
Dim i As Integer

      lblStatus.ForeColor = vbBlack
      cmdExit.Enabled = False
        
      For Each aprinter In Printers
         If aprinter.DeviceName = coPrinter.Text Then
            Set Printer = aprinter
         End If
      Next
      Printer.Copies = Val(txtCopies)
      Printer.ScaleMode = 5
      On Error GoTo PrinterError
      If optAllPages.Value = True Then
         If bMultipage = True Then
            DoEvents
            frmMain.imgPage.PrintImage , , 2, False, Printer.DeviceName, Printer.DriverName, Printer.Port
         Else
            For i = 1 To iPages
               lblStatus.Caption = "Loading Page " & i
               DoEvents
               frmMain.imgPage.Image = sFiles(i)
               frmMain.imgPage.Display
               frmMain.imgPage.Refresh
               frmMain.imgPage.PrintImage 1, 1, 2, False, Printer.DeviceName, Printer.DriverName, Printer.Port
            Next
            frmMain.LoadPage
         End If
      Else
         DoEvents
         If bMultipage = True Then
            frmMain.imgPage.PrintImage CInt(txtPage), CInt(txtPage), 2, False, Printer.DeviceName, Printer.DriverName, Printer.Port
         Else
            lblStatus.Caption = "Loading Page " & i
            DoEvents
            frmMain.imgPage.Image = sFiles(txtPage)
            frmMain.imgPage.Display
            frmMain.imgPage.Refresh
            frmMain.imgPage.PrintImage 1, 1, 2, False, Printer.DeviceName, Printer.DriverName, Printer.Port
         End If
         frmMain.LoadPage
      End If
      cmdExit.Enabled = True
      Printer.EndDoc
      DoEvents
      lblStatus.Caption = "Document Printed"
      DoEvents
      Exit Sub
   
PrinterError:
      Printer.KillDoc
      lblStatus.ForeColor = vbRed
      lblStatus.Caption = "Printer Error"
      cmdExit.Enabled = True

End Sub
Private Sub coPrinter_KeyPress(KeyAscii As Integer)
      coPrinter.Text = coPrinter.Text
End Sub
Private Sub Form_Load()
Dim aprinter As Printer
Dim Returned As Integer
Dim i As Integer
Dim StrTemp As String
Dim lpBuf As String * 255
Dim flag As Boolean
Dim foo As Integer

      coPrinter.Clear
      For Each aprinter In Printers
         If StrComp(aprinter.DeviceName, "Rendering Subsystem") Then
            coPrinter.AddItem aprinter.DeviceName
         End If
      Next
      If coPrinter.ListCount <= 0 Then
         Returned = MsgBox("No Printers Found", 0)
         Me.Hide
         Exit Sub
      End If
      coPrinter.ListIndex = 0
      lblStatus.Caption = "No. of Pages = " & iPages
      udPage.Max = iPages
      udPage.Min = 1
      GoTo Exit_PrintFax
      
Err_PrintFax:
      MsgBox "No Pages To Print", vbOKOnly
      Unload Me
      End
      
Exit_PrintFax:
    
End Sub
Private Sub optAllPages_Click()
      If optAllPages.Value = True Then
         txtPage.Enabled = False
         udPage.Enabled = False
      End If
End Sub
Private Sub optIndividualPage_Click()
      If optIndividualPage.Value = True Then
         txtPage.Enabled = True
         udPage.Enabled = True
      End If
End Sub
Private Sub txtCopies_Change()
      If Val(txtCopies) Then txtCopies = Val(txtCopies)
      If Val(txtCopies) = 0 Then txtCopies = ""
End Sub
Private Sub txtCopies_Validate(Cancel As Boolean)
      If Val(txtCopies) = 0 Then Cancel = True
End Sub
Private Sub txtPage_Change()
      If Val(txtPage) Then txtPage = Val(txtPage)
      If Val(txtPage) = 0 Then txtPage = "1"
      If Val(txtPage) > iPages Then txtPage = iPages
End Sub
Private Sub txtPage_Validate(Cancel As Boolean)
      If Val(txtPage) = 0 Then Cancel = True
End Sub
