Attribute VB_Name = "modFunctions"
Option Explicit

Global ZoomPercent As Long
Global Rotation As Long
Global Page As Long
Global bMultipage As Boolean
Global sFiles() As String
Global iPages As Integer
Global iCurrentPage As Integer

Public Sub main()
      Load frmMain
      DoEvents
      frmMain.Show
      If frmMain.WindowState = vbMinimized Then
         frmMain.WindowState = vbNormal
      End If
End Sub
Function ChangeZoom(ImgControl As ImgEdit, Offset As Long)
      On Error GoTo ErrBeep:
      ZoomPercent = ImgControl.Zoom
      ZoomPercent = ZoomPercent + Offset
    
      If ZoomPercent < 2 Then
         Beep
         ZoomPercent = 2
      End If
      If ZoomPercent > 500 Then
         Beep
         ZoomPercent = 500
      End If
      ImgControl.Zoom = ZoomPercent
      ImgControl.Refresh
      Call UpdateTitle
      Exit Function
ErrBeep:
      On Error GoTo 0
      Beep
End Function
Function ChangePage(ImgControl As ImgEdit, Offset As Long)
      On Error GoTo ErrBeep:
      Page = ImgControl.Page
      Page = Page + Offset
      If Page < 1 Then
         Beep
         Page = 1
         Exit Function
      End If
      If Page > ImgControl.PageCount Then
         Beep
         Page = ImgControl.PageCount
         Exit Function
      End If
      ImgControl.Page = Page
      ImgControl.Display
      Call UpdateTitle
      Exit Function
ErrBeep:
      On Error GoTo 0
      Beep
End Function
Public Sub UpdateTitle()
      If bMultipage Then
         frmMain.Caption = "Fax Viewer - Page " & iCurrentPage & " Of " & frmMain.imgPage.PageCount & " - Current Zoom (" & frmMain.imgPage.Zoom & "%)"
      Else
         frmMain.Caption = "Fax Viewer - Page " & iCurrentPage & " Of " & iPages & " - Current Zoom (" & frmMain.imgPage.Zoom & "%)"
      End If
End Sub

Public Sub StartupWindowpos(f As Form, _
      Optional bSave As Boolean = False, _
      Optional bWidth As Boolean = False, _
      Optional bHeight As Boolean = False, _
      Optional bVisible As Boolean = False, _
      Optional bWindowState As Boolean)
    
Dim sSection As String

      sSection = "Window Startup\" & f.Tag

      If bSave Then
         If bWindowState Then
            Call SaveSetting(App.Comments, sSection, "State", f.WindowState)
         End If
         If f.WindowState <> vbMinimized Then
            Call SaveSetting(App.Comments, sSection, "Top", f.Top)
            Call SaveSetting(App.Comments, sSection, "Left", f.Left)
            Call SaveSetting(App.Comments, sSection, "State", f.WindowState)
            If bWidth Then
               Call SaveSetting(App.Comments, sSection, "Width", f.Width)
            End If
            If bHeight Then
               Call SaveSetting(App.Comments, sSection, "Height", f.Height)
            End If
         End If
         If bVisible Then
            Call SaveSetting(App.Comments, sSection, "Visible", f.Visible)
         End If

      Else
         f.Top = GetSetting(App.Comments, sSection, "Top", f.Top)
         f.Left = GetSetting(App.Comments, sSection, "Left", f.Left)
         f.WindowState = GetSetting(App.Comments, sSection, "State", vbNormal)
         If bWidth Then
            f.Width = GetSetting(App.Comments, sSection, "Width", f.Width)
         End If
         If bHeight Then
            f.Height = GetSetting(App.Comments, sSection, "Height", f.Height)
         End If
         If bVisible Then
            f.Visible = GetSetting(App.Comments, sSection, "Visible", f.Visible)
         End If
         If bWindowState Then
            f.WindowState = GetSetting(App.Comments, sSection, "State", f.WindowState)
         End If
      End If
End Sub


