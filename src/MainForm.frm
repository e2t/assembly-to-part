VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Сохранить сборки как деталь"
   ClientHeight    =   9180.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12705
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    ExitApp
End Sub

Private Sub btnRun_Click()
    RunAndExit
End Sub

Private Sub UserForm_Initialize()
    '''The resize must be first!
    Me.Width = MaximizedWidth
    Me.Height = MaximizedHeight
    '''The resize must be first!
End Sub

Private Sub UserForm_Resize()
    Me.btnCancel.Left = Me.Width - 81
    Me.btnCancel.Top = Me.Height - 51
    
    Me.btnRun.Left = Me.btnCancel.Left - 78
    Me.btnRun.Top = Me.btnCancel.Top

    Me.txtTargetDir.Width = Me.Width - 180
    Me.txtTargetDir.Top = Me.Height - 45
    
    Me.labTargetDir.Left = Me.txtTargetDir.Left
    Me.labTargetDir.Top = Me.txtTargetDir.Top - 12
    
    Me.lstConfs.Width = Me.Width - 15
    Me.lstConfs.Height = Me.Height - 65
End Sub

Private Sub lstConfs_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode

        Case vbKeyReturn
            If IsShiftPressed Or IsCtrlPressed Then
                RunAndExit
            End If

        Case vbKeyA
            If IsCtrlPressed Then
                SetSelectionAllRows
            End If
            
        Case vbKeyTab
            If IsCtrlPressed Then
                InvertSelectionAllRows
            End If

    End Select
End Sub
