VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgras 
   Caption         =   "进度"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   OleObjectBlob   =   "frmProgras.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmProgras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim giMax As Integer
Dim giSkip As Integer
Dim giCurrentPos As Double
Dim gdScale As Double
Dim gdBarLength As Double

Private m_iPauseOrBreak As Integer '0:正常，1:暂停，2:中断


Sub initProgras(argMax As Integer, Optional argStart As Integer = 1, Optional argInfo As String = "")
  giMax = argMax
  giCurrentPos = argStart
  m_iPauseOrBreak = 0
  With Me
    .lbBar.Width = argStart
    gdBarLength = Me.lbBack.Width
    gdScale = gdBarLength / argMax
    .lbMax.Caption = CStr(argMax)
    .lbSkip.Caption = CStr(argStart)
    .lbInfo.Caption = argInfo
  End With
  switchBt True
End Sub

Sub goSkip(argSkip As Integer, argInfo As String)
  With Me
    .lbBar.Width = gdScale * argSkip
    .lbSkip.Caption = CStr(argSkip)
    .lbInfo.Caption = argInfo
  End With
End Sub

Sub showMe(Optional argMod As Integer = vbModeless)
  'switchBt True
  frmProgras.Show argMod
End Sub

Sub closeMe()
  Unload frmProgras
End Sub


Private Sub btClose_Click()
  closeMe
End Sub

Sub switchBt(argRunning As Boolean)
  With Me
      .btBreak.Enabled = argRunning
      .btPause.Enabled = argRunning
  End With
End Sub

Private Sub btPause_Click()
  PauseOrBreak = 1
End Sub

Private Sub btBreak_Click()
  PauseOrBreak = 2
End Sub


Public Property Get PauseOrBreak() As Integer

    PauseOrBreak = m_iPauseOrBreak

End Property

Public Property Let PauseOrBreak(ByVal argPauseOrBreak As Integer)

    m_iPauseOrBreak = argPauseOrBreak
    switchBt m_iPauseOrBreak = 0
End Property


