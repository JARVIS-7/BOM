VERSION 5.00
Begin VB.Form frmCountdown 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CountDown"
   ClientHeight    =   3030
   ClientLeft      =   6540
   ClientTop       =   5685
   ClientWidth     =   5550
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmCountdown.frx":0000
   LinkTopic       =   "frmCountdown"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCountdown.frx":000C
   ScaleHeight     =   3030
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSep 
      Height          =   45
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   4575
   End
   Begin VB.CommandButton btnAction 
      Caption         =   "Abbrechen"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton btnAction 
      Caption         =   "Sofort"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Timer tmrCountDown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2400
   End
   Begin VB.Label lblZeit 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2905
      TabIndex        =   4
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   630
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "heruntergefahren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3515
      TabIndex        =   5
      Top             =   1970
      UseMnemonic     =   0   'False
      Width           =   2050
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Sekunden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   1970
      UseMnemonic     =   0   'False
      Width           =   1890
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "der Rechner wird in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   20
      UseMnemonic     =   0   'False
      Width           =   4560
   End
End
Attribute VB_Name = "frmCountdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'
' BOM default header
'
' this file is OpenSource.
' license model: GPL
' please respect the limitations
'
' main language: german
' compiled under VB6 SP5 german
'
' $author: hjs$
' $id: V 2.0.2 date 030303 hjs$
' $version: 2.0.2$
' $file: $
'
' last modified:
' &date: 030303$
'
' contact: visit http://de.groups.yahoo.com/group/BOMInfo
'
'*******************************************************
Option Explicit
Private mlRestzeit As Integer
Private mbFrmRetValue As Boolean

Public Enum FrmTypEnum
    [ftCountDown1] = 0& 'CountDown mit 1 Button
    [ftCountDown2] = 1& ' dto. mit 2 Button´s
    [ftUpdateBox] = 2&
End Enum

Private Type udtInitFrm
    FrmTyp As FrmTypEnum
    IconResNr As Long
    FrmPos As Long
    FrmTitle As String
    Cnt As Long
    Txt1 As String
    Txt2 As String
    Txt3 As String
    Btn1 As String
    Btn2 As String
    DefRetValue As Boolean
End Type
Private mudtInitFrm As udtInitFrm

Private Sub btnAction_Click(Index As Integer)
    
    tmrCountDown.Enabled = False
    mbFrmRetValue = CBool(Index = 0) '0= Sofort , 1= Abbrechen
    Me.Hide
           
End Sub

Private Sub Form_Load()
    
    With mudtInitFrm
        Me.Caption = .FrmTitle
        Set Me.Icon = MyLoadResPicture(.IconResNr, 16)
        Call SendMessage(Me.hWnd, WM_SETICON, 0, Me.Icon)
        Call PositionSetzen(.FrmPos)
        '...
        lblMsg(0).Caption = .Txt1
        lblMsg(1).Caption = .Txt2
        lblMsg(2).Caption = .Txt3
        btnAction(0).Caption = .Btn1
        btnAction(1).Caption = .Btn2
        lblZeit.Caption = CStr(.Cnt)
        mlRestzeit = .Cnt
        mbFrmRetValue = .DefRetValue
    End With
        
    Call AdjustCtrls
        
    With tmrCountDown
        .Interval = 1000
        .Enabled = True
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    With tmrCountDown
        .Interval = 0
        .Enabled = False
    End With
    
    If UnloadMode = vbFormControlMenu Then
        mbFrmRetValue = False
        Me.Hide
    Else
        Unload Me: Set frmCountdown = Nothing
    End If
    
End Sub
Private Sub PositionSetzen(lPosition As Long)
        
    With Me
        Select Case lPosition
            Case 1 'Bildschirmmitte
                Call .Move((Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2, .Width, .Height)
            Case 2 'unten Rechts
                If GetTaskBarProps("ALIGN") = 2 Then
                    Call .Move((GetTaskBarProps("LEFT") * Screen.TwipsPerPixelX) - .Width, .Top, .Width, .Height)
                Else
                    Call .Move(Screen.Width - .Width, .Top, .Width, .Height)
                End If
                
                If GetTaskBarProps("ALIGN") = 4 Then
                    Call .Move(.Left, (GetTaskBarProps("TOP") * Screen.TwipsPerPixelY) - .Height, .Width, .Height)
                Else
                    Call .Move(.Left, Screen.Height - .Height, .Width, .Height)
                End If
            Case Else
        End Select
    End With
    
End Sub

Private Sub tmrCountDown_Timer()
        
    mlRestzeit = mlRestzeit - 1
    lblZeit.Caption = CStr(mlRestzeit) & " "
    If mlRestzeit < 0 Then
        tmrCountDown.Enabled = False
        Me.Hide
    End If
    
End Sub
Public Property Get FrmRetValue() As Boolean
FrmRetValue = mbFrmRetValue
End Property

Public Sub InitFrm(eFrmTyp As FrmTypEnum, lIconResNr As Long, lFrmPos, sFrmTitle As String, lCnt As Long, sTxt1 As String, sTxt2 As String, sTxt3 As String, sBtn1 As String, sBtn2 As String, bDefRetValue As Boolean)

    With mudtInitFrm
        .FrmTyp = eFrmTyp
        .IconResNr = lIconResNr
        .FrmPos = lFrmPos
        .FrmTitle = sFrmTitle
        .Cnt = lCnt
        .Txt1 = sTxt1
        .Txt2 = sTxt2
        .Txt3 = sTxt3
        .Btn1 = sBtn1
        .Btn2 = sBtn2
        .DefRetValue = bDefRetValue
    End With
End Sub

Private Sub AdjustCtrls()
    
    Select Case mudtInitFrm.FrmTyp
        Case [ftCountDown1], [ftCountDown2]
            If mudtInitFrm.FrmTyp = [ftCountDown1] Then
                btnAction(0).Visible = False
                With btnAction(1)
                    Call .Move(((fraSep.Width - .Width) \ 2) + fraSep.Left, .Top, .Width, .Height)
                    .Refresh
                End With
            End If
            '...
            With lblMsg(0)
                If InDevelopment Then
                    .BorderStyle = vbFixedSingle
                Else
                    .BorderStyle = vbBSNone
                End If
                .Alignment = vbCenter
                .Font.Size = 14
                .Font.bold = True
                .AutoSize = True
                .AutoSize = False
                Call .Move(1000, Me.ScaleHeight \ 8, Me.ScaleWidth - 1000, .Height)
                .Refresh
                '...
                Call lblZeit.Move(.Left + 1120, .Top + (1.3 * .Height), lblZeit.Width, lblZeit.Height)
                lblZeit.Refresh
                '...
                lblMsg(1).BorderStyle = .BorderStyle
                lblMsg(1).Font.Size = .Font.Size
                lblMsg(1).Font.bold = .Font.bold
                lblMsg(1).AutoSize = True
                Call lblMsg(1).Move(lblZeit.Left + (lblZeit.Width + 80), lblZeit.Top + 40, lblMsg(1).Width, lblMsg(1).Height)
                lblMsg(1).Refresh
    
                lblMsg(2).BorderStyle = .BorderStyle
                lblMsg(2).Alignment = .Alignment
                lblMsg(2).Font.Size = .Font.Size
                lblMsg(2).Font.bold = .Font.bold
                lblMsg(2).AutoSize = True
                lblMsg(2).AutoSize = False
                Call lblMsg(2).Move(.Left, lblZeit.Top + (1.3 * lblZeit.Height), .Width, lblMsg(2).Height)
                lblMsg(2).Refresh
            End With
            '...
        Case Else
    End Select
    
End Sub
