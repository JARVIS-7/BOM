VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmInfo 
   ClientHeight    =   2070
   ClientLeft      =   7260
   ClientTop       =   4395
   ClientWidth     =   3090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      ExtentX         =   5318
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbCodeResize As Boolean

Private Sub GetClientSize(ByRef lBWidth As Long, ByRef lTHeight As Long, ByRef lCWidth As Long, ByRef lCHeight As Long)
        
    Dim tRC As RECT
    Dim tPA As POINTAPI
    Dim lMenuHeight As Long
    Dim lMenu As Long
    
    Call GetClientRect(frmHaupt.hWnd, tRC)
    Call ClientToScreen(frmHaupt.hWnd, tPA)
    
    With tRC
        .Left = tPA.X
        .Top = tPA.Y
        tPA.X = .Right
        tPA.Y = .Bottom
        Call ClientToScreen(frmHaupt.hWnd, tPA)
        .Right = tPA.X
        .Bottom = tPA.Y
        lCHeight = (.Bottom - .Top) * Screen.TwipsPerPixelY
        lTHeight = .Top * Screen.TwipsPerPixelY - frmHaupt.Top
        lCWidth = (.Right - .Left) * Screen.TwipsPerPixelX
        lBWidth = (frmHaupt.Width - lCWidth) / 2
    End With
    
    lMenu = GetMenu(frmHaupt.hWnd)
    If lMenu Then
        Call GetMenuItemRect(frmHaupt.hWnd, lMenu, 0, tRC)
        lMenuHeight = (tRC.Bottom - tRC.Top + 1) * Screen.TwipsPerPixelY
        lCHeight = lCHeight + lMenuHeight
        lTHeight = lTHeight - lMenuHeight
    End If
    
End Sub

Private Sub Form_Activate()
        
    On Error Resume Next
    
    If frmHaupt.WindowState <> vbMinimized Then
        Call SetSize
        
        With WebBrowser1
            If Not gsGlobalUrl = .LocationURL Then
                .Stop
                .Navigate "about:blank"
                .Navigate gsGlobalUrl
            End If
        End With
        Me.SetFocus
    End If
    
End Sub

Private Sub Form_Deactivate()

  frmHaupt.HideInfo

End Sub

Private Sub Form_Load()
    
    Call SetSize
    WebBrowser1.Navigate "about:blank"
    
End Sub

Private Sub SetSize()
    
    Dim lBorderWidth As Long
    Dim lTitelHeight As Long
    Dim lClientWidth As Long
    Dim lClientHeight As Long
    
    Call GetClientSize(lBorderWidth, lTitelHeight, lClientWidth, lClientHeight)
    
    mbCodeResize = True
    
    Call Me.Move(frmHaupt.Left + lBorderWidth + (gfInfoLeft / 100 * lClientWidth), _
        frmHaupt.Top + lTitelHeight + (gfInfoTop / 100 * lClientHeight), _
        gfInfoWidth / 100 * lClientWidth, gfInfoHeight / 100 * lClientHeight)
        
    mbCodeResize = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'MD-Marker , 20090406
    WebBrowser1.Stop
    
    Unload Me: Set frmInfo = Nothing
    
End Sub

Private Sub Form_Resize()
    
    Dim lBorderWidth As Long
    Dim lTitelHeight As Long
    Dim lClientWidth As Long
    Dim lClientHeight As Long
        
    Dim fInfoTop As Double
    Dim fInfoLeft As Double
    Dim fInfoHeight As Double
    Dim fInfoWidth As Double
        
    If Not mbCodeResize Then
        
        Call GetClientSize(lBorderWidth, lTitelHeight, lClientWidth, lClientHeight)
        
        fInfoTop = Round(CDbl(Me.Top - frmHaupt.Top - lTitelHeight) / lClientHeight * 100, 3)
        fInfoLeft = Round(CDbl(Me.Left - frmHaupt.Left - lBorderWidth) / lClientWidth * 100, 3)
        fInfoHeight = Round(CDbl(Me.Height) / lClientHeight * 100, 3)
        fInfoWidth = Round(CDbl(Me.Width) / lClientWidth * 100, 3)
         
        If fInfoTop < 1000 And fInfoLeft < 1000 And fInfoHeight > 5 And fInfoWidth > 5 Then
            
            gfInfoTop = fInfoTop
            gfInfoLeft = fInfoLeft
            gfInfoHeight = fInfoHeight
            gfInfoWidth = fInfoWidth
            
        End If
         
    End If
    
    Call WebBrowser1.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)

End Sub
