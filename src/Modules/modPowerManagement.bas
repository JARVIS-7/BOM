Attribute VB_Name = "modPowerManagement"
Option Explicit

Private mdatCurrWakeUpTime As Date
Private mlHTimer As Long


Private Declare Function CreateWaitableTimer Lib "kernel32" _
    Alias "CreateWaitableTimerA" ( _
    ByVal lpSemaphoreAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal lpName As String) As Long

'Private Declare Function OpenWaitableTimer Lib "kernel32" _
    Alias "OpenWaitableTimerA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long

Private Declare Function SetWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long, _
    lpDueTime As FILETIME, _
    ByVal lPeriod As Long, _
    ByVal pfnCompletionRoutine As Long, _
    ByVal lpArgToCompletionRoutine As Long, _
    ByVal fResume As Long) As Long

'Private Declare Function CancelWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long)

Private Declare Function IsSystemResumeAutomatic Lib "kernel32.dll" () As Long

'Private Declare Function SetSystemPowerState Lib "kernel32.dll" ( _
    ByVal fSuspend As Long, _
    ByVal fForce As Long) As Long
                 
Private Declare Function SetSuspendState Lib "Powrprof" ( _
    ByVal Hibernate As Long, _
    ByVal ForceCritical As Long, _
    ByVal DisableWakeEvent As Long) As Long
                         
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const ERROR_ALREADY_EXISTS As Long = 183&
'
'Private Const ES_CONTINUOUS As Long = &H80000000
Private Const ES_SYSTEM_REQUIRED As Long = &H1&
'Private Const ES_DISPLAY_REQUIRED As Long = &H2&
'
'
Public Const WM_POWERBROADCAST As Long = &H218
Public Const PBT_APMRESUMEAUTOMATIC As Long = &H12
Public Const PBT_APMRESUMESUSPEND As Long = &H7
Public Const PBT_APMSUSPEND As Long = &H4

Private Declare Function SetThreadExecutionState Lib "kernel32.dll" (ByVal esFlags As Long) As Long

Public Sub ResetSystemIdleTimer()
    
    If giPreventSuspend > 0 Then
        Call SetThreadExecutionState(ES_SYSTEM_REQUIRED)
    End If
    
End Sub

Public Sub ResetWakupTime()
    
    Dim lRet As Long
    Dim ft As FILETIME
    
    If giWakeOnAuction > 0 Then
        If mlHTimer <> 0 And mdatCurrWakeUpTime <> 0 Then
            'Close the handles when you are done with them.
            mlHTimer = CreateWaitableTimer(0, True, App.EXEName & "_Wakeup_Timer_" & CStr(glThreadID))
            ft.dwLowDateTime = -1
            ft.dwHighDateTime = -1
            lRet = SetWaitableTimer(mlHTimer, ft, 0, 0, 0, True)
            mdatCurrWakeUpTime = 0
            Call DebugPrint("Aufwachzeit gelöscht", 2)
            'frmHaupt.PanelText frmHaupt.StatusBar1, 1, "Aufwachzeit gelöscht", True
        End If
    End If
    
End Sub
Public Sub SetWakeupTime(datNewWakeupTime As Date)
  Dim ft As FILETIME
  Dim lRet As Long
  Dim dblDelay As Double
  Dim dblDelayLow As Double
  Dim dblUnits As Double
  Dim lSecondsToGo As Long
    
  If giWakeOnAuction > 0 Then
    If (mdatCurrWakeUpTime <> 0 And Abs(DateDiff("s", mdatCurrWakeUpTime, datNewWakeupTime)) < 3) Or datNewWakeupTime < MyNow Then
      Exit Sub
    End If
    mdatCurrWakeUpTime = datNewWakeupTime
    
    lSecondsToGo = DateDiff("s", MyNow, mdatCurrWakeUpTime)
        
    mlHTimer = CreateWaitableTimer(0, True, App.EXEName & "_Wakeup_Timer_" & CStr(glThreadID))
    'DebugPrint "hTimer: " & mlHTimer
    If Err.LastDllError = ERROR_ALREADY_EXISTS Then
      ' If the timer already exists, it does not hurt to open it
      ' as long as the person who is trying to open it has the
      ' proper access rights.
    Else
      ft.dwLowDateTime = -1
      ft.dwHighDateTime = -1
      lRet = SetWaitableTimer(mlHTimer, ft, 0, 0, 0, True)
    End If

    ' Convert the Units to nanoseconds.
    dblUnits = CDbl(&H10000) * CDbl(&H10000)
    dblDelay = CDbl(lSecondsToGo) * 1000 * 10000
    ' By setting the high/low time to a negative number, it tells
    ' the Wait (in SetWaitableTimer) to use an offset time as
    ' opposed to a hardcoded time. If it were positive, it would
    ' try to convert the value to GMT.
    ft.dwHighDateTime = -CLng(dblDelay / dblUnits) - 1
    dblDelayLow = -dblUnits * (dblDelay / dblUnits - Fix(dblDelay / dblUnits))

    If dblDelayLow < CDbl(&H80000000) Then
      ' &H80000000 is MAX_LONG, so you are just making sure
      ' that you don't overflow when you try to stick it into
      ' the FILETIME structure.
      dblDelayLow = dblUnits + dblDelayLow
      ft.dwHighDateTime = ft.dwHighDateTime + 1
    End If

    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlHTimer, ft, 0, 0, 0, True)
    'DebugPrint "lRet: " & lRet

    DebugPrint "Neue Aufwachzeit: " & Date2Str(datNewWakeupTime), 2
'    frmHaupt.PanelText frmHaupt.StatusBar1, 1, "Neue Aufwachzeit: " & Date2Str(datNewWakeupTime), True
  End If

End Sub

Public Sub Resuspend(ByRef oFrm As Form)
    
    If gbResuspendAfterEnd Then
        If IsSystemResumeAutomatic() > 0 Or gbForceResuspendAfterEnd Then
            'wir geben dem System noch 10 Sekunden, um die Mails rauszuhauen, _
            z.B. Virenscanner können die Zustellung verzögern
            If ShowUpdateBox(oFrm, [ftCountDown1], _
                103, IIf(oFrm.WindowState = vbMinimized Or oFrm.Visible = False, 2, 1), _
                "Suspend", 10, gsarrLangTxt(208), gsarrLangTxt(209), _
                gsarrLangTxt(210), "-", gsarrLangTxt(359), True) Then
                    
                Call Suspend
            End If
        End If
    End If
    
End Sub

Public Sub Suspend()
    
    Call DebugPrint("Suspend!", 2)
    Call NewWndProc(0, WM_POWERBROADCAST, PBT_APMSUSPEND, 0)
    Call SetSuspendState(gbHibernate, 0, 0)
    
End Sub

