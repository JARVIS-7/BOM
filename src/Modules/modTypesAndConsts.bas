Attribute VB_Name = "modTypesAndConsts"
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
'Alle Konstanten, Globals und Typen
Option Explicit

Public Const glMINRESICONS As Long = 101&
Public Const glMAXRESICONS As Long = 116&

'XOR- Key
'Private Const C_Key As String = "Ç¦ÚÃÉÆòª¶ÚËÌÍý´Ç¦ÚÃÉÆòª¶ÚËÌÍý´Ç¦ÚÃÉÆòª¶ÚËÌÍý´"
'Für Ebay- und POP-PW
Public Const gsCKEY2 As String = "zedn32jo45fslnczehaö127jncrksl"

'Registrierung
Public Const gsBETASTRING As String = "" ' "beta"
'Private Const BadList As String = "null;"

Public gbShowSplash As Boolean ' False
'Private RegName As String
Public gbAboutIsUp As Boolean
'Private TimeIsUp As Boolean
'Private BestBefore   As String
'Private NervZeit As Integer
Public gbNewBOMVersionAvailable As Boolean

'Allgemeines
Public gbTest As Boolean     'Tesdrucke

Public Const gsBOMUrlHP As String = "https://github.com/JARVIS-7/BOM/"
Public Const gsBOMUrlSF As String = "https://ssl.schnapper.de/"

'Stati
'MD-Marker 20090328 Neu angelegt
Public Enum AuktionStatusEnum
    [asSellerAway] = -4& 'Artikel nicht verfügbar, Verkäufer abwesend
    [asAccessErr] = -3& 'Fehler beim Zugriff auf den Artikel, nur Warnung
    [asPower] = -2& 'Powerauktion, nur Warnung
    [asLowBid] = -1&   'Gebot zu niedrig, nur Warnung
    [asNixLos] = 0&   'nix los
    [asErr] = 1&    'Fehlerhaft beboten
    [asOK] = 2& 'erfolgreich beboten
    [asBieten] = 3& 'Bieten geht los!
    [asCancelGroup] = 4& 'Bieten gecancelt wg. Bietgruppe
    [asCancelBid] = 5&  'Bieten gecancelt wg. fehlendem Gebot
    [asBuyOnly] = 6&    'nur Sofortkaufen möglich
    [asUeberboten] = 7& 'Pech gehabt!
    [asNotFound] = 8& 'Artikel nicht auf dem Server gefunden
    [asHoldGroup] = 9& 'Artikel auf hold wg. Bietgruppe
    [asDelegatedBom] = 10& 'Artikel an anderen BOM delegiert
    [asBuyOnlyOnHold] = 11& 'Sofortkauf-Artikel auf hold wg. Bietgruppe
    [asBuyOnlyCanceled] = 12& 'Sofortkauf-Artikel gecancelt wg. Bietgruppe
    [asBuyOnlyBuyItNow] = 13& 'Sofortkauf-Artikel jetzt kaufen
    [asBuyOnlyDelegated] = 14& 'Sofortkauf-Artikel an anderen BOM delegiert
    [asAdvertisement] = 19& 'Preisanzeige, kein Kauf/Bieten möglich
    [asEnde] = 100& 'Ende ohne Gebote + Abgelaufen
End Enum
'...
'Public Const giSTATSELLERAWAY As Integer = -4 'Artikel nicht verfügbar, Verkäufer abwesend
'Public Const giSTATACCESSERROR As Integer = -3 'Fehler beim Zugriff auf den Artikel, nur Warnung
'Public Const giSTATPOWER As Integer = -2 'Powerauktion, nur Warnung
'Public Const giSTATLOW As Integer = -1 'Gebot zu niedrig, nur Warnung
'Public Const giSTAT0 As Integer = 0   'nix los
'Public Const giSTATERR As Integer = 1 'Fehlerhaft beboten
'Public Const giSTATOK As Integer = 2  'erfolgreich beboten
'Public Const giSTATBIETEN As Integer = 3 'Bieten geht los!
'Public Const giSTATCANCEL As Integer = 4 'Bieten gecancelt wg. Bietgruppe
'Public Const giSTATGEBOT As Integer = 5 'Bieten gecancelt wg. fehlendem Gebot
'Public Const giSTATBUYONLY As Integer = 6 'nur Sofortkaufen möglich
'Public Const giSTATUEBERBOTEN As Integer = 7 'Pech gehabt!
'Public Const giSTATNOTFOUND As Integer = 8 'Artikel nicht auf dem Server gefunden
'Public Const giSTATHOLD As Integer = 9 'Artikel auf hold wg. Bietgruppe
'Public Const giSTATDELEGATED As Integer = 10 'Artikel an anderen BOM delegiert
'Public Const giSTATBUYONLYONHOLD As Integer = 11 'Sofortkauf-Artikel auf hold wg. Bietgruppe
'Public Const giSTATBUYONLYCANCELED As Integer = 12 'Sofortkauf-Artikel gecancelt wg. Bietgruppe
'Public Const giSTATBUYONLYBUYITNOW As Integer = 13 'Sofortkauf-Artikel jetzt kaufen
'Public Const giSTATBUYONLYDELEGATED As Integer = 14 'Sofortkauf-Artikel an anderen BOM delegiert
'Public Const giSTATADVERTISEMENT As Integer = 19 'Preisanzeige, kein Kauf/Bieten möglich
'Public Const giSTATENDE As Integer = 100 'Ende ohne Gebote

'Währungsfaktoren zum Euro, aktualisiert am 21.10.2007
Public Const grWEDEFAULTUSD As Single = 0.69952!
Public Const grWEDEFAULTGBP As Single = 1.43493!
Public Const grWEDEFAULTCHF As Single = 0.59924!
Public Const grWEDEFAULTAUD As Single = 0.62346!
Public Const grWEDEFAULTCAD As Single = 0.72425!

Public gcolWeValues As Collection
Public gcolWeNames As Collection

'WIN- Platformen
Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public gsWinVersion As String
'Private giWinPlatform As Integer

'Fensteraufloesung
'Public Const Hight_Small = 12000
'Public Const Hight_High = 9000
'Public Const Hight_Limit = 24000
'Public Const SourceWidth = 12000

Public Const gfRESTZEITEWIG As Double = 999#
Public Const gdatENDEZEITNOTFOUND As Date = #11/11/2111#

'Variable Laufzeitdaten
Public glThreadID As Long
Public gbLastDialupWasManually As Boolean
Public gbKeepDialupAlive As Boolean
Public giDialupRequestTimeout As Integer
Public gbUsesModem As Boolean
Public gbGeboteAktualisieren As Boolean
Public gbEBayWartung As Boolean
Public glVorlaufGebot As Long 'Sekunden
Public glVorlaufModem As Long 'Minuten
Public gfVorlaufSnipe As Double 'Sekunden
Public gsUser As String
Public gsPass As String
Public gsGlobalUrl As String
Public gbAutoStart As Boolean
Public gsTempPfad As String
Public gbFileWinShutdown As Boolean
Public gbUseWinShutdown As Boolean
'Public gbUseIntBrowser As Boolean 'MD-Marker 20090325 , interner Browser entfernt
Public giVorlaufLan As Integer 'Minuten
Public giReLogin As Integer 'Minuten (mae 050718)
Public gsGlobFontName As String
Public giStartupSize As Integer
Public gbSettingsIsUp As Boolean
Public gbExplicitEnd As Boolean
Public gbStopTests As Boolean
Public gbTrayAction As Boolean
Public gsEbayLocalPass As String
Public gbAutoMode As Boolean
Public gbKeinHinweisNachZeitsync As Boolean
Public gfRestzeitBerechner As Double ' wird benutzt um die Restzeit zu berechnen
Public gfRestzeitZaehler As Double ' nur hier darf die Restzeit ausgelesen werden
Public glFileNrDebug As Long
Public gbAutoLogged As Boolean 'AutoLogin vor Auktion durchgeführt? [mae 050515]
Public gsNextUser As String 'UserAccount, über den der nächste Artikel läuft [mae 050515]
Public gfLzMittel As Double
Public giLastGebotEditedIndex As Integer

'Public gewarnt As String
Public gbWarningflag As Boolean

'1.7.3 UseProxy
Public gbUseProxy As Boolean
Public gsProxyName As String
Public giProxyPort As Integer
Public gbUseProxyAuthentication As Boolean
Public gsProxyUser As String
Public gsProxyPass As String
Public gbUseDirectConnect As Boolean

'Vorlaufzeit gewandelt
Public gfVorlaufGebotTimeVal As Double

'1.8.0
Public gsBrowserIdString As String
Public giMaxRow As Integer   ' Anz. Zeilen auf dem Schirm, 0 .. x
Public giMaxRowSetting As Integer   ' Anz. Zeilen in den Settings, wird erst beim Neustart aktiv, 0 .. x
Public giDefaultHeight As Integer '= 440
Public giDefaultFontSize As Integer '= 8
Public gbPassAtStart As Boolean
Public gbArtikelRefreshPost As Boolean
Public gbArtikelRefreshPost2 As Boolean
Public giArtikelRefreshCycle As Integer 'Secs
Public gbAutoAktualisieren As Boolean
'sh 25.10.03 nur nächste + xminvor
Public gbAutoAktualisierenNext As Boolean
Public gbAktualisierenXvor As Boolean
Public giAktXminvor As Integer
Public giAktXminvorCycle As Integer
Public gbAAnext As Boolean
Public giArtAktOptions As Integer
Public giArtAktOptionsValue As Integer
Public giAktualisierenOpt As Integer
Public gbAutoWarnNoBid As Boolean
Public gbConcurrentUpdates As Boolean
Public gbQuietAfterManualBid As Boolean
Public gbUpdateAfterManualBid As Boolean
Public gbUpdateAnonymous As Boolean
Public gsWatchListType As String

Public giUseTimeSync  As Integer '0: kein, 1: einmal, 2: Gebot, 4: Start, 8: Regelmäßig alle x Minuten
Public gbSendAuctionEnd As Boolean
Public gbSendAuctionEndNoSuccess As Boolean
Public gbSendIfLow As Boolean
Public gsSendEndTo As String
Public gsSendEndFrom As String
Public gsSendEndFromRealname As String
Public gbCheckForUpdate As Boolean 'BOM- Updateprüfung
Public gbCheckForUpdateBeta As Boolean
Public glCheckForUpdateInterval  As Long
Public gbAutoUpdateCurrencies As Boolean
Public gbUseWheel As Boolean
Public gbWheelUsed As Boolean
Public gbShowToolbar As Boolean
Public gsTestArtikel As String
Public gbTestConnect As Boolean
Public gbOperaField As Boolean
Public giToolbarSize As Integer
Public gsSeparator As String
Public gsDelimiter As String
Public gbMinToTray As Boolean
Public gbNewItemWindowAlwaysOnTop As Boolean
Public gbNewItemWindowOpenOnStartup As Boolean
Public gbNewItemWindowKeepsValues As Boolean
Public gsNewItemWindowWidgetOrdner As String
Public gbWarnenBeimBeenden As Boolean
Public gbBeendenNachAuktion As Boolean
Public gbBeendenNachAuktionAktiv As Boolean
Public gbShowTitleDateTime As Boolean
Public gbShowTitleTimeLeft As Boolean
Public gbShowTitleVersion As Boolean
Public gbShowTitleAuctionHome As Boolean
Public gbShowTitleDefaultUser As Boolean
Public glTimeSyncIntervall As Long
Public glTimeSyncCounter As Long
Public gbCleanStatus As Boolean
Public glCleanStatusTime As Long
Public gbBrowseInNewWindow As Boolean
Public gbBrowseInline As Boolean
Public giInlineBrowserDelay As Integer
Public giInlineBrowserModifierKey As Integer
Public gbCountDownInAutomodeOnly As Boolean
Public giReloadTimes As Integer
  
'ODBC
Public gbUsesOdbc As Boolean
Public giOdbcZyklus As Integer
Public gsOdbcProvider As String
Public gsOdbcDb As String
Public gsOdbcUser As String
Public gsOdbcPass As String
Public gsOdbcStopRead As Boolean

'1.8.2
Public gbAutoLogin As Boolean

'2.0.1
Public gsAktLanguage As String

'Fensterposition Hauptfenster
Public glPosTop As Long
Public glPosLeft As Long
Public glPosWidth As Long
Public glPosHeight As Long

'MD-Marker 20090325 , frmBrowser aus dem Projekt entfernt
'Fensterposition Browser
'Public glBrowserLeft As Long
'Public glBrowserTop As Long
'Public glBrowserWidth As Long
'Public glBrowserHeight As Long

'Fensterposition NeuerArtikelFenster
Public glNeuerArtikelLeft As Long
Public glNeuerArtikelTop As Long
Public glNeuerArtikelWidth As Long
Public glNeuerArtikelHeight As Long

'Fensterposition DebugFenster
Public glDebugWindowLeft As Long
Public glDebugWindowTop As Long
Public glDebugWindowWidth As Long
Public glDebugWindowHeight As Long

'Fensterposition InlineBrowser
Public gfInfoLeft As Double
Public gfInfoTop As Double
Public gfInfoWidth As Double
Public gfInfoHeight As Double

'NTP- Service
Public gsNtpServer As String
Public giUseNtp As Integer
Public gsNtpData As String
Public giNtpErr As Integer
Public gfNtpDelay As Double

'** für POP- Zugriffe
Public gbUsePop As Boolean
Public giPopZyklus As Integer
Public gsPopUser As String
Public gsPopPass As String
Public gsPopServer As String
Public giPopTimeOut As Integer
Public giPopPort As Integer
Public gsSmtpServer As String
Public giSmtpPort As Integer
Public gbSessionClosed As Boolean
Public gbFatalError As Boolean
Public gbTimeOutOccurs As Boolean
Public gsAbsender As String
Public gbPopTestIsOk As Boolean
Public gbUseSmtpAuth As Boolean
Public gbPopEncryptedOnly As Boolean
Public gbPopSendEncryptedAcknowledgment As Boolean
Public gbPopNeedsUsername As Boolean
Public gsPopSubjectDelimiter As String

'INET Connect
Public glConnectID As Long
Public gsConnectName As String

Public gbPlaySoundOnBid As Boolean
Public gsSoundOnBid As String
Public gsSoundOnBidSuccess As String
Public gsSoundOnBidFail As String

'Variables and constant für test und allg. TCP
Public gsOutText As String
Public gsResponseState As String
Public giSmtpResponse As Integer
Public gsThisChunk As String
Public gsWholeThing As String
Public gsDotLine As String
Public gsStatusTxt As String

Public giPopTimeOutCount As Integer
Public giTimeOutTimerTimeOut As Integer

Public gdatarrPanelTimes(1 To 10) As Date
Public gsarrPanelFixText(1 To 10) As String
Public glarrPanelFixBackColor(1 To 10) As Long
Public glarrPanelFixForeColor(1 To 10) As Long

'Rausgezogen ;-)
Public giAktAnzArtikel As Integer

'MD-Marker
'Datentypen möglichst nach Typen anordnen, _
Len(gtarrArtikel(1)) sollte LenB(gtarrArtikel(1)) entsprechen, _
siehe Hilfe -> Padding oder Alignment
Public Type udtArtikelZeile
    Artikel As String
    EndeZeit As Date
    Titel As String
    Gebot As Double
    MinGebot As Double
    AktPreis As Double
    Gruppe As String
    Status As Long 'Integer
    WE As String
    AnzGebote As Integer
    Bieter As String
    PostUpdateDone As Boolean
    Kommentar As String
    Versand As String
    Verkaeufer As String
    eBayUser As String
    eBayPass As String
    UseToken As Boolean
    UserAccount As String
    NotFound As Integer
    Bewertung As String
    Standort As String
    MindestpreisNichtErreicht As Boolean
    Ueberarbeitet As Boolean
    UpdateInProgressSince As Date
    ExtCmdPreDone As Boolean
    ExtCmdPostDone As Boolean
    LastChangedId As Long
    TimeZone As Double
End Type
'Achtung: Screen- Zeilen: 0 .. Max
'Array: 1..x
Public gtarrArtikelArray() As udtArtikelZeile
Public gtarrRemovedArtikelArray() As udtArtikelZeile

Public Type Sema
    sema_Name As String
    is_requested As Boolean
    request_Count As Integer
End Type

Public gtBrowserSema As Sema
Public gtTcpSema As Sema

Public Type udtUserPass
    UaUser As String
    UaPass As String
    UaToken As Boolean
End Type
'Benutzerverwaltung
Public gtarrUserArray() As udtUserPass

Public giUserAnzahl As Integer
Public giDefaultUser As Integer 'index userarray
Public giArtChoose As Integer 'globartikel
Public gbEmpUserEnd As Boolean

Public gbStripJS As Boolean
Public gbLogHtml As Boolean
Public gsBrowserLanguage As String
Public gsSortOrder As String
Public giShippingMode As Integer
Public gbShowDebugWindow As Boolean
Public giDebugLevel As Integer

Public gsIconSet As String
Public gsarrIconSet() As String
Public giTrayIconDisplayTimeOnlineMode1 As Integer
Public giTrayIconDisplayTimeOnlineMode2 As Integer
Public giTrayIconDisplayTimeOfflineMode1 As Integer
Public giTrayIconDisplayTimeOfflineMode2 As Integer
Public gsarrServerStrArr() As String
Public gsServerStringsFile As String
Public gfTimeDeviation As Double
Public goCookieHandler As clsCookieHandler
Public gbUseIECookies As Boolean
Public gbUseCurl As Boolean
Public glHttpTimeOut As Long
Public gbUseSecurityToken As Boolean
Public gsAppDataPath As String
Public gsToolTipSeparator As String
Public gsReservedPriceMarker As String
Public gbConfirmDelete As Boolean
Public gsSendItemTo As String
Public gbSendItemEncrypted As Boolean

Public gbBuyItNow As Boolean

Public gsPopCmdSSL As String
Public gsSmtpCmdSSL As String
Public gbPopUseSSL As Boolean
Public gbSmtpUseSSL As Boolean
Public gbHideSSLWindow As Boolean
Public glSSLStartupDelay As Long

Public gsSearchTerm As String
Public glSearchPosition As Long

Public glLogfileMaxSize As Long
Public glLogfileShrinkPercent As Long
Public gbLogDeletedItems As Boolean
Public gbBlacklistDeletedItems As Boolean

Public gbUseIsoDate As Boolean
Public gbUseUnixDate As Boolean
Public gbNoEnumFonts As Boolean
Public gbSuppressHeader As Boolean
Public gbIgnoreItemErrorsOnStartup As Boolean

Public gbEditShippingOnClick As Boolean
Public gbOpenBrowserOnClick As Boolean
Public gbCommentInTitle As Boolean
Public gbRevisedInTitle As Boolean

Public giPreventSuspend As Integer
Public giWakeOnAuction As Integer
Public gbResuspendAfterEnd As Boolean
Public gbForceResuspendAfterEnd As Boolean
Public giSleepAfterWakeup As Integer
Public gdatFallAsleepDate As Date
Public giSuspendState As Integer ' 0 = wach, 1 = einschlafend, 2 = aufwachend
Public gbWarSchonWach As Boolean
Public gbSuspendNachAuktionAktiv As Boolean
Public gbHibernate As Boolean

Public gsLastSavedCrc As String
Public gbShowShippingCosts As Boolean
Public gbShowWeekday As Boolean
Public gbShowFocusRect As Boolean
Public glFocusRectColor As Long
Public gsSpecialDateFormat As String

Public glExtCmdTimeWindow As Long
Public gsExtCmdPreCmd As String
Public gsExtCmdPostCmd As String
Public gsExtCmdPeriodicCmd  As String
Public glExtCmdPreTime As Long
Public glExtCmdPostTime As Long
Public glExtCmdPeriodicTime As Long
Public giExtCmdWindowStyle As Integer

Public glSendCsvInterval As Long
Public gsSendCsvTo As String

Public gbReadEndedItems As Boolean
Public gbBeepBeforeAuction As Boolean
Public gbBlockEndedItems As Boolean
Public gbBlockBuyItNowItems As Boolean

Public glbJARVISstate As String
