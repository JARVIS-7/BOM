Attribute VB_Name = "modKeywords"
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

Private Const msREQUIREDKEYWORDSVERSION As String = "1.0.147"

'
' Liste der Keywords und die entsprechenden Zugriffe
'
Public gsAuctionHome As String

Public gsMainUrl As String
Public gsScript1 As String
Public gsScript2 As String
Public gsScript3 As String
Public gsScript4 As String
Public gsScript5 As String
Public gsScript9 As String

'CGI- Strings
Public gsScriptCommand1 As String
Public gsScriptCommand2 As String
Public gsScriptCommand3 As String
Public gsScriptCommand4 As String
Public gsScriptCommand5 As String
Public gsScriptCommand9 As String
Public gsSiteEncoding As String
Public gbUseSSL As Boolean

'Datum und Zeit f�r Artikelseite
Public gsAnsDateFormat1 As String
Public gsAnsOffsetLocal1_1 As String
Public gsAnsOffsetLocal1_2 As String
Public gsarrMonthNames1(12) As String

'Datum und Zeit f�r Zeitsync
Public gsAnsDateFormat2 As String
Public gsAnsOffsetLocal2_1 As String
Public gsAnsOffsetLocal2_2 As String
Public gsarrMonthNames2(12) As String

'Login
Public gsCmdLogOff As String
Public gsCmdLogIn As String
Public gsCmdLogIn2 As String
Public gsAnsLoginOk As String
Public gsAnsLoginOk2 As String
Public gsAnsLoginFrm As String
Public gsAnsUserField As String
Public gsAnsUserField2 As String
Public gsAnsUserField3 As String
Public gsAnsPassField As String
Public gsAnsTokenField As String
Public gsAnsLoginSubmitImage1 As String
Public gsAnsLoginSubmitImage2 As String
Public gsAnsLoginSubmitImage3 As String

'Beobachtete lesen
Public gsAnsWatchList As String
Public gsCmdWatchList As String
Public gsCmdWatchList2 As String
Public gsCmdBidList As String
Public gsCmdSummary As String

Public gsAnsSummary As String
Public gsAnsBidStart As String
Public gsAnsBidEnd As String
Public gsAnsBidItemStart1 As String
Public gsAnsBidItemStart2 As String
Public gsAnsBidItemEnd1 As String
Public gsAnsBidItemEnd2 As String
Public gsAnsBidItemPreEnd1 As String
Public ansBidItemPreEnd2 As String
Public gsAnsBidItemEnded As String

Public gsAnsWatchStart As String
Public gsAnsWatchEnd As String
Public gsAnsWatchItemStart1 As String
Public gsAnsWatchItemStart2 As String
Public gsAnsWatchItemEnd1 As String
Public gsAnsWatchItemEnd2 As String
Public gsAnsWatchItemPreEnd1 As String
Public gsAnsWatchItemPreEnd2 As String
Public gsAnsWatchItemEnded As String

'Bieten
Public gsCmdDecSeparator As String
Public gsCmdMakeBid As String
Public gsCmdBuyItNow As String
Public gsAnsBidForm As String
Public gsAnsBidForm2 As String
Public gsAnsBuyForm As String
Public gsAnsBuyForm2 As String
Public gsAnsBidAccepted As String
Public gsAnsBidAccepted2 As String
Public gsAnsBidAccepted3 As String
Public gsAnsBidAccepted4 As String
Public gsAnsBidOutBid As String
Public gsAnsBidReserveNotMet As String
Public gsAnsBidErrMinBid As String
Public gsAnsBidErrEnded As String
Public gsAnsBidErrEnded2 As String
Public gsAnsBidErrNotAvail  As String
Public gsAnsBidErrGeneral As String
Public gsAnsBidConfirm As String
Public gsAnsLinkChangeUser As String
Public gsAnsSignInError As String
Public gsAnsTimeLeft As String
Public gsAnsTimeLeftStart As String
Public gsAnsTimeLeftEnd As String
Public gsAnsBidField As String
Public gsAnsBidStep1SubmitImage As String
Public gsAnsBidStep2SubmitImage As String
Public gsAnsBuyStep1SubmitImage As String
Public gsAnsBuyStep2SubmitImage As String

'Artikelinfo
Public gsCmdViewItem As String
'Public ansCurrBid As String
'Public ansWinnBid As String
'Public ansApproxBid As String
'Public ansStartingBid As String
Public gsAnsHistory As String
'Public ansPrice As String
'Public ansSoldFor As String
Public gsarrAnsPriceA() As String
Public gsarrAnsPriceStartA() As String
Public gsarrAnsPriceEndA() As String
Public gsarrAnsPriceTypeA() As String
Public gsAnsNumBids As String
Public gsAnsNumBidsStart As String
Public gsAnsNumBidsEnd As String
Public gsAnsNumBids2 As String
Public gsAnsNumBidsStart2 As String
Public gsAnsNumBidsEnd2 As String
Public gsAnsNumBids3 As String
Public gsAnsNumBidsStart3 As String
Public gsAnsNumBidsEnd3 As String
Public gsAnsMinBid  As String
Public gsAnsMinBidStart As String
Public gsAnsMinBidEnd As String
Public gsAnsDutch As String
Public gsAnsBuyOnly As String
Public gsAnsAdvertisement As String
Public gsAnsShipping1 As String
Public gsAnsShippingStart1 As String
Public gsAnsShippingEnd1 As String
Public gsAnsShipping2 As String
Public gsAnsShippingStart2 As String
Public gsAnsShippingEnd2 As String
Public gsAnsShipping3 As String
Public gsAnsShippingStart3 As String
Public gsAnsShippingEnd3 As String
Public gsAnsTitle As String
Public gsAnsTitleStart As String
Public gsAnsTitleEnd As String
Public gsAnsInvalid As String
Public gsAnsSellerAway As String
Public gsAnsEndTime As String
Public gsAnsEndTime2 As String
Public gsAnsEndTimeEpoch As String
Public gsAnsEndTimeMaxLen As Long
Public gsAnsTime1_1 As String
Public gsAnsTime1_2 As String
Public gsAnsTime2_1 As String
Public gsAnsTime2_2 As String
Public gsAnsDescriptionBegin As String
Public gsAnsDescriptionEnd As String
Public gsAnsAskSeller As String
Public gsAnsAskSellerStart As String
Public gsAnsAskSellerEnd As String
Public gsAnsQuant As String
Public gsAnsQuantStart As String
Public gsAnsQuantEnd As String
Public gsAnsPrivat As String
Public gsAnsBidder As String
Public gsAnsWinner As String
Public gsAnsBuyer As String
Public gsAnsUserIDStart  As String
Public gsAnsUserIdEnd As String
Public gsAnsUserIdEnd2 As String
Public gsAnsBuyerReserve As String
Public gsAnsRevised As String
Public gsAnsSwitchToAnonymous As String

'Misc
Public gsAnsWatchItem As String
Public gsAnsWatchItem2 As String
Public gsAnsWatchItem3 As String
Public gsCmdTimeShow As String
Public gsCmdTimeShowFormat As String
Public gsAnsMaintenance As String
Public gsAnsMailToFriend As String
Public gsAnsMailToFriendAddressStart As String
Public gsAnsMailToFriendAddressEnd As String
Public gsAnsMailToFriendItemStart As String
Public gsAnsMailToFriendItemEnd As String
Public gsAnsLinkStart As String
Public gsAnsLinkEnd As String

Public gsCmdUpdateCurrencies As String
Public gsCmdUpdateCurrReferer As String
Public gsAnsCurrencyStart As String
Public gsAnsCurrencyEnd As String
Public gsAnsCurrency1 As String
Public gsAnsCurrency2 As String
Public gsAnsCurrency3 As String
Public gsAnsCurrency4 As String

Public gsarrAnsPsNameA() As String
Public gsarrAnsPsLinkA() As String
Public gsarrAnsPsEncodingA() As String
Public gbarrAnsPsEditA() As Boolean

Public gsarrAnsToolNameA() As String
Public gsarrAnsToolLinkA() As String
Public gsarrAnsToolEncodingA() As String
Public gbarrAnsToolEditA() As Boolean

Public gsAnsAssessment As String
Public gsAnsAssessmentAnzStart  As String
Public gsAnsAssessmentAnzEnd As String
Public gsAnsAssessmentPercentStart As String
Public gsAnsAssessmentPercentEnd As String

Public gsAnsLocation As String
Public gsAnsLocationStart As String
Public gsAnsLocationEnd As String

Public gsAnsAnonBidder As String
Public gsAnsAnonBidderStart As String
Public gsAnsAnonBidderEnd As String

Public gsAnsNoteStart As String
Public gsAnsNoteEnd As String
Public gsAnsNoteTextStart As String
Public gsAnsNoteTextEnd As String
Public gsAnsNoteLineIDStart As String
Public gsAnsNoteLineIDEnd As String

Public Sub ReadAllKeywords()

Dim i As Integer
Dim sTmp As String

On Error Resume Next
Dim oIni As clsIni

Set oIni = New clsIni

oIni.ReadIni App.Path & "\" & gsServerStringsFile

'sect Server
LocCINIGetValue oIni, "Server", "AuctionHome", gsAuctionHome
LocCINIGetValue oIni, "Server", "Webpage", gsMainUrl
LocCINIGetValue oIni, "Server", "Script1", gsScript1
LocCINIGetValue oIni, "Server", "Script2", gsScript2
LocCINIGetValue oIni, "Server", "Script3", gsScript3
LocCINIGetValue oIni, "Server", "Script4", gsScript4
LocCINIGetValue oIni, "Server", "Script5", gsScript5
LocCINIGetValue oIni, "Server", "Script9", gsScript9
LocCINIGetValue oIni, "Server", "ScriptCommand1", gsScriptCommand1
LocCINIGetValue oIni, "Server", "ScriptCommand2", gsScriptCommand2
LocCINIGetValue oIni, "Server", "ScriptCommand3", gsScriptCommand3
LocCINIGetValue oIni, "Server", "ScriptCommand4", gsScriptCommand4
LocCINIGetValue oIni, "Server", "ScriptCommand5", gsScriptCommand5
LocCINIGetValue oIni, "Server", "ScriptCommand9", gsScriptCommand9
LocCINIGetValue oIni, "Server", "SiteEncoding", gsSiteEncoding
gbUseSSL = oIni.GetValue("Server", "UseSSL")

'sect DateTime1
LocCINIGetValue oIni, "DateTime1", "ansDateFormat", gsAnsDateFormat1
LocCINIGetValue oIni, "DateTime1", "ansOffsetLocal1", gsAnsOffsetLocal1_1
LocCINIGetValue oIni, "DateTime1", "ansOffsetLocal2", gsAnsOffsetLocal1_2
LocCINIGetValue oIni, "DateTime1", "ansTime1", gsAnsTime1_1
LocCINIGetValue oIni, "DateTime1", "ansTime2", gsAnsTime1_2

For i = 1 To 12
    LocCINIGetValue oIni, "DateTime1", "month" & CStr(i), gsarrMonthNames1(i)
Next i

'sect DateTime2
LocCINIGetValue oIni, "DateTime2", "ansDateFormat", gsAnsDateFormat2
LocCINIGetValue oIni, "DateTime2", "ansOffsetLocal1", gsAnsOffsetLocal2_1
LocCINIGetValue oIni, "DateTime2", "ansOffsetLocal2", gsAnsOffsetLocal2_2
LocCINIGetValue oIni, "DateTime2", "ansTime1", gsAnsTime2_1
LocCINIGetValue oIni, "DateTime2", "ansTime2", gsAnsTime2_2

For i = 1 To 12
    LocCINIGetValue oIni, "DateTime2", "month" & CStr(i), gsarrMonthNames2(i)
Next i

'sect Login
LocCINIGetValue oIni, "Login", "cmdLogOff", gsCmdLogOff
LocCINIGetValue oIni, "Login", "cmdLogIn", gsCmdLogIn
LocCINIGetValue oIni, "Login", "cmdLogIn2", gsCmdLogIn2
LocCINIGetValue oIni, "Login", "ansLoginOk", gsAnsLoginOk
LocCINIGetValue oIni, "Login", "ansLoginOk2", gsAnsLoginOk2
LocCINIGetValue oIni, "Login", "ansLoginForm", gsAnsLoginFrm
LocCINIGetValue oIni, "Login", "ansUserField", gsAnsUserField
LocCINIGetValue oIni, "Login", "ansUserField2", gsAnsUserField2
LocCINIGetValue oIni, "Login", "ansUserField3", gsAnsUserField3
LocCINIGetValue oIni, "Login", "ansPassField", gsAnsPassField
LocCINIGetValue oIni, "Login", "ansTokenField", gsAnsTokenField
LocCINIGetValue oIni, "Login", "ansLoginSubmitImage1", gsAnsLoginSubmitImage1
LocCINIGetValue oIni, "Login", "ansLoginSubmitImage2", gsAnsLoginSubmitImage2
LocCINIGetValue oIni, "Login", "ansLoginSubmitImage3", gsAnsLoginSubmitImage3

'beobachtete lesen
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchlist", gsAnsWatchList
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "cmdWatchlist", gsCmdWatchList
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "cmdWatchlist2", gsCmdWatchList2
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "cmdBidlist", gsCmdBidList
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "cmdSummary", gsCmdSummary
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansSummary", gsAnsSummary
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidStart", gsAnsBidStart
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidEnd", gsAnsBidEnd
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidItemStart1", gsAnsBidItemStart1
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidItemStart2", gsAnsBidItemStart2
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidItemEnd1", gsAnsBidItemEnd1
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidItemEnd2", gsAnsBidItemEnd2
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidItemPreEnd1", gsAnsBidItemPreEnd1
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidItemPreEnd2", ansBidItemPreEnd2
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansBidItemEnded", gsAnsBidItemEnded
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchStart", gsAnsWatchStart
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchEnd", gsAnsWatchEnd
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchItemStart1", gsAnsWatchItemStart1
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchItemStart2", gsAnsWatchItemStart2
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchItemEnd1", gsAnsWatchItemEnd1
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchItemEnd2", gsAnsWatchItemEnd2
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchItemPreEnd1", gsAnsWatchItemPreEnd1
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchItemPreEnd2", gsAnsWatchItemPreEnd2
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansWatchItemEnded", gsAnsWatchItemEnded

LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansNoteStart", gsAnsNoteStart
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansNoteEnd", gsAnsNoteEnd
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansNoteTextStart", gsAnsNoteTextStart
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansNoteTextEnd", gsAnsNoteTextEnd
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansNoteLineIDStart", gsAnsNoteLineIDStart
LocCINIGetValue oIni, "Watchlist" & gsWatchListType, "ansNoteLineIDEnd", gsAnsNoteLineIDEnd


'Bieten
LocCINIGetValue oIni, "Bidding", "cmdDecSeparator", gsCmdDecSeparator
LocCINIGetValue oIni, "Bidding", "cmdMakeBid", gsCmdMakeBid
LocCINIGetValue oIni, "Bidding", "cmdBuyItNow", gsCmdBuyItNow
LocCINIGetValue oIni, "Bidding", "ansBidForm", gsAnsBidForm
LocCINIGetValue oIni, "Bidding", "ansBidForm2", gsAnsBidForm2
LocCINIGetValue oIni, "Bidding", "ansBuyForm", gsAnsBuyForm
LocCINIGetValue oIni, "Bidding", "ansBuyForm2", gsAnsBuyForm2
LocCINIGetValue oIni, "Bidding", "ansBidAccepted", gsAnsBidAccepted
LocCINIGetValue oIni, "Bidding", "ansBidAccepted2", gsAnsBidAccepted2
LocCINIGetValue oIni, "Bidding", "ansBidAccepted3", gsAnsBidAccepted3
LocCINIGetValue oIni, "Bidding", "ansBidAccepted4", gsAnsBidAccepted4
LocCINIGetValue oIni, "Bidding", "ansBidOutbid", gsAnsBidOutBid
LocCINIGetValue oIni, "Bidding", "ansBidReserveNotMet", gsAnsBidReserveNotMet
LocCINIGetValue oIni, "Bidding", "ansBidErrMinBid", gsAnsBidErrMinBid
LocCINIGetValue oIni, "Bidding", "ansBidErrEnded", gsAnsBidErrEnded
LocCINIGetValue oIni, "Bidding", "ansBidErrEnded2", gsAnsBidErrEnded2
LocCINIGetValue oIni, "Bidding", "ansBidErrNotAvail", gsAnsBidErrNotAvail
LocCINIGetValue oIni, "Bidding", "ansBidErrGeneral", gsAnsBidErrGeneral
LocCINIGetValue oIni, "Bidding", "ansBidConfirm", gsAnsBidConfirm
LocCINIGetValue oIni, "Bidding", "ansLinkChangeUser", gsAnsLinkChangeUser
LocCINIGetValue oIni, "Bidding", "ansSignInError", gsAnsSignInError
LocCINIGetValue oIni, "Bidding", "ansTimeLeft", gsAnsTimeLeft
LocCINIGetValue oIni, "Bidding", "ansTimeLeftStart", gsAnsTimeLeftStart
LocCINIGetValue oIni, "Bidding", "ansTimeLeftEnd", gsAnsTimeLeftEnd
LocCINIGetValue oIni, "Bidding", "ansBidField", gsAnsBidField
LocCINIGetValue oIni, "Bidding", "ansBidStep1SubmitImage", gsAnsBidStep1SubmitImage
LocCINIGetValue oIni, "Bidding", "ansBidStep2SubmitImage", gsAnsBidStep2SubmitImage
LocCINIGetValue oIni, "Bidding", "ansBuyStep1SubmitImage", gsAnsBuyStep1SubmitImage
LocCINIGetValue oIni, "Bidding", "ansBuyStep2SubmitImage", gsAnsBuyStep2SubmitImage

'Artikelinfo
LocCINIGetValue oIni, "ItemInfo", "cmdViewItem", gsCmdViewItem
'LocCINIGetValue oIni, "ItemInfo", "ansCurrBid", ansCurrBid
'LocCINIGetValue oIni, "ItemInfo", "ansWinnBid", ansWinnBid
'LocCINIGetValue oIni, "ItemInfo", "ansApproxBid", ansApproxBid
'LocCINIGetValue oIni, "ItemInfo", "ansStartingBid", ansStartingBid
LocCINIGetValue oIni, "ItemInfo", "ansHistory", gsAnsHistory
'LocCINIGetValue oIni, "ItemInfo", "ansPrice", ansPrice
'LocCINIGetValue oIni, "ItemInfo", "ansSoldFor", ansSoldFor
LocCINIGetValue oIni, "ItemInfo", "ansMinBid", gsAnsMinBid
LocCINIGetValue oIni, "ItemInfo", "ansMinBidStart", gsAnsMinBidStart
LocCINIGetValue oIni, "ItemInfo", "ansMinBidEnd", gsAnsMinBidEnd
LocCINIGetValue oIni, "ItemInfo", "ansNumBids", gsAnsNumBids
LocCINIGetValue oIni, "ItemInfo", "ansNumBidsStart", gsAnsNumBidsStart
LocCINIGetValue oIni, "ItemInfo", "ansNumBidsEnd", gsAnsNumBidsEnd
LocCINIGetValue oIni, "ItemInfo", "ansNumBids2", gsAnsNumBids2
LocCINIGetValue oIni, "ItemInfo", "ansNumBidsStart2", gsAnsNumBidsStart2
LocCINIGetValue oIni, "ItemInfo", "ansNumBidsEnd2", gsAnsNumBidsEnd2
LocCINIGetValue oIni, "ItemInfo", "ansNumBids3", gsAnsNumBids3
LocCINIGetValue oIni, "ItemInfo", "ansNumBidsStart3", gsAnsNumBidsStart3
LocCINIGetValue oIni, "ItemInfo", "ansNumBidsEnd3", gsAnsNumBidsEnd3
LocCINIGetValue oIni, "ItemInfo", "ansDutch", gsAnsDutch
LocCINIGetValue oIni, "ItemInfo", "ansBuyOnly", gsAnsBuyOnly
LocCINIGetValue oIni, "ItemInfo", "ansAdvertisement", gsAnsAdvertisement
LocCINIGetValue oIni, "ItemInfo", "ansShipping1", gsAnsShipping1
LocCINIGetValue oIni, "ItemInfo", "ansShippingStart1", gsAnsShippingStart1
LocCINIGetValue oIni, "ItemInfo", "ansShippingEnd1", gsAnsShippingEnd1
LocCINIGetValue oIni, "ItemInfo", "ansShipping2", gsAnsShipping2
LocCINIGetValue oIni, "ItemInfo", "ansShippingStart2", gsAnsShippingStart2
LocCINIGetValue oIni, "ItemInfo", "ansShippingEnd2", gsAnsShippingEnd2
LocCINIGetValue oIni, "ItemInfo", "ansShipping3", gsAnsShipping3
LocCINIGetValue oIni, "ItemInfo", "ansShippingStart3", gsAnsShippingStart3
LocCINIGetValue oIni, "ItemInfo", "ansShippingEnd3", gsAnsShippingEnd3

LocCINIGetValue oIni, "ItemInfo", "ansNumPrice", sTmp
If Val(sTmp) <= 0 Then sTmp = 1
ReDim gsarrAnsPriceA(1 To Val(sTmp))
ReDim gsarrAnsPriceStartA(1 To Val(sTmp))
ReDim gsarrAnsPriceEndA(1 To Val(sTmp))
ReDim gsarrAnsPriceTypeA(1 To Val(sTmp))
For i = 1 To UBound(gsarrAnsPriceA)
  LocCINIGetValue oIni, "ItemInfo", "ansPrice" & CStr(i), gsarrAnsPriceA(i)
  LocCINIGetValue oIni, "ItemInfo", "ansPriceStart" & CStr(i), gsarrAnsPriceStartA(i)
  LocCINIGetValue oIni, "ItemInfo", "ansPriceEnd" & CStr(i), gsarrAnsPriceEndA(i)
  LocCINIGetValue oIni, "ItemInfo", "ansPriceType" & CStr(i), gsarrAnsPriceTypeA(i)
Next i

'titel
LocCINIGetValue oIni, "ItemInfo", "ansTitle", gsAnsTitle
LocCINIGetValue oIni, "ItemInfo", "ansTitleStart", gsAnsTitleStart
LocCINIGetValue oIni, "ItemInfo", "ansTitleEnd", gsAnsTitleEnd
'Endezeit
LocCINIGetValue oIni, "ItemInfo", "ansInvalid", gsAnsInvalid
LocCINIGetValue oIni, "ItemInfo", "ansSellerAway", gsAnsSellerAway
LocCINIGetValue oIni, "ItemInfo", "ansEndtime", gsAnsEndTime
LocCINIGetValue oIni, "ItemInfo", "ansEndtime2", gsAnsEndTime2
LocCINIGetValue oIni, "ItemInfo", "ansEndtimeEpoch", gsAnsEndTimeEpoch
LocCINIGetValue oIni, "ItemInfo", "ansEndtimeMaxlen", sTmp
gsAnsEndTimeMaxLen = IIf(Val(sTmp) > 0 And Val(sTmp) < 10000, Val(sTmp), 200)
'Beschreibung
LocCINIGetValue oIni, "ItemInfo", "ansDescriptionBegin", gsAnsDescriptionBegin
LocCINIGetValue oIni, "ItemInfo", "ansDescriptionEnd", gsAnsDescriptionEnd
'VK, K und Mengen
LocCINIGetValue oIni, "ItemInfo", "ansAskSeller", gsAnsAskSeller
LocCINIGetValue oIni, "ItemInfo", "ansAskSellerStart", gsAnsAskSellerStart
LocCINIGetValue oIni, "ItemInfo", "ansAskSellerEnd", gsAnsAskSellerEnd
LocCINIGetValue oIni, "ItemInfo", "ansQuant", gsAnsQuant
LocCINIGetValue oIni, "ItemInfo", "ansQuantStart", gsAnsQuantStart
LocCINIGetValue oIni, "ItemInfo", "ansQuantEnd", gsAnsQuantEnd
LocCINIGetValue oIni, "ItemInfo", "ansPrivat", gsAnsPrivat
LocCINIGetValue oIni, "ItemInfo", "ansAnonBidder", gsAnsAnonBidder
LocCINIGetValue oIni, "ItemInfo", "ansAnonBidderStart", gsAnsAnonBidderStart
LocCINIGetValue oIni, "ItemInfo", "ansAnonBidderEnd", gsAnsAnonBidderEnd
LocCINIGetValue oIni, "ItemInfo", "ansBidder", gsAnsBidder
LocCINIGetValue oIni, "ItemInfo", "ansWinner", gsAnsWinner
LocCINIGetValue oIni, "ItemInfo", "ansBuyer", gsAnsBuyer
LocCINIGetValue oIni, "ItemInfo", "ansUserIDStart", gsAnsUserIDStart
LocCINIGetValue oIni, "ItemInfo", "ansUserIdEnd", gsAnsUserIdEnd
LocCINIGetValue oIni, "ItemInfo", "ansUserIdEnd2", gsAnsUserIdEnd2
LocCINIGetValue oIni, "ItemInfo", "ansAssessment", gsAnsAssessment
LocCINIGetValue oIni, "ItemInfo", "ansAssessmentAnzStart", gsAnsAssessmentAnzStart
LocCINIGetValue oIni, "ItemInfo", "ansAssessmentAnzEnd", gsAnsAssessmentAnzEnd
LocCINIGetValue oIni, "ItemInfo", "ansAssessmentPercentStart", gsAnsAssessmentPercentStart
LocCINIGetValue oIni, "ItemInfo", "ansAssessmentPercentEnd", gsAnsAssessmentPercentEnd
LocCINIGetValue oIni, "ItemInfo", "ansLocation", gsAnsLocation
LocCINIGetValue oIni, "ItemInfo", "ansLocationStart", gsAnsLocationStart
LocCINIGetValue oIni, "ItemInfo", "ansLocationEnd", gsAnsLocationEnd
LocCINIGetValue oIni, "ItemInfo", "ansBuyerReserve", gsAnsBuyerReserve
LocCINIGetValue oIni, "ItemInfo", "ansRevised", gsAnsRevised
LocCINIGetValue oIni, "ItemInfo", "ansSwitchToAnonymous", gsAnsSwitchToAnonymous

'Diverses
LocCINIGetValue oIni, "Misc", "ansWatchItem", gsAnsWatchItem
LocCINIGetValue oIni, "Misc", "ansWatchItem2", gsAnsWatchItem2
LocCINIGetValue oIni, "Misc", "ansWatchItem3", gsAnsWatchItem3
LocCINIGetValue oIni, "Misc", "cmdTimeShow", gsCmdTimeShow
LocCINIGetValue oIni, "Misc", "cmdTimeShowFormat", gsCmdTimeShowFormat
LocCINIGetValue oIni, "Misc", "ansMaintenance", gsAnsMaintenance
LocCINIGetValue oIni, "Misc", "ansMailToFriend", gsAnsMailToFriend
LocCINIGetValue oIni, "Misc", "ansMailToFriendAddressStart", gsAnsMailToFriendAddressStart
LocCINIGetValue oIni, "Misc", "ansMailToFriendAddressEnd", gsAnsMailToFriendAddressEnd
LocCINIGetValue oIni, "Misc", "ansMailToFriendItemStart", gsAnsMailToFriendItemStart
LocCINIGetValue oIni, "Misc", "ansMailToFriendItemEnd", gsAnsMailToFriendItemEnd
LocCINIGetValue oIni, "Misc", "ansLinkStart", gsAnsLinkStart
LocCINIGetValue oIni, "Misc", "ansLinkEnd", gsAnsLinkEnd
LocCINIGetValue oIni, "Misc", "cmdUpdateCurrencies", gsCmdUpdateCurrencies
LocCINIGetValue oIni, "Misc", "cmdUpdateCurrReferer", gsCmdUpdateCurrReferer
LocCINIGetValue oIni, "Misc", "ansCurrencyStart", gsAnsCurrencyStart
LocCINIGetValue oIni, "Misc", "ansCurrencyEnd", gsAnsCurrencyEnd
LocCINIGetValue oIni, "Misc", "ansCurrency1", gsAnsCurrency1
LocCINIGetValue oIni, "Misc", "ansCurrency2", gsAnsCurrency2
LocCINIGetValue oIni, "Misc", "ansCurrency3", gsAnsCurrency3
LocCINIGetValue oIni, "Misc", "ansCurrency4", gsAnsCurrency4

For i = 2 To UBound(gsarrAnsPsNameA)
  frmHaupt.mnuls(i).Visible = False
Next i
frmHaupt.mnuProductSearch.Enabled = False

LocCINIGetValue oIni, "ProductSearch", "ansNumPsLinks", sTmp
If Val(sTmp) < 0 Then sTmp = 0
ReDim gsarrAnsPsNameA(1 To Val(sTmp))
ReDim gsarrAnsPsLinkA(1 To Val(sTmp))
ReDim gsarrAnsPsEncodingA(1 To Val(sTmp))
ReDim gbarrAnsPsEditA(1 To Val(sTmp))
For i = 1 To Val(sTmp)
  LocCINIGetValue oIni, "ProductSearch", "ansPsName" & CStr(i), gsarrAnsPsNameA(i)
  LocCINIGetValue oIni, "ProductSearch", "ansPsLink" & CStr(i), gsarrAnsPsLinkA(i)
  LocCINIGetValue oIni, "ProductSearch", "ansPsEncoding" & CStr(i), gsarrAnsPsEncodingA(i)
  LocCINIGetValue oIni, "ProductSearch", "ansPsEdit" & CStr(i), sTmp
  gbarrAnsPsEditA(i) = sTmp

  If gsarrAnsPsNameA(i) > "" And gsarrAnsPsLinkA(i) > "" Then
    frmHaupt.mnuProductSearch.Enabled = True

    With frmHaupt.mnuls(i)
      .Caption = gsarrAnsPsNameA(i) & IIf(gbarrAnsPsEditA(i), "...", "")
      .Enabled = True
      .Visible = True
    End With
  End If

Next i

For i = 2 To UBound(gsarrAnsToolNameA)
  frmHaupt.mnult(i).Visible = False
Next i
frmHaupt.mnuTools.Enabled = False

LocCINIGetValue oIni, "Tools", "ansNumTools", sTmp
If Val(sTmp) < 0 Then sTmp = 0
ReDim gsarrAnsToolNameA(1 To Val(sTmp))
ReDim gsarrAnsToolLinkA(1 To Val(sTmp))
ReDim gsarrAnsToolEncodingA(1 To Val(sTmp))
ReDim gbarrAnsToolEditA(1 To Val(sTmp))
For i = 1 To Val(sTmp)
  LocCINIGetValue oIni, "Tools", "ansToolName" & CStr(i), gsarrAnsToolNameA(i)
  LocCINIGetValue oIni, "Tools", "ansToolLink" & CStr(i), gsarrAnsToolLinkA(i)
  LocCINIGetValue oIni, "Tools", "ansToolEncoding" & CStr(i), gsarrAnsToolEncodingA(i)
  LocCINIGetValue oIni, "Tools", "ansToolEdit" & CStr(i), sTmp
  gbarrAnsToolEditA(i) = sTmp

  If gsarrAnsToolNameA(i) > "" And gsarrAnsToolLinkA(i) > "" Then
    frmHaupt.mnuTools.Enabled = True

    With frmHaupt.mnult(i)
      .Caption = gsarrAnsToolNameA(i) & IIf(gbarrAnsToolEditA(i), "...", "")
      .Enabled = True
      .Visible = True
    End With
  End If

Next i

Set oIni = Nothing

End Sub

Public Sub CheckVersionOfKeywordsFile()

    Dim sMsg As String

    sMsg = gsarrLangTxt(25) & gsarrLangTxt(27)
    sMsg = Replace(sMsg, "%FILE%", gsServerStringsFile)
    sMsg = Replace(sMsg, "%REQVER%", msREQUIREDKEYWORDSVERSION)
    sMsg = Replace(sMsg, "\n", vbCrLf)

    If VersionValue(msREQUIREDKEYWORDSVERSION) > VersionValue(GetKeywordsFileVersion()) Then
        If vbNo = MsgBox(sMsg, vbYesNo Or vbDefaultButton2 Or vbQuestion) Then
            End 'MD-Marker
        End If
    End If

End Sub
Public Function GetKeywordsFileVersion() As String

    Dim iTmp As Integer
    Dim sVersion As String
    Dim sFile As String

    sFile = App.Path & "\" & gsServerStringsFile

    iTmp = INIGetValue(sFile, "Keyfile", "Version", sVersion)

    GetKeywordsFileVersion = sVersion

End Function

Private Sub LocIniGetValue(ByVal sSect As String, ByVal sKey As String, ByRef sValue As String)
Dim iTmp As Integer
Dim sFile As String

On Error Resume Next

'es werden nur die Daten aus dem HSP gespeichert
sFile = App.Path & "\" & gsServerStringsFile

'String lesen
iTmp = INIGetValue(sFile, sSect, sKey, sValue)

If iTmp = 0 Then
    MsgBox "Sector '" & sSect & "' Keyword '" & sKey & "' not found." _
    & vbCrLf & "please check " & gsServerStringsFile 'ServerStrings.ini
Else
    'evtl Linefeed tauschen
    sValue = Replace(sValue, "[LF]", vbCrLf)
    'Leerwerte
    sValue = Replace(sValue, "[NUL]", "")
    'Nicht zu findende Werte
    sValue = Replace(sValue, "[N/A]", "~*+#-:,!�$%&/(){[]}=")
End If
End Sub

Private Sub LocCINIGetValue(oIni As clsIni, ByVal sSect As String, ByVal sKey As String, ByRef sValue As String)
Dim iTmp As Integer

On Error Resume Next

'String lesen
iTmp = CINIGetValue(oIni, sSect, sKey, sValue)

If iTmp = 0 Then
    MsgBox "Sector '" & sSect & "' Keyword '" & sKey & "' not found." _
    & vbCrLf & "please check " & gsServerStringsFile 'ServerStrings.ini
Else
    'evtl Linefeed tauschen
    sValue = Replace(sValue, "[LF]", vbCrLf)
    'Leerwerte
    sValue = Replace(sValue, "[NUL]", "")
    'Nicht zu findende Werte
    sValue = Replace(sValue, "[N/A]", "~*+#-:,!�$%&/(){[]}=")
End If
End Sub

Public Sub KeyAutoSwitch(Optional sTxt As String = "")

    Static sCurrentKeys As String ' welche Keys haben wir derzeit wirklich geladen
    Static colKeySwitchCriteria As New Collection
    Dim i As Integer
    Dim sScript1Tmp As String
    Dim sScriptCommand1Tmp As String
    Dim sansTitleTmp As String
    Dim ansDescriptionBeginTmp As String
    Dim bytFoundItem As Byte
    Dim bytFoundCnt As Byte
    Dim lScriptCommandPos As Long
    Dim lDescriptionBeginPos As Long
    Dim lTitlePos As Long

    If colKeySwitchCriteria.Count = 0 Then ' erstmal alles initialisieren
        sCurrentKeys = gsServerStringsFile

        For i = LBound(gsarrServerStrArr) To UBound(gsarrServerStrArr)
            gsServerStringsFile = gsarrServerStrArr(i)
            If gsServerStringsFile > "" Then
                Call LocIniGetValue("Server", "Script1", sScript1Tmp)
                Call LocIniGetValue("Server", "ScriptCommand1", sScriptCommand1Tmp)
                Call LocIniGetValue("ItemInfo", "ansTitle", sansTitleTmp)
                Call LocIniGetValue("ItemInfo", "ansDescriptionBegin", ansDescriptionBeginTmp)
                Call colKeySwitchCriteria.Add(Array(sScript1Tmp & sScriptCommand1Tmp, ansDescriptionBeginTmp, sansTitleTmp), gsServerStringsFile)
            End If
        Next i
        gsServerStringsFile = sCurrentKeys
    End If

    For i = LBound(gsarrServerStrArr) To UBound(gsarrServerStrArr)
        If ExistCollectionKey(colKeySwitchCriteria, gsarrServerStrArr(i)) Then

            lScriptCommandPos = InStr(1, sTxt, colKeySwitchCriteria(gsarrServerStrArr(i))(0))
            lDescriptionBeginPos = InStr(1, sTxt, colKeySwitchCriteria(gsarrServerStrArr(i))(1))
            If lDescriptionBeginPos = 0 Then lDescriptionBeginPos = Len(sTxt)

            lTitlePos = InStr(1, sTxt, colKeySwitchCriteria(gsarrServerStrArr(i))(2))

            If lScriptCommandPos > 0 And _
               lScriptCommandPos < lDescriptionBeginPos And _
               lTitlePos > 0 Then
                bytFoundCnt = bytFoundCnt + 1
                bytFoundItem = i
            End If
        End If
    Next i

    If bytFoundCnt = 0 And sCurrentKeys <> gsServerStringsFile Then
        Call ReadAllKeywords
    End If

    If bytFoundCnt = 1 Then
        sCurrentKeys = gsServerStringsFile ' Originaleinstellung merken
        gsServerStringsFile = gsarrServerStrArr(bytFoundItem) ' umbiegen
        If sCurrentKeys <> gsServerStringsFile Then Call ReadAllKeywords ' und einlesen
        gsServerStringsFile = sCurrentKeys ' wieder zur�ckbiegen
        sCurrentKeys = gsarrServerStrArr(bytFoundItem) ' und merken was wir derzeit geladen haben
    End If

End Sub
