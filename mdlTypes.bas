Attribute VB_Name = "mdlTypes"
Option Explicit
Enum ePlaybackEngine
    pMediaPlayer = 0
    pUnused = 1
    pMp3 = 2
End Enum
Private Type gPlayback
    pCurrentFile As String
    pCurrentEngine As ePlaybackEngine
    pPlaying As Boolean
    pRepeatCurrentTrack As Boolean
End Type
Private Type gFormConst
    fWidth As Integer
    fHeight As Integer
    fLeft As Integer
    fTop As Integer
End Type
Private Type gLastPositions
    lParentForm As gFormConst
End Type
'Private Type gRegInfo
'    rName As String
'    rPassword As String
    'rRegistered As Boolean
'End Type
Private Type gChanFolder
    lEnabled As Boolean
    lChannel As String
End Type
Private Type gIdent
    iEnabled As Boolean
    iShow As Boolean
    iPort As String
    iUserID As String
    iSystem As String
End Type
Private Type gOptions
    oShowChannelFolder As Boolean
    oShowQuit As Boolean
    oShowJoinPart As Boolean
    oShowModes As Boolean
    oShowTopics As Boolean
    oShowKicks As Boolean
    oWhois As Boolean
    oSkipMOTD As Boolean
    oShowMOTD As Boolean
    oReJoin As Boolean
    oShowAddress As Boolean
    oShowNotifyInActiveWindow As Boolean
    oWhoisNotify As Boolean
    oInvisable As Boolean
    oServerMessages As Boolean
    oOpMessages As Boolean
End Type
Private Type gSettings
    sBanNickname As String
    sRetrieveAddressFromWhoisForKickBan As Boolean
    sRetrieveAddressFromWhoisForBan As Boolean
    sBanChannel As String
    sLatestVersionData As String
    sShowExtraProgress As Boolean
    
    sBalloons As Boolean
    sConnectionManagerVisible As Boolean
    sSystemStatsConsoleVisible As Boolean
    sSetupWizardVisible As Boolean
    sSplashVisible As Boolean
    sChannelFolderVisible As Boolean
    sAdvancedSystemStatsVisisble As Boolean
    sMainVisisble As Boolean
    sCustomizeVisible As Boolean
    sAddMediaVisible As Boolean
    sPlaylistVisible As Boolean
    sMobileMixerVisible As Boolean
    sChannelListVisible As Boolean
    sIRCServerVisible As Boolean
    sWebVisible As Boolean
    sMOTDVisible As Boolean
    sByPassStartupScreen As Boolean
    sDownloadManager As Boolean
    sBorderlessObjects As Boolean
    sReconnectOnDisconnect As Boolean
    sAutoSelectAlternateNickname As Boolean
    sRefreshPictureColors As Boolean
    sDCCEnabled As Boolean
    sUseNickCompletor As Boolean
    sAutoJoinOnInvite As Boolean
    sLastWindowPos As gLastPositions
    sOptions As gOptions
    sNickname As String
    sPassword As String
    sEMail As String
    sRealName As String
    sNetwork As String
    sPort As String
    sIdent As gIdent
    sServer As String
    sShowOptionsOnStartup As Boolean
    sShowSplashOnStartup As Boolean
    sSkipMOTD As Boolean
    sHomepage As String
    sISON As Boolean
    sConnectOnStartup As Boolean
    sColors As String
    sSetupActivated As Boolean
    sTextCount As Integer
    sActiveServerForm As Form
    sBGPicture As String
    sBGColor As Integer
    sBackgroundWebpage As Boolean
    sGeneralPrompts As Boolean
    sDCCPrompts As Boolean
    sAutoJoinActivated As Boolean
    sShuffle As Boolean
    sContinuousPlay As Boolean
    sAddJoinedChannelsToChannelFolder As Boolean
    sShowQuickmix As Boolean
    sAutoJoinEnabled As Boolean
    sNavigateOnStartup As Boolean
    sLogoTwitchOnPeaks As Boolean
    sSearchForMedia As Boolean
    sChannelCount As Integer
    sPlayNext As String
    sExlusiveToMp3InPlaylist As Boolean
    sAlwaysShowAudioSettings As Boolean
    sNotifyVisible As Boolean
    sShowNotifyWindow As Boolean
    sShowQuickNotify As Boolean
    sShowServerOnStartup As Boolean
    sHandleErrors As Boolean
    sApplyThemeToIRCColors As Boolean
    sSaveIRCColorsToTheme As Boolean
    sButtonType As Integer
    sAudioServer As Boolean
    sFileOfferInChannel As Boolean
    sOfferWhenPlayed As Boolean
    sEnableList As Boolean
    sEnableSearch As Boolean
    sPlaySounds As Boolean
    sShowTips As Boolean
    sAutosizeStatusbarItems As Boolean
    sSecureQuery As Boolean
    sColoredNicklist As Boolean
    sTimeStamping As Boolean
    sUpdateCheck As Boolean
    sShowSmallNetworks As Boolean
    sServerMinimum As Integer
    sAutoPortScanner As Boolean
    sTestConnectionsLoaded As Boolean
    sShowWhoisInChannel As Boolean
    sEnding As Boolean
End Type
Private Type gNetwork
    nDescription As String
End Type
Private Type gModes
    mI As Boolean
    mW As Boolean
    mS As Boolean
End Type
Private Type gServer
    sDescription As String
    sServer As String
    sPortRange As String
    sPassword As String
    sNetwork As Integer
End Type
Private Type gServers
    sNetwork(1000) As gNetwork
    sServer(1000) As gServer
    sNetworkCount As Integer
    sNetworkUBound As Integer
    sServerCount As Integer
    sServerUBound As Integer
End Type
Private Type gInitialAudioValues
    iInitialBassEnabled As Boolean
    iInitialTrebleEnabled As Boolean
    iInitialCDAudioEnabled As Boolean
    iInitialLineInEnabled As Boolean
    iInitialMicEnabled As Boolean
    iInitialWaveEnabled As Boolean
    iBass As Long
    iTreble As Long
    iMic As Long
    iCDAudio As Long
    iLineIN As Long
    iWave As Long
End Type
Private Type gSecureQuery
    sAccepted As Boolean
    sAddToIgnore As Boolean
    sAddToNotify As Boolean
End Type


Global lSettings As gSettings
Global lServers As gServers
Global lPlayback As gPlayback
'Global lRegInfo As gRegInfo
Global lModes As gModes
Global lRedColor As Long
Global lGreenColor As Long
Global lBlueColor As Long
Global lInitialAudioValues As gInitialAudioValues
Global lSecureQuery As gSecureQuery
Global ShowFontType As Integer
Global lReturnColor As String
Global bDocked As Boolean
Global lDockedWidth As Long
Global lDockedHeight As Long
Global ProgressScrolling As Boolean
Global ACTION_CHANNEL As String
Global lDirReturn As String
