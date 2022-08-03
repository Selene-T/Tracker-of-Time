Option Explicit On
Option Strict On

Public Class frmTrackerOfTime

    'Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    'Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    'Public Declare Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Integer, ByVal lpBaseAddress As Integer, ByRef lpBuffer As Integer, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer

    ' Constant variables used throughout the app. The most important here is the 'IS_64BIT' as this needs to be set if compiling in x64
    Private Const PROCESS_ALL_ACCESS As Integer = &H1F0FFF
    Private Const CHECK_COUNT As Byte = 124
    Private Const IS_64BIT As Boolean = True
    Private Const VER As String = "4.0.7"
    Private p As Process = Nothing

    ' Variables used to determine what emulator is connected, its state, and its starting memory address
    Private romAddrStart As Integer = &HDFE40000
    Private romAddrStart64 As Int64 = 0
    Private emulator As String = String.Empty
    Private keepRunning As Boolean = False
    Private zeldazFails As Integer = 0
    Private isSoH As Boolean = False
    Private wasSoH As Boolean = False
    Private locSwap(30) As String

    ' Variables for a variety of rom info used in the scan
    Private pedestalRead As Byte = 0
    Private playerName As String = String.Empty
    Private randoVer As String = String.Empty
    Private rainbowBridge(1) As Byte
    Private aGetQuantity(14) As Boolean
    Private aRandoSet() As Byte
    Private lastFirstEnum As Byte = 255

    ' Arrays for tracking checks
    Private keyCount As Integer = 337
    Private aKeys(keyCount) As keyCheck
    Private aKeysDungeons(11)() As keyCheck
    Private canDungeon(11) As Boolean
    Private aQuestRewardsCollected(22) As Boolean
    Private aEquipment(31) As Boolean
    Private aQuestItems(31) As Boolean
    Private aUpgrades(31) As Boolean
    Private knifeCheck As Boolean

    ' RTB variables
    Private emboldenList As New List(Of String)
    Private rtbRefresh As Byte = 0
    Private rtbLines As Byte = 14
    Private lastArea As String = String.Empty
    Private lastOutput As New List(Of String)

    ' Arrays for location scanning and their settings
    Public arrLocation(CHECK_COUNT) As Integer
    Private arrChests(CHECK_COUNT) As Long
    Private arrHigh(CHECK_COUNT) As Byte
    Private arrLow(CHECK_COUNT) As Byte

    ' Variables for from settings
    Private firstRun As Boolean = True
    Private showSetting As Boolean = False

    ' Variable for the colour of the highlight, blended from both front and background colour
    Private cBlend As New Color

    ' Arrays for the displaying of the dungeon name over a quest reward
    Private aQuestRewardsText(22) As Byte
    Private aDungeonLetters() As String = {"", "FREE", "DEKU", "DC", "JABU", "FRST", "FIRE", "WTR", "SPRT", "SHDW"}

    ' Variables for tracking logic checks
    Private allItems As String = String.Empty
    Private goldSkulltulas As Byte = 0
    Private canMagic As Boolean = False
    Private canAdult As Boolean = False
    Private canYoung As Boolean = False
    Private isAdult As Boolean = False
    Private magicBeans As Byte = 0
    Private aDungeonKeys(7) As Byte
    Private aBossKeys(7) As Boolean
    Private aWarps(7) As String
    Private aDungeonRewards(7) As Byte
    Private aReachA(255) As Boolean
    Private aReachY(255) As Boolean
    Private aExitMap(37)() As Byte
    Private aVisited(37) As Boolean
    Private iER As Byte = 0
    ' Old ER is for detecting if there is an ER change, thus triggering the clearing of exits.
    ' It will not be reset when stopping the scan, as I do not want people's ER progress to reset randomly.
    ' This may be used later for an 'Entrance Reset' option.
    Private iOldER As Byte = 255
    Private iLastMinimap As Integer = 101
    Private iRoom As Byte = 0
    Private aIconLoc(30) As String
    Private aIconName(30) As String
    Private lRegions As New List(Of Rectangle)
    Private justTheTip As New ToolTip
    Private lastTip As Byte = 255

    'Private aRegions(30) As Region
    Private bSpawnWarps As Boolean = False
    Private bSongWarps As Boolean = False
    Private maxLife As Byte = 0
    Private aAddresses(19) As Integer

    ' Variables for detecting room info
    Private Const CUR_ROOM_ADDR As Integer = &H1C8544
    Private lastRoomScan As Long = 0

    ' Cheat menu vars
    Private iCheat As Byte = 0
    Private aCheat() As Keys = {Keys.Up, Keys.Up, Keys.Down, Keys.Down, Keys.Left, Keys.Right, Keys.Left, Keys.Right}


    ' Arrays of MQ settings, a current and old to compare to so updating only happens on changes
    Private aMQ(11) As Boolean
    Private aMQOld(11) As Boolean

    ' Array of Objects
    Private aoLabels(23) As Label
    Private aoDungeonLabels(11) As Label
    Private aoSmallKeys(7) As PictureBox
    Private aoBossKeys(10) As PictureBox
    Private aoCompasses(9) As PictureBox
    Private aoMaps(9) As PictureBox
    Private aoQuestItems(22) As PictureBox
    Private aoQuestItemImages(22) As Image
    Private aoQuestItemImagesEmpty(22) As Image
    Private aoEquipment(20) As PictureBox
    Private aoInventory(23) As PictureBox

    Private Sub frmTrackerOfTime_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        stopScanning()
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = aCheat(iCheat) Then
            incB(iCheat)
            If iCheat = aCheat.Length Then
                btnTest.Visible = True
                'Button2.Visible = True
                iCheat = 0
            End If
        Else
            iCheat = 0
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    ' On load, populate the locations array
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.setFirstTime Then
            showSetting = True
            updateShowSettings()
        End If
        Me.Text = "Tracker of Time v" & VER
        For i = 0 To aKeys.Length - 1
            aKeys(i) = New keyCheck
        Next

        makeArrayObjects()
        redimArrayExits()
        custMenu.highlight = Me.ForeColor
        custMenu.backColour = Me.BackColor
        custMenu.foreColour = Me.ForeColor
        mnuOptions.Renderer = New custMenu

        loadSettings()

        For i As Integer = 0 To arrLocation.Length - 1
            arrChests(i) = 0
            arrHigh(i) = 0
            arrLow(i) = 31
        Next

        ' Used to handle the click event for all of the label checkboxes
        For Each Label In pnlSettings.Controls.OfType(Of Label)().Where(Function(lbl As Label) Mid(lbl.Name, 1, 3) = "lcx" Or Mid(lbl.Name, 1, 3) = "lbl")
            AddHandler Label.MouseClick, AddressOf handleLCXMouseClick
        Next
        ' Used to give all label trackbars arrows and handle the click events
        For Each Label In pnlSettings.Controls.OfType(Of Label)().Where(Function(lbl As Label) Mid(lbl.Name, 1, 3) = "ltb")
            AddHandler Label.MouseClick, AddressOf handleLTBMouseClick
            AddHandler Label.MouseDoubleClick, AddressOf handleLTBMouseClick
            AddHandler Label.Paint, AddressOf drawArrows
        Next
        ' Used to give a border to each panel below the item display
        For Each Panel In pnlMain.Controls.OfType(Of Panel)().Where(Function(pnl As Panel) pnl.Height < 50)
            AddHandler Panel.Paint, AddressOf pnlDrawBorder
        Next
        ' Used to give a border to each panel below the item display
        For Each Label In pnlWorldMap.Controls.OfType(Of Label)()
            AddHandler Label.Paint, AddressOf lblDrawBorder
        Next

        populateLocations()

    End Sub

    Private Sub makeArrayObjects()
        ' Link the labels for displaying progress into an array
        aoLabels(0) = lblKokiriForest
        aoLabels(1) = lblLostWoods
        aoLabels(2) = lblSacredForestMeadow
        aoLabels(3) = lblHyruleField
        aoLabels(4) = lblLonLonRanch
        aoLabels(5) = lblMarket
        aoLabels(6) = lblTempleOfTime
        aoLabels(7) = lblHyruleCastle
        aoLabels(8) = lblKakarikoVillage
        aoLabels(9) = lblGraveyard
        aoLabels(10) = lblDMTrail
        aoLabels(11) = lblDMCrater
        aoLabels(12) = lblGoronCity
        aoLabels(13) = lblZorasRiver
        aoLabels(14) = lblZorasDomain
        aoLabels(15) = lblZorasFountain
        aoLabels(16) = lblLakeHylia
        aoLabels(17) = lblGerudoValley
        aoLabels(18) = lblGerudoFortress
        aoLabels(19) = lblHauntedWasteland
        aoLabels(20) = lblDesertColossus
        aoLabels(21) = lblOutsideGanonsCastle
        aoLabels(22) = lblQuestGoldSkulltulas
        aoLabels(23) = lblQuestMasks

        ' Link the dungeon labels for displaying progress into an array
        aoDungeonLabels(0) = lblDekuTree
        aoDungeonLabels(1) = lblDodongosCavern
        aoDungeonLabels(2) = lblJabuJabusBelly
        aoDungeonLabels(3) = lblForestTemple
        aoDungeonLabels(4) = lblFireTemple
        aoDungeonLabels(5) = lblWaterTemple
        aoDungeonLabels(6) = lblSpiritTemple
        aoDungeonLabels(7) = lblShadowTemple
        aoDungeonLabels(8) = lblBottomOfTheWell
        aoDungeonLabels(9) = lblIceCavern
        aoDungeonLabels(10) = lblGerudoTrainingGround
        aoDungeonLabels(11) = lblGanonsCastle

        ' Link the pictureboxes of each small key into an array
        aoSmallKeys(0) = pbxFoTSmallKey
        aoSmallKeys(1) = pbxFiTSmallKey
        aoSmallKeys(2) = pbxWTSmallKey
        aoSmallKeys(3) = pbxSpTSmallKey
        aoSmallKeys(4) = pbxShTSmallKey
        aoSmallKeys(5) = pbxBotWSmallKey
        aoSmallKeys(6) = pbxGTGSmallKey
        aoSmallKeys(7) = pbxGCSmallKey

        ' Link the pictureboxes of each boss key into an array
        aoBossKeys(3) = pbxFoTBossKey
        aoBossKeys(4) = pbxFiTBossKey
        aoBossKeys(5) = pbxWTBossKey
        aoBossKeys(6) = pbxSpTBossKey
        aoBossKeys(7) = pbxShTBossKey
        aoBossKeys(10) = pbxGCBossKey

        ' Link the pictureboxes of each compass into an array
        aoCompasses(0) = pbxDTCompass
        aoCompasses(1) = pbxDCCompass
        aoCompasses(2) = pbxJBCompass
        aoCompasses(3) = pbxFoTCompass
        aoCompasses(4) = pbxFiTCompass
        aoCompasses(5) = pbxWTCompass
        aoCompasses(6) = pbxSpTCompass
        aoCompasses(7) = pbxShTCompass
        aoCompasses(8) = pbxBotWCompass
        aoCompasses(9) = pbxICCompass

        ' Link the pictureboxes of each map into an array
        aoMaps(0) = pbxDTMap
        aoMaps(1) = pbxDCMap
        aoMaps(2) = pbxJBMap
        aoMaps(3) = pbxFotMap
        aoMaps(4) = pbxFiTMap
        aoMaps(5) = pbxWTMap
        aoMaps(6) = pbxSpTMap
        aoMaps(7) = pbxShTMap
        aoMaps(8) = pbxBotWMap
        aoMaps(9) = pbxICMap

        ' Link the pictureboxes of each quest item into an array
        aoQuestItems(0) = pbxMedalForest
        aoQuestItems(1) = pbxMedalFire
        aoQuestItems(2) = pbxMedalWater
        aoQuestItems(3) = pbxMedalSpirit
        aoQuestItems(4) = pbxMedalShadow
        aoQuestItems(5) = pbxMedalLight
        aoQuestItems(6) = pbxMinuetOfForest
        aoQuestItems(7) = pbxBoleroOfFire
        aoQuestItems(8) = pbxSerenadeOfWater
        aoQuestItems(9) = pbxRequiemOfSpirit
        aoQuestItems(10) = pbxNocturneOfShadow
        aoQuestItems(11) = pbxPreludeOfLight
        aoQuestItems(12) = pbxZeldasLullaby
        aoQuestItems(13) = pbxEponasSong
        aoQuestItems(14) = pbxSariasSong
        aoQuestItems(15) = pbxSunsSong
        aoQuestItems(16) = pbxSongOfTime
        aoQuestItems(17) = pbxSongOfStorms
        aoQuestItems(18) = pbxStoneKokiri
        aoQuestItems(19) = pbxStoneGoron
        aoQuestItems(20) = pbxStoneZora
        aoQuestItems(21) = pbxStoneOfAgony
        aoQuestItems(22) = pbxGerudosCard

        ' Link the pictureboxes of each equipment item into an array
        aoEquipment(0) = pbxQuiver
        aoEquipment(1) = pbxBulletBag
        aoEquipment(2) = pbxKokiriSword
        aoEquipment(3) = pbxMasterSword
        aoEquipment(4) = pbxBiggoronsSword
        aoEquipment(5) = pbxWallet
        aoEquipment(6) = pbxBombBag
        aoEquipment(7) = pbxDekuShield
        aoEquipment(8) = pbxHylianShield
        aoEquipment(9) = pbxMirrorShield
        aoEquipment(10) = pbxStoneOfAgony
        aoEquipment(11) = pbxGauntlet
        aoEquipment(12) = pbxKokiriTunic
        aoEquipment(13) = pbxGoronTunic
        aoEquipment(14) = pbxZoraTunic
        aoEquipment(15) = pbxGerudosCard
        aoEquipment(16) = pbxScale
        aoEquipment(17) = pbxKokiriBoots
        aoEquipment(18) = pbxIronBoots
        aoEquipment(19) = pbxHoverBoots
        aoEquipment(20) = pbxBrokenKnife

        ' Link the pictureboxes of each inventory item into an array
        aoInventory(0) = pbx01
        aoInventory(1) = pbx02
        aoInventory(2) = pbx03
        aoInventory(3) = pbx04
        aoInventory(4) = pbx05
        aoInventory(5) = pbx06
        aoInventory(6) = pbx07
        aoInventory(7) = pbx08
        aoInventory(8) = pbx09
        aoInventory(9) = pbx10
        aoInventory(10) = pbx11
        aoInventory(11) = pbx12
        aoInventory(12) = pbx13
        aoInventory(13) = pbx14
        aoInventory(14) = pbx15
        aoInventory(15) = pbx16
        aoInventory(16) = pbx17
        aoInventory(17) = pbx18
        aoInventory(18) = pbx19
        aoInventory(19) = pbx20
        aoInventory(20) = pbx21
        aoInventory(21) = pbx22
        aoInventory(22) = pbx23
        aoInventory(23) = pbx24

        For i As Byte = 0 To 22
            refreshQuestItemImages(i)
        Next
    End Sub
    Private Sub refreshQuestItemImages(ByVal img As Byte)
        Select Case img
            Case 0
                aoQuestItemImages(img) = My.Resources.medalForest
                aoQuestItemImagesEmpty(img) = My.Resources.medalForestEmpty
            Case 1
                aoQuestItemImages(img) = My.Resources.medalFire
                aoQuestItemImagesEmpty(img) = My.Resources.medalFireEmpty
            Case 2
                aoQuestItemImages(img) = My.Resources.medalWater
                aoQuestItemImagesEmpty(img) = My.Resources.medalWaterEmpty
            Case 3
                aoQuestItemImages(img) = My.Resources.medalSpirit
                aoQuestItemImagesEmpty(img) = My.Resources.medalSpiritEmpty
            Case 4
                aoQuestItemImages(img) = My.Resources.medalShadow
                aoQuestItemImagesEmpty(img) = My.Resources.medalShadowEmpty
            Case 5
                aoQuestItemImages(img) = My.Resources.medalLight
                aoQuestItemImagesEmpty(img) = My.Resources.medalLightEmpty
            Case 6
                aoQuestItemImages(img) = My.Resources.songForest
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyWarp
            Case 7
                aoQuestItemImages(img) = My.Resources.songFire
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyWarp
            Case 8
                aoQuestItemImages(img) = My.Resources.songWater
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyWarp
            Case 9
                aoQuestItemImages(img) = My.Resources.songSpirit
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyWarp
            Case 10
                aoQuestItemImages(img) = My.Resources.songShadow
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyWarp
            Case 11
                aoQuestItemImages(img) = My.Resources.songLight
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyWarp
            Case 12 To 17
                aoQuestItemImages(img) = My.Resources.songNormal
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyNormal
            Case 18
                aoQuestItemImages(img) = My.Resources.stoneKokiri
                aoQuestItemImagesEmpty(img) = My.Resources.stoneKokiriEmpty
            Case 19
                aoQuestItemImages(img) = My.Resources.stoneGoron
                aoQuestItemImagesEmpty(img) = My.Resources.stoneGoronEmpty
            Case 20
                aoQuestItemImages(img) = My.Resources.stoneZora
                aoQuestItemImagesEmpty(img) = My.Resources.stoneZoraEmpty
            Case 21
                aoQuestItemImages(img) = My.Resources.stoneOfAgony
                aoQuestItemImagesEmpty(img) = My.Resources.stoneOfAgonyEmpty
            Case 22
                aoQuestItemImages(img) = My.Resources.gerudosCard
                aoQuestItemImagesEmpty(img) = My.Resources.gerudosCardEmpty
        End Select
    End Sub
    Private Sub loadSettings()
        'ddThemes.SelectedIndex = My.Settings.setTheme
        changeTheme(My.Settings.setTheme)
        subMenuCheck(My.Settings.setTheme)
        updateShowSettings()
        For Each Label In pnlSettings.Controls.OfType(Of Label)().Where(Function(lbl As Label) Mid(lbl.Name, 1, 3) = "ltb")
            updateLTB(Label.Name)
        Next
    End Sub
    Private Sub updateShowSettings()
        ShowSettingsToolStripMenuItem.Text = "Settings " & IIf(showSetting, "<", ">").ToString
        resizeForm()
    End Sub
    Private Sub redimArrayExits()
        ' Create default arrays. 255 is blank because 0 is used for KF Main. Also start with all true because we will false out the ones after depending on the ER settings
        For i = 0 To aExitMap.Length - 1
            ReDim aExitMap(i)(6)
            aVisited(i) = True
            For ii = 0 To 6
                aExitMap(i)(ii) = 255
            Next
        Next
        createArrayExits()
    End Sub

    Private Sub createArrayExits(Optional typeOfER As Byte = 0)
        ' Populates the exits with the default exits

        ' Overworld Exits
        ' MK Entrance (27-29)
        aExitMap(0)(0) = 7      ' HF
        aExitMap(0)(1) = 10     ' MK
        ' MK Back Alley (30-31)
        aExitMap(1)(0) = 10     ' MK
        aExitMap(1)(1) = 10     ' MK
        ' MK (Young) (32-33)
        aExitMap(2)(0) = 13     ' HC
        aExitMap(2)(1) = 197    ' MK Entrance
        aExitMap(2)(2) = 11     ' Outside ToT
        aExitMap(2)(3) = 199    ' MK Back Alley (right)
        aExitMap(2)(4) = 198    ' MK Back Alley (left)
        ' MK (Adult) (34)
        aExitMap(3)(0) = 51     ' OGC
        aExitMap(3)(1) = 197    ' MK Entrance
        aExitMap(3)(2) = 11     ' Outside ToT
        ' Outside ToT (35-37)
        aExitMap(4)(0) = 200    ' ToT Front
        aExitMap(4)(1) = 10     ' MK
        ' ToT (67)
        aExitMap(5)(0) = 11     ' Outside ToT
        ' HF (81)
        aExitMap(6)(0) = 15     ' KV Main
        aExitMap(6)(1) = 4      ' LW Bridge
        aExitMap(6)(2) = 34     ' ZR Front
        aExitMap(6)(3) = 44     ' GV Hyrule Side
        aExitMap(6)(4) = 197    ' MK Entrance
        aExitMap(6)(5) = 9      ' LLR
        aExitMap(6)(6) = 42     ' LH Main
        ' KV (82)
        aExitMap(7)(1) = 20     ' DMT Lower
        aExitMap(7)(2) = 7      ' HF
        aExitMap(7)(3) = 18     ' GY Lower
        ' GY (83)
        aExitMap(8)(1) = 15     ' KV Main
        ' ZR (84)
        aExitMap(9)(0) = 7      ' HF
        aExitMap(9)(1) = 37     ' ZD Main
        aExitMap(9)(2) = 2      ' LW Front
        ' KF (85)
        aExitMap(10)(1) = 8      ' LW Between Bridge
        aExitMap(10)(2) = 2      ' LW Front
        ' SFM (86)
        aExitMap(11)(1) = 3     ' LW Behind Mido
        ' LH (87)
        aExitMap(12)(0) = 7     ' HF
        aExitMap(12)(2) = 37    ' ZD Main
        ' ZD (88)
        aExitMap(13)(0) = 40    ' ZF Main
        aExitMap(13)(1) = 36    ' ZR Behind Waterfall
        aExitMap(13)(2) = 42    ' LH Main
        ' ZF (89)
        aExitMap(14)(2) = 38    ' ZD Behind King
        ' GV (90)
        aExitMap(15)(0) = 42    ' LH Main
        aExitMap(15)(1) = 7     ' HF
        aExitMap(15)(2) = 46    ' GF Main
        ' LW (91)
        aExitMap(16)(0) = 5     ' SFM Main
        aExitMap(16)(1) = 0     ' KF Main
        aExitMap(16)(2) = 35    ' ZR Main
        aExitMap(16)(3) = 30    ' GC Shortcut
        aExitMap(16)(4) = 0     ' KF Main
        aExitMap(16)(5) = 7     ' HF
        ' DC (92)
        aExitMap(17)(1) = 49    ' HW Colossus Side
        ' GF (93)
        aExitMap(18)(1) = 45    ' GV Gerudo Side
        aExitMap(18)(2) = 48    ' HW Gerudo Side
        ' HW (94)
        aExitMap(19)(0) = 50    ' DC
        aExitMap(19)(1) = 47    ' GF Behind Gate
        ' HC (95)
        aExitMap(20)(0) = 10    ' MK
        ' DMT (96)
        aExitMap(21)(1) = 29    ' GC Main
        aExitMap(21)(2) = 17    ' KV Behind Gate
        aExitMap(21)(3) = 22    ' DMC Upper
        ' DMC (97)
        aExitMap(22)(1) = 31    ' GC Darunia
        aExitMap(22)(2) = 21    ' DMT Upper
        ' GC (98)
        aExitMap(23)(0) = 25    ' DMC Lower Local
        aExitMap(23)(1) = 20    ' DMT Lower
        aExitMap(23)(2) = 2     ' LW Front
        ' LLR (99)
        aExitMap(24)(0) = 7     ' HF
        ' OGC (100)
        aExitMap(25)(1) = 10    ' MK

        ' Dungeon Exits
        aExitMap(10)(0) = 60    ' Deku Tree
        aExitMap(26)(0) = 1
        aExitMap(21)(0) = 69    ' Dodongo's Cavern
        aExitMap(27)(0) = 20
        aExitMap(14)(0) = 79    ' Jabu-Jabu's Belly
        aExitMap(28)(0) = 40
        aExitMap(11)(0) = 88    ' Forest Temple
        aExitMap(29)(0) = 6
        aExitMap(22)(0) = 108   ' Fire Temple
        aExitMap(30)(0) = 28
        aExitMap(12)(0) = 119   ' Water Temple
        aExitMap(31)(0) = 42
        aExitMap(17)(0) = 132   ' Spirit Temple
        aExitMap(32)(0) = 50
        aExitMap(8)(0) = 150    ' Shadow Temple
        aExitMap(33)(0) = 19
        aExitMap(7)(0) = 174    ' BotW
        aExitMap(34)(0) = 15
        aExitMap(14)(1) = 178   ' Ice Cavern
        aExitMap(35)(0) = 41
        aExitMap(18)(0) = 179   ' Gerudo Training Ground
        aExitMap(36)(0) = 46
        aExitMap(25)(0) = 193   ' Ganon's Castle
        aExitMap(37)(0) = 51
        aExitMap(37)(1) = 196
    End Sub
    Private Sub clearArrayExitsOverworld()
        ' Make sure Overworld ER is active
        If iER Mod 2 = 0 Then Exit Sub

        ' Clear visits to overworld maps
        For i = 0 To 26
            aVisited(i) = False
        Next

        ' Unlink all the overworld map exits
        ' MK Entrance (27-29)
        aExitMap(0)(0) = 255    ' HF
        aExitMap(0)(1) = 255    ' MK
        ' MK Back Alley (30-31)
        aExitMap(1)(0) = 255    ' MK
        aExitMap(1)(1) = 255    ' MK
        ' MK (Young) (32-33)
        aExitMap(2)(0) = 255    ' HC
        aExitMap(2)(1) = 255    ' MK Entrance
        aExitMap(2)(2) = 255    ' Outside ToT
        aExitMap(2)(3) = 255    ' MK Back Alley (right)
        aExitMap(2)(4) = 255    ' MK Back Alley (left)
        ' MK (Adult) (34)
        aExitMap(3)(0) = 255    ' OGC
        aExitMap(3)(1) = 255    ' MK Entrance
        aExitMap(3)(2) = 255    ' Outside ToT
        ' Outside ToT (35-37)
        aExitMap(4)(0) = 255    ' ToT Front
        aExitMap(4)(1) = 255    ' MK
        ' ToT (67)
        aExitMap(5)(0) = 255    ' Outside ToT
        ' HF (81)
        aExitMap(6)(0) = 255    ' KV Main
        aExitMap(6)(1) = 255    ' LW Bridge
        aExitMap(6)(2) = 255    ' ZR Front
        aExitMap(6)(3) = 255    ' GV Hyrule Side
        aExitMap(6)(4) = 255    ' MK Entrance
        aExitMap(6)(5) = 255    ' LLR
        aExitMap(6)(6) = 255    ' LH Main
        ' KV (82)
        aExitMap(7)(1) = 255    ' DMT Lower
        aExitMap(7)(2) = 255    ' HF
        aExitMap(7)(3) = 255    ' GY Lower
        ' GY (83)
        aExitMap(8)(1) = 255    ' KV Main
        ' ZR (84)
        aExitMap(9)(0) = 255    ' HF
        aExitMap(9)(1) = 255    ' ZD Main
        aExitMap(9)(2) = 255    ' LW Front
        ' KF (85)
        aExitMap(10)(1) = 255    ' LW Bridge
        aExitMap(10)(2) = 255    ' LW Front
        ' SFM (86)
        aExitMap(11)(1) = 255    ' LW Behind Mido
        ' LH (87)
        aExitMap(12)(1) = 255    ' HF
        aExitMap(12)(2) = 255    ' ZD Main
        ' ZD (88)
        aExitMap(13)(0) = 255    ' ZF Main
        aExitMap(13)(1) = 255    ' ZR Behind Waterfall
        aExitMap(13)(2) = 255    ' LH Main
        ' ZF (89)
        aExitMap(14)(2) = 255    ' ZD Behind King
        ' GV (90)
        aExitMap(15)(0) = 255    ' LH Main
        aExitMap(15)(1) = 255    ' HF
        aExitMap(15)(2) = 255    ' GF Main
        ' LW (91)
        aExitMap(16)(0) = 255    ' SFM Main
        aExitMap(16)(1) = 255    ' KF Main
        aExitMap(16)(2) = 255    ' ZR Main
        aExitMap(16)(3) = 255    ' GC Shortcut
        aExitMap(16)(4) = 255    ' KF Main
        aExitMap(16)(5) = 255    ' HF
        ' DC (92)
        aExitMap(17)(1) = 255    ' HW C olossus Side
        ' GF (93)
        aExitMap(18)(1) = 255    ' GV Gerudo Side
        aExitMap(18)(2) = 255    ' HW Gerudo Side
        ' HW (94)
        aExitMap(19)(0) = 255    ' DC
        aExitMap(19)(1) = 255    ' GF Behind Gate
        ' HC (95)
        aExitMap(20)(0) = 255    ' MK
        ' DMT (96)
        aExitMap(21)(1) = 255    ' GC Main
        aExitMap(21)(2) = 255    ' KV Behind Gate
        aExitMap(21)(3) = 255    ' DMC Upper
        ' DMC (97)
        aExitMap(22)(1) = 255    ' GC Darunia
        aExitMap(22)(2) = 255    ' DMT Upper
        ' GC (98)
        aExitMap(23)(0) = 255    ' DMC Lower Local
        aExitMap(23)(1) = 255    ' DMT Lower
        aExitMap(23)(2) = 255    ' LW Front
        ' LLR (99)
        aExitMap(24)(0) = 255    ' HF
        ' OGC (100)
        aExitMap(25)(1) = 255    ' MK
    End Sub
    Private Sub clearArrayExitsDungeons()
        ' Make sure Dungeon ER is active
        If iER < 2 Then Exit Sub

        ' Clear visits to overworld maps
        For i = 27 To 37
            aVisited(i) = False
        Next

        ' Unlink all the dungeon related exits
        aExitMap(10)(0) = 255   ' Deku Tree
        aExitMap(26)(0) = 255
        aExitMap(21)(0) = 255   ' Dodongo's Cavern
        aExitMap(27)(0) = 255
        aExitMap(14)(0) = 255   ' Jabu-Jabu's Belly
        aExitMap(28)(0) = 255
        aExitMap(11)(0) = 255   ' Forest Temple
        aExitMap(29)(0) = 255
        aExitMap(22)(0) = 255   ' Fire Temple
        aExitMap(30)(0) = 255
        aExitMap(12)(0) = 255   ' Water Temple
        aExitMap(31)(0) = 255
        aExitMap(17)(0) = 255   ' Spirit Temple
        aExitMap(32)(0) = 255
        aExitMap(8)(0) = 255    ' Shadow Temple
        aExitMap(33)(0) = 255
        aExitMap(7)(0) = 255    ' BotW
        aExitMap(34)(0) = 255
        aExitMap(14)(1) = 255   ' Ice Cavern
        aExitMap(35)(0) = 255
        aExitMap(18)(0) = 255   ' Gerudo Training Ground
        aExitMap(36)(0) = 255
        aExitMap(25)(0) = 193   ' Ganon's Castle
        aExitMap(37)(0) = 255
        aExitMap(37)(1) = 196
        aVisited(37) = True     ' Since GaC is not randomized yet, we will always enabled it
    End Sub


    Private Sub populateLocations()
        ' Populate location's addresses and clear up the chest data

        ' Unset all MQs
        For i = 0 To aMQ.Length - 1
            aMQ(i) = False
            aMQOld(i) = True
        Next
        updateMQs()

        ' Reset some global variables
        randoVer = String.Empty
        playerName = String.Empty
        canMagic = False
        canAdult = False
        canYoung = False
        isAdult = False
        rainbowBridge(0) = 4
        rainbowBridge(1) = 64
        magicBeans = 0
        goldSkulltulas = 0
        maxLife = 0
        iER = 0
        iLastMinimap = 101
        lastArea = String.Empty
        lastOutput.Clear()
        lastTip = 255
        lastFirstEnum = 255
        'isTriforceHunt = False

        For i = 0 To 7
            aDungeonKeys(i) = 0
            aBossKeys(i) = False
            aWarps(i) = String.Empty
            aDungeonRewards(i) = 100
        Next
        For i = 0 To 11
            canDungeon(i) = False
        Next
        For i = 0 To aReachA.Length - 1
            aReachA(i) = False
            aReachY(i) = False
        Next
        For i = 0 To aAddresses.Length - 1
            aAddresses(i) = 0
        Next
        For i = 0 To aEquipment.Length - 1
            aEquipment(i) = False
            aQuestItems(i) = False
            aUpgrades(i) = False
        Next

        bSpawnWarps = False
        bSongWarps = False
        'clearArrayExits()
        createArrayExits()

        ' Clean up the UI
        clearItems()
        updateLabels()
        updateLabelsDungeons()
        pbxSpawnYoung.Image = My.Resources.spawnLocations
        pbxSpawnAdult.Image = My.Resources.spawnLocations

        ' Process the high/lows, but only the first time
        If firstRun Then
            redirectChecks(True)
            firstRun = False

            arrLocation(0) = &H11AD18 + 4       ' DMC/DMT/OGC Great Fairy Fountain (Events)
            arrLocation(1) = &H11AF80 + 4       ' Hyrule Field (Events) Big Poes Captured and Ocarina of Time
            arrLocation(2) = &H11B028 + 4       ' Lake Hylia (Events) Ruto's Letter, Open Water Temple, and Bean Plant
            arrLocation(3) = &H11A714 + 12      ' Fire Temple (Standing)
            arrLocation(4) = &H11A730 + 12      ' Water Temple (Standing)
            arrLocation(5) = &H11A768 + 12      ' Shadow Temple (Standing)
            arrLocation(6) = &H11A784 + 12      ' Bottom of the Well (Standing)
            arrLocation(7) = &H11A7A0 + 12      ' Ice Cavern (Standing)
            arrLocation(8) = &H11A7D8 + 12      ' Gerudo Training Ground (Standing)
            arrLocation(9) = &H11A810 + 12      ' Ganon’s Castle #1 (Standing)
            arrLocation(10) = &H11A880 + 12     ' Deku Tree Boss Room (Standing)
            arrLocation(11) = &H11A89C + 12     ' Dodongo's Cavern Boss Room (Standing)
            arrLocation(12) = &H11A8B8 + 12     ' Jabu-Jabu's Belly Boss Room (Standing)
            arrLocation(13) = &H11A8D4 + 12     ' Forest Temple Boss Room (Standing)
            arrLocation(14) = &H11A8F0 + 12     ' Fire Temple Boss Room (Standing)
            arrLocation(15) = &H11A90C + 12     ' Water Temple Boss Room (Standing)
            arrLocation(16) = &H11A928 + 12     ' Spirit Temple Boss Room (Standing)
            arrLocation(17) = &H11A944 + 12     ' Shadow Temple Boss Room (Standing)
            arrLocation(18) = &H11ACA8 + 12     ' Impa’s House (Standing)
            arrLocation(19) = &H11AD6C + 12     ' All Grottos (Standing)
            arrLocation(20) = &H11AE84 + 12     ' Windmill / Dampe's Grave (Standing)
            arrLocation(21) = &H11AEF4 + 12     ' Lon Lon Tower Item (Standing)
            arrLocation(22) = &H11AFB8 + 12     ' Graveyard (Standing)
            arrLocation(23) = &H11AFD4 + 12     ' Zora's River (Standing)
            arrLocation(24) = &H11B028 + 12     ' Lake Hylia (Standing)
            arrLocation(25) = &H11B060 + 12     ' Zora's Fountain (Standing)
            arrLocation(26) = &H11B07C + 12     ' Gerudo Valley (Standing)
            arrLocation(27) = &H11B0B4 + 12     ' Desert Colossus (Standing)
            arrLocation(28) = &H11B124 + 12     ' Death Mountain Trail (Standing)
            arrLocation(29) = &H11B140 + 12     ' Death Mountain Crater (Standing)
            arrLocation(30) = &H11B15C + 12     ' Goron City (Standing)
            arrLocation(31) = &H11A6A4          ' Deku Tree
            arrLocation(32) = &H11A6C0          ' Dodongo's Cavern
            arrLocation(33) = &H11A6DC          ' Jabu-Jabu's Belly
            arrLocation(34) = &H11A6F8          ' Forest Temple
            arrLocation(35) = &H11A714          ' Fire Temple
            arrLocation(36) = &H11A730          ' Water Temple
            arrLocation(37) = &H11A74C          ' Spirit Temple
            arrLocation(38) = &H11A768          ' Shadow Temple
            arrLocation(39) = &H11A784          ' Bottom of the Well
            arrLocation(40) = &H11A7A0          ' Ice Cavern
            arrLocation(41) = &H11A7BC          ' Ganon’s Castle #2
            arrLocation(42) = &H11A7D8          ' Gerudo Training Ground
            arrLocation(43) = &H11A810          ' Ganon’s Castle #1
            arrLocation(44) = &H11A89C          ' Dodongo's Cavern Boss Room
            arrLocation(45) = &H11AB04          ' Mido’s House
            arrLocation(46) = &H11AD6C          ' All Grottos
            arrLocation(47) = &H11AD88          ' Grave with Sun Song Chest
            arrLocation(48) = &H11ADA4          ' Graveyard Under Grave
            arrLocation(49) = &H11ADC0          ' Royal Grave
            arrLocation(50) = &H11AE84          ' Windmill / Dampe
            arrLocation(51) = &H11AFF0          ' Kokiri Forest
            arrLocation(52) = &H11B028          ' Lake Hylia
            arrLocation(53) = &H11B044          ' Zora's Domain
            arrLocation(54) = &H11B07C          ' Gerudo Valley
            arrLocation(55) = &H11B0B4          ' Desert Colossus
            arrLocation(56) = &H11B0D0          ' Gerudo’s Fortress
            arrLocation(57) = &H11B0EC          ' Haunted Wasteland
            arrLocation(58) = &H11B124          ' Death Mountain Trail
            arrLocation(59) = &H11B15C          ' Goron City
            arrLocation(60) = &H11A640          ' *Biggoron Check
            arrLocation(61) = &H11B490          ' *Big Fish
            arrLocation(62) = &H11B4A4          ' *Events 1: Egg from Malon, Obtained Epona, Won Cow
            arrLocation(63) = &H11B4A8          ' *Events 2: Zora Diving Game, Darunia’s Joy
            arrLocation(64) = &H11B4AC          ' *Events 3: Zelda’s Letter, Song from Impa, Sun Song??, opened Temple of Time, Rainbow Bridge
            arrLocation(65) = &H11B4B4          ' *Events 5: Scarecrow as Adult
            arrLocation(66) = &H11B4B8          ' *Events 6: Song at Colossus, Trials
            arrLocation(67) = &H11B4BC          ' *Events 7: Saria Gift, Skulltula trades, Barrier Lowered
            arrLocation(68) = &H11B4C0          ' *Item Collect #1
            arrLocation(69) = &H11B4C4          ' *Item Collection #2
            arrLocation(70) = &H11B4E8          ' *Item: Rolling Goron as Young + Adult Link
            arrLocation(71) = &H11B4EC          ' *Thaw Zora King
            arrLocation(72) = &H11B4F8          ' *Items: 1st and 2nd Scrubs, Lost Dog
            arrLocation(73) = &H11B894          ' *Scarecrow Song
            arrLocation(74) = &H11A66C          ' *Equipment checks, figured this would be easier
            arrLocation(75) = &H11A60C          ' *Check for Biggoron's Sword
            arrLocation(76) = &H11A670          ' *Upgrades
            arrLocation(77) = &H11A674          ' *Quest Items and Songs
            arrLocation(78) = &H11B46C          ' **Gold Skulltulas 1
            arrLocation(79) = &H11B470          ' **Gold Skulltulas 2
            arrLocation(80) = &H11B474          ' **Gold Skulltulas 3
            arrLocation(81) = &H11B478          ' **Gold Skulltulas 4
            arrLocation(82) = &H11B47C          ' **Gold Skulltulas 5
            arrLocation(83) = &H11B480          ' **Gold Skulltulas 6
            arrLocation(84) = &H11A6B4          ' ***Scrub Shuffle (Deku Tree)
            arrLocation(85) = &H11A6D0          ' ***Scrub Shuffle (Dodongo's Cavern)
            arrLocation(86) = &H11A6EC          ' ***Scrub Shuffle (Jabu-Jabu's Belly)
            arrLocation(87) = &H11A820          ' ***Scrub Shuffle (Ganon's Castle)
            arrLocation(88) = &H11A874          ' ***Scrub Shuffle (Hyrule Field Grotto)
            arrLocation(89) = &H11A900          ' ***Scrub Shuffle (Zora's River Grotto)
            arrLocation(90) = &H11A954          ' ***Scrub Shuffle (Sacred Forest Meadow Grotto)
            arrLocation(91) = &H11A970          ' ***Scrub Shuffle (Lake Hylia Grotto)
            arrLocation(92) = &H11A98C          ' ***Scrub Shuffle (Gerudo Valley Grotto)
            arrLocation(93) = &H11AA18          ' ***Scrub Shuffle (Lost Woods Grotto)
            arrLocation(94) = &H11AA88          ' ***Scrub Shuffle (Death Mountain Crater Grotto)
            arrLocation(95) = &H11AAC0          ' ***Scrub Shuffle (Goron City)
            arrLocation(96) = &H11AADC          ' ***Scrub Shuffle (Lon Lon Ranch)
            arrLocation(97) = &H11AAF8          ' ***Scrub Shuffle (Desert Colossus)
            arrLocation(98) = &H11B0A8          ' ***Scrub Shuffle (Lost Woods)
            arrLocation(99) = &H11AB84         ' *Shopsanity Checks
            arrLocation(100) = &H11AC54 + 12    ' Link's House (Standing)
            arrLocation(101) = &H11AC8C + 12    ' Lon Lon Ranch Stables (Standing)
            arrLocation(102) = &H11A6DC + 12    ' Jabu-Jabu's Belly (Standing)
            arrLocation(103) = &H11A6A4 + 4     ' Deku Tree (Events)
            arrLocation(104) = &H11A6C0 + 4     ' Dodongo's Cavern (Events)
            arrLocation(105) = &H11A6DC + 4     ' Jabu-Jabu's Belly (Events)
            arrLocation(106) = &H11A6F8 + 4     ' Forest Temple (Events)
            arrLocation(107) = &H11A714 + 4     ' Fire Temple (Events)
            arrLocation(108) = &H11A730 + 4     ' Water Temple (Events)
            arrLocation(109) = &H11A74C + 4     ' Spirit Temple (Events)
            arrLocation(110) = &H11A768 + 4     ' Shadow Temple (Events)
            arrLocation(111) = &H11A7A0 + 4     ' Ice Cavern (Events)
            arrLocation(112) = &H11A7D8 + 4     ' Gerudo Training Ground (Events)
            arrLocation(113) = &H11A810 + 4     ' Ganon’s Castle #1 (Events)
            arrLocation(114) = &H11B0EC + 12    ' Haunted Wasteland (Standing)
            arrLocation(115) = &H11B15C + 4     ' Goron City (Events)
            arrLocation(116) = &H11AFD4 + 4     ' Zora's River (Events)
            arrLocation(117) = &H11A784 + 4     ' BotW (Events)
            arrLocation(118) = &H11B178         ' Lon Lon Ranch
            arrLocation(119) = &H11B00C         ' Sacred Forest Meadow
            arrLocation(120) = &H11ADDC         ' Shooting Gallery
            arrLocation(121) = &H11AD50         ' GF DC/HC/ZF
            arrLocation(122) = &H11AD18         ' DMC/DMT/OGC Great Fairy Fountain
            arrLocation(123) = &H11ADF8         ' Temple of Time
            arrLocation(124) = &H11AF80         ' Hyrule Field
        End If

        For i As Integer = 0 To arrLocation.Length - 1
            arrChests(i) = 0
        Next
    End Sub

    Private Sub getAge()
        ' Checks for the current age variable, starts off with setting ages to not accessable
        canAdult = False
        canYoung = False

        Dim addrAge As Integer = If(isSoH, SAV(&H4), &H11A5D4)

        ' If 0 or 1 , set the age as accessable
        Select Case CByte(goRead(addrAge, 1))
            Case 0
                canAdult = True
                isAdult = True
                If aReachA(12) Then canYoung = True
            Case 1
                canYoung = True
                isAdult = False
                If aReachY(12) Then canAdult = True
        End Select
    End Sub
    Private Sub getDungeonItems()
        'arrLocation(104) = &H11A678         ' *Boss Key/Compass/Map 1
        'arrLocation(105) = &H11A67C         ' *Boss Key/Compass/Map 2
        'arrLocation(106) = &H11A680         ' *Boss Key/Compass/Map 3

        Dim stringItems As String = String.Empty
        Dim tempItems As String = String.Empty
        Dim valueItems As Integer = 0

        For i = 0 To 2
            ' Scans at &H11A678, &H11A67C, and &H11A680 
            tempItems = Hex(goRead(If(isSoH, soh.SAV(&HAC), &H11A678) + (i * 4)))

            ' Make sure all leading 0's are put back
            While tempItems.Length < 8
                tempItems = "0" & tempItems
            End While

            endianFlip(tempItems)

            ' Add info into main string
            stringItems = stringItems & tempItems
        Next

        For i = 0 To 10
            ' Grab only the lower hex of each paired check as the 16's place does not matter, store as value
            valueItems = CInt("&H" & Mid(stringItems, (i + 1) * 2, 1))

            ' Again, something that should never happen, but best cover for anyone hacking. This works by lowering the everything down to 3 bits (1-7)
            While valueItems > 7
                dec(valueItems, 8)
            End While

            If i <= 9 Then
                ' Map and compass checks are only for the first 10 dungeons
                If valueItems >= 4 Then
                    ' This is the Map check, 3rd bit (4). Make sure coloured map is visible
                    dec(valueItems, 4)
                    If Not aoMaps(i).Visible Then aoMaps(i).Visible = True
                Else
                    ' Or not if not found
                    If aoMaps(i).Visible Then aoMaps(i).Visible = False
                End If

                If valueItems >= 2 Then
                    ' This is the compass check, 2nd bit (2). Make sure coloured compass is visible
                    dec(valueItems, 2)
                    If Not aoCompasses(i).Visible Then aoCompasses(i).Visible = True
                Else
                    ' Or not if not found
                    If aoCompasses(i).Visible Then aoCompasses(i).Visible = False
                End If
            End If
            Select Case i
                Case 3 To 7, 10
                    If valueItems = 1 Then
                        ' This is the Boss Key check, 1st bit (1). Make sure coloured boss key is visible
                        If Not aoBossKeys(i).Visible Then aoBossKeys(i).Visible = True
                        aBossKeys(i - 3) = True
                    Else
                        ' Or not if not found
                        If aoBossKeys(i).Visible Then aoBossKeys(i).Visible = False
                        aBossKeys(i - 3) = False
                    End If

            End Select
        Next

        If Not aAddresses(9) = 0 Then scanDungeonRewards()
    End Sub
    Private Sub getGoldSkulltulas()
        ' Get gold skulltula count
        Dim bGS As Byte = CByte(goRead(If(isSoH, soh.SAV(&HD4), &H11A6A0 + 2), 1))

        ' Checks to see if the number of gold skulltula's have changed, if not then do not bother to do all that
        If bGS = goldSkulltulas Then Exit Sub

        ' Update your gold skulltulas
        goldSkulltulas = bGS

        With pbxGoldSkulltula
            ' If not visible, make visible, and reset image to default to remove drawn numbers
            If Not .Visible Then .Visible = True
            .Image = My.Resources.goldSkulltula

            ' If single digit
            Dim xPos As Byte = 34
            ' If double digits
            If bGS > 9 Then xPos = 19
            ' If triple digits
            If bGS > 99 Then xPos = 3
            ' Font for items numbers
            Dim fontGS = New Font("Lucida Console", 24, FontStyle.Bold, GraphicsUnit.Pixel)

            ' Draw the value over the lower right of the gold skulltula picturebox, first in black to give it some definition, then in white
            Graphics.FromImage(.Image).DrawString(bGS.ToString, fontGS, New SolidBrush(Color.Black), xPos - 6, 28)
            Graphics.FromImage(.Image).DrawString(bGS.ToString, fontGS, New SolidBrush(Color.White), xPos - 5, 29)
        End With
    End Sub
    Private Sub getHearts()
        'If isSoH Then Exit Sub
        Dim enhFlagAddr As Integer = If(isSoH, soh.SAV(&H36 - 2), &H11A60C)
        Dim bHeartsAddr As Integer = If(isSoH, soh.SAV(&H28), &H11A5FC)
        Dim bPOHAddr As Integer = If(isSoH, soh.SAV(&HA8), &H11A674)

        ' Check if player has enhanced defence
        Dim isEnhanced As Boolean = CBool(IIf(goRead(enhFlagAddr + 2, 1) > 0, True, False))

        ' Get max heart container value, divide by 16 to undo their 16x multiplyer
        Dim bHearts As Byte = CByte(goRead(bHeartsAddr, 15) / 16)
        maxLife = bHearts

        ' Just a percaution, limit hearts to 99, even though 20 is the max. Never know what people may do to their save file 
        If bHearts > 99 Then bHearts = 99

        ' 29 is the starting position for single digit numbers
        Dim xPos As Byte = 29
        ' If double digits, set starting position to 14
        If bHearts > 9 Then xPos = 14

        ' Set up the font for the number of hearts
        Dim fontHearts = New Font("Lucida Console", 24, FontStyle.Bold, GraphicsUnit.Pixel)

        With pbxHeartContainer
            ' If not already visible, make it so. This is left here to let it stay greyed out before the first scan
            If Not .Visible Then .Visible = True

            If isEnhanced Then
                ' Set image to enhanced defence if active
                .Image = My.Resources.enhancedDefence
            Else
                ' Set image to normal heart container, still refreshing it to remove drawn numbers
                .Image = My.Resources.heartContainer
            End If

            ' Draw the value over the lower right of the heart container picturebox, first in black to give it some definition, then in white
            Graphics.FromImage(.Image).DrawString(bHearts.ToString, fontHearts, New SolidBrush(Color.Black), xPos - 1, 28)
            Graphics.FromImage(.Image).DrawString(bHearts.ToString, fontHearts, New SolidBrush(Color.White), xPos, 29)
            '.Invalidate()
        End With

        ' Now for the heart pieces
        If isSoH Then
            bHearts = CByte(goRead(bPOHAddr) >> 24)
        Else
            bHearts = CByte(goRead(bPOHAddr + 3, 1))
        End If

        With pbxPoH
            ' If not visible, make visible
            If Not .Visible Then .Visible = True
            Select Case bHearts
                Case Is >= 64
                    ' This is only a temporary chance check where all 4 pieces of heart are detected, but before the game combines them into a full piece
                    .Image = My.Resources.poh4
                Case Is >= 48
                    ' 3 pieces of heart
                    .Image = My.Resources.poh3
                Case Is >= 32
                    ' 2 pieces of heart
                    .Image = My.Resources.poh2
                Case Is >= 16
                    ' 1 piece of heart
                    .Image = My.Resources.poh1
                Case Else
                    ' No pieces of heart
                    .Image = My.Resources.poh0
            End Select
        End With
    End Sub

    Private Sub getMagic()
        ' Get magic check. This one is a curious one as there are many areas to check, and none perfect. Should someone hack their file to have magic, this should hopefully find it though
        Dim bMagic As Byte = CByte(goRead(If(isSoH, soh.SAV(&H2C), &H11A600 + 1), 1))

        With pbxMagicBar
            ' If not visible, make visible
            If Not .Visible Then .Visible = True
            Select Case bMagic
                Case 2
                    ' 2 for both magic upgrades
                    .Image = My.Resources.magicBar2
                    canMagic = True
                Case 1
                    ' 1 for one magic upgrade
                    .Image = My.Resources.magicBar1
                    canMagic = True
                Case Else
                    ' 0, or whatever else it finds, for no magic
                    .Image = My.Resources.magicBar0
                    canMagic = False
            End Select
        End With
    End Sub
    Private Sub getPedestalRead()
        pedestalRead = CByte(goRead(&H11B4FC, 1))
    End Sub
    Private Sub getSmallKeys()
        ' If(isSoH, soh.SAV(&HAC), &H11A678) (start of dugeonkeys array)

        ' Variables used to grab values and store the needed parts into an easy to use string
        Dim tempKeys As String = String.Empty
        Dim stringKeys As String = String.Empty

        '  Grab keys for just Forest Temple
        tempKeys = Hex(goRead(If(isSoH, soh.SAV(&HC3), &H11A68C), 1))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys, 2)

        ' Set string to the Forest Temple keys
        stringKeys = tempKeys

        ' Grab keys for Fire, Water, Spirit, and Shadow Temple
        tempKeys = Hex(goRead(If(isSoH, soh.SAV(&HC4), &H11A690)))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys)
        If isSoH Then endianFlip(tempKeys)

        ' Add all four of the grabbed Temple keys
        stringKeys = stringKeys & tempKeys

        ' Grab keys for Bottom of the Well and Gerudo Training Ground
        tempKeys = Hex(goRead(If(isSoH, soh.SAV(&HC8), &H11A694)))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys)
        If isSoH Then endianFlip(tempKeys)

        ' Add Bottom of the Well and Gerudo Training Ground keys
        stringKeys = stringKeys & Mid(tempKeys, 1, 2) & Mid(tempKeys, 7, 2)

        ' Grab keys for Ganon's Castle
        tempKeys = Hex(goRead(If(isSoH, soh.SAV(&HC9), &H11A698)))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys)
        If isSoH Then endianFlip(tempKeys)

        ' Add Ganon's Castle keys
        stringKeys = stringKeys & Mid(tempKeys, 3, 2)

        ' Set up the font for the number of keys
        Dim fontSmallKeys = New Font("Lucida Console", 20, FontStyle.Bold, GraphicsUnit.Pixel)
        ' Variable for temp storage for keys
        Dim valKeys As Byte = 0
        ' Position for x: 2 for double digits, 11 for single
        Dim xPos As Byte = 0

        ' Step through each stored small keys
        For i = 0 To 7
            With aoSmallKeys(i)
                ' Default position 20 since technically no key should reach double digits
                xPos = 20

                ' Move hex value in string to a value
                valKeys = CByte("&H" & Mid(stringKeys, (i * 2) + 1, 2))

                If valKeys = 255 Then
                    ' 255 (FF) is used as no keys found at all for the dungeon, make sure it is greyed out
                    If .Visible Then .Visible = False
                    valKeys = 0
                Else
                    ' For any real value, makue sure key is visible and reload the default key image to remove old numbers drawn on it
                    If Not .Visible Then .Visible = True
                    .Image = My.Resources.dungeonSmallKey()

                    ' The number of keys should never be over 9, much less 99, but just in case, keep it down to double digits
                    If valKeys > 99 Then valKeys = 99

                    ' Again, keys should never be double digit, but just in case, move the starting x position to draw 2 digits
                    If valKeys > 9 Then xPos = 7

                    ' Draw the value over the lower right of the key picturebox, first in black to give it some definition, then in white
                    Graphics.FromImage(.Image).DrawString(valKeys.ToString, fontSmallKeys, New SolidBrush(Color.Black), xPos - 1, 19)
                    Graphics.FromImage(.Image).DrawString(valKeys.ToString, fontSmallKeys, New SolidBrush(Color.White), xPos, 20)
                    '.Invalidate()
                End If
                aDungeonKeys(i) = valKeys
            End With
        Next
    End Sub
    Private Sub getTriforce()
        ' soh doesn't have Triforce Hunt as of 3.0.0
        If isSoH Then Exit Sub

        ' Get Triforce count
        Dim iTriforce As Byte = CByte(goRead(&H11AE94, 1))

        With pbxTriforce
            If iTriforce > 0 Then
                If Not .Visible Then .Visible = True
                ' Reset the image to the default triforce piece to remove drawn numbers
                .Image = My.Resources.triforce

                ' If single digit
                Dim xPos As Byte = 34
                ' If double digits
                If iTriforce > 9 Then xPos = 19
                ' If triple digits
                If iTriforce > 99 Then xPos = 3
                ' Font for triforce numbers
                Dim fontTriforce = New Font("Lucida Console", 24, FontStyle.Bold, GraphicsUnit.Pixel)

                ' Draw the value over the lower right of the triforce picturebox, first in black to give it some definition, then in white
                Graphics.FromImage(.Image).DrawString(iTriforce.ToString, fontTriforce, New SolidBrush(Color.Black), xPos - 6, 28)
                Graphics.FromImage(.Image).DrawString(iTriforce.ToString, fontTriforce, New SolidBrush(Color.White), xPos - 5, 29)
            Else
                If .Visible Then .Visible = False
            End If
        End With
    End Sub

    Private Sub cccCarpenters(chx As Object, e As EventArgs)
        checkCarpenters()
    End Sub
    Private Sub cccEquipment(chx As Object, e As EventArgs)
        checkEquipment()
    End Sub
    Private Sub cccUpgrades(chx As Object, e As EventArgs)
        checkUpgrades()
    End Sub

    Private Function checkBit(ByVal address As Integer, ByVal bit As Byte) As Boolean
        checkBit = False
        ' Reads word from memory
        Dim read As String = Hex(goRead(address))

        ' Fixes hex size
        While read.Length < 8
            read = "0" & read
        End While

        ' Grab only the hex digit we want
        read = Mid(read, CInt(8 - Math.Floor(bit / 4)), 1)

        Select Case (bit Mod 4)
            Case 0
                Select Case CInt("&H" & read)
                    Case 1, 3, 5, 7, 9, 11, 13, 150
                        Return True
                End Select
            Case 1
                Select Case CInt("&H" & read)
                    Case 2, 3, 6, 7, 10, 11, 14, 15
                        Return True
                End Select
            Case 2
                Select Case CInt("&H" & read)
                    Case 4 To 7, 12 To 15
                        Return True
                End Select
            Case 3
                Select Case CInt("&H" & read)
                    Case 8 To 15
                        Return True
                End Select
        End Select
    End Function

    Private Sub scanER()
        ' ER scanning
        Dim isOverworld As Boolean = False
        Dim isDungeon As Boolean = False
        Dim doTrials As Boolean = False

        ' Split the ER value into separate checks
        If iER Mod 2 = 1 Then isOverworld = True
        If iER > 1 Then isDungeon = True

        ' For building the arrays for scenes with Dungeon entrances
        Dim ent As Byte = 0

        Dim locationCode As Integer = CInt(IIf(isSoH, GDATA(&H200, 1), goRead(CUR_ROOM_ADDR + 2, 15)))
        Dim locationArray As Byte = 255
        Dim readExits(6) As Integer     ' Exits read for current map
        Dim aAlign(6) As Byte           ' Sets text alignment: 0 = Left | 1 = Centre | 2 = Right
        Dim lPoints As New List(Of Point)
        For i = 0 To aAlign.Length - 1
            aAlign(i) = 0
            readExits(i) = 0
        Next

        getAge()

        'Dim addrRoom As Integer = If(isSoH, &HD16D9C, &H1D8BEE)
        Dim addrRoom As Integer = If(isSoH, SAV(-&H1B17C4), &H1D8BEE)
        If locationCode <= 9 Then
            iRoom = CByte(goRead(addrRoom, 1))
        End If

        ' First load up the map, I keep this separated so that it does not matter for ER settings
        For i = 0 To 1
            Dim exitLoop = True
            Select Case locationCode
                Case 0
                    Select Case iRoom
                        Case 10, 12
                            pbxMap.Image = My.Resources.mapDT4  ' 3F
                        Case 1, 2, 11
                            pbxMap.Image = My.Resources.mapDT3  ' 2F
                        Case 0
                            pbxMap.Image = My.Resources.mapDT2  ' 1F
                        Case 3 To 8
                            pbxMap.Image = My.Resources.mapDT1  ' B1
                        Case 9
                            pbxMap.Image = My.Resources.mapDT0  ' B2
                    End Select
                Case 1
                    Select Case iRoom
                        Case 5, 6, 9, 10, 12, 16, 17, 18
                            pbxMap.Image = My.Resources.mapDDC1 ' 2F
                        Case 0 To 4, 7, 8, 11, 13 To 15
                            pbxMap.Image = My.Resources.mapDDC0 ' 1F
                    End Select
                Case 2
                    Select Case iRoom
                        Case 0 To 2, 4 To 12
                            pbxMap.Image = My.Resources.mapJB1  ' 1F
                        Case 3, 13 To 16
                            pbxMap.Image = My.Resources.mapJB0  ' B1
                    End Select
                Case 3
                    Select Case iRoom
                        Case 10, 12 To 14, 19, 20, 23 To 26
                            pbxMap.Image = My.Resources.mapFoT3 ' 2F
                        Case 0 To 8, 11, 15, 16, 18, 21, 22
                            pbxMap.Image = My.Resources.mapFoT2 ' 1F
                        Case 9
                            pbxMap.Image = My.Resources.mapFoT1 ' B1
                        Case 17
                            pbxMap.Image = My.Resources.mapFoT0 ' B2
                    End Select
                Case 4
                    Select Case iRoom
                        Case 8, 30, 34, 35
                            pbxMap.Image = My.Resources.mapFiT4 ' 5F
                        Case 7, 12 To 14, 27, 32, 33, 37
                            pbxMap.Image = My.Resources.mapFiT3 ' 4F
                        Case 5, 9, 11, 16, 23 To 26, 28, 31
                            pbxMap.Image = My.Resources.mapFiT2 ' 3F
                        Case 4, 10, 36
                            pbxMap.Image = My.Resources.mapFiT1 ' 2F
                        Case 0 To 3, 15, 17 To 22
                            pbxMap.Image = My.Resources.mapFiT0 ' 1F
                    End Select
                Case 5
                    Select Case iRoom
                        Case 0, 1, 4 To 7, 10, 11, 13, 17, 19, 20, 30, 31, 43
                            pbxMap.Image = My.Resources.mapWaT3 ' 3F
                        Case 22, 25, 29, 32, 35, 39, 41
                            pbxMap.Image = My.Resources.mapWaT2 ' 2F
                        Case 3, 8, 9, 12, 14 To 16, 18, 21, 23, 24, 26, 28, 33, 34, 36 To 38, 40, 42
                            pbxMap.Image = My.Resources.mapWaT1 ' 1F
                        Case 2, 27
                            pbxMap.Image = My.Resources.mapWaT0 ' B1
                    End Select
                Case 6
                    Select Case iRoom
                        Case 22, 24 To 26, 31
                            pbxMap.Image = My.Resources.mapSpT3  ' 4F
                        Case 7 To 11, 16 To 21, 23, 29
                            pbxMap.Image = My.Resources.mapSpT2  ' 3F
                        Case 5, 6, 28, 30
                            pbxMap.Image = My.Resources.mapSpT1  ' 2F
                        Case 0 To 4, 12 To 15, 27
                            pbxMap.Image = My.Resources.mapSpT0  ' 1F
                    End Select
                Case 7
                    Select Case iRoom
                        Case 0 To 2, 4
                            pbxMap.Image = My.Resources.mapShT3 ' B1
                        Case 5 To 8
                            pbxMap.Image = My.Resources.mapShT2 ' B2
                        Case 9, 16, 22
                            pbxMap.Image = My.Resources.mapShT1 ' B3
                        Case 3, 10 To 15, 17 To 21, 23 To 26
                            pbxMap.Image = My.Resources.mapShT0 ' B4
                    End Select
                Case 8
                    Select Case iRoom
                        Case 0 To 6
                            pbxMap.Image = My.Resources.mapBotW2 ' B1
                        Case 7, 8
                            pbxMap.Image = My.Resources.mapBotW1 ' B2
                        Case 9
                            pbxMap.Image = My.Resources.mapBotW0 ' B3
                    End Select
                Case 9
                    pbxMap.Image = My.Resources.mapIC0  ' F1
                Case 10
                    pbxMap.Image = My.Resources.mapGAC0
                    doTrials = True
                    iRoom = 0
                Case 11
                    pbxMap.Image = My.Resources.mapGTG    ' GTG
                Case 13
                    iRoom = getGanonMap()
                    Select Case iRoom
                        Case 1
                            pbxMap.Image = My.Resources.mapGAC1
                            doTrials = True
                        Case 2
                            pbxMap.Image = My.Resources.mapGAC2
                        Case 3
                            pbxMap.Image = My.Resources.mapGAC3
                        Case 4
                            pbxMap.Image = My.Resources.mapGAC4
                        Case 5
                            pbxMap.Image = My.Resources.mapGAC5
                        Case 6
                            pbxMap.Image = My.Resources.mapGAC6
                        Case 7
                            pbxMap.Image = My.Resources.mapGAC7
                        Case Else
                            pbxMap.Image = My.Resources.mapGAC0
                            doTrials = True
                    End Select
                Case 27 To 29
                    pbxMap.Image = My.Resources.mapMKE
                    'If locationCode = 29 Then
                    'pbxMap.Image = My.Resources.mapMKEntrance2
                    'Else
                    'pbxMap.Image = My.Resources.mapMKEntrance
                    'End If
                Case 30, 31
                    pbxMap.Image = My.Resources.mapMKBA
                Case 32, 33
                    pbxMap.Image = My.Resources.mapMK
                Case 34
                    pbxMap.Image = My.Resources.mapMK2
                Case 35 To 37
                    pbxMap.Image = My.Resources.mapOToT
                    'pbxMap.Image = My.Resources.mapToTOutside
                    'Case 37
                    'pbxMap.Image = My.Resources.mapToTOutside2
                Case 67
                    pbxMap.Image = My.Resources.mapToT
                Case 81
                    pbxMap.Image = My.Resources.mapHF
                Case 82
                    pbxMap.Image = My.Resources.mapKV
                Case 83
                    pbxMap.Image = My.Resources.mapGY
                Case 84
                    pbxMap.Image = My.Resources.mapZR
                Case 85
                    pbxMap.Image = My.Resources.mapKF
                Case 86
                    pbxMap.Image = My.Resources.mapSFM
                Case 87
                    pbxMap.Image = My.Resources.mapLH
                Case 88
                    pbxMap.Image = My.Resources.mapZD
                Case 89
                    pbxMap.Image = My.Resources.mapZF
                Case 90
                    pbxMap.Image = My.Resources.mapGV
                Case 91
                    pbxMap.Image = My.Resources.mapLW
                Case 92
                    pbxMap.Image = My.Resources.mapDC
                Case 93, 12
                    pbxMap.Image = My.Resources.mapGF
                    locationCode = 93
                Case 94
                    pbxMap.Image = My.Resources.mapHW2
                Case 95
                    pbxMap.Image = My.Resources.mapHC
                Case 96
                    pbxMap.Image = My.Resources.mapDMT
                Case 97
                    pbxMap.Image = My.Resources.mapDMC
                Case 98
                    pbxMap.Image = My.Resources.mapGC
                Case 99
                    pbxMap.Image = My.Resources.mapLLR
                Case 100
                    pbxMap.Image = My.Resources.mapOGC
                Case Else
                    locationCode = iLastMinimap
                    exitLoop = False
            End Select
            If exitLoop Then Exit For
        Next
        iLastMinimap = locationCode

        If doTrials Then
            updateTrials()
        Else
            pbxTrialFire.Visible = False
            pbxTrialForest.Visible = False
            pbxTrialWater.Visible = False
            pbxTrialSpirit.Visible = False
            pbxTrialShadow.Visible = False
            pbxTrialLight.Visible = False
        End If

        If iER = 0 Then Exit Sub

        ' Dungeon only checks
        If isDungeon Then
            Select Case locationCode
                Case 0
                    readExits(0) = &H377116     ' KF from Deku Tree
                    If aMQ(0) Then readExits(0) = &H3770B6 ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H377114     ' Queen Gohma from Deku Tree
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    locationArray = 26
                Case 1
                    readExits(0) = &H36FABE     ' DMT from Dodongo's Cavern
                    If aMQ(1) Then readExits(0) = &H36FA8E ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H36FABC     ' King Dodongo from Dodongo's Cavern
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    locationArray = 27
                Case 2
                    readExits(0) = &H36F43E     ' ZD from Jabu-Jabu's Belly
                    If aMQ(2) Then readExits(0) = &H36F40E ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H36F43C     ' Barinade from Jabu-Jabu's Belly
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    locationArray = 28
                Case 3
                    readExits(0) = &H36ED12     ' SFM from Forest Temple
                    If aMQ(3) Then readExits(0) = &H36ED12 ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H36ED10     ' Phantom Ganon from Forest Temple
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    locationArray = 29
                Case 4
                    readExits(0) = &H36A3C6     ' DMC from Fire Temple
                    If aMQ(4) Then readExits(0) = &H36A376 ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H36A3C4     ' Volvagia from Fire Temple
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    locationArray = 30
                Case 5
                    readExits(0) = &H36EF76     ' LH from Water Temple
                    If aMQ(5) Then readExits(0) = &H36EF36 ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H36EF74     ' Morpha from Temple
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    locationArray = 31
                Case 6
                    readExits(0) = &H36B1EA     ' DC from Spirit Temple
                    If aMQ(6) Then readExits(0) = &H36B12A ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H36B1E8     ' Twinrova from Spirit Temple
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    'readExits(2) = &H36B1EC     ' DC Statue Hand Right from Spirit Temple
                    'lPoints.Add(New Point(99, 332))
                    'aAlign(2) = 2
                    'readExits(3) = &H36B1EE     ' DC Statue Hand Left from Spirit Temple
                    'lPoints.Add(New Point(482, 357))
                    locationArray = 32
                Case 7
                    readExits(0) = &H36C86E     ' GY from Shadow Temple
                    If aMQ(7) Then readExits(0) = &H36C86E ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    'readExits(1) = &H36C86C     ' Bongo Bongo from Shadow Temple
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(1) = 1
                    locationArray = 33
                Case 8
                    readExits(0) = &H3785C6     ' KV from BotW
                    If aMQ(8) Then readExits(0) = &H378566 ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    locationArray = 34
                Case 9
                    readExits(0) = &H37352E     ' ZF from Ice Cavern
                    If aMQ(9) Then readExits(0) = &H37344E ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    locationArray = 35
                Case 10
                    'readExits(0) = &H374348     ' Ganon's Castle from Ganon's Tower
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(0) = 1
                    'readExits(1) = &H37434A     ' Ganon from Ganon's Tower
                    'lPoints.Add(New Point(278, 390))
                    'aAlign(1) = 1
                Case 11
                    readExits(0) = &H373686     ' GF from GTG
                    If aMQ(10) Then readExits(0) = &H373686 ' MQ version
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    locationArray = 36
                Case 13
                    'readExits(0) = &H3634BC     ' Ganon's Tower from Ganon's Castle
                    'lPoints.Add(New Point(278, 3))
                    'aAlign(0) = 1
                    readExits(0) = &H3634BE     ' OGC from Ganon's Castle
                    lPoints.Add(New Point(278, 390))
                    aAlign(0) = 1
                    locationArray = 37
            End Select
        End If

        Select Case locationCode
            Case 82
                lPoints.Add(New Point(369, 223))
                aAlign(0) = 1
                ent = 1
                If isDungeon Then readExits(0) = &H368A82 ' BotW from KV
                locationArray = 7
            Case 83
                lPoints.Add(New Point(492, 202))
                ent = 1
                If isDungeon Then readExits(0) = &H378E44 ' Shadow Temple from GY
                locationArray = 8
            Case 85
                lPoints.Add(New Point(440, 234))
                aAlign(0) = 1
                ent = 1
                If isDungeon Then readExits(0) = &H3738F4 ' Deku Tree from KF
                locationArray = 10
            Case 86
                lPoints.Add(New Point(278, 3))
                aAlign(0) = 1
                ent = 1
                If isDungeon Then readExits(0) = &H36FCD4 ' Forest Temple from SFM
                locationArray = 11
            Case 87
                lPoints.Add(New Point(311, 278))
                aAlign(0) = 1
                ent = 1
                If isDungeon Then readExits(0) = &H3696A6 ' Water Temple from LH
                locationArray = 12
            Case 89
                lPoints.Add(New Point(262, 152))
                lPoints.Add(New Point(346, 8))
                aAlign(1) = 1
                ent = 2
                If isDungeon Then
                    readExits(0) = &H3733CC     ' Jabu-Jabu's Belly from ZF
                    readExits(1) = &H3733D0     ' Ice Cavern from ZF
                End If
                locationArray = 14
            Case 92
                lPoints.Add(New Point(85, 163))
                aAlign(0) = 1
                ent = 1
                If isDungeon Then
                    readExits(0) = &H36B5AC     ' Spirit Temple from DC
                    'readExits(2) = &H36B5B0     ' Spirit Temple from Statue Hand Left
                    'lPoints.Add(New Point(84, 257))
                    'aAlign(2) = 1
                    'readExits(3) = &H36B5B2     ' Spirit Temple from Statue Hand Right
                    'lPoints.Add(New Point(84, 118))
                    'aAlign(3) = 1
                End If
                locationArray = 17
            Case 93, 12
                lPoints.Add(New Point(275, 242))
                aAlign(0) = 1
                ent = 1
                If isDungeon Then readExits(0) = &H374D02 ' GTG from GF
                locationArray = 18
            Case 96
                lPoints.Add(New Point(260, 174))
                aAlign(0) = 2
                ent = 1
                If isDungeon Then readExits(0) = &H365FF0 ' Dodongo's Cavern from DMT
                locationArray = 21
            Case 97
                lPoints.Add(New Point(308, 3))
                aAlign(0) = 1
                ent = 1
                If isDungeon Then readExits(0) = &H374B9E ' Fire Temple from DMC
                locationArray = 22
            Case 100
                lPoints.Add(New Point(278, 3))
                aAlign(1) = 1
                ent = 1
                If isDungeon Then readExits(0) = &H37FEC0 ' Ganon's Castle from OGC
                locationArray = 25
        End Select

        If iER Mod 2 = 1 Then
            Select Case locationCode
                Case 27 To 29
                    readExits(0) = &H384650     ' HF from MK Entrance Day
                    lPoints.Add(New Point(482, 73))
                    aAlign(0) = 1
                    readExits(1) = &H384652     ' MK from MK Entrance Day
                    lPoints.Add(New Point(67, 346))
                    aAlign(1) = 1
                    ' For night, adjust by -0x48
                    If locationCode = 28 Then
                        readExits(0) = readExits(0) - &H48
                        readExits(1) = readExits(1) - &H48
                    End If
                    locationArray = 0
                Case 30, 31
                    readExits(0) = &H383824     ' MK from Back Alley Left Day
                    lPoints.Add(New Point(266, 307))
                    readExits(1) = &H383826     ' MK from Back Alley Right Day
                    lPoints.Add(New Point(266, 70))
                    ' For night, adjust by -0x98
                    If locationCode = 31 Then
                        readExits(0) = readExits(0) - &H98
                        readExits(1) = readExits(1) - &H98
                    End If
                    locationArray = 1
                Case 32, 33
                    readExits(0) = &H3824A8     ' HC from MK Day
                    lPoints.Add(New Point(278, 3))
                    aAlign(0) = 1
                    readExits(1) = &H3824AA     ' MK Entrance from MK Day
                    lPoints.Add(New Point(278, 390))
                    aAlign(1) = 1
                    readExits(2) = &H3824AE     ' Outside ToT from MK Day
                    lPoints.Add(New Point(456, 81))
                    readExits(3) = &H3824AC     ' Back Alley Right from MK Day
                    lPoints.Add(New Point(98, 56))
                    aAlign(3) = 2
                    readExits(4) = &H3824B2     ' Back Alley Left from MK Day
                    lPoints.Add(New Point(98, 330))
                    aAlign(4) = 2
                    ' For night, adjust by +0x40
                    If locationCode = 33 Then
                        readExits(0) = readExits(0) + &H40
                        readExits(1) = readExits(1) + &H40
                        readExits(2) = readExits(2) + &H40
                        readExits(3) = readExits(3) + &H40
                        readExits(4) = readExits(4) + &H40
                    End If
                    locationArray = 2
                Case 34
                    readExits(0) = &H383478     ' OGC from MK Day
                    lPoints.Add(New Point(278, 3))
                    aAlign(0) = 1
                    readExits(1) = &H38347A     ' MK Entrance from MK
                    lPoints.Add(New Point(278, 390))
                    aAlign(1) = 1
                    readExits(2) = &H38347E     ' Outside ToT from MK
                    lPoints.Add(New Point(456, 70))
                    locationArray = 3
                Case 35, 36
                    readExits(0) = &H383524     ' ToT from Outside ToT Day
                    lPoints.Add(New Point(114, 18))
                    aAlign(0) = 1
                    readExits(1) = &H383526     ' MK from Outside ToT Day
                    lPoints.Add(New Point(442, 377))
                    aAlign(1) = 1
                    ' For night, adjust by -0x18
                    If locationCode = 36 Then
                        readExits(0) = readExits(0) - &H18
                        readExits(1) = readExits(1) - &H18
                    End If
                    locationArray = 4
                Case 37
                    readExits(0) = &H383574     ' ToT from Outside ToT
                    lPoints.Add(New Point(114, 18))
                    aAlign(0) = 1
                    readExits(1) = &H383576     ' MK from Outside ToT 
                    lPoints.Add(New Point(442, 377))
                    aAlign(1) = 1
                    locationArray = 4
                Case 67
                    readExits(0) = &H372332     ' Outside ToT from ToT
                    lPoints.Add(New Point(278, 385))
                    aAlign(0) = 1
                    locationArray = 5
                Case 81
                    readExits(0) = &H36BF9C     ' KV from HF
                    lPoints.Add(New Point(418, 64))
                    readExits(1) = &H36BFA0     ' LW Bridge from HF
                    lPoints.Add(New Point(437, 215))
                    readExits(2) = &H36BFA2     ' ZR from HF
                    lPoints.Add(New Point(447, 121))
                    readExits(3) = &H36BFA4     ' GV from HF
                    lPoints.Add(New Point(104, 183))
                    aAlign(3) = 2
                    readExits(4) = &H36BFA8     ' MK from HF
                    lPoints.Add(New Point(313, 14))
                    aAlign(4) = 1
                    readExits(5) = &H36BFAA     ' LLR from HF
                    lPoints.Add(New Point(280, 151))
                    aAlign(5) = 1
                    readExits(6) = &H36BFAE     ' LH from HF
                    lPoints.Add(New Point(195, 380))
                    aAlign(6) = 1
                    locationArray = 6
                Case 82
                    readExits(ent) = &H368A78     ' DMT from KV
                    lPoints.Add(New Point(298, 11))
                    aAlign(ent) = 1
                    incB(ent)
                    readExits(ent) = &H368A7A     ' HF from KV
                    lPoints.Add(New Point(57, 294))
                    aAlign(ent) = 2
                    incB(ent)
                    readExits(ent) = &H368A7E     ' GY from KV
                    lPoints.Add(New Point(493, 312))
                    locationArray = 7
                Case 83
                    readExits(ent) = &H378E46     ' KV from GY
                    lPoints.Add(New Point(57, 225))
                    aAlign(ent) = 2
                    locationArray = 8
                Case 84
                    readExits(0) = &H37951C     ' HF from ZR
                    lPoints.Add(New Point(46, 319))
                    aAlign(0) = 2
                    readExits(1) = &H37951E     ' ZD from ZR
                    lPoints.Add(New Point(509, 92))
                    readExits(2) = &H379522     ' LW from ZR
                    lPoints.Add(New Point(483, 157))
                    aAlign(2) = 1
                    locationArray = 9
                Case 85
                    readExits(ent) = &H3738FA     ' LW Bridge from KF
                    lPoints.Add(New Point(79, 187))
                    aAlign(ent) = 2
                    incB(ent)
                    readExits(ent) = &H373902     ' LW from KF
                    lPoints.Add(New Point(172, 84))
                    aAlign(ent) = 1
                    locationArray = 10
                Case 86
                    readExits(ent) = &H36FCD6     ' LW from SFM
                    lPoints.Add(New Point(278, 389))
                    aAlign(ent) = 1
                    locationArray = 11
                Case 87
                    readExits(ent) = &H3696A2     ' HF from LH
                    lPoints.Add(New Point(276, 5))
                    aAlign(ent) = 1
                    incB(ent)
                    readExits(ent) = &H3696AE     ' ZD from LH
                    lPoints.Add(New Point(313, 126))
                    aAlign(ent) = 1
                    locationArray = 12
                Case 88
                    readExits(0) = &H37B26C     ' ZF from ZD
                    lPoints.Add(New Point(291, 24))
                    aAlign(0) = 1
                    readExits(1) = &H37B26E     ' ZR from ZD
                    lPoints.Add(New Point(182, 291))
                    aAlign(1) = 1
                    readExits(2) = &H37B270     ' LH from ZD
                    lPoints.Add(New Point(352, 387))
                    aAlign(2) = 2
                    locationArray = 13
                Case 89
                    readExits(ent) = &H3733D2     ' ZD from ZF
                    lPoints.Add(New Point(189, 268))
                    aAlign(ent) = 2
                    locationArray = 14
                Case 90
                    readExits(0) = &H373904     ' LH from GV
                    lPoints.Add(New Point(328, 387))
                    aAlign(0) = 1
                    readExits(1) = &H373906     ' HF from GV
                    lPoints.Add(New Point(464, 218))
                    readExits(2) = &H373908     ' GF from GV
                    lPoints.Add(New Point(84, 117))
                    aAlign(2) = 2
                    locationArray = 15
                Case 91
                    readExits(0) = &H37475C     ' SFM from LW
                    lPoints.Add(New Point(299, 14))
                    aAlign(0) = 1
                    readExits(1) = &H37475E     ' KF from LW
                    lPoints.Add(New Point(248, 251))
                    aAlign(1) = 1
                    readExits(2) = &H374768     ' ZR from LW
                    lPoints.Add(New Point(423, 158))
                    readExits(3) = &H37476A     ' GC from LW
                    If My.Settings.setScrub Then
                        ' Have to move it down because the scrub icons crash right over it
                        lPoints.Add(New Point(299, 128))
                    Else
                        lPoints.Add(New Point(299, 117))
                    End If
                    aAlign(3) = 1
                    readExits(4) = &H37476C     ' KF from LW Bridge
                    lPoints.Add(New Point(192, 341))
                    readExits(5) = &H37476E     ' HF from LW Bridge
                    lPoints.Add(New Point(127, 321))
                    aAlign(5) = 2
                    locationArray = 16
                Case 92
                    readExits(ent) = &H36B5AE     ' HW from DC
                    lPoints.Add(New Point(476, 179))
                    locationArray = 17
                Case 93
                    readExits(ent) = &H374CE6     ' GV from GF
                    lPoints.Add(New Point(277, 350))
                    incB(ent)
                    readExits(ent) = &H374D00     ' HW from GF
                    lPoints.Add(New Point(145, 133))
                    aAlign(ent) = 2
                    locationArray = 18
                Case 94
                    readExits(0) = &H37EBFC     ' DC from HW
                    lPoints.Add(New Point(10, 99))
                    aAlign(0) = 2
                    readExits(1) = &H37EBFE     ' GF from HW
                    lPoints.Add(New Point(545, 336))
                    locationArray = 19
                Case 95
                    readExits(0) = &H36C55E     ' MK from HC
                    lPoints.Add(New Point(162, 365))
                    aAlign(0) = 1
                    'readExits(1) = &H36C562     ' GFF from HC
                    'lPoints.Add(New Point(346, 243))
                    'readExits(2) = &H36C55C     ' Castle Courtyard from HC
                    'lPoints.Add(New Point(89, 32))
                    locationArray = 20
                Case 96
                    readExits(ent) = &H365FEC     ' GC from DMT
                    lPoints.Add(New Point(320, 149))
                    incB(ent)
                    readExits(ent) = &H365FEE     ' KV from DMT
                    lPoints.Add(New Point(259, 384))
                    aAlign(ent) = 1
                    incB(ent)
                    readExits(ent) = &H365FF2     ' DMC from DMT
                    lPoints.Add(New Point(315, 3))
                    aAlign(ent) = 1
                    locationArray = 21
                Case 97
                    readExits(ent) = &H374B98     ' GC from DMC
                    lPoints.Add(New Point(129, 192))
                    aAlign(ent) = 2
                    incB(ent)
                    readExits(ent) = &H374B9A     ' DMT from DMC
                    lPoints.Add(New Point(199, 387))
                    aAlign(ent) = 1
                    locationArray = 22
                Case 98
                    readExits(0) = &H37A64C     ' DMC from GC
                    lPoints.Add(New Point(281, 3))
                    aAlign(0) = 1
                    readExits(1) = &H37A64E     ' DMT from GC
                    lPoints.Add(New Point(278, 390))
                    aAlign(1) = 1
                    readExits(2) = &H37A650     ' LW from GC
                    lPoints.Add(New Point(351, 361))
                    locationArray = 23
                Case 99
                    readExits(0) = &H377C12     ' HF from LLR
                    lPoints.Add(New Point(348, 14))
                    aAlign(0) = 1
                    locationArray = 24
                Case 100
                    readExits(ent) = &H37FEC2     ' MK from OGC
                    lPoints.Add(New Point(278, 390))
                    aAlign(ent) = 1
                    locationArray = 25
            End Select
        End If

        ' Sets that you have been to the current
        If locationArray = 255 Then Exit Sub
        aVisited(locationArray) = True

        Dim exitCode As String = String.Empty
        Dim fontGS = New Font("Lucida Console", 24, FontStyle.Bold, GraphicsUnit.Pixel)
        Dim ptX As Integer = 0
        Dim ptY As Integer = 0
        Dim iNewReach As Byte = 0
        Dim iVisited As Byte = 0
        Dim doDisplay As Boolean = False
        For i = 0 To readExits.Length - 1
            If Not readExits(i) = 0 Then

                iNewReach = 255
                exitCode = Hex(goRead(readExits(i), 15))
                fixHex(exitCode, 3)
                exitCode = exit2label(exitCode, iNewReach)


                ' If aReachExit is not 255, set exit to the new iNewReach
                If Not iNewReach = 255 Then
                    aExitMap(locationArray)(i) = iNewReach
                End If

                ' Only run the display part if the panel is actually visable
                If pnlER.Visible Then
                    ' Convert the iNewReach into the location for aVisited
                    iVisited = zone2map(iNewReach)
                    doDisplay = True
                    If iVisited = 255 Then
                        doDisplay = True
                    Else
                        If Not aVisited(iVisited) Then exitCode = "?"
                    End If
                    If doDisplay Then
                        ptX = lPoints(i).X
                        ptY = lPoints(i).Y

                        ' Make sure the text will not run off the display area
                        Select Case aAlign(i)
                            Case 0
                                If ptX + ((exitCode.Length) * 15) + 4 > 548 Then
                                    ptX = 548
                                    aAlign(i) = 2
                                End If
                            Case 1
                                Dim limitSize As Double = 0
                                If ptX > (pbxMap.Width / 2) Then
                                    limitSize = ptX + (((exitCode.Length) * 15) / 2) + 4
                                    If limitSize > 548 Then
                                        ptX = 548
                                        aAlign(i) = 2
                                    End If
                                Else
                                    limitSize = ptX - (((exitCode.Length) * 15) / 2) + 4
                                    If limitSize < 0 Then
                                        ptX = 0
                                        aAlign(i) = 0
                                    End If
                                End If
                            Case 2
                                If ptX - ((exitCode.Length) * 15) - 4 < 0 Then
                                    ptX = 0
                                    aAlign(i) = 0
                                End If
                        End Select

                        ' Adjust the starting position for the exit's alignment
                        Select Case aAlign(i)
                            Case 1
                                ptX = CInt(ptX - ((exitCode.Length * 15) / 2) - 4)
                            Case 2
                                ptX = ptX - (exitCode.Length * 15) - 4
                        End Select

                        Graphics.FromImage(pbxMap.Image).DrawString(exitCode, fontGS, New SolidBrush(Color.Black), ptX + 1, ptY + 1)
                        Graphics.FromImage(pbxMap.Image).DrawString(exitCode, fontGS, New SolidBrush(Color.White), ptX, ptY)
                    End If
                End If
            End If
        Next
    End Sub
    Private Sub updateMiniMap()
        If Not pnlER.Visible Then Exit Sub

        Dim aIconPos As New List(Of Point)
        For i = 0 To aIconLoc.Length - 1
            aIconLoc(i) = String.Empty
            aIconName(i) = String.Empty
        Next
        lRegions.Clear()

        Select Case iLastMinimap
            Case 0
                If Not aMQ(0) Then
                    Select Case iRoom
                        Case 10, 12 ' 3F
                            aIconLoc(0) = "3102"
                            aIconPos.Add(New Point(213, 178))
                            aIconLoc(1) = "3106"
                            aIconPos.Add(New Point(273, 208))
                            aIconLoc(2) = "7803"
                            aIconPos.Add(New Point(273, 224))
                        Case 1, 2, 11 '2F
                            aIconLoc(0) = "3101"
                            aIconPos.Add(New Point(254, 350))
                            aIconLoc(1) = "3105"
                            aIconPos.Add(New Point(270, 362))
                        Case 0 '1F
                            aIconLoc(0) = "3103"
                            aIconPos.Add(New Point(466, 214))
                        Case 3 To 8 'B1
                            aIconLoc(0) = "3104"
                            aIconPos.Add(New Point(419, 131))
                            aIconLoc(1) = "7802"
                            aIconPos.Add(New Point(429, 172))
                            aIconLoc(2) = "7801"
                            aIconPos.Add(New Point(379, 120))
                            aIconLoc(3) = "7800"
                            aIconPos.Add(New Point(64, 84))
                        Case 9 'B2
                            aIconLoc(0) = "1031"
                            aIconPos.Add(New Point(300, 33))
                    End Select
                Else
                    Select Case iRoom
                        Case 10, 12 ' 3F
                            aIconLoc(0) = "3102"
                            aIconPos.Add(New Point(273, 215))
                            aIconLoc(1) = "3106"
                            aIconPos.Add(New Point(213, 178))
                        Case 1, 2, 11 '2F
                            aIconLoc(0) = "3101"
                            aIconPos.Add(New Point(254, 350))
                            aIconLoc(1) = "7803"
                            aIconPos.Add(New Point(270, 362))
                        Case 0 '1F
                            aIconLoc(0) = "3103"
                            aIconPos.Add(New Point(466, 214))
                            aIconLoc(1) = "7801"
                            aIconPos.Add(New Point(458, 230))
                        Case 3 To 8 'B1
                            aIconLoc(0) = "3104"
                            aIconPos.Add(New Point(419, 131))
                            aIconLoc(1) = "3105"
                            aIconPos.Add(New Point(362, 337))
                            aIconLoc(2) = "3100"
                            aIconPos.Add(New Point(246, 323))
                            aIconLoc(3) = "7802"
                            aIconPos.Add(New Point(156, 185))
                            aIconLoc(4) = "7800"
                            aIconPos.Add(New Point(63, 83))
                            aIconLoc(5) = "8405"
                            aIconPos.Add(New Point(333, 208))
                        Case 9 'B2
                            aIconLoc(0) = "1031"
                            aIconPos.Add(New Point(300, 33))
                    End Select
                End If
            Case 1
                If Not aMQ(1) Then
                    Select Case iRoom
                        Case 5, 6, 9, 10, 12, 16, 17, 18 ' 2F
                            aIconLoc(0) = "3206"
                            aIconPos.Add(New Point(307, 314))
                            aIconLoc(1) = "3204"
                            aIconPos.Add(New Point(289, 242))
                            aIconLoc(2) = "3210"
                            aIconPos.Add(New Point(156, 223))
                            aIconLoc(3) = "7810"
                            aIconPos.Add(New Point(109, 205))
                            aIconLoc(4) = "7808"
                            aIconPos.Add(New Point(109, 253))
                            aIconLoc(5) = "8501"
                            aIconPos.Add(New Point(293, 200))
                            aIconLoc(6) = "8504"
                            aIconPos.Add(New Point(313, 200))
                        Case 0 To 4, 7, 8, 11, 13 To 15 ' 1F
                            aIconLoc(0) = "3208"
                            aIconPos.Add(New Point(158, 239))
                            aIconLoc(1) = "3205"
                            aIconPos.Add(New Point(109, 319))
                            aIconLoc(2) = "4400"
                            aIconPos.Add(New Point(179, 140))
                            aIconLoc(3) = "1131"
                            aIconPos.Add(New Point(179, 156))
                            aIconLoc(4) = "7812"
                            aIconPos.Add(New Point(344, 361))
                            aIconLoc(5) = "7809"
                            aIconPos.Add(New Point(334, 291))
                            aIconLoc(6) = "7811"
                            aIconPos.Add(New Point(289, 39))
                            aIconLoc(7) = "8505"
                            aIconPos.Add(New Point(158, 304))
                            aIconLoc(8) = "8502"
                            aIconPos.Add(New Point(339, 129))
                    End Select
                Else
                    Select Case iRoom
                        Case 5, 6, 9, 10, 12, 16, 17, 18 ' 2F
                            aIconLoc(0) = "3203"
                            aIconPos.Add(New Point(293, 279))
                            aIconLoc(1) = "3202"
                            aIconPos.Add(New Point(303, 207))
                            aIconLoc(2) = "3205"
                            aIconPos.Add(New Point(107, 323))
                            aIconLoc(3) = "7812"
                            aIconPos.Add(New Point(319, 202))
                            aIconLoc(4) = "7810"
                            aIconPos.Add(New Point(390, 202))
                            aIconLoc(5) = "8505"
                            aIconPos.Add(New Point(104, 205))
                        Case 0 To 4, 7, 8, 11, 13 To 15 ' 1F
                            aIconLoc(0) = "3200"
                            aIconPos.Add(New Point(261, 224))
                            aIconLoc(1) = "3204"
                            aIconPos.Add(New Point(269, 259))
                            aIconLoc(2) = "3201"
                            aIconPos.Add(New Point(287, 45))
                            aIconLoc(3) = "4400"
                            aIconPos.Add(New Point(179, 140))
                            aIconLoc(4) = "1131"
                            aIconPos.Add(New Point(179, 156))
                            aIconLoc(5) = "7809"
                            aIconPos.Add(New Point(330, 129))
                            aIconLoc(6) = "7811"
                            aIconPos.Add(New Point(117, 319))
                            aIconLoc(7) = "7808"
                            aIconPos.Add(New Point(294, 105))
                            aIconLoc(8) = "8504"
                            aIconPos.Add(New Point(158, 296))
                            aIconLoc(9) = "8502"
                            aIconPos.Add(New Point(158, 312))
                            aIconLoc(10) = "8508"
                            aIconPos.Add(New Point(345, 356))
                    End Select
                End If
            Case 2
                If Not aMQ(2) Then
                    Select Case iRoom
                        Case 0 To 2, 4 To 12 ' 1F
                            aIconLoc(0) = "3301"
                            aIconPos.Add(New Point(350, 107))
                            aIconLoc(1) = "3302"
                            aIconPos.Add(New Point(184, 107))
                            aIconLoc(2) = "3304"
                            aIconPos.Add(New Point(229, 53))
                            aIconLoc(3) = "1231"
                            aIconPos.Add(New Point(341, 232))
                            aIconLoc(4) = "7818"
                            aIconPos.Add(New Point(347, 270))
                        Case 3, 13 To 16 ' B1
                            aIconLoc(0) = "7819"
                            aIconPos.Add(New Point(324, 224))
                            aIconLoc(1) = "7817"
                            aIconPos.Add(New Point(254, 163))
                            aIconLoc(2) = "7816"
                            aIconPos.Add(New Point(270, 159))
                            aIconLoc(3) = "8601"
                            aIconPos.Add(New Point(221, 264))
                    End Select
                Else
                    Select Case iRoom
                        Case 0 To 2, 4 To 12 ' 1F
                            aIconLoc(0) = "3303"
                            aIconPos.Add(New Point(269, 339))
                            aIconLoc(1) = "3305"
                            aIconPos.Add(New Point(285, 331))
                            aIconLoc(2) = "3309"
                            aIconPos.Add(New Point(311, 43))
                            aIconLoc(3) = "3307"
                            aIconPos.Add(New Point(293, 265))
                            aIconLoc(4) = "3310"
                            aIconPos.Add(New Point(332, 248))
                            aIconLoc(5) = "1231"
                            aIconPos.Add(New Point(341, 232))
                            aIconLoc(6) = "7818"
                            aIconPos.Add(New Point(240, 50))
                            aIconLoc(7) = "7817"
                            aIconPos.Add(New Point(337, 264))
                            aIconLoc(8) = "10224"
                            aIconPos.Add(New Point(198, 265))
                        Case 3, 13 To 16 ' B1
                            aIconLoc(0) = "3302"
                            aIconPos.Add(New Point(285, 258))
                            aIconLoc(1) = "3300"
                            aIconPos.Add(New Point(224, 263))
                            aIconLoc(2) = "3304"
                            aIconPos.Add(New Point(258, 202))
                            aIconLoc(3) = "3308"
                            aIconPos.Add(New Point(286, 163))
                            aIconLoc(4) = "3306"
                            aIconPos.Add(New Point(317, 224))
                            aIconLoc(5) = "3301"
                            aIconPos.Add(New Point(338, 211))
                            aIconLoc(6) = "7816"
                            aIconPos.Add(New Point(333, 227))
                            aIconLoc(7) = "7819"
                            aIconPos.Add(New Point(211, 233))
                    End Select
                End If
            Case 3
                If Not aMQ(3) Then
                    Select Case iRoom
                        Case 10, 12 To 14, 19, 20, 23 To 26 ' 2F
                            aIconLoc(0) = "3403"
                            aIconPos.Add(New Point(258, 335))
                            aIconLoc(1) = "3401"
                            aIconPos.Add(New Point(277, 100))
                            aIconLoc(2) = "3404"
                            aIconPos.Add(New Point(145, 198))
                            aIconLoc(3) = "3414"
                            aIconPos.Add(New Point(149, 51))
                            aIconLoc(4) = "3413"
                            aIconPos.Add(New Point(235, 44))
                            aIconLoc(5) = "3412"
                            aIconPos.Add(New Point(277, 56))
                            aIconLoc(6) = "3415"
                            aIconPos.Add(New Point(315, 44))
                            aIconLoc(7) = "7825"
                            aIconPos.Add(New Point(296, 335))
                            aIconLoc(8) = "7826"
                            aIconPos.Add(New Point(186, 87))

                        Case 0 To 8, 11, 15, 16, 18, 21, 22 ' 1F
                            aIconLoc(0) = "3403"
                            aIconPos.Add(New Point(258, 335))
                            aIconLoc(1) = "3400"
                            aIconPos.Add(New Point(277, 59))
                            aIconLoc(2) = "3405"
                            aIconPos.Add(New Point(384, 110))
                            aIconLoc(3) = "3402"
                            aIconPos.Add(New Point(114, 153))
                            aIconLoc(4) = "3407"
                            aIconPos.Add(New Point(419, 133))
                            aIconLoc(5) = "7825"
                            aIconPos.Add(New Point(296, 335))
                            aIconLoc(6) = "7827"
                            aIconPos.Add(New Point(295, 125))
                            aIconLoc(7) = "7824"
                            aIconPos.Add(New Point(359, 100))
                        Case 9 ' B1
                            aIconLoc(0) = "3409"
                            aIconPos.Add(New Point(195, 169))
                        Case 17 ' B2
                            aIconLoc(0) = "3411"
                            aIconPos.Add(New Point(243, 219))
                            aIconLoc(1) = "1331"
                            aIconPos.Add(New Point(279, 104))
                            aIconLoc(2) = "7828"
                            aIconPos.Add(New Point(243, 203))
                    End Select
                Else
                    Select Case iRoom
                        Case 10, 12 To 14, 19, 20, 23 To 26 ' 2F
                            aIconLoc(0) = "3403"
                            aIconPos.Add(New Point(296, 322))
                            aIconLoc(1) = "3405"
                            aIconPos.Add(New Point(348, 76))
                            aIconLoc(2) = "3414"
                            aIconPos.Add(New Point(149, 51))
                            aIconLoc(3) = "3413"
                            aIconPos.Add(New Point(235, 44))
                            aIconLoc(4) = "3412"
                            aIconPos.Add(New Point(277, 44))
                            aIconLoc(5) = "3415"
                            aIconPos.Add(New Point(315, 44))
                        Case 0 To 8, 11, 15, 16, 18, 21, 22 ' 1F
                            aIconLoc(0) = "3403"
                            aIconPos.Add(New Point(296, 322))
                            aIconLoc(1) = "3400"
                            aIconPos.Add(New Point(277, 48))
                            aIconLoc(2) = "3401"
                            aIconPos.Add(New Point(361, 98))
                            aIconLoc(3) = "3406"
                            aIconPos.Add(New Point(426, 116))
                            aIconLoc(4) = "3402"
                            aIconPos.Add(New Point(113, 153))
                            aIconLoc(5) = "7825"
                            aIconPos.Add(New Point(271, 261))
                            aIconLoc(6) = "7824"
                            aIconPos.Add(New Point(342, 129))
                            aIconLoc(7) = "7826"
                            aIconPos.Add(New Point(185, 170))
                            aIconLoc(8) = "7828"
                            aIconPos.Add(New Point(160, 223))
                        Case 9 ' B1
                            aIconLoc(0) = "3409"
                            aIconPos.Add(New Point(343, 171))
                            aIconLoc(1) = "7827"
                            aIconPos.Add(New Point(191, 163))
                        Case 17 ' B2
                            aIconLoc(0) = "3411"
                            aIconPos.Add(New Point(240, 148))
                            aIconLoc(1) = "1331"
                            aIconPos.Add(New Point(279, 96))
                    End Select
                End If
            Case 4
                If Not aMQ(4) Then
                    Select Case iRoom
                        Case 8, 30, 34, 35 ' 5F
                            aIconLoc(0) = "3513"
                            aIconPos.Add(New Point(419, 163))
                            aIconLoc(1) = "3505"
                            aIconPos.Add(New Point(123, 221))
                            aIconLoc(2) = "7903"
                            aIconPos.Add(New Point(395, 185))
                        Case 7, 12 To 14, 27, 32, 33, 37 ' 4F
                            aIconLoc(0) = "7904"
                            aIconPos.Add(New Point(370, 117))
                        Case 5, 9, 11, 16, 23 To 26, 28, 31 ' 3F
                            aIconLoc(0) = "3503"
                            aIconPos.Add(New Point(432, 271))
                            aIconLoc(1) = "3508"
                            aIconPos.Add(New Point(374, 57))
                            aIconLoc(2) = "3510"
                            aIconPos.Add(New Point(360, 174))
                            aIconLoc(3) = "3506"
                            aIconPos.Add(New Point(418, 294))
                            aIconLoc(4) = "3507"
                            aIconPos.Add(New Point(236, 96))
                            aIconLoc(5) = "3509"
                            aIconPos.Add(New Point(166, 191))
                            aIconLoc(6) = "7902"
                            aIconPos.Add(New Point(440, 105))
                        Case 4, 10, 36 ' 2F
                            aIconLoc(0) = "3511"
                            aIconPos.Add(New Point(458, 188))
                        Case 0 To 3, 15, 17 To 22 ' 1F
                            aIconLoc(0) = "3501"
                            aIconPos.Add(New Point(193, 263))
                            aIconLoc(1) = "3500"
                            aIconPos.Add(New Point(262, 87))
                            aIconLoc(2) = "3512"
                            aIconPos.Add(New Point(262, 151))
                            aIconLoc(3) = "3504"
                            aIconPos.Add(New Point(430, 69))
                            aIconLoc(4) = "3502"
                            aIconPos.Add(New Point(393, 343))
                            aIconLoc(5) = "1431"
                            aIconPos.Add(New Point(166, 191))
                            aIconLoc(6) = "7901"
                            aIconPos.Add(New Point(330, 68))
                            aIconLoc(7) = "7900"
                            aIconPos.Add(New Point(362, 66))
                    End Select
                Else
                    Select Case iRoom
                        Case 8, 30, 34, 35 ' 5F
                            aIconLoc(0) = "3505"
                            aIconPos.Add(New Point(126, 226))
                            aIconLoc(1) = "7902"
                            aIconPos.Add(New Point(421, 162))
                        Case 7, 12 To 14, 27, 32, 33, 37 ' 4F
                            aIconLoc(0) = "7901"
                            aIconPos.Add(New Point(196, 179))
                        Case 5, 9, 11, 16, 23 To 26, 28, 31 ' 3F
                            aIconLoc(0) = "3503"
                            aIconPos.Add(New Point(429, 274))
                            aIconLoc(1) = "3508"
                            aIconPos.Add(New Point(374, 57))
                            aIconLoc(2) = "3506"
                            aIconPos.Add(New Point(418, 295))
                            aIconLoc(3) = "328"
                            aIconPos.Add(New Point(72, 201))
                            aIconLoc(4) = "7903"
                            aIconPos.Add(New Point(166, 190))
                            aIconLoc(5) = "7904"
                            aIconPos.Add(New Point(237, 95))
                        Case 4, 10, 36 ' 2F
                            aIconLoc(0) = "3511"
                            aIconPos.Add(New Point(458, 188))
                        Case 0 To 3, 15, 17 To 22 ' 1F
                            aIconLoc(0) = "3502"
                            aIconPos.Add(New Point(262, 169))
                            aIconLoc(1) = "3500"
                            aIconPos.Add(New Point(262, 89))
                            aIconLoc(2) = "3512"
                            aIconPos.Add(New Point(262, 153))
                            aIconLoc(3) = "3507"
                            aIconPos.Add(New Point(193, 263))
                            aIconLoc(4) = "3501"
                            aIconPos.Add(New Point(360, 61))
                            aIconLoc(5) = "3504"
                            aIconPos.Add(New Point(393, 344))
                            aIconLoc(6) = "1431"
                            aIconPos.Add(New Point(165, 191))
                            aIconLoc(7) = "7900"
                            aIconPos.Add(New Point(432, 69))
                    End Select
                End If
            Case 5
                If Not aMQ(5) Then
                    Select Case iRoom
                        Case 0, 1, 4 To 7, 10, 11, 13, 17, 19, 20, 30, 31, 43 ' 3F
                            aIconLoc(0) = "3602"
                            aIconPos.Add(New Point(431, 311))
                            aIconLoc(1) = "3609"
                            aIconPos.Add(New Point(416, 246))
                            aIconLoc(2) = "3608"
                            aIconPos.Add(New Point(375, 356))
                            aIconLoc(3) = "3607"
                            aIconPos.Add(New Point(154, 57))
                            aIconLoc(4) = "1531"
                            aIconPos.Add(New Point(323, 171))
                            aIconLoc(5) = "7910"
                            aIconPos.Add(New Point(316, 279))
                            aIconLoc(6) = "7909"
                            aIconPos.Add(New Point(221, 274))
                        Case 22, 25, 29, 32, 35, 39, 41 ' 2F
                            aIconLoc(0) = "3600"
                            aIconPos.Add(New Point(415, 311))
                            aIconLoc(1) = "3603"
                            aIconPos.Add(New Point(230, 139))
                            aIconLoc(2) = "7909"
                            aIconPos.Add(New Point(221, 274))
                            aIconLoc(3) = "7912"
                            aIconPos.Add(New Point(173, 105))
                        Case 3, 8, 9, 12, 14 To 16, 18, 21, 23, 24, 26, 28, 33, 34, 36 To 38, 40, 42 ' 1F
                            aIconLoc(0) = "3601"
                            aIconPos.Add(New Point(431, 311))
                            aIconLoc(1) = "3603"
                            aIconPos.Add(New Point(230, 139))
                            aIconLoc(2) = "3610"
                            aIconPos.Add(New Point(206, 181))
                            aIconLoc(3) = "3605"
                            aIconPos.Add(New Point(253, 120))
                            aIconLoc(4) = "7908"
                            aIconPos.Add(New Point(204, 346))
                            aIconLoc(5) = "7909"
                            aIconPos.Add(New Point(221, 274))
                            aIconLoc(6) = "7912"
                            aIconPos.Add(New Point(173, 105))
                            aIconLoc(7) = "7911"
                            aIconPos.Add(New Point(266, 158))
                        Case 2, 27 ' B1
                            aIconLoc(0) = "3606"
                            aIconPos.Add(New Point(369, 333))
                            aIconLoc(1) = "3610"
                            aIconPos.Add(New Point(206, 181))
                            aIconLoc(2) = "7908"
                            aIconPos.Add(New Point(204, 346))
                            aIconLoc(3) = "7909"
                            aIconPos.Add(New Point(221, 274))
                            aIconLoc(4) = "7911"
                            aIconPos.Add(New Point(256, 158))
                    End Select
                Else
                    Select Case iRoom
                        Case 0, 1, 4 To 7, 10, 11, 13, 17, 19, 20, 30, 31, 43 ' 3F
                            aIconLoc(0) = "3602"
                            aIconPos.Add(New Point(431, 311))
                            aIconLoc(1) = "1531"
                            aIconPos.Add(New Point(323, 171))
                            aIconLoc(2) = "7910"
                            aIconPos.Add(New Point(269, 304))
                        Case 22, 25, 29, 32, 35, 39, 41 ' 2F
                            aIconLoc(0) = "3600"
                            aIconPos.Add(New Point(418, 311))
                            aIconLoc(1) = "7908"
                            aIconPos.Add(New Point(382, 333))
                            aIconLoc(2) = "7909"
                            aIconPos.Add(New Point(156, 147))
                        Case 3, 8, 9, 12, 14 To 16, 18, 21, 23, 24, 26, 28, 33, 34, 36 To 38, 40, 42 ' 1F
                            aIconLoc(0) = "3601"
                            aIconPos.Add(New Point(438, 311))
                            aIconLoc(1) = "3605"
                            aIconPos.Add(New Point(262, 238))
                            aIconLoc(2) = "401"
                            aIconPos.Add(New Point(247, 126))
                            aIconLoc(3) = "7909"
                            aIconPos.Add(New Point(156, 147))
                            aIconLoc(4) = "7911"
                            aIconPos.Add(New Point(321, 115))
                            aIconLoc(5) = "7912"
                            aIconPos.Add(New Point(202, 346))
                        Case 2, 27 ' B1
                            aIconLoc(0) = "3606"
                            aIconPos.Add(New Point(370, 333))
                    End Select
                End If
            Case 6
                If Not aMQ(6) Then
                    Select Case iRoom
                        Case 22, 24 To 26, 31 ' 4F
                            aIconLoc(0) = "3710"
                            aIconPos.Add(New Point(322, 75))
                            aIconLoc(1) = "3718"
                            aIconPos.Add(New Point(197, 131))
                        Case 7 To 11, 16 To 21, 23, 29 ' 3F
                            aIconLoc(0) = "3701"
                            aIconPos.Add(New Point(136, 244))
                            aIconLoc(1) = "3511"
                            aIconPos.Add(New Point(162, 349))
                            aIconLoc(2) = "3715"
                            aIconPos.Add(New Point(315, 100))
                            aIconLoc(3) = "3705"
                            aIconPos.Add(New Point(428, 151))
                            aIconLoc(4) = "3721"
                            aIconPos.Add(New Point(441, 296))
                            aIconLoc(5) = "3720"
                            aIconPos.Add(New Point(424, 296))
                            aIconLoc(6) = "5509"
                            aIconPos.Add(New Point(379, 349))
                            aIconLoc(7) = "1631"
                            aIconPos.Add(New Point(261, 64))
                            aIconLoc(8) = "7916"
                            aIconPos.Add(New Point(108, 306))
                            aIconLoc(9) = "7918"
                            aIconPos.Add(New Point(213, 83))
                        Case 5, 6, 28, 30 ' 2F
                            aIconLoc(0) = "3706"
                            aIconPos.Add(New Point(176, 108))
                            aIconLoc(1) = "3712"
                            aIconPos.Add(New Point(184, 124))
                            aIconLoc(2) = "3703"
                            aIconPos.Add(New Point(261, 119))
                            aIconLoc(3) = "3713"
                            aIconPos.Add(New Point(370, 136))
                            aIconLoc(4) = "3714"
                            aIconPos.Add(New Point(370, 152))
                            aIconLoc(5) = "3702"
                            aIconPos.Add(New Point(234, 110))
                            aIconLoc(6) = "7919"
                            aIconPos.Add(New Point(160, 108))
                        Case 0 To 4, 12 To 15, 27 ' 1F
                            aIconLoc(0) = "3708"
                            aIconPos.Add(New Point(102, 97))
                            aIconLoc(1) = "3700"
                            aIconPos.Add(New Point(221, 104))
                            aIconLoc(2) = "3704"
                            aIconPos.Add(New Point(297, 97))
                            aIconLoc(3) = "3707"
                            aIconPos.Add(New Point(405, 55))
                            aIconLoc(4) = "7920"
                            aIconPos.Add(New Point(221, 124))
                            aIconLoc(5) = "7917"
                            aIconPos.Add(New Point(393, 118))
                    End Select
                Else
                    Select Case iRoom
                        Case 22, 24 To 26, 31 ' 4F
                            aIconLoc(0) = "3718"
                            aIconPos.Add(New Point(261, 169))
                            aIconLoc(1) = "7918"
                            aIconPos.Add(New Point(304, 91))
                            aIconLoc(2) = "7920"
                            aIconPos.Add(New Point(321, 73))
                        Case 7 To 11, 16 To 21, 23, 29 ' 3F
                            aIconLoc(0) = "3701"
                            aIconPos.Add(New Point(156, 184))
                            aIconLoc(1) = "5511"
                            aIconPos.Add(New Point(161, 349))
                            aIconLoc(2) = "3702"
                            aIconPos.Add(New Point(314, 83))
                            aIconLoc(3) = "3725"
                            aIconPos.Add(New Point(392, 203))
                            aIconLoc(4) = "3724"
                            aIconPos.Add(New Point(437, 219))
                            aIconLoc(5) = "3705"
                            aIconPos.Add(New Point(428, 151))
                            aIconLoc(6) = "5509"
                            aIconPos.Add(New Point(380, 349))
                            aIconLoc(7) = "1631"
                            aIconPos.Add(New Point(261, 65))
                            aIconLoc(8) = "7916"
                            aIconPos.Add(New Point(148, 244))
                        Case 5, 6, 28, 30 ' 2F
                            aIconLoc(0) = "3706"
                            aIconPos.Add(New Point(180, 124))
                            aIconLoc(1) = "3712"
                            aIconPos.Add(New Point(162, 157))
                            aIconLoc(2) = "3728"
                            aIconPos.Add(New Point(262, 178))
                            aIconLoc(3) = "3703"
                            aIconPos.Add(New Point(262, 117))
                            aIconLoc(4) = "3715"
                            aIconPos.Add(New Point(300, 83))
                        Case 0 To 4, 12 To 15, 27 ' 1F
                            aIconLoc(0) = "3730"
                            aIconPos.Add(New Point(251, 204))
                            aIconLoc(1) = "3731"
                            aIconPos.Add(New Point(267, 204))
                            aIconLoc(2) = "3726"
                            aIconPos.Add(New Point(251, 220))
                            aIconLoc(3) = "3727"
                            aIconPos.Add(New Point(267, 220))
                            aIconLoc(4) = "3708"
                            aIconPos.Add(New Point(113, 114))
                            aIconLoc(5) = "3700"
                            aIconPos.Add(New Point(113, 151))
                            aIconLoc(6) = "3729"
                            aIconPos.Add(New Point(164, 209))
                            aIconLoc(7) = "3704"
                            aIconPos.Add(New Point(296, 96))
                            aIconLoc(8) = "3707"
                            aIconPos.Add(New Point(405, 42))
                            aIconLoc(9) = "7917"
                            aIconPos.Add(New Point(307, 133))
                            aIconLoc(10) = "7919"
                            aIconPos.Add(New Point(414, 64))
                    End Select
                End If
            Case 7
                If Not aMQ(7) Then
                    Select Case iRoom
                        Case 0 To 2, 4 ' B1
                            aIconLoc(0) = "3801"
                            aIconPos.Add(New Point(235, 115))
                            aIconLoc(1) = "3807"
                            aIconPos.Add(New Point(175, 144))
                        Case 5 To 8 ' B2
                            aIconLoc(0) = "3803"
                            aIconPos.Add(New Point(382, 200))
                            aIconLoc(1) = "3802"
                            aIconPos.Add(New Point(399, 122))
                        Case 9, 16, 22 ' B3
                            aIconLoc(0) = "3812"
                            aIconPos.Add(New Point(459, 265))
                            aIconLoc(1) = "3822"
                            aIconPos.Add(New Point(459, 281))
                            aIconLoc(2) = "501"
                            aIconPos.Add(New Point(269, 224))
                            aIconLoc(3) = "7927"
                            aIconPos.Add(New Point(475, 273))
                            aIconLoc(4) = "7924"
                            aIconPos.Add(New Point(253, 224))
                            aIconLoc(5) = "7928"
                            aIconPos.Add(New Point(404, 99))
                            aIconLoc(6) = "7926"
                            aIconPos.Add(New Point(81, 103))
                        Case 3, 10 To 15, 17 To 21, 23 To 26 ' B4
                            aIconLoc(0) = "3805"
                            aIconPos.Add(New Point(220, 334))
                            aIconLoc(1) = "3804"
                            aIconPos.Add(New Point(250, 334))
                            aIconLoc(2) = "3806"
                            aIconPos.Add(New Point(205, 359))
                            aIconLoc(3) = "3809"
                            aIconPos.Add(New Point(324, 220))
                            aIconLoc(4) = "501"
                            aIconPos.Add(New Point(269, 224))
                            aIconLoc(5) = "3821"
                            aIconPos.Add(New Point(463, 157))
                            aIconLoc(6) = "3808"
                            aIconPos.Add(New Point(431, 110))
                            aIconLoc(7) = "3820"
                            aIconPos.Add(New Point(439, 126))
                            aIconLoc(8) = "3810"
                            aIconPos.Add(New Point(130, 53))
                            aIconLoc(9) = "3811"
                            aIconPos.Add(New Point(153, 53))
                            aIconLoc(10) = "3813"
                            aIconPos.Add(New Point(141, 163))
                            aIconLoc(11) = "1731"
                            aIconPos.Add(New Point(195, 213))
                            aIconLoc(12) = "7925"
                            aIconPos.Add(New Point(235, 359))
                            aIconLoc(13) = "7924"
                            aIconPos.Add(New Point(253, 224))
                            aIconLoc(14) = "7928"
                            aIconPos.Add(New Point(404, 99))
                            aIconLoc(15) = "7926"
                            aIconPos.Add(New Point(81, 103))
                    End Select
                Else
                    Select Case iRoom
                        Case 0 To 2, 4 ' B1
                            aIconLoc(0) = "3801"
                            aIconPos.Add(New Point(235, 115))
                            aIconLoc(1) = "3807"
                            aIconPos.Add(New Point(175, 144))
                        Case 5 To 8 ' B2
                            aIconLoc(0) = "3802"
                            aIconPos.Add(New Point(400, 121))
                            aIconLoc(1) = "3803"
                            aIconPos.Add(New Point(382, 199))
                        Case 9, 16, 22 ' B3
                            aIconLoc(0) = "3814"
                            aIconPos.Add(New Point(404, 99))
                            aIconLoc(1) = "3812"
                            aIconPos.Add(New Point(469, 265))
                            aIconLoc(2) = "3822"
                            aIconPos.Add(New Point(469, 281))
                            aIconLoc(3) = "3816"
                            aIconPos.Add(New Point(254, 224))
                            aIconLoc(4) = "506"
                            aIconPos.Add(New Point(78, 104))
                        Case 3, 10 To 15, 17 To 21, 23 To 26 ' B4
                            aIconLoc(0) = "3814"
                            aIconPos.Add(New Point(404, 99))
                            aIconLoc(1) = "3815"
                            aIconPos.Add(New Point(321, 329))
                            aIconLoc(2) = "3805"
                            aIconPos.Add(New Point(220, 334))
                            aIconLoc(3) = "3806"
                            aIconPos.Add(New Point(205, 359))
                            aIconLoc(4) = "3804"
                            aIconPos.Add(New Point(250, 334))
                            aIconLoc(5) = "3809"
                            aIconPos.Add(New Point(324, 220))
                            aIconLoc(6) = "3816"
                            aIconPos.Add(New Point(254, 224))
                            aIconLoc(7) = "3821"
                            aIconPos.Add(New Point(466, 165))
                            aIconLoc(8) = "3808"
                            aIconPos.Add(New Point(431, 115))
                            aIconLoc(9) = "3820"
                            aIconPos.Add(New Point(439, 131))
                            aIconLoc(10) = "3810"
                            aIconPos.Add(New Point(129, 52))
                            aIconLoc(11) = "3811"
                            aIconPos.Add(New Point(153, 52))
                            aIconLoc(12) = "506"
                            aIconPos.Add(New Point(78, 104))
                            aIconLoc(13) = "3813"
                            aIconPos.Add(New Point(141, 163))
                            aIconLoc(14) = "1731"
                            aIconPos.Add(New Point(195, 213))
                            aIconLoc(15) = "7925"
                            aIconPos.Add(New Point(235, 359))
                            aIconLoc(16) = "7924"
                            aIconPos.Add(New Point(466, 149))
                            aIconLoc(17) = "7927"
                            aIconPos.Add(New Point(423, 99))
                            aIconLoc(18) = "7928"
                            aIconPos.Add(New Point(207, 122))
                            aIconLoc(19) = "7926"
                            aIconPos.Add(New Point(185, 190))
                    End Select
                End If
            Case 8
                If Not aMQ(8) Then
                    Select Case iRoom
                        Case 0 To 6 ' B1
                            aIconLoc(0) = "3908"
                            aIconPos.Add(New Point(261, 151))
                            aIconLoc(1) = "3905"
                            aIconPos.Add(New Point(359, 151))
                            aIconLoc(2) = "3901"
                            aIconPos.Add(New Point(269, 129))
                            aIconLoc(3) = "3914"
                            aIconPos.Add(New Point(351, 129))
                            aIconLoc(4) = "3904"
                            aIconPos.Add(New Point(192, 21))
                            aIconLoc(5) = "601"
                            aIconPos.Add(New Point(78, 106))
                            aIconLoc(6) = "3903"
                            aIconPos.Add(New Point(444, 196))
                            aIconLoc(7) = "3920"
                            aIconPos.Add(New Point(471, 196))
                            aIconLoc(8) = "3910"
                            aIconPos.Add(New Point(432, 70))
                            aIconLoc(9) = "3912"
                            aIconPos.Add(New Point(415, 100))
                            aIconLoc(10) = "8000"
                            aIconPos.Add(New Point(431, 100))
                            aIconLoc(11) = "8002"
                            aIconPos.Add(New Point(273, 55))
                            aIconLoc(12) = "8001"
                            aIconPos.Add(New Point(341, 53))
                        Case 7, 8 ' B2
                            aIconLoc(0) = "3902"
                            aIconPos.Add(New Point(295, 143))
                            aIconLoc(1) = "3916"
                            aIconPos.Add(New Point(312, 194))
                            aIconLoc(2) = "3909"
                            aIconPos.Add(New Point(177, 96))
                        Case 9 ' B3
                            aIconLoc(0) = "3907"
                            aIconPos.Add(New Point(386, 171))
                    End Select
                Else
                    Select Case iRoom
                        Case 0 To 6 ' B1
                            aIconLoc(0) = "3903"
                            aIconPos.Add(New Point(312, 97))
                            aIconLoc(1) = "601"
                            aIconPos.Add(New Point(340, 52))
                            aIconLoc(2) = "3902"
                            aIconPos.Add(New Point(444, 187))
                            aIconLoc(3) = "602"
                            aIconPos.Add(New Point(471, 169))
                            aIconLoc(4) = "8001"
                            aIconPos.Add(New Point(273, 66))
                            aIconLoc(5) = "8002"
                            aIconPos.Add(New Point(69, 84))
                        Case 9 ' B3
                            aIconLoc(0) = "3901"
                            aIconPos.Add(New Point(387, 169))
                            aIconLoc(1) = "8000"
                            aIconPos.Add(New Point(222, 20))
                    End Select
                End If
            Case 9 ' IC
                If Not aMQ(9) Then
                    aIconLoc(0) = "4000"
                    aIconPos.Add(New Point(355, 26))
                    aIconLoc(1) = "4001"
                    aIconPos.Add(New Point(410, 237))
                    aIconLoc(2) = "701"
                    aIconPos.Add(New Point(404, 196))
                    aIconLoc(3) = "4002"
                    aIconPos.Add(New Point(178, 228))
                    aIconLoc(4) = "6402"
                    aIconPos.Add(New Point(194, 228))
                    aIconLoc(5) = "8009"
                    aIconPos.Add(New Point(279, 152))
                    aIconLoc(6) = "8010"
                    aIconPos.Add(New Point(420, 204))
                    aIconLoc(7) = "8008"
                    aIconPos.Add(New Point(198, 104))
                Else ' 1F
                    aIconLoc(0) = "4001"
                    aIconPos.Add(New Point(411, 237))
                    aIconLoc(1) = "4000"
                    aIconPos.Add(New Point(355, 26))
                    aIconLoc(2) = "701"
                    aIconPos.Add(New Point(413, 45))
                    aIconLoc(3) = "4002"
                    aIconPos.Add(New Point(178, 228))
                    aIconLoc(4) = "6402"
                    aIconPos.Add(New Point(194, 228))
                    aIconLoc(5) = "8009"
                    aIconPos.Add(New Point(363, 66))
                    aIconLoc(6) = "8010"
                    aIconPos.Add(New Point(158, 126))
                    aIconLoc(7) = "8008"
                    aIconPos.Add(New Point(180, 35))
                End If
            Case 11 ' GTG
                If Not aMQ(10) Then
                    aIconLoc(0) = "4219"
                    aIconPos.Add(New Point(252, 320))
                    aIconLoc(1) = "4207"
                    aIconPos.Add(New Point(289, 320))
                    aIconLoc(2) = "4200"
                    aIconPos.Add(New Point(134, 314))
                    aIconLoc(3) = "4201"
                    aIconPos.Add(New Point(401, 339))
                    aIconLoc(4) = "4217"
                    aIconPos.Add(New Point(163, 115))
                    aIconLoc(5) = "4215"
                    aIconPos.Add(New Point(133, 35))
                    aIconLoc(6) = "4214"
                    aIconPos.Add(New Point(149, 35))
                    aIconLoc(7) = "4220"
                    aIconPos.Add(New Point(149, 19))
                    aIconLoc(8) = "4202"
                    aIconPos.Add(New Point(133, 19))
                    aIconLoc(9) = "4203"
                    aIconPos.Add(New Point(278, 120))
                    aIconLoc(10) = "4204"
                    aIconPos.Add(New Point(271, 210))
                    aIconLoc(11) = "4218"
                    aIconPos.Add(New Point(399, 103))
                    aIconLoc(12) = "4216"
                    aIconPos.Add(New Point(399, 121))
                    aIconLoc(13) = "4205"
                    aIconPos.Add(New Point(291, 226))
                    aIconLoc(14) = "4208"
                    aIconPos.Add(New Point(303, 210))
                    aIconLoc(15) = "801"
                    aIconPos.Add(New Point(372, 224))
                    aIconLoc(16) = "4213"
                    aIconPos.Add(New Point(461, 220))
                    aIconLoc(17) = "4211"
                    aIconPos.Add(New Point(274, 247))
                    aIconLoc(18) = "4206"
                    aIconPos.Add(New Point(226, 223))
                    aIconLoc(19) = "4210"
                    aIconPos.Add(New Point(236, 187))
                    aIconLoc(20) = "4209"
                    aIconPos.Add(New Point(252, 187))
                    aIconLoc(21) = "4212"
                    aIconPos.Add(New Point(271, 226))
                Else ' F1
                    aIconLoc(0) = "4219"
                    aIconPos.Add(New Point(252, 320))
                    aIconLoc(1) = "4207"
                    aIconPos.Add(New Point(289, 320))
                    aIconLoc(2) = "4211"
                    aIconPos.Add(New Point(274, 248))
                    aIconLoc(3) = "4206"
                    aIconPos.Add(New Point(226, 223))
                    aIconLoc(4) = "4210"
                    aIconPos.Add(New Point(236, 187))
                    aIconLoc(5) = "4209"
                    aIconPos.Add(New Point(252, 187))
                    aIconLoc(6) = "4201"
                    aIconPos.Add(New Point(401, 339))
                    aIconLoc(7) = "4213"
                    aIconPos.Add(New Point(461, 220))
                    aIconLoc(8) = "4200"
                    aIconPos.Add(New Point(134, 313))
                    aIconLoc(9) = "4217"
                    aIconPos.Add(New Point(141, 21))
                    aIconLoc(10) = "4202"
                    aIconPos.Add(New Point(162, 115))
                    aIconLoc(11) = "4203"
                    aIconPos.Add(New Point(270, 79))
                    aIconLoc(12) = "4218"
                    aIconPos.Add(New Point(398, 103))
                    aIconLoc(13) = "4214"
                    aIconPos.Add(New Point(398, 121))
                    aIconLoc(14) = "4205"
                    aIconPos.Add(New Point(291, 226))
                    aIconLoc(15) = "4208"
                    aIconPos.Add(New Point(304, 210))
                    aIconLoc(16) = "4204"
                    aIconPos.Add(New Point(280, 192))
                End If
            Case 10, 13 ' GAT & GAC
                If Not aMQ(11) Then
                    Select Case iRoom
                        Case 0 ' Main Upper
                            aIconLoc(0) = "4111"
                            aIconPos.Add(New Point(270, 112))
                        Case 1 ' Main Lower
                            aIconLoc(0) = "8709"
                            aIconPos.Add(New Point(292, 255))
                            aIconLoc(1) = "8706"
                            aIconPos.Add(New Point(278, 271))
                            aIconLoc(2) = "8704"
                            aIconPos.Add(New Point(262, 271))
                            aIconLoc(3) = "8708"
                            aIconPos.Add(New Point(248, 255))
                        Case 2 ' Forest Trial
                            aIconLoc(0) = "4309"
                            aIconPos.Add(New Point(213, 111))
                        Case 3 ' Water Trial
                            aIconLoc(0) = "4307"
                            aIconPos.Add(New Point(196, 200))
                            aIconLoc(1) = "4306"
                            aIconPos.Add(New Point(196, 245))
                        Case 4 ' Shadow Trial
                            aIconLoc(0) = "4308"
                            aIconPos.Add(New Point(183, 292))
                            aIconLoc(1) = "4305"
                            aIconPos.Add(New Point(305, 173))
                        Case 6 ' Light Trial
                            aIconLoc(0) = "4312"
                            aIconPos.Add(New Point(430, 217))
                            aIconLoc(1) = "4311"
                            aIconPos.Add(New Point(410, 221))
                            aIconLoc(2) = "4313"
                            aIconPos.Add(New Point(390, 217))
                            aIconLoc(3) = "4314"
                            aIconPos.Add(New Point(430, 181))
                            aIconLoc(4) = "4310"
                            aIconPos.Add(New Point(410, 177))
                            aIconLoc(5) = "4315"
                            aIconPos.Add(New Point(390, 181))
                            aIconLoc(6) = "4316"
                            aIconPos.Add(New Point(410, 199))
                            aIconLoc(7) = "4317"
                            aIconPos.Add(New Point(351, 191))
                        Case 7 ' Spirit Trial
                            aIconLoc(0) = "4318"
                            aIconPos.Add(New Point(219, 140))
                            aIconLoc(1) = "4320"
                            aIconPos.Add(New Point(196, 182))
                    End Select
                Else
                    Select Case iRoom
                        Case 0 ' Main Upper
                            aIconLoc(0) = "4111"
                            aIconPos.Add(New Point(270, 112))
                        Case 1 ' Main Lower
                            aIconLoc(0) = "8709"
                            aIconPos.Add(New Point(292, 255))
                            aIconLoc(1) = "8706"
                            aIconPos.Add(New Point(278, 271))
                            aIconLoc(2) = "8704"
                            aIconPos.Add(New Point(262, 271))
                            aIconLoc(3) = "8708"
                            aIconPos.Add(New Point(248, 255))
                            aIconLoc(4) = "8701"
                            aIconPos.Add(New Point(248, 239))
                        Case 2 ' Forest Trial
                            aIconLoc(0) = "4302"
                            aIconPos.Add(New Point(257, 166))
                            aIconLoc(1) = "4302"
                            aIconPos.Add(New Point(304, 212))
                            aIconLoc(2) = "901"
                            aIconPos.Add(New Point(246, 145))
                        Case 3 ' Water Trial
                            aIconLoc(0) = "4301"
                            aIconPos.Add(New Point(167, 189))
                        Case 4 ' Shadow Trial
                            aIconLoc(0) = "4300"
                            aIconPos.Add(New Point(183, 292))
                            aIconLoc(1) = "4305"
                            aIconPos.Add(New Point(305, 173))
                        Case 6 ' Light Trial
                            aIconLoc(0) = "4304"
                            aIconPos.Add(New Point(351, 191))
                        Case 7 ' Spirit Trial
                            aIconLoc(0) = "4310"
                            aIconPos.Add(New Point(219, 140))
                            aIconLoc(1) = "4320"
                            aIconPos.Add(New Point(196, 182))
                            aIconLoc(2) = "4309"
                            aIconPos.Add(New Point(262, 202))
                            aIconLoc(3) = "4308"
                            aIconPos.Add(New Point(284, 214))
                            aIconLoc(4) = "4307"
                            aIconPos.Add(New Point(272, 236))
                            aIconLoc(5) = "4306"
                            aIconPos.Add(New Point(250, 224))
                    End Select
                End If
            Case 27 To 29   ' MK Entrance
                aIconLoc(0) = "8119"
                aIconPos.Add(New Point(186, 101))
            Case 30, 31 ' MK Back Alley
                aIconLoc(0) = "7201"
                aIconPos.Add(New Point(160, 337))
                If My.Settings.setShop > 0 Then
                    aIconLoc(1) = name2loc("Bombchu Shop: Lower-Left", "MK")
                    aIconPos.Add(New Point(201, 393))
                    aIconLoc(2) = name2loc("Bombchu Shop: Lower-Right", "MK")
                    aIconPos.Add(New Point(217, 393))
                    aIconLoc(3) = name2loc("Bombchu Shop: Upper-Left", "MK")
                    aIconPos.Add(New Point(201, 377))
                    aIconLoc(4) = name2loc("Bombchu Shop: Upper-Right", "MK")
                    aIconPos.Add(New Point(217, 377))
                End If
            Case 32, 33 ' MK Young
                aIconLoc(0) = "6829"
                aIconPos.Add(New Point(187, 28))
                aIconLoc(1) = "6801"
                aIconPos.Add(New Point(100, 191))
                aIconLoc(2) = "6802"
                aIconPos.Add(New Point(100, 207))
                aIconLoc(3) = "7201"
                aIconPos.Add(New Point(369, 324))
                aIconLoc(4) = "6811"
                aIconPos.Add(New Point(134, 371))
                If My.Settings.setShop > 0 Then
                    aIconLoc(5) = name2loc("Potion Shop: Lower-Left", "MK")
                    aIconPos.Add(New Point(458, 208))
                    aIconLoc(6) = name2loc("Bazaar: Lower-Left", "MK")
                    aIconPos.Add(New Point(458, 314))
                    aIconLoc(7) = name2loc("Potion Shop: Lower-Right", "MK")
                    aIconPos.Add(New Point(474, 208))
                    aIconLoc(8) = name2loc("Bazaar: Lower-Right", "MK")
                    aIconPos.Add(New Point(474, 314))
                    aIconLoc(9) = name2loc("Potion Shop: Upper-Left", "MK")
                    aIconPos.Add(New Point(458, 192))
                    aIconLoc(10) = name2loc("Bazaar: Upper-Left", "MK")
                    aIconPos.Add(New Point(458, 298))
                    aIconLoc(11) = name2loc("Potion Shop: Upper-Right", "MK")
                    aIconPos.Add(New Point(474, 192))
                    aIconLoc(12) = name2loc("Bazaar: Upper-Right", "MK")
                    aIconPos.Add(New Point(474, 298))
                End If
            Case 67 ' ToT
                aIconLoc(0) = "6405"
                aIconPos.Add(New Point(270, 115))
                aIconLoc(1) = locSwap(5)
                aIconPos.Add(New Point(270, 299))
            Case 81 ' HF
                aIconLoc(0) = "4600"
                aIconPos.Add(New Point(274, 66))
                aIconLoc(1) = "4602"
                aIconPos.Add(New Point(295, 295))
                aIconLoc(2) = "4603"
                aIconPos.Add(New Point(224, 332))
                aIconLoc(3) = "6827"
                aIconPos.Add(New Point(208, 332))
                aIconLoc(4) = "1901"
                aIconPos.Add(New Point(204, 99))
                aIconLoc(5) = "103"
                aIconPos.Add(New Point(315, 66))
                aIconLoc(6) = locSwap(9)
                aIconPos.Add(New Point(299, 66))
                aIconLoc(7) = "8016"
                aIconPos.Add(New Point(132, 185))
                aIconLoc(8) = "8017"
                aIconPos.Add(New Point(352, 43))
                aIconLoc(9) = "1925"
                aIconPos.Add(New Point(148, 185))
                aIconLoc(10) = "8803"
                aIconPos.Add(New Point(208, 332))
                ' Sell Bunny Hood
                aIconLoc(11) = "6911"
                aIconPos.Add(New Point(244, 137))
                ' Big Poe Hunt
                aIconLoc(12) = "124"
                aIconPos.Add(New Point(307, 89))
                aIconLoc(13) = "123"
                aIconPos.Add(New Point(260, 141))
                aIconLoc(14) = "122"
                aIconPos.Add(New Point(193, 75))
                aIconLoc(15) = "130"
                aIconPos.Add(New Point(186, 180))
                aIconLoc(16) = "131"
                aIconPos.Add(New Point(179, 237))
                aIconLoc(17) = "128"
                aIconPos.Add(New Point(298, 274))
                aIconLoc(18) = "129"
                aIconPos.Add(New Point(311, 295))
                aIconLoc(19) = "127"
                aIconPos.Add(New Point(324, 213))
                aIconLoc(20) = "126"
                aIconPos.Add(New Point(336, 156))
                aIconLoc(21) = "125"
                aIconPos.Add(New Point(387, 80))
            Case 82 ' KV
                aIconLoc(0) = "4610"
                aIconPos.Add(New Point(265, 225))
                aIconLoc(1) = "4608"
                aIconPos.Add(New Point(357, 171))
                aIconLoc(2) = "6828"
                aIconPos.Add(New Point(350, 324))
                aIconLoc(3) = "6928"
                aIconPos.Add(New Point(350, 340))
                aIconLoc(4) = "6805"
                aIconPos.Add(New Point(332, 198))
                aIconLoc(5) = "1801"
                aIconPos.Add(New Point(306, 345))
                aIconLoc(6) = locSwap(14)
                aIconPos.Add(New Point(315, 288))
                aIconLoc(7) = "2001"
                aIconPos.Add(New Point(412, 240))
                aIconLoc(8) = "6411"
                aIconPos.Add(New Point(428, 240))
                aIconLoc(9) = "6404"
                aIconPos.Add(New Point(312, 244))
                aIconLoc(10) = "8205"
                aIconPos.Add(New Point(234, 241))
                aIconLoc(11) = "8203"
                aIconPos.Add(New Point(331, 288))
                aIconLoc(12) = "8204"
                aIconPos.Add(New Point(245, 303))
                aIconLoc(13) = "8201"
                aIconPos.Add(New Point(274, 120))
                aIconLoc(14) = "8202"
                aIconPos.Add(New Point(308, 181))
                aIconLoc(15) = "8206"
                aIconPos.Add(New Point(290, 345))
                aIconLoc(16) = "1824"
                aIconPos.Add(New Point(274, 345))
                ' Sell Keaton Mask
                aIconLoc(17) = "6908"
                aIconPos.Add(New Point(290, 81))
                If My.Settings.setShop > 0 Then
                    aIconLoc(18) = name2loc("Bazaar: Lower-Left", "KV")
                    aIconPos.Add(New Point(255, 152))
                    aIconLoc(19) = name2loc("Potion Shop: Lower-Left", "KV")
                    aIconPos.Add(New Point(321, 158))
                    aIconLoc(20) = name2loc("Bazaar: Lower-Right", "KV")
                    aIconPos.Add(New Point(271, 152))
                    aIconLoc(21) = name2loc("Potion Shop: Lower-Right", "KV")
                    aIconPos.Add(New Point(337, 158))
                    aIconLoc(22) = name2loc("Bazaar: Upper-Left", "KV")
                    aIconPos.Add(New Point(255, 136))
                    aIconLoc(23) = name2loc("Potion Shop: Upper-Left", "KV")
                    aIconPos.Add(New Point(321, 142))
                    aIconLoc(24) = name2loc("Bazaar: Upper-Right", "KV")
                    aIconPos.Add(New Point(271, 136))
                    aIconLoc(25) = name2loc("Potion Shop: Upper-Right", "KV")
                    aIconPos.Add(New Point(337, 142))
                End If
            Case 83 ' GY
                aIconLoc(0) = locSwap(2)
                aIconPos.Add(New Point(156, 203))
                aIconLoc(1) = "4800"
                aIconPos.Add(New Point(178, 214))
                aIconLoc(2) = "2204"
                aIconPos.Add(New Point(110, 160))
                aIconLoc(3) = "4700"
                aIconPos.Add(New Point(226, 231))
                aIconLoc(4) = "4900"
                aIconPos.Add(New Point(286, 201))
                aIconLoc(5) = locSwap(8)
                aIconPos.Add(New Point(286, 217))
                aIconLoc(6) = "5000"
                aIconPos.Add(New Point(171, 152))
                aIconLoc(7) = "2007"
                aIconPos.Add(New Point(155, 152))
                aIconLoc(8) = "8200"
                aIconPos.Add(New Point(139, 160))
                aIconLoc(9) = "8027"
                aIconPos.Add(New Point(226, 273))
            Case 84 ' ZR
                aIconLoc(0) = "2304"
                aIconPos.Add(New Point(202, 111))
                aIconLoc(1) = "4609"
                aIconPos.Add(New Point(198, 211))
                aIconLoc(2) = "2311"
                aIconPos.Add(New Point(438, 78))
                aIconLoc(3) = "2301"
                aIconPos.Add(New Point(117, 179))
                aIconLoc(4) = "8209"
                aIconPos.Add(New Point(50, 238))
                aIconLoc(5) = "8208"
                aIconPos.Add(New Point(454, 82))
                aIconLoc(6) = "8212"
                aIconPos.Add(New Point(176, 216))
                aIconLoc(7) = "8211"
                aIconPos.Add(New Point(381, 85))
                aIconLoc(8) = "8909"
                aIconPos.Add(New Point(48, 187))
                aIconLoc(9) = "8908"
                aIconPos.Add(New Point(64, 187))
                aIconLoc(10) = "6700"
                aIconPos.Add(New Point(251, 107))
            Case 85 ' KF
                aIconLoc(0) = "4500"
                aIconPos.Add(New Point(144, 160))
                aIconLoc(1) = "4501"
                aIconPos.Add(New Point(160, 160))
                aIconLoc(2) = "4502"
                aIconPos.Add(New Point(144, 176))
                aIconLoc(3) = "4503"
                aIconPos.Add(New Point(160, 176))
                aIconLoc(4) = "5100"
                aIconPos.Add(New Point(162, 345))
                aIconLoc(5) = "4612"
                aIconPos.Add(New Point(154, 133))
                aIconLoc(6) = "8101"
                aIconPos.Add(New Point(104, 229))
                aIconLoc(7) = "8102"
                aIconPos.Add(New Point(247, 243))
                aIconLoc(8) = "10024"
                aIconPos.Add(New Point(177, 274))
                ' Moved the Soil Gold Skultulla down to combine it into the shopsanity check because it needs to be moved if shopsanity is on
                aIconLoc(9) = "8100"
                If My.Settings.setShop = 0 Then
                    aIconPos.Add(New Point(249, 175))
                Else
                    aIconPos.Add(New Point(254, 178))
                    aIconLoc(10) = name2loc("Shop: Lower-Left", "KF")
                    aIconPos.Add(New Point(222, 186))
                    aIconLoc(11) = name2loc("Shop: Lower-Right", "KF")
                    aIconPos.Add(New Point(238, 186))
                    aIconLoc(12) = name2loc("Shop: Upper-Left", "KF")
                    aIconPos.Add(New Point(222, 170))
                    aIconLoc(13) = name2loc("Shop: Upper-Right", "KF")
                    aIconPos.Add(New Point(238, 170))
                End If
            Case 86 ' SFM
                aIconLoc(0) = "4617"
                aIconPos.Add(New Point(262, 345))
                aIconLoc(1) = locSwap(6)
                aIconPos.Add(New Point(265, 64))
                aIconLoc(2) = "6400"
                aIconPos.Add(New Point(281, 64))
                aIconLoc(3) = "8111"
                aIconPos.Add(New Point(301, 246))
                aIconLoc(4) = "9009"
                aIconPos.Add(New Point(281, 94))
                aIconLoc(5) = "9008"
                aIconPos.Add(New Point(297, 94))
            Case 87 ' LH
                aIconLoc(0) = "7316"
                aIconPos.Add(New Point(334, 144))
                aIconLoc(1) = "6512"
                aIconPos.Add(New Point(350, 144))
                aIconLoc(2) = "211"
                aIconPos.Add(New Point(306, 196))
                aIconLoc(3) = "6110"
                aIconPos.Add(New Point(393, 169))
                aIconLoc(4) = "6111"
                aIconPos.Add(New Point(393, 185))
                aIconLoc(5) = "2430"
                aIconPos.Add(New Point(231, 172))
                aIconLoc(6) = "6800"
                aIconPos.Add(New Point(247, 172))
                aIconLoc(7) = locSwap(3)
                aIconPos.Add(New Point(361, 310))
                aIconLoc(8) = "8216"
                aIconPos.Add(New Point(239, 156))
                aIconLoc(9) = "8218"
                aIconPos.Add(New Point(231, 188))
                aIconLoc(10) = "8217"
                aIconPos.Add(New Point(377, 310))
                aIconLoc(11) = "8219"
                aIconPos.Add(New Point(247, 188))
                aIconLoc(12) = "8220"
                aIconPos.Add(New Point(315, 315))
                aIconLoc(13) = "9101"
                aIconPos.Add(New Point(208, 265))
                aIconLoc(14) = "9104"
                aIconPos.Add(New Point(224, 265))
                aIconLoc(15) = "9106"
                aIconPos.Add(New Point(240, 265))
            Case 88 ' ZD
                aIconLoc(0) = "5300"
                aIconPos.Add(New Point(309, 207))
                aIconLoc(1) = "6308"
                aIconPos.Add(New Point(309, 191))
                aIconLoc(2) = "7109"
                aIconPos.Add(New Point(398, 150))
                aIconLoc(3) = "8214"
                aIconPos.Add(New Point(325, 207))
                If My.Settings.setShop > 0 Then
                    aIconLoc(4) = name2loc("Shop: Lower-Left", "ZD")
                    aIconPos.Add(New Point(382, 321))
                    aIconLoc(5) = name2loc("Shop: Lower-Right", "ZD")
                    aIconPos.Add(New Point(398, 321))
                    aIconLoc(6) = name2loc("Shop: Upper-Left", "ZD")
                    aIconPos.Add(New Point(382, 305))
                    aIconLoc(7) = name2loc("Shop: Upper-Right", "ZD")
                    aIconPos.Add(New Point(398, 305))
                End If
            Case 89 ' ZF
                aIconLoc(0) = locSwap(10)
                aIconPos.Add(New Point(378, 380))
                aIconLoc(1) = "2501"
                aIconPos.Add(New Point(424, 172))
                aIconLoc(2) = "2520"
                aIconPos.Add(New Point(353, 172))
                aIconLoc(3) = "8210"
                aIconPos.Add(New Point(221, 311))
                aIconLoc(4) = "8215"
                aIconPos.Add(New Point(361, 338))
                aIconLoc(5) = "8213"
                aIconPos.Add(New Point(430, 302))
            Case 90 ' GV
                aIconLoc(0) = "2602"
                aIconPos.Add(New Point(267, 259))
                aIconLoc(1) = "2601"
                aIconPos.Add(New Point(293, 24))
                aIconLoc(2) = "5400"
                aIconPos.Add(New Point(210, 218))
                aIconLoc(3) = "8225"
                aIconPos.Add(New Point(386, 157))
                aIconLoc(4) = "8224"
                aIconPos.Add(New Point(261, 177))
                aIconLoc(5) = "8227"
                aIconPos.Add(New Point(223, 105))
                aIconLoc(6) = "8226"
                aIconPos.Add(New Point(216, 195))
                aIconLoc(7) = "2624"
                aIconPos.Add(New Point(261, 193))
                aIconLoc(8) = "9209"
                aIconPos.Add(New Point(202, 124))
                aIconLoc(9) = "9208"
                aIconPos.Add(New Point(218, 124))
            Case 91 ' LW
                aIconLoc(0) = "7202"
                aIconPos.Add(New Point(163, 391))
                aIconLoc(1) = "4620"
                aIconPos.Add(New Point(304, 149))
                aIconLoc(2) = "7203"
                aIconPos.Add(New Point(280, 47))
                aIconLoc(3) = "6717"
                aIconPos.Add(New Point(166, 332))
                aIconLoc(4) = "6813"
                aIconPos.Add(New Point(307, 210))
                aIconLoc(5) = "6807"
                aIconPos.Add(New Point(318, 239))
                aIconLoc(6) = "6806"
                aIconPos.Add(New Point(169, 217))
                aIconLoc(7) = locSwap(13)
                aIconPos.Add(New Point(227, 118))
                aIconLoc(8) = "6815"
                aIconPos.Add(New Point(243, 118))
                aIconLoc(9) = "8108"
                aIconPos.Add(New Point(155, 278))
                aIconLoc(10) = "8109"
                aIconPos.Add(New Point(280, 98))
                aIconLoc(11) = "8110"
                aIconPos.Add(New Point(296, 98))
                aIconLoc(12) = "9810"
                aIconPos.Add(New Point(163, 391))
                aIconLoc(13) = "9801"
                aIconPos.Add(New Point(296, 114))
                aIconLoc(14) = "9802"
                aIconPos.Add(New Point(280, 114))
                aIconLoc(15) = "9311"
                aIconPos.Add(New Point(272, 47))
                aIconLoc(16) = "9304"
                aIconPos.Add(New Point(288, 47))
                ' Sell Skull Mask
                aIconLoc(17) = "6909"
                aIconPos.Add(New Point(185, 217))
            Case 92 ' DC
                aIconLoc(0) = locSwap(12)
                aIconPos.Add(New Point(321, 96))
                aIconLoc(1) = "6628"
                aIconPos.Add(New Point(124, 190))
                aIconLoc(2) = "2713"
                aIconPos.Add(New Point(154, 190))
                aIconLoc(3) = "8308"
                aIconPos.Add(New Point(117, 209))
                aIconLoc(4) = "8310"
                aIconPos.Add(New Point(283, 166))
                aIconLoc(5) = "8311"
                aIconPos.Add(New Point(219, 317))
                aIconLoc(6) = "9709"
                aIconPos.Add(New Point(196, 109))
                aIconLoc(7) = "9708"
                aIconPos.Add(New Point(212, 109))
            Case 93, 12 ' GF + Thieves' Hideout
                aIconLoc(0) = "5600"
                aIconPos.Add(New Point(311, 177))
                aIconLoc(1) = "7200"
                aIconPos.Add(New Point(454, 289))
                aIconLoc(2) = "6831"
                aIconPos.Add(New Point(454, 305))
                aIconLoc(3) = "6500"
                aIconPos.Add(New Point(289, 192))
                aIconLoc(4) = "6502"
                aIconPos.Add(New Point(289, 260))
                aIconLoc(5) = "6503"
                aIconPos.Add(New Point(327, 184))
                aIconLoc(6) = "6501"
                aIconPos.Add(New Point(307, 193))
                aIconLoc(7) = "3801"
                aIconPos.Add(New Point(342, 206))
                aIconLoc(8) = "8300"
                aIconPos.Add(New Point(440, 52))
            Case 94 ' HW
                aIconLoc(0) = "5700"
                aIconPos.Add(New Point(299, 79))
                aIconLoc(1) = "8309"
                aIconPos.Add(New Point(299, 95))
                aIconLoc(2) = "11401"
                aIconPos.Add(New Point(341, 362))
            Case 95 ' HC
                aIconLoc(0) = "6202"
                aIconPos.Add(New Point(171, 287))
                aIconLoc(1) = "6416"
                aIconPos.Add(New Point(139, 87))
                aIconLoc(2) = "6409"
                aIconPos.Add(New Point(155, 87))
                aIconLoc(3) = locSwap(11)
                aIconPos.Add(New Point(334, 246))
                aIconLoc(4) = "8117"
                aIconPos.Add(New Point(223, 141))
                aIconLoc(5) = "8118"
                aIconPos.Add(New Point(155, 279))
            Case 96 ' DMT
                aIconLoc(0) = "5801"
                aIconPos.Add(New Point(286, 239))
                aIconLoc(1) = "2830"
                aIconPos.Add(New Point(241, 198))
                aIconLoc(2) = locSwap(1)
                aIconPos.Add(New Point(294, 31))
                aIconLoc(3) = "4623"
                aIconPos.Add(New Point(288, 185))
                aIconLoc(4) = "6008"
                aIconPos.Add(New Point(336, 41))
                aIconLoc(5) = "8125"
                aIconPos.Add(New Point(233, 214))
                aIconLoc(6) = "8126"
                aIconPos.Add(New Point(222, 303))
                aIconLoc(7) = "8127"
                aIconPos.Add(New Point(249, 214))
                aIconLoc(8) = "8128"
                aIconPos.Add(New Point(307, 116))
                aIconLoc(9) = "1924"
                aIconPos.Add(New Point(270, 233))
            Case 97 'DMC
                aIconLoc(0) = "4626"
                aIconPos.Add(New Point(297, 344))
                aIconLoc(1) = "2902"
                aIconPos.Add(New Point(300, 280))
                aIconLoc(2) = "2908"
                aIconPos.Add(New Point(242, 198))
                aIconLoc(3) = locSwap(0)
                aIconPos.Add(New Point(187, 269))
                aIconLoc(4) = "6401"
                aIconPos.Add(New Point(237, 160))
                aIconLoc(5) = "8131"
                aIconPos.Add(New Point(216, 351))
                aIconLoc(6) = "8124"
                aIconPos.Add(New Point(285, 176))
                aIconLoc(7) = "C01"
                aIconPos.Add(New Point(234, 296))
                aIconLoc(8) = "9401"
                aIconPos.Add(New Point(141, 153))
                aIconLoc(9) = "9404"
                aIconPos.Add(New Point(157, 153))
                aIconLoc(10) = "9406"
                aIconPos.Add(New Point(173, 153))
            Case 98 ' GC
                aIconLoc(0) = "5900"
                aIconPos.Add(New Point(102, 72))
                aIconLoc(1) = "5902"
                aIconPos.Add(New Point(118, 72))
                aIconLoc(2) = "5901"
                aIconPos.Add(New Point(134, 72))
                aIconLoc(3) = "7014"
                aIconPos.Add(New Point(257, 254))
                aIconLoc(4) = "7025"
                aIconPos.Add(New Point(273, 254))
                aIconLoc(5) = "3031"
                aIconPos.Add(New Point(266, 196))
                aIconLoc(6) = locSwap(4)
                aIconPos.Add(New Point(273, 58))
                aIconLoc(7) = "3001"
                aIconPos.Add(New Point(148, 345))
                aIconLoc(8) = "8130"
                aIconPos.Add(New Point(134, 56))
                aIconLoc(9) = "8129"
                aIconPos.Add(New Point(266, 212))
                aIconLoc(10) = "9501"
                aIconPos.Add(New Point(363, 69))
                aIconLoc(11) = "9504"
                aIconPos.Add(New Point(379, 69))
                aIconLoc(12) = "9506"
                aIconPos.Add(New Point(395, 69))
                If My.Settings.setShop > 0 Then
                    aIconLoc(13) = name2loc("Shop: Lower-Left", "GC")
                    aIconPos.Add(New Point(228, 203))
                    aIconLoc(14) = name2loc("Shop: Lower-Right", "GC")
                    aIconPos.Add(New Point(244, 203))
                    aIconLoc(15) = name2loc("Shop: Upper-Left", "GC")
                    aIconPos.Add(New Point(228, 187))
                    aIconLoc(16) = name2loc("Shop: Upper-Right", "GC")
                    aIconPos.Add(New Point(244, 187))
                End If
            Case 99 ' LLR
                aIconLoc(0) = "2101"
                aIconPos.Add(New Point(164, 364))
                aIconLoc(1) = "6818"
                aIconPos.Add(New Point(340, 85))
                aIconLoc(2) = locSwap(7)
                aIconPos.Add(New Point(251, 229))
                aIconLoc(3) = "6208"
                aIconPos.Add(New Point(267, 229))
                aIconLoc(4) = "8024"
                aIconPos.Add(New Point(129, 315))
                aIconLoc(5) = "8025"
                aIconPos.Add(New Point(312, 289))
                aIconLoc(6) = "8026"
                aIconPos.Add(New Point(323, 106))
                aIconLoc(7) = "8027"
                aIconPos.Add(New Point(336, 131))
                aIconLoc(8) = "10124"
                aIconPos.Add(New Point(277, 103))
                aIconLoc(9) = "10125"
                aIconPos.Add(New Point(277, 87))
                aIconLoc(10) = "2125"
                aIconPos.Add(New Point(156, 380))
                aIconLoc(11) = "2124"
                aIconPos.Add(New Point(172, 380))
                aIconLoc(12) = "6214"
                aIconPos.Add(New Point(259, 245))
                aIconLoc(13) = "9601"
                aIconPos.Add(New Point(362, 354))
                aIconLoc(14) = "9604"
                aIconPos.Add(New Point(378, 354))
                aIconLoc(15) = "9606"
                aIconPos.Add(New Point(394, 354))
            Case 100 ' OGC
                aIconLoc(0) = "008"
                aIconPos.Add(New Point(460, 165))
                aIconLoc(1) = "8116"
                aIconPos.Add(New Point(355, 168))
        End Select

        Dim key As New keyCheck
        Dim fillColour As Color = Color.Lime
        Dim shutupLambda As Integer = 0
        Dim addCheck As Boolean = False
        Dim prefix As String = String.Empty
        Dim suffix As String = String.Empty

        For i = 0 To aIconLoc.Length - 1
            If aIconLoc(i) = String.Empty Then Exit For
            If Not aIconLoc(i) = "0" Then
                ' This shuts VB.NET up about the /!\ Warning about iteration variable in the lambda expression. It is dumb, I know.
                shutupLambda = i

                If iLastMinimap > 25 Then

                    ' Checks normal keys
                    For Each thisKey In aKeys.Where(Function(k As keyCheck) k.loc.Equals(aIconLoc(shutupLambda)))
                        key = thisKey
                    Next
                Else
                    Dim ii As Byte = CByte(iLastMinimap)
                    Select Case ii
                        Case 11
                            ii = 10
                        Case 10, 13
                            ii = 11
                    End Select
                    For Each thiskey In aKeysDungeons(ii).Where(Function(k As keyCheck) k.loc.Equals(aIconLoc(shutupLambda)))
                        key = thiskey
                    Next
                End If

                addCheck = True
                prefix = String.Empty
                suffix = String.Empty
                If key.gs Then
                    prefix = "GS: "
                    If My.Settings.setGSLoc >= 1 Then
                        Select Case My.Settings.setSkulltula
                            Case 0
                                addCheck = False
                            Case 1
                                addCheck = True
                            Case Else
                                addCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                        End Select
                    End If
                ElseIf key.cow Then
                    prefix = "Cow: "
                    addCheck = My.Settings.setCow
                ElseIf key.scrub Then
                    prefix = "Scrub: "
                    addCheck = My.Settings.setScrub
                    'ElseIf key.shop Then
                    'If My.Settings.setShop > 0 Then addCheck = True
                End If

                If My.Settings.setHideQuests And key.area = "QBPH" Then addCheck = False

                With aIconPos(i)
                    If (key.scan Or Mid(key.loc, 1, 3) = "650") And addCheck Then
                        lRegions.Add(New Rectangle(.X, .Y, 15, 15))

                        Graphics.FromImage(pbxMap.Image).DrawRectangle(Pens.Black, .X, .Y, 15, 15)
                        Graphics.FromImage(pbxMap.Image).DrawRectangle(Pens.White, .X + 1, .Y + 1, 13, 13)

                        If Not key.forced And Not key.checked Then
                            fillColour = Color.Lime
                            Select Case checkLogic(key.logic, key.zone)
                                Case 0
                                    fillColour = Color.Red
                                Case 1
                                    suffix = " (Y)"
                                Case 2
                                    suffix = " (A)"
                            End Select
                            Graphics.FromImage(pbxMap.Image).FillRectangle(New SolidBrush(fillColour), .X + 2, .Y + 2, 12, 12)
                        End If
                        aIconName(i) = prefix & key.name & suffix
                    Else
                        ' Have to create an unreachable location to at least populate the list to make it match up
                        lRegions.Add(New Rectangle(-2, -2, 1, 1))
                    End If
                End With
            End If
        Next
        If iLastMinimap = 94 Then
            ' Slow down the slow scan, and speed up the fast scan
            tmrAutoScan.Interval = 7500
            tmrFastScan.Interval = 333
            wastelandPOS()
        Else
            ' Return them to normal
            tmrAutoScan.Interval = 5000
            tmrFastScan.Interval = 1000
        End If
        'pbxMap.Invalidate()
        'pbxMap.Update()
    End Sub
    Private Function exit2label(ByVal exitCode As String, ByRef reachMap As Byte) As String
        exit2label = String.Empty
        Select Case exitCode
            Case "09C", "0BB", "0C1", "0C9", "209", "211", "266", "26A", "272", "286", "33C", "433", "437", "443", "447"
                ' KF Main
                exit2label = "KF"
                reachMap = 0
            Case "20D"
                ' KF Trapped
                exit2label = "KF"
                reachMap = 0
            Case "11E", "4D6", "4DA"
                ' LW Front
                exit2label = "LW Front"
                reachMap = 2
            Case "1A9"
                ' LW Behind Mido
                exit2label = "LW Back"
                reachMap = 3
            Case "4DE"
                ' LW Bridge
                exit2label = "LW Bridge"
                reachMap = 4
            Case "5E0"
                ' LW Between Bridge (For Gift from Saria)
                exit2label = "LW Bridge"
                reachMap = 8
            Case "0FC", "215", "600"
                ' SFM Main
                exit2label = "SFM"
                reachMap = 5
            Case "17D", "181", "185", "189", "18D", "1F9", "1FD", "27E", "311"
                ' HF
                exit2label = "HF"
                reachMap = 7
            Case "04F", "157", "2F9", "378", "42F", "5D0", "5D4"
                ' LLR
                exit2label = "LLR"
                reachMap = 9
            Case "033", "034", "035", "036", "26E", "26F", "270", "271", "276", "277", "278", "279"
                ' MK Entrance
                exit2label = "MK Entrance"
                reachMap = 197
            Case "063", "067", "07E", "0B1", "16D", "1CD", "1D1", "1D5", "25A", "25E", "262", "263", "29E", "29F", "2A2", "388",
                    "3B8", "3BC", "3C0", "43B", "507", "528", "52C", "530"
                ' MK
                exit2label = "MK"
                reachMap = 10
            Case "29A", "29B", "29C", "29D"
                ' MK Alley Back (left)
                exit2label = "MK Back Alley"
                reachMap = 198
            Case "0AD", "0AE", "0AF", "0B0"
                ' MK Alley Back (right)
                exit2label = "MK Back Alley"
                reachMap = 199
            Case "171", "172", "472"
                ' ToT Front
                exit2label = "Outside ToT"
                reachMap = 1
            Case "053", "054", "5F4"
                ' ToT
                exit2label = "ToT"
                reachMap = 200
            Case "138", "23D", "340"
                ' HC & OGC
                If isAdult Then
                    exit2label = "OGC"
                    reachMap = 51
                Else
                    exit2label = "HC"
                    reachMap = 13
                End If
            Case "07A", "296"
                ' HC Castle Courtyard
                exit2label = "Castle Courtyard"
                reachMap = 13
            Case "400", "5F0"
                ' HC Zelda's Courtyard
                exit2label = "Zelda's Courtyard"
                reachMap = 13
            Case "03B", "072", "0B7", "0DB", "195", "201", "2FD", "2A6", "345", "349", "34D", "351", "384", "39C", "3EC", "44B",
                    "453", "463", "4EE", "4FF", "550", "5C8", "5DC"
                ' KV Main
                exit2label = "KV"
                reachMap = 15
            Case "554"
                ' KV Rooftops
                exit2label = "KV Roofs"
                reachMap = 16
            Case "191"
                ' KV Behind Gate
                exit2label = "KV"
                reachMap = 17
            Case "0E4", "30D", "355"
                ' GY Main
                exit2label = "GY"
                reachMap = 18
            Case "205", "568"
                ' GY Upper
                exit2label = "GY near Temple"
                reachMap = 19
            Case "13D", "1B9", "242"
                ' DMT Lower
                exit2label = "DMT Lower"
                reachMap = 20
            Case "1BD", "315", "45B"
                ' DMT Upper 
                exit2label = "DMT Upper"
                reachMap = 21
            Case "147"
                ' DMC Upper Local
                exit2label = "DMC Upper"
                reachMap = 22
            Case "246", "482", "4BE"
                ' DMC Lower Nearby
                exit2label = "DMC near GC"
                reachMap = 24
            Case "4F6"
                ' DMC Central Local
                exit2label = "DMC Warp"
                reachMap = 27
            Case "24A"
                ' DMC near Temple
                exit2label = "DMC near Temple"
                reachMap = 28
            Case "14D", "3FC"
                ' GC Main
                exit2label = "GC"
                reachMap = 29
            Case "4E2"
                ' GC Shortcut
                exit2label = "GC"
                reachMap = 30
            Case "1C1"
                ' GC Darunia
                exit2label = "GC Darunia"
                reachMap = 31
            Case "37C"
                ' GC Shoppe
                exit2label = "GC Shop"
                reachMap = 32
            Case "0EA"
                ' ZR Front
                exit2label = "ZR Front"
                reachMap = 34
            Case "199", "1DD"
                ' ZR Main
                exit2label = "ZR"
                reachMap = 35
            Case "19D"
                ' ZR Behind Waterfall
                exit2label = "ZR"
                reachMap = 36
            Case "108", "153", "328", "3C4"
                ' ZD Main
                exit2label = "ZD"
                reachMap = 37
            Case "1A1"
                ' ZD Behind King
                exit2label = "ZD Behind King"
                reachMap = 38
            Case "221", "225", "371", "394"
                ' ZF Main
                exit2label = "ZF"
                reachMap = 40
            Case "3D4"
                ' ZF Ledge
                exit2label = "ZF Ledge"
                reachMap = 41
            Case "043", "102", "219", "21D", "3CC", "560", "604"
                ' LH Main
                exit2label = "LH"
                reachMap = 42
            Case "309", "45F"
                ' LH Fishing Ledge
                exit2label = "LH"
                reachMap = 43
            Case "117"
                ' GV Hyrule Side
                exit2label = "GV Hyrule Side"
                reachMap = 44
            Case "22D", "3A0", "3D0"
                ' GV Gerudo Side
                exit2label = "GV Gerudo Side"
                reachMap = 45
            Case "129", "3A8"
                ' GF Main
                exit2label = "GF"
                reachMap = 46
            Case "3AC"
                ' GF Behind Gate
                exit2label = "GF Behind Gate"
                reachMap = 47
            Case "130"
                ' HW Gerudo Side
                exit2label = "HW Gerudo Side"
                reachMap = 48
            Case "365"
                ' HW Colossus Side
                exit2label = "HW Colossus Side"
                reachMap = 49
            Case "123", "1F1", "1E1", "57C", "588"
                ' DC Main
                exit2label = "DC"
                reachMap = 50
            Case "1E5"
                ' DC Statue Hand Right
                exit2label = "DC Statue Hand Right"
                reachMap = 50
            Case "1E9"
                ' DC Statue Hand Left
                exit2label = "DC Statue Hand left"
                reachMap = 50
            Case "4C2"
                ' Schrodinger's Fairy Adult
                exit2label = "OGC Fairy"
                reachMap = 51
            Case "578"
                ' Schrodinger's Fairy Young
                exit2label = "HC Fairy"
                reachMap = 13
            Case "000"
                ' Deku Tree
                exit2label = "Deku Tree"
                reachMap = 60
            Case "40F"
                ' Queen Gohma
                exit2label = "Queen Gohma"
            Case "004"
                ' Dodongo's Cavern
                exit2label = "Dodongo's Cavern"
                reachMap = 69
            Case "40B"
                ' King Dodongo
                exit2label = "King Dodongo"
            Case "028"
                ' Jabu-Jabu's Belly
                exit2label = "Jabu-Jabu's Belly"
                reachMap = 79
            Case "301"
                ' Barinade
                exit2label = "Barinade"
            Case "169"
                ' Forest Temple
                exit2label = "Forest Temple"
                reachMap = 88
            Case "00C"
                ' Phantom Ganon
                exit2label = "Phantom Ganon"
            Case "165"
                ' Fire Temple
                exit2label = "Fire Temple"
                reachMap = 108
            Case "305"
                ' Volvagia
                exit2label = "Volvagia"
            Case "010"
                ' Water Temple
                exit2label = "Water Temple"
                reachMap = 119
            Case "417"
                ' Morpha
                exit2label = "Morpha"
            Case "082", "3F0", "3F4"
                ' Spirit Temple
                exit2label = "Spirit Temple"
                reachMap = 132
            Case "08D"
                ' Twinrova
                exit2label = "Twinrova"
            Case "037"
                ' Shadow Temple
                exit2label = "Shadow Temple"
                reachMap = 150
            Case "413"
                ' Bongo Bongo
                exit2label = "Bongo Bongo"
            Case "098"
                ' Bottom of the Well
                exit2label = "BotW"
                reachMap = 174
            Case "088"
                ' Ice Cavern
                exit2label = "Ice Cavern"
                reachMap = 178
            Case "008"
                ' Gerudo Training Ground
                exit2label = "GTG"
                reachMap = 179
            Case "467", "534"
                ' Ganon's Castle
                exit2label = "Ganon's Castle"
            Case "41B"
                ' Ganon's Tower
                exit2label = "Ganon's Tower"
            Case "41F"
                ' Ganondorf
                exit2label = "Ganondorf"
            Case Else
                ' Unknown entry
                exit2label = exitCode & "?"
        End Select
    End Function
    Private Function zone2map(ByVal zone As Byte) As Byte
        ' Converts zone into map array for ER usage. Default 255 as null
        zone2map = 255

        Select Case zone
            Case 0, 1
                zone2map = 10   ' KF
            Case 2 To 4, 8
                zone2map = 16   ' LW
            Case 5, 6
                zone2map = 11   ' SFM
            Case 7
                zone2map = 6    ' HF
            Case 9
                zone2map = 24   ' LLR
            Case 197
                zone2map = 0    ' MK Entrance
            Case 10
                If isAdult Then
                    zone2map = 3    ' MK Adult
                Else
                    zone2map = 2    ' MK Young
                End If
            Case 198, 199
                zone2map = 1    ' MK Back Alley
            Case 11
                zone2map = 4    ' Outside ToT
            Case 12, 200
                zone2map = 5    ' ToT
            Case 13
                zone2map = 20   ' HC
            Case 15 To 17
                zone2map = 7    ' KV
            Case 18, 19
                zone2map = 8    ' GY
            Case 20, 21
                zone2map = 21   ' DMT
            Case 22 To 28
                zone2map = 22   ' DMC
            Case 29 To 33
                zone2map = 23   ' GC
            Case 34 To 36
                zone2map = 9    ' ZR
            Case 37 To 39
                zone2map = 13   ' ZD
            Case 40, 41
                zone2map = 14   ' ZF
            Case 42, 43
                zone2map = 12   ' LH
            Case 44, 45, 56, 57
                zone2map = 15   ' GV
            Case 46, 47
                zone2map = 18   ' GF
            Case 48, 49, 55
                zone2map = 19   ' HW
            Case 50, 54
                zone2map = 17   ' DC
            Case 51, 53
                zone2map = 25   ' OGC
            Case 60 To 68
                zone2map = 26   ' Deku Tree
            Case 69 To 78
                zone2map = 27   ' Dodongo's Cavern
            Case 79 To 87
                zone2map = 28   ' Jabu-Jabu's Belly
            Case 88 To 107
                zone2map = 29   ' Forest Temple
            Case 108 To 118
                zone2map = 30   ' Fire Temple
            Case 119 To 131
                zone2map = 31   ' Water Temple
            Case 132 To 149
                zone2map = 32   ' Spirit Temple
            Case 150 To 173
                zone2map = 33   ' Shadow Temple
            Case 174 To 177
                zone2map = 34   ' Bottom of the Well
            Case 178, 201 To 204
                zone2map = 35   ' Ice Cavern
            Case 179 To 192
                zone2map = 36   ' Gerudo Training Ground
            Case 193 To 196
                zone2map = 37   ' Ganon's Castle and Tower
        End Select
    End Function
    ' Scan each of the chests data
    Private Sub readChestData()
        If emulator = String.Empty Then
            isSoH = False
            If IS_64BIT = False Then
                attachToProject64()
                If emulator = String.Empty Then attachToM64PY()
            Else
                attachToBizHawk()
                If emulator = String.Empty Then attachToRMG()
                If emulator = String.Empty Then attachToM64P()
                If emulator = String.Empty Then attachToRetroArch()
                If emulator = String.Empty Then attachToModLoader64()
                If emulator = String.Empty Then attachToSoH()
            End If
            If Not emulator = String.Empty Then
                Me.Text = "Tracker of Time v" & VER & " (" & emulator & ")"
                Select Case LCase(emulator)
                    Case "emuhawk", "rmg", "mupen64plus-gui", "retroarch - mupen64plus", "retroarch - parallel", "modloader64-gui", "soh"
                        emulator = "variousX64"
                End Select
            End If
        End If
        If emulator = String.Empty Then Exit Sub
        If isLoadedGame() = False Then Exit Sub

        ' Get current room code
        Dim roomCode As String = String.Empty
        Dim tempVar As Integer = 0
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If

        Dim locationCode As Integer = CInt(IIf(isSoH, GDATA(&H200, 1), goRead(CUR_ROOM_ADDR + 2, 15)))

        'Me.Text = locationCode.ToString
        Dim doMath As Integer = 0
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If

        Dim chestCheck As Long = 0
        Dim foundChests As Double = 0
        Dim compareTo As Double = 0
        Dim strI As String = String.Empty
        Dim strII As String = String.Empty
        Dim doCheck As Boolean = False
        Dim checkAgain As Boolean = False

        For i = 0 To arrLocation.Length - 1
            ' 118 is only for SoH
            If i = 118 And Not isSoH Then Exit For
            If i Mod 5 = 0 Then Application.DoEvents()
            checkAgain = True
            Select Case i
                Case 0 To 59, Is >= 100
                    ' These are the area checks, either chest, standing items, area events, as they will need to be checked as they happen

                    doMath = (locationCode * 28) + &H11A6A4
                    If isSoH Then doMath = doMath + &HDADF94
                    tempVar = CInt(IIf(isSoH, &H2388, &H1CA1D8))

                    Dim doFlip As Boolean = False
                    Select Case i
                        Case 0 To 2, 103 To 113, 115 To 117
                            ' Scene Checks
                            inc(doMath, 4)
                            ' -0x10 from default
                            dec(tempVar, 16)
                        Case 3 To 30, 100 To 102, 114
                            ' Standing Checks
                            inc(doMath, 12)
                            inc(tempVar, 12)
                    End Select

                    If doMath = arrLocation(i) Then
                        checkAgain = False
                        If isSoH Then
                            chestCheck = GDATA(tempVar)
                        Else
                            chestCheck = goRead(tempVar, arrHigh(i))
                        End If

                        If Not keepRunning Then
                            stopScanning()
                            Exit Sub
                        End If
                        If Not chestCheck = lastRoomScan Then
                            arrChests(i) = chestCheck
                            parseChestData(i)
                        End If
                        lastRoomScan = chestCheck
                    End If
                Case 78 To 83
                    ' If skulltula check is off, do not check again
                    Select Case My.Settings.setSkulltula
                        Case 0
                            checkAgain = False
                        Case 1
                            checkAgain = True
                        Case Else
                            checkAgain = CBool(IIf(goldSkulltulas < 50, True, False))
                    End Select
                Case 84 To 98
                    ' If scrub shuffle check is off, do not check again
                    checkAgain = My.Settings.setScrub
            End Select

            If checkAgain Then
                If Not keepRunning Then
                    stopScanning()
                    Exit Sub
                End If

                If isSoH Then
                    chestCheck = goRead(arrLocation(i))
                Else
                    chestCheck = goRead(arrLocation(i), arrHigh(i))
                End If

                If isSoH Then
                    Select Case i
                        Case 61 To 70
                            Dim tempHex As String = Hex(chestCheck)
                            fixHex(tempHex)
                            tempHex = Mid(tempHex, 5) & Mid(tempHex, 1, 4)
                            chestCheck = CUInt("&H" & tempHex)
                    End Select
                End If

                If Not chestCheck = arrChests(i) Then
                    arrChests(i) = chestCheck
                    parseChestData(i)
                End If
            End If
        Next
        scanSingleChecks()

        ' Check to see if Bonooru's flag is flipped, this is for when a player has a free scarecrow song
        ' If it is, then force Pierre's check so that it does not show up
        If checkLoc("6512") And Not checkLoc("7316") Then
            flipKeyForced("6512", "LH", 2)
        End If

        If rtbRefresh = 0 Then
            Dim newArea As String = String.Empty
            Select Case locationCode
                Case 0 To 9
                    ' Dungeons from DT to IC
                    newArea = "DUN" & locationCode.ToString
                Case 11
                    ' Gerudo Training Grounds
                    newArea = "DUN10"
                Case 10, 13
                    ' Ganon's Castle 
                    newArea = "DUN11"
                Case 17 To 24
                    ' Bosses from DT to ShT
                    newArea = "BOSS" & (locationCode - 17).ToString
                Case 25
                    ' Ganondorf
                    newArea = "BOSS8"
                Case 79
                    ' Ganon
                    newArea = "BOSS9"
                Case 26
                    ' Falling Castle
                    newArea = "ESCAPE"
                Case 38 To 41, 45, 52, 85
                    ' Kakori Forest
                    newArea = "KF"
                Case 91
                    ' Lost Woods
                    newArea = "LW"
                Case 86
                    ' Sacred Forest Meadow
                    newArea = "SFM"
                Case 81
                    ' Hyrule Field
                    newArea = "HF"
                Case 54, 76, 99
                    ' Lon Lon Ranch
                    newArea = "LLR"
                Case 16, 27, 28, 30 To 33, 43, 44, 49 To 51, 53, 75, 77
                    ' The Market
                    newArea = "MK"
                Case 35 To 37, 67
                    ' Temple of Time
                    newArea = "TT"
                Case 69, 74, 95
                    ' Hyrule Castle
                    newArea = "HC"
                Case 42, 48, 55, 78, 80, 82
                    ' Kakakori Village
                    newArea = "KV"
                Case 58, 63 To 65, 83
                    ' Graveyard
                    newArea = "GY"
                Case 96
                    ' Death Mountain Trail
                    newArea = "DMT"
                Case 97
                    ' Death Mountain Crater
                    newArea = "DMC"
                Case 46, 98
                    ' Goron City
                    newArea = "GC"
                Case 84
                    ' Zora's River
                    newArea = "ZR"
                Case 47, 88
                    ' Zora's Domain
                    newArea = "ZD"
                Case 89
                    ' Zora's Fountain
                    newArea = "ZF"
                Case 56, 73, 87
                    ' Lake Hylia
                    newArea = "LH"
                Case 57, 90
                    ' Gerudo Valley
                    newArea = "GV"
                Case 12, 93
                    ' Gerudo Fortress
                    newArea = "GF"
                Case 94
                    ' Haunted Wasteland
                    newArea = "HW"
                Case 92
                    ' Desert Colossus
                    newArea = "DC"
                Case 29, 34, 100
                    ' Outside Ganon's Castle
                    newArea = "OGC"
                Case 62
                    newArea = "GROTTO"
                Case 72
                    newArea = "GRAVE"
                Case Else
                    newArea = "OTHER"
            End Select

            If Not newArea = lastArea Then
                'changeSong(newArea)
                Select Case newArea
                    Case "GROTTO", "GRAVE", "OTHER"
                        newArea = lastArea
                    Case "BOSS8", "BOSS9", "ESCAPE"
                        newArea = "DUN11"
                    Case Else
                        newArea = newArea.Replace("BOSS", "DUN")
                End Select
                lastArea = newArea
            End If

            If Not lastArea = String.Empty Then
                If Mid(lastArea, 1, 3) = "DUN" Then
                    displayChecksDungeons(CByte(Mid(lastArea, 4)), False, False)
                Else
                    displayChecks(lastArea, False, False)
                End If
            End If
        Else
            decB(rtbRefresh)
        End If
    End Sub
    Private Sub scanSingleChecks()
        Dim arrSingles(9) As Integer

        If isSoH Then
            arrSingles(0) = SAV(&HAD0)  ' &HEC9030
            arrSingles(1) = SAV(&HAEC)  ' &HEC904C
            arrSingles(2) = SAV(&HB0C)  ' &HEC906C
            arrSingles(3) = SAV(&HB78)  ' &HEC90D8
            arrSingles(4) = SAV(&H9F0)  ' &HEC8F50
            arrSingles(5) = SAV(&HB08)  ' &HEC9068
            arrSingles(6) = 0
            arrSingles(7) = SAV(&HEE4)  ' &HEC9444
            arrSingles(8) = SAV(&HF08)  ' &HEC9468
        Else
            arrSingles(0) = &H11B09C
            arrSingles(1) = &H11B0B8
            arrSingles(2) = &H11B128
            arrSingles(3) = &H11B144
            arrSingles(4) = &H11AFBC
            arrSingles(5) = &H11B0D4
            arrSingles(6) = &H11B150
            arrSingles(7) = &H11B4B0
            arrSingles(8) = &H11B4D4
        End If

        ' LW Bean Planted
        If My.Settings.setSkulltula > 0 And My.Settings.setGSLoc >= 1 Then setLoc("B0", checkBit(arrSingles(0), 22))
        ' DC Bean Planted
        setLoc("B1", checkBit(arrSingles(1), 24))
        ' DMT Bean Planted
        setLoc("B2", checkBit(arrSingles(2), 6))
        ' DMC Bean Planted
        setLoc("B3", checkBit(arrSingles(3), 3))
        ' GY Bean Planted
        setLoc("B4", checkBit(arrSingles(4), 3))

        ' GV Opened Gate to Haunted Wasteland
        setLoc("C00", checkBit(arrSingles(5), 3))
        ' DMC Deku Near Ladder
        If My.Settings.setScrub Then
            setLoc("C01", checkBit(arrSingles(6), 6))
        End If
        ' EV: KV Well Drained
        setLoc("C02", checkBit(arrSingles(7), 23))
        ' EV: LH Restored
        setLoc("C04", checkBit(arrSingles(7), 25))
        ' Deliver Zelda's Letter | Unlock Mask Shoppe
        setLoc("C05", checkBit(arrSingles(8), 6))

        ' Bombchu's in Logic setting
        Dim updateSetting As Boolean = False
        If Not aAddresses(7) = 0 Then
            If goRead(aAddresses(7), 1) = 1 Then updateSetting = True
            If Not My.Settings.setBombchus = updateSetting Then
                My.Settings.setBombchus = updateSetting
                updateSettingsPanel()
            End If
        End If

        ' Cow Shuffle setting
        If Not aAddresses(8) = 0 Then
            updateSetting = False
            If goRead(aAddresses(8), 1) = 1 Then updateSetting = True
            If Not My.Settings.setCow = updateSetting Then
                My.Settings.setCow = updateSetting
                updateSettingsPanel()
            End If
        End If

        ' Scrub Shuffle setting
        If Not aAddresses(10) = 0 Then
            updateSetting = False
            If goRead(aAddresses(10), 1) = 1 Then updateSetting = True
            If Not My.Settings.setScrub = updateSetting Then
                My.Settings.setScrub = updateSetting
                updateSettingsPanel()
            End If
        End If

        ' Small Keys setting
        If Not aAddresses(11) = 0 Then
            Dim tempRead As Byte = CByte(goRead(aAddresses(11), 1))
            ' Autotracker info uses 2 for keysanity and 1 for removed, I use 1 for keysanity and 2 for removed. 
            ' This flips those two so it fits into the tracker correctly
            Select Case tempRead
                Case 1
                    tempRead = 2
                Case 2
                    tempRead = 1
            End Select
            If Not My.Settings.setSmallKeys = tempRead Then
                My.Settings.setSmallKeys = tempRead
                updateLTB("ltbKeys")
                updateSmallKeys()
            End If
        End If
        ' Kokiri Forest setting
        If Not aAddresses(15) = 0 Then
            Dim tempRead As Byte = CByte(goRead(aAddresses(15), 1))
            ' 0 and 1 are 'Open' and 'Closed Deku', both resulting in the KF being open for out settings, since the tracker auto-detects Mido at the Deku Tree
            ' 2 is 'Closed' and the only one where the kid by the exit is there
            Select Case tempRead
                Case 0, 1
                    updateSetting = True
                Case Else
                    updateSetting = False
            End Select
            If Not My.Settings.setOpenKF = updateSetting Then
                My.Settings.setOpenKF = updateSetting
                updateSettingsPanel()
            End If
        End If

        ' Zora's Fountain setting
        If Not aAddresses(16) = 0 Then
            Dim tempRead As Byte = CByte(goRead(aAddresses(16), 1))
            ' 2 is 'Closed' and the only one where adult Link cannot get through
            ' 0 and 1 are 'Open' and 'Adult', both resulting in ZF being open for adult
            Select Case tempRead
                Case 2
                    updateSetting = False
                Case Else
                    updateSetting = True
            End Select
            If Not My.Settings.setOpenZF = updateSetting Then
                My.Settings.setOpenZF = updateSetting
                updateSettingsPanel()
            End If
        End If
    End Sub
    Private Sub parseChestData(ByVal loc As Integer)
        Dim foundChests As Double = arrChests(loc)
        Dim compareTo As Double = 2147483648
        Dim strI As String = loc.ToString
        Dim strII As String = String.Empty
        Dim doCheck As Boolean = False
        Dim gotHit As Boolean = False
        Dim startI As Byte = 31

        If Not isSoH Then
            Select Case arrHigh(loc)
                Case 0 To 7
                    startI = 7
                    compareTo = 128
                Case 8 To 15
                    startI = 15
                    compareTo = 32768
                Case Else
                    startI = 31
                    compareTo = 2147483648
            End Select
        End If

        If foundChests < 0 Then foundChests = foundChests + (compareTo * 2)
        For i = startI To 0 Step -1
            strII = i.ToString
            doCheck = False
            gotHit = False

            While strII.Length < 2
                strII = "0" & strII
            End While
            If foundChests >= compareTo And foundChests > 0 Then
                foundChests = foundChests - compareTo
                doCheck = True
            End If

            For Each key In aKeys.Where(Function(k As keyCheck) k.loc.Equals(strI & strII))
                With key
                    .checked = doCheck
                    Select Case .loc
                        Case "6500", "6501", "6502", "6503"
                            checkCarpenters()
                    End Select
                    If .area = "INV" Then data2vars(.loc, .checked)
                    gotHit = True
                End With
            Next

            If gotHit = True Then
                Select Case strI
                    Case "74", "75"
                        checkEquipment()
                    Case "76"
                        checkUpgrades()
                End Select
            Else
                ' Else, keep searching
                For ii = 0 To 11
                    For j = 0 To aKeysDungeons(ii).Length - 1
                        With aKeysDungeons(ii)(j)
                            If .loc = strI & strII Then
                                .checked = doCheck
                                gotHit = True
                            End If
                        End With
                        If gotHit Then Exit For
                    Next
                    If gotHit Then Exit For
                Next
            End If

            If i <= arrLow(loc) Then Exit Sub
            compareTo = compareTo / 2
        Next
    End Sub

    Private Sub data2vars(ByVal var As String, ByVal val As Boolean)
        ' 7508 is not an array, just a single variable
        If var = "7508" Then
            knifeCheck = val
            Exit Sub
        End If

        ' Grab the bit offset as the array location
        Dim bit As Byte = CByte(Mid(var, 3))
        Dim thisArray() As Boolean

        ' Depending on the arrLocation used, select the appropriate array
        Select Case Mid(var, 1, 2)
            Case "74"
                thisArray = aEquipment
            Case "76"
                thisArray = aUpgrades
            Case "77"
                thisArray = aQuestItems
            Case Else
                Exit Sub
        End Select

        ' Assign value
        thisArray(bit) = val
    End Sub

    Private Sub updateLabels()
        Dim aCheck(24) As Integer
        Dim aTotal(24) As Integer
        Dim outputLabel(24) As String
        Dim countCheck As Boolean = False
        Dim aBoldLabels(24) As Boolean

        For i = 0 To aTotal.Length - 1
            aTotal(i) = 0
            aCheck(i) = 0
            outputLabel(i) = String.Empty
            aBoldLabels(i) = False
        Next

        For i = 0 To aKeys.Length - 1
            With aKeys(i)
                If .scan = True And Not .area = "EVENT" Then
                    countCheck = False
                    If .gs Then
                        If My.Settings.setGSLoc >= 1 Then
                            Select Case My.Settings.setSkulltula
                                Case 0
                                    countCheck = False
                                Case 1
                                    countCheck = True
                                Case Else
                                    countCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                            End Select
                        End If
                    ElseIf .cow Then
                        If My.Settings.setCow Then countCheck = True
                    ElseIf .scrub Then
                        If My.Settings.setScrub Then countCheck = True
                    ElseIf .shop Then
                        If My.Settings.setShop > 0 Then countCheck = True
                    Else
                        countCheck = True
                    End If

                    If countCheck Then
                        inc(aTotal(area2num(.area)))
                        If .checked Or .forced Then inc(aCheck(area2num(.area)))

                        ' If logic setting is on and area is not yet set to bold, check if there is a logic check, and if so, set it to bold
                        If My.Settings.setLogic And Not aBoldLabels(area2num(.area)) Then
                            If Not .checked And Not .forced Then
                                ' TESTLOGIC If checkLogic(.logic, .zone) Then aBoldLabels(area2num(.area)) = True
                                If Not checkLogic(.logic, .zone) = 0 Then aBoldLabels(area2num(.area)) = True
                            End If
                        End If
                    End If
                End If
            End With
        Next

        Dim workLabel As Label = zKF
        If My.Settings.setMap Then
            ' Connect The Market, Masks, and Temple of Time together
            inc(aCheck(5), aCheck(23))
            inc(aCheck(5), aCheck(6))
            inc(aTotal(5), aTotal(23))
            inc(aTotal(5), aTotal(6))
            If aBoldLabels(23) Or aBoldLabels(6) Then aBoldLabels(5) = True

            ' Connect Hyrule Castle and Outside Ganon's Castle together
            inc(aCheck(7), aCheck(21))
            inc(aTotal(7), aTotal(21))
            If aBoldLabels(21) Then aBoldLabels(7) = True

            ' Connect Kakariko Village and Gold Skulltulas together
            inc(aCheck(8), aCheck(22))
            inc(aTotal(8), aTotal(22))
            If aBoldLabels(22) Then aBoldLabels(8) = True
            For i = 0 To 20
                Select Case i
                    Case 0
                        workLabel = zKF
                    Case 1
                        workLabel = zLW
                    Case 2
                        workLabel = zSFM
                    Case 3
                        workLabel = zHF
                    Case 4
                        workLabel = zLLR
                    Case 5, 6
                        workLabel = zMK
                    Case 7, 21
                        workLabel = zCastle
                    Case 8
                        workLabel = zKV
                    Case 9
                        workLabel = zGY
                    Case 10
                        workLabel = zDMT
                    Case 11
                        workLabel = zDMC
                    Case 12
                        workLabel = zGC
                    Case 13
                        workLabel = zZR
                    Case 14
                        workLabel = zZD
                    Case 15
                        workLabel = zZF
                    Case 16
                        workLabel = zLH
                    Case 17
                        workLabel = zGV
                    Case 18
                        workLabel = zGF
                    Case 19
                        workLabel = zHW
                    Case 20
                        workLabel = zDC
                End Select

                If aCheck(i) = aTotal(i) Then
                    workLabel.BackColor = Color.Transparent
                Else
                    workLabel.Visible = True
                    If aBoldLabels(i) Then
                        workLabel.BackColor = Color.Lime
                    Else
                        workLabel.BackColor = Color.Red
                    End If

                End If
                If i = 5 Then inc(i)
            Next
        Else
            For i = 0 To 23
                If i >= 22 Then
                    ' 22, 23, 24, and 25 are Quests, add the quest title and grab the second part of the split
                    If aoLabels(i).Text.Contains(":") Then outputLabel(i) = "Quest:" & aoLabels(i).Text.Split(CChar(":"))(1) & ": "
                Else
                    ' All others, just use the first grab of the split
                    If aoLabels(i).Text.Contains(":") Then outputLabel(i) = aoLabels(i).Text.Split(CChar(":"))(0) & ": "
                End If

                outputLabel(i) = outputLabel(i) & aCheck(i).ToString & "/" & aTotal(i).ToString
            Next

            Dim normalFont As New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            Dim boldFont As New Font("Microsoft Sans Serif", 8, FontStyle.Bold)


            If Not lblKokiriForest.Text = outputLabel(0) Then lblKokiriForest.Text = outputLabel(0)
            lblKokiriForest.Font = CType(IIf(aBoldLabels(0), boldFont, normalFont), Drawing.Font)
            If Not lblLostWoods.Text = outputLabel(1) Then lblLostWoods.Text = outputLabel(1)
            lblLostWoods.Font = CType(IIf(aBoldLabels(1), boldFont, normalFont), Drawing.Font)
            If Not lblSacredForestMeadow.Text = outputLabel(2) Then lblSacredForestMeadow.Text = outputLabel(2)
            lblSacredForestMeadow.Font = CType(IIf(aBoldLabels(2), boldFont, normalFont), Drawing.Font)
            If Not lblHyruleField.Text = outputLabel(3) Then lblHyruleField.Text = outputLabel(3)
            lblHyruleField.Font = CType(IIf(aBoldLabels(3), boldFont, normalFont), Drawing.Font)
            If Not lblLonLonRanch.Text = outputLabel(4) Then lblLonLonRanch.Text = outputLabel(4)
            lblLonLonRanch.Font = CType(IIf(aBoldLabels(4), boldFont, normalFont), Drawing.Font)
            If Not lblMarket.Text = outputLabel(5) Then lblMarket.Text = outputLabel(5)
            lblMarket.Font = CType(IIf(aBoldLabels(5), boldFont, normalFont), Drawing.Font)
            If Not lblTempleOfTime.Text = outputLabel(6) Then lblTempleOfTime.Text = outputLabel(6)
            lblTempleOfTime.Font = CType(IIf(aBoldLabels(6), boldFont, normalFont), Drawing.Font)
            If Not lblHyruleCastle.Text = outputLabel(7) Then lblHyruleCastle.Text = outputLabel(7)
            lblHyruleCastle.Font = CType(IIf(aBoldLabels(7), boldFont, normalFont), Drawing.Font)
            If Not lblKakarikoVillage.Text = outputLabel(8) Then lblKakarikoVillage.Text = outputLabel(8)
            lblKakarikoVillage.Font = CType(IIf(aBoldLabels(8), boldFont, normalFont), Drawing.Font)
            If Not lblGraveyard.Text = outputLabel(9) Then lblGraveyard.Text = outputLabel(9)
            lblGraveyard.Font = CType(IIf(aBoldLabels(9), boldFont, normalFont), Drawing.Font)
            If Not lblDMTrail.Text = outputLabel(10) Then lblDMTrail.Text = outputLabel(10)
            lblDMTrail.Font = CType(IIf(aBoldLabels(10), boldFont, normalFont), Drawing.Font)
            If Not lblDMCrater.Text = outputLabel(11) Then lblDMCrater.Text = outputLabel(11)
            lblDMCrater.Font = CType(IIf(aBoldLabels(11), boldFont, normalFont), Drawing.Font)
            If Not lblGoronCity.Text = outputLabel(12) Then lblGoronCity.Text = outputLabel(12)
            lblGoronCity.Font = CType(IIf(aBoldLabels(12), boldFont, normalFont), Drawing.Font)
            If Not lblZorasRiver.Text = outputLabel(13) Then lblZorasRiver.Text = outputLabel(13)
            lblZorasRiver.Font = CType(IIf(aBoldLabels(13), boldFont, normalFont), Drawing.Font)
            If Not lblZorasDomain.Text = outputLabel(14) Then lblZorasDomain.Text = outputLabel(14)
            lblZorasDomain.Font = CType(IIf(aBoldLabels(14), boldFont, normalFont), Drawing.Font)
            If Not lblZorasFountain.Text = outputLabel(15) Then lblZorasFountain.Text = outputLabel(15)
            lblZorasFountain.Font = CType(IIf(aBoldLabels(15), boldFont, normalFont), Drawing.Font)
            If Not lblLakeHylia.Text = outputLabel(16) Then lblLakeHylia.Text = outputLabel(16)
            lblLakeHylia.Font = CType(IIf(aBoldLabels(16), boldFont, normalFont), Drawing.Font)
            If Not lblGerudoValley.Text = outputLabel(17) Then lblGerudoValley.Text = outputLabel(17)
            lblGerudoValley.Font = CType(IIf(aBoldLabels(17), boldFont, normalFont), Drawing.Font)
            If Not lblGerudoFortress.Text = outputLabel(18) Then lblGerudoFortress.Text = outputLabel(18)
            lblGerudoFortress.Font = CType(IIf(aBoldLabels(18), boldFont, normalFont), Drawing.Font)
            If Not lblHauntedWasteland.Text = outputLabel(19) Then lblHauntedWasteland.Text = outputLabel(19)
            lblHauntedWasteland.Font = CType(IIf(aBoldLabels(19), boldFont, normalFont), Drawing.Font)
            If Not lblDesertColossus.Text = outputLabel(20) Then lblDesertColossus.Text = outputLabel(20)
            lblDesertColossus.Font = CType(IIf(aBoldLabels(20), boldFont, normalFont), Drawing.Font)
            If Not lblOutsideGanonsCastle.Text = outputLabel(21) Then lblOutsideGanonsCastle.Text = outputLabel(21)
            lblOutsideGanonsCastle.Font = CType(IIf(aBoldLabels(21), boldFont, normalFont), Drawing.Font)
            If Not lblQuestGoldSkulltulas.Text = outputLabel(22) Then lblQuestGoldSkulltulas.Text = outputLabel(22)
            lblQuestGoldSkulltulas.Font = CType(IIf(aBoldLabels(22), boldFont, normalFont), Drawing.Font)
            If Not lblQuestMasks.Text = outputLabel(23) Then lblQuestMasks.Text = outputLabel(23)
            lblQuestMasks.Font = CType(IIf(aBoldLabels(23), boldFont, normalFont), Drawing.Font)
        End If

    End Sub
    Private Sub updateLabelsDungeons()
        Dim aTotal(11) As Integer
        Dim aCheck(11) As Integer
        Dim outputLabel(11) As String
        Dim countCheck As Boolean = False
        Dim aBoldLabels(11) As Boolean

        For i = 0 To aTotal.Length - 1
            aTotal(i) = 0
            aCheck(i) = 0
            outputLabel(i) = String.Empty
            aBoldLabels(i) = False

            For ii = 0 To aKeysDungeons(i).Length - 1
                With aKeysDungeons(i)(ii)
                    If .scan = True And Not .area = "EVENT" Then
                        countCheck = False
                        If .gs Then
                            If My.Settings.setGSLoc <= 1 Then
                                Select Case My.Settings.setSkulltula
                                    Case 0
                                        countCheck = False
                                    Case 1
                                        countCheck = True
                                    Case Else
                                        countCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                                End Select
                            End If
                        ElseIf .cow Then
                            If My.Settings.setCow Then countCheck = True
                        ElseIf .scrub Then
                            If My.Settings.setScrub Then countCheck = True
                        Else
                            countCheck = True
                        End If
                        If countCheck Then
                            inc(aTotal(i))
                            If .checked Or .forced Then inc(aCheck(i))

                            ' If logic setting is on and area is not yet set to bold, check if there is a logic check, and if so, set it to bold
                            If My.Settings.setLogic And Not aBoldLabels(i) Then
                                If Not .checked Then
                                    ' TESTLOGIC If checkLogic(.logic, 99, .area) Then aBoldLabels(i) = True
                                    If Not checkLogic(.logic, .zone) = 0 Then aBoldLabels(i) = True
                                End If
                            End If
                        End If
                    End If
                End With
            Next
        Next
        Dim workLabel As Label = zKF
        If My.Settings.setMap Then
            For i = 0 To 11
                Select Case i
                    Case 0
                        workLabel = zDT
                    Case 1
                        workLabel = zDDC
                    Case 2
                        workLabel = zJB
                    Case 3
                        workLabel = zFoT
                    Case 4
                        workLabel = zFiT
                    Case 5
                        workLabel = zWaT
                    Case 6
                        workLabel = zSpT
                    Case 7
                        workLabel = zShT
                    Case 8
                        workLabel = zBotW
                    Case 9
                        workLabel = zIC
                    Case 10
                        workLabel = zGTG
                    Case 11
                        workLabel = zIGC
                End Select

                If aCheck(i) = aTotal(i) Then
                    workLabel.BackColor = Color.Transparent
                Else
                    workLabel.Visible = True
                    If aBoldLabels(i) Then
                        workLabel.BackColor = Color.Lime
                        If canDungeon(i) Then
                            workLabel.CreateGraphics.DrawRectangle(New Pen(Color.Yellow, 2), 2, 2, workLabel.Width - 4, workLabel.Height - 4)
                            workLabel.CreateGraphics.DrawRectangle(New Pen(Color.White, 1), 3, 3, workLabel.Width - 7, workLabel.Height - 7)
                        End If
                    Else
                        workLabel.BackColor = Color.Red
                    End If
                End If
            Next
        Else
            For i = 0 To aTotal.Length - 1
                outputLabel(i) = aoDungeonLabels(i).Text.Split(CChar(":"))(0) & ": "
                outputLabel(i) = outputLabel(i) & aCheck(i).ToString & "/" & aTotal(i).ToString
            Next

            Dim normalFont As New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            Dim boldFont As New Font("Microsoft Sans Serif", 8, FontStyle.Bold)

            If Not lblDekuTree.Text = outputLabel(0) Then lblDekuTree.Text = outputLabel(0)
            lblDekuTree.Font = CType(IIf(aBoldLabels(0), boldFont, normalFont), Drawing.Font)
            If Not lblDodongosCavern.Text = outputLabel(1) Then lblDodongosCavern.Text = outputLabel(1)
            lblDodongosCavern.Font = CType(IIf(aBoldLabels(1), boldFont, normalFont), Drawing.Font)
            If Not lblJabuJabusBelly.Text = outputLabel(2) Then lblJabuJabusBelly.Text = outputLabel(2)
            lblJabuJabusBelly.Font = CType(IIf(aBoldLabels(2), boldFont, normalFont), Drawing.Font)
            If Not lblForestTemple.Text = outputLabel(3) Then lblForestTemple.Text = outputLabel(3)
            lblForestTemple.Font = CType(IIf(aBoldLabels(3), boldFont, normalFont), Drawing.Font)
            If Not lblFireTemple.Text = outputLabel(4) Then lblFireTemple.Text = outputLabel(4)
            lblFireTemple.Font = CType(IIf(aBoldLabels(4), boldFont, normalFont), Drawing.Font)
            If Not lblWaterTemple.Text = outputLabel(5) Then lblWaterTemple.Text = outputLabel(5)
            lblWaterTemple.Font = CType(IIf(aBoldLabels(5), boldFont, normalFont), Drawing.Font)
            If Not lblSpiritTemple.Text = outputLabel(6) Then lblSpiritTemple.Text = outputLabel(6)
            lblSpiritTemple.Font = CType(IIf(aBoldLabels(6), boldFont, normalFont), Drawing.Font)
            If Not lblShadowTemple.Text = outputLabel(7) Then lblShadowTemple.Text = outputLabel(7)
            lblShadowTemple.Font = CType(IIf(aBoldLabels(7), boldFont, normalFont), Drawing.Font)
            If Not lblBottomOfTheWell.Text = outputLabel(8) Then lblBottomOfTheWell.Text = outputLabel(8)
            lblBottomOfTheWell.Font = CType(IIf(aBoldLabels(8), boldFont, normalFont), Drawing.Font)
            If Not lblIceCavern.Text = outputLabel(9) Then lblIceCavern.Text = outputLabel(9)
            lblIceCavern.Font = CType(IIf(aBoldLabels(9), boldFont, normalFont), Drawing.Font)
            If Not lblGerudoTrainingGround.Text = outputLabel(10) Then lblGerudoTrainingGround.Text = outputLabel(10)
            lblGerudoTrainingGround.Font = CType(IIf(aBoldLabels(10), boldFont, normalFont), Drawing.Font)
            If Not lblGanonsCastle.Text = outputLabel(11) Then lblGanonsCastle.Text = outputLabel(11)
            lblGanonsCastle.Font = CType(IIf(aBoldLabels(11), boldFont, normalFont), Drawing.Font)
        End If
    End Sub

    Private Function area2num(ByVal area As String) As Integer
        ' Converts the area code into a numeric value. Starts at 100 to cause error for unexpected codes
        area2num = 100
        Select Case area
            Case "KF"
                Return 0
            Case "LW"
                Return 1
            Case "SFM"
                Return 2
            Case "HF", "QBPH"
                Return 3
            Case "LLR"
                Return 4
            Case "MK"
                Return 5
            Case "TT"
                Return 6
            Case "HC"
                Return 7
            Case "KV"
                Return 8
            Case "GY"
                Return 9
            Case "DMT"
                Return 10
            Case "DMC"
                Return 11
            Case "GC"
                Return 12
            Case "ZR", "QF"
                Return 13
            Case "ZD"
                Return 14
            Case "ZF"
                Return 15
            Case "LH"
                Return 16
            Case "GV"
                Return 17
            Case "GF"
                Return 18
            Case "HW"
                Return 19
            Case "DC"
                Return 20
            Case "OGC"
                Return 21
            Case "QGS"
                Return 22
            Case "QM"
                Return 23
            Case "INV", "EVENT", ""
                Return 24
        End Select
    End Function
    Private Function area2code(ByVal area As String) As String
        ' Converts the area name into a area code
        area2code = String.Empty
        Select Case area
            Case "Kokiri Forest"
                Return "KF"
            Case "Lost Woods"
                Return "LW"
            Case "Sacred Forest Meadow"
                Return "SFM"
            Case "Hyrule Field"
                Return "HF"
            Case "Lon Lon Ranch"
                Return "LLR"
            Case "Market", "Market & Temple of Time"
                Return "MK"
            Case "Temple of Time"
                Return "TT"
            Case "Hyrule Castle", "Hyrule Castle & Outside Ganon's Castle"
                Return "HC"
            Case "Kakariko Village"
                Return "KV"
            Case "Graveyard"
                Return "GY"
            Case "Death Mountain Trail"
                Return "DMT"
            Case "Death Mountain Crater"
                Return "DMC"
            Case "Goron City"
                Return "GC"
            Case "Zora's River"
                Return "ZR"
            Case "Zora's Domain"
                Return "ZD"
            Case "Zora's Fountain"
                Return "ZF"
            Case "Lake Hylia"
                Return "LH"
            Case "Gerudo Valley"
                Return "GV"
            Case "Gerudo Fortress"
                Return "GF"
            Case "Haunted Wasteland"
                Return "HW"
            Case "Desert Colossus"
                Return "DC"
            Case "Outside Ganon's Castle"
                Return "OGC"
            Case "Quest Big Poe Hunt"
                Return "QBPH"
            Case "Quest Frogs"
                Return "QF"
            Case "Quest Gold Skulltulas"
                Return "QGS"
            Case "Quest Masks"
                Return "QM"
            Case "Deku Tree"
                Return "DUN0"
            Case "Dodongo's Cavern"
                Return "DUN1"
            Case "Jabu-Jabu's Belly"
                Return "DUN2"
            Case "Forest Temple"
                Return "DUN3"
            Case "Fire Temple"
                Return "DUN4"
            Case "Water Temple"
                Return "DUN5"
            Case "Spirit Temple"
                Return "DUN6"
            Case "Shadow Temple"
                Return "DUN7"
            Case "Bottom of the Well"
                Return "DUN8"
            Case "Ice Cavern"
                Return "DUN9"
            Case "Gerudo Training Ground"
                Return "DUN10"
            Case "Ganon's Castle"
                Return "DUN11"
        End Select
    End Function
    Private Function dungeonNumber2name(ByVal dunNum As Byte) As String
        dungeonNumber2name = String.Empty
        Select Case dunNum
            Case 0
                dungeonNumber2name = "Deku Tree"
            Case 1
                dungeonNumber2name = "Dodongo's Cavern"
            Case 2
                dungeonNumber2name = "Jabu-Jabu's Belly"
            Case 3
                dungeonNumber2name = "Forest Temple"
            Case 4
                dungeonNumber2name = "Fire Temple"
            Case 5
                dungeonNumber2name = "Water Temple"
            Case 6
                dungeonNumber2name = "Spirit Temple"
            Case 7
                dungeonNumber2name = "Shadow Temple"
            Case 8
                dungeonNumber2name = "Bottom of the Well"
            Case 9
                dungeonNumber2name = "Ice Cavern"
            Case 10
                dungeonNumber2name = "Gerudo Training Ground"
            Case 11
                dungeonNumber2name = "Ganon's Castle"
        End Select
    End Function

    Private Sub stopScanning()
        keepRunning = False
        tmrAutoScan.Enabled = False
        tmrFastScan.Enabled = False
        AutoScanToolStripMenuItem.Text = "Auto Scan"
        Me.Text = "Tracker of Time v" & VER
        emulator = String.Empty
    End Sub

    Private Function goRead(ByVal offsetAddress As Integer, Optional bitType As Byte = 31) As Integer
        goRead = 0
        Select Case emulator
            Case String.Empty
                Exit Function
            Case "variousX64"
                Try
                    Select Case bitType
                        Case 0 To 7
                            goRead = ReadMemory(Of Byte)(romAddrStart64 + offsetAddress)
                        Case 8 To 15
                            goRead = ReadMemory(Of Int16)(romAddrStart64 + offsetAddress)
                        Case Else
                            goRead = ReadMemory(Of Integer)(romAddrStart64 + offsetAddress)
                    End Select
                Catch ex As Exception
                    stopScanning()
                    If Not ex.Message = "External component has thrown an exception." Then
                        MessageBox.Show("goRead Problem: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End Try
            Case Else
                If IS_64BIT Then Exit Function
                Select Case bitType
                    Case 0 To 7
                        goRead = quickRead8(romAddrStart + offsetAddress, emulator)
                    Case 8 To 15
                        goRead = quickRead16(romAddrStart + offsetAddress, emulator)
                    Case Else
                        goRead = quickRead32(romAddrStart + offsetAddress, emulator)
                End Select
        End Select
    End Function
    Private Function quickRead8(ByVal readAddress As Integer, ByVal sTarget As String) As Integer
        quickRead8 = 0

        Try
            quickRead8 = Memory.ReadInt8(p, readAddress)
        Catch ex As Exception
            stopScanning()
            MessageBox.Show("quickRead Problem: " & vbCrLf & ex.Message & vbCrLf & readAddress.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function
    Private Function quickRead16(ByVal readAddress As Integer, ByVal sTarget As String) As Integer
        quickRead16 = 0

        Try
            quickRead16 = Memory.ReadInt16(p, readAddress)
        Catch ex As Exception
            stopScanning()
            MessageBox.Show("quickRead Problem: " & vbCrLf & ex.Message & vbCrLf & readAddress.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function
    Private Function quickRead32(ByVal readAddress As Integer, ByVal sTarget As String, Optional doStopScanning As Boolean = True) As Integer
        quickRead32 = 0

        Try
            quickRead32 = Memory.ReadInt32(p, readAddress)
        Catch ex As Exception
            If doStopScanning Then stopScanning()
            MessageBox.Show("quickRead Problem: " & vbCrLf & ex.Message & vbCrLf & readAddress.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function
    Private Sub quickWrite16(ByVal writeAddress As Integer, ByVal writeValue As Int16, ByVal sTarget As String)
        writeAddress = romAddrStart + writeAddress

        Try
            WriteInt16(p, writeAddress, writeValue)
        Catch ex As Exception
            MessageBox.Show("quickWrite Problem: " & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    Private Sub quickWrite(ByVal writeAddress As Integer, ByVal writeValue As Integer, ByVal sTarget As String)
        Try
            WriteInt32(p, writeAddress, writeValue)
        Catch ex As Exception
            MessageBox.Show("quickWrite Problem: " & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub tmrAutoScan_Tick(sender As Object, e As EventArgs) Handles tmrAutoScan.Tick
        ' Timer is ran every 3 seconds. Often enough for accuracy but not so much that it bogs things down. May Lower to 5 seconds.

        ' Checks that things should still be running and ready
        If Not keepRunning Or emulator = String.Empty Then
            stopScanning()
            Exit Sub
        End If

        ' Runs a check for "ZELDAZ" in the memory where it is expected
        If checkZeldaz() = 0 Then
            ' On fail, count up the fails, allowing for 5 fails (15 seconds) before disconnecting
            inc(zeldazFails)
            If zeldazFails >= 5 Then
                stopScanning()
                Exit Sub
            End If
        Else
            ' Only do a scan with the zeldaCheck was successful, and reset the fail counter
            zeldazFails = 0
            getAge()
            If My.Settings.setNavi Then shutupNavi()
            checkMQs()
            ' Do not bother scanning the all the checks if we are focused on the ER panel
            updateEverything()
            readChestData()
            If Not keepRunning Then
                stopScanning()
                Exit Sub
            End If
        End If
    End Sub
    Private Sub goScan(ByVal auto As Boolean)
        zeldazFails = 0
        keepRunning = True
        updateFast()
        updateEverything()
        readChestData()
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If
        getAge()
        checkMQs()
        updateFast()
        updateEverything()
        readChestData()
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If
        If Not auto Then Exit Sub
        If tmrAutoScan.Enabled = False Then
            tmrFastScan.Enabled = True
            tmrAutoScan.Enabled = True
            ' checkMQs()
        Else
            stopScanning()
        End If
    End Sub

    Private Function checkZeldaz() As Byte
        If isSoH Then Return 2

        ' Checks for the 'ZELDAZ' within the memory to make sure you are playing Ocarina of Time, and that it is still reading the correct memory region
        checkZeldaz = 0
        Dim zeldaz1 As Integer = goRead(&H11A5EC)
        Dim zeldaz2 As Integer = goRead(&H11A5F0 + 2, 15)

        ' If both checks are zero, and menu screen varaible is there, set to 1 for half-true
        If zeldaz1 = 0 And zeldaz2 = 0 Then
            ' Checks if on the game menu screen
            If goRead(&H11B92C, 1) = 2 Then checkZeldaz = 1
        End If

        ' Check both and if they are proper, 2 for full-true
        If zeldaz1 = 1514490948 And zeldaz2 = 16730 Then checkZeldaz = 2

        ' Check for the rando version only if a full-true
        If checkZeldaz = 2 Then getRandoVer()
    End Function
    Private Function isLoadedGame() As Boolean
        Dim addrLoaded As Integer = &H11B92C
        If isSoH Then addrLoaded = soh.SAV(&H1320)
        ' Checks the game state (2=game menu, 1=title screen, 0=gameplay), if 0 and a successful ZELDAZ check, then true
        isLoadedGame = False
        If goRead(addrLoaded, 1) = 0 And checkZeldaz() = 2 Then
            isLoadedGame = True
        Else
            If isSoH Then
                lastFirstEnum = 255 ' SOH is loaded but we're on the main menu, so clear the rando settings
                ResetToolStripMenuItem_Click(Nothing, Nothing)
            End If
        End If
    End Function
    Private Sub debugInfo()
        emulator = String.Empty
        If IS_64BIT = False Then
            attachToProject64()
            If emulator = String.Empty Then attachToM64PY()
        Else
            attachToBizHawk()
            If emulator = String.Empty Then attachToRMG()
            If emulator = String.Empty Then attachToM64P()
            If emulator = String.Empty Then attachToRetroArch()
            If emulator = String.Empty Then attachToModLoader64()
            If emulator = String.Empty Then attachToSoH()
        End If
        If Not emulator = String.Empty Then
            Me.Text = "Tracker of Time v" & VER & " (" & emulator & ")"
            Select Case LCase(emulator)
                Case "emuhawk", "rmg", "mupen64plus-gui", "retroarch - mupen64plus", "retroarch - parallel", "modloader64-gui"
                    emulator = "variousX64"
            End Select
        End If
        If emulator = String.Empty Then Exit Sub
        getRandoVer()
        rtbOutputLeft.Clear()
        rtbOutputLeft.AppendText("Attached to " & emulator & vbCrLf & "Starting address: 0x" & Hex(CInt(IIf(IS_64BIT, romAddrStart64, romAddrStart))) & vbCrLf & "Randomizer Version: " & randoVer & vbCrLf & vbCrLf)
        scanEmulator(emulator)
        Dim zeldaz1 As Integer = goRead(&H11A5EC)
        Dim zeldaz2 As Integer = goRead(&H11A5F0 + 2, 15)
        rtbOutputLeft.AppendText("ZELDAZ check: " & Hex(zeldaz1) & Hex(zeldaz2) & vbCrLf & "Game State: " & goRead(&H11B92C, 1).ToString & vbCrLf & vbCrLf)
    End Sub
    Private Sub scanEmulator(Optional emuName As String = "rmg")
        Dim target As Process = Nothing
        Try
            target = Process.GetProcessesByName(emuName)(0)
            rtbOutputLeft.AppendText(target.ProcessName & vbCrLf)
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                rtbOutputLeft.AppendText(emuName & " not found!" & vbCrLf)
                'MessageBox.Show("BizHawk not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                rtbOutputLeft.AppendText("Problem: " & ex.Message & vbCrLf)
            End If
            Return
        End Try
        Dim addressDLL As Int64 = 0
        For Each mo As ProcessModule In target.Modules
            'If LCase(mo.ModuleName) = "mupen64plus.dll" Then
            'addressDLL = mo.BaseAddress.ToInt64
            'Exit For
            'End If
            rtbOutputLeft.AppendText(mo.ModuleName & ":" & Hex(mo.BaseAddress.ToInt64) & vbCrLf)
        Next
        rtbOutputLeft.AppendText(vbCrLf)
    End Sub
    Private Sub checkMQs()
        If Not isLoadedGame() Then Exit Sub

        ' soh 3.0.0 has no MQ dungeons
        If isSoH Then Exit Sub

        For i = 0 To aMQ.Length - 1
            aMQ(i) = False
        Next

        If aAddresses(0) = 0 Then
            updateMQs()
            Exit Sub
        End If

        ' Update the MQ Dungeons
        Dim MQs As String = String.Empty
        Dim tempRead As String = String.Empty
        Dim startAt As Byte = 1

        For i = 0 To 12 Step 4
            tempRead = Hex(goRead(aAddresses(0) + i))
            fixHex(tempRead)
            MQs = MQs & tempRead
        Next
        If Mid(MQs, 1, 4) = "FFFF" Then startAt = 5
        MQs = Mid(MQs, startAt, 28)

        If Not MQs.Replace("0", "").Replace("1", "") = "" Then Exit Sub
        Dim addExtra As Byte = 1
        For i = 0 To 11
            If i > 9 Then incB(addExtra, 2)
            Select Case Mid(MQs, (i * 2) + addExtra, 2)
                Case "00"
                    aMQ(i) = False
                Case "01"
                    aMQ(i) = True
            End Select
        Next
        updateMQs()
    End Sub
    Private Sub updateEverything()
        If checkZeldaz() = 2 And isLoadedGame() Then
            If isSoH Then getSoHRandoSettings()
            getWarps()
            getRainbowBridge()
            changeScrubs()
            'updateItems()
            'updateQuestItems()
            'updateDungeonItems()
            If Not pnlER.Visible Then
                updateLabels()
                updateLabelsDungeons()
            Else
                tmrAutoScan.Interval = 5000
                tmrFastScan.Interval = 1000
                'updateMiniMap()
            End If
        End If
    End Sub

    Private Sub updateFast()
        If checkZeldaz() = 2 And isLoadedGame() Then
            getER()
            updateItems()
            updateQuestItems()
            updateDungeonItems()
            If pnlER.Visible Then updateMiniMap()
        End If
    End Sub

    Private Function checkLoc(ByVal cloc As String) As Boolean
        If firstRun Then Return False
        ' Checks for a specific key by location to see if it is checked
        checkLoc = False

        ' Checks normal keys
        For Each key In aKeys.Where(Function(k As keyCheck) k.loc.Equals(cloc))
            Return key.checked
        Next

        ' Checks dungeon keys
        For i = 0 To 11
            For Each key In aKeysDungeons(i).Where(Function(k As keyCheck) k.loc.Equals(cloc))
                Return key.checked
            Next
        Next
    End Function
    Private Sub setLoc(ByVal loc As String, ByVal checked As Boolean, Optional isDungeon As Boolean = False)
        ' Sets a specific key's checked status
        If isDungeon Then
            ' Checks dungeon keys
            For i = 0 To 11
                For Each key In aKeysDungeons(i).Where(Function(k As keyCheck) k.loc.Equals(loc))
                    key.checked = checked
                Next
            Next
        Else
            ' Checks normal keys
            For Each key In aKeys.Where(Function(k As keyCheck) k.loc.Equals(loc))
                key.checked = checked
            Next
        End If
    End Sub
    Private Function name2loc(ByVal keyName As String, ByVal keyArea As String) As String
        name2loc = "0"
        If firstRun Then Exit Function
        ' Checks for a specific key by name to get the loc

        ' Checks normal keys only because this is used only for shops
        For Each key In aKeys.Where(Function(k As keyCheck) k.name.Equals(keyName))
            If key.area = keyArea Then Return key.loc
        Next
    End Function
    Private Sub shutupNavi()
        ' With great power, comes little care for what others have to day. Shut up Navi's timed complaints.

        Dim addrNavi As Integer = &H11A60A
        Select Case emulator
            Case String.Empty
                Exit Sub
            Case "variousX64"
                If isSoH Then addrNavi = SAV(&H32)
                WriteMemory(Of Int16)(romAddrStart64 + addrNavi, 0)
            Case Else
                quickWrite16(addrNavi, 0, emulator)
        End Select
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        ' Move the trial icon locations
        'pbxTrialSpirit.Height = pbxMap.Height + 216
        'pbxTrialForest.Height = pbxMap.Height + 216
        'pbxTrialFire.Height = pbxMap.Height
        'pbxTrialShadow.Height = pbxMap.Height
        'pbxTrialLight.Height = pbxMap.Height + 108
        'pbxTrialWater.Height = pbxMap.Height + 108

        'scanEmulator("modloader64-gui")
        'pnlER.Visible = Not pnlER.Visible
        'goScan(False)
        'rtbOutputLeft.Clear()
        'Dim linkRot As Double = ((goRead(&H1DAA74, 15) / 65535 * 360) - 90) * -1

        'Me.Text = linkRot.ToString
        'MsgBox(Hex(goRead(&H400CEB, 1)))
        'MsgBox(My.Settings.setSmallKeys.ToString)
        'For Each i As Integer In aAddresses
        'rtbAddLine(Hex(i))
        'Next
        'debugInfo()

        '        MsgBox(Hex(goRead(&HEC85FE)))
        'MsgBox(checkLoc("11806").ToString)
        'Dim test As Integer = goRead(arrLocation(118))
        'MsgBox(Hex(test))
        'dump()
        Dim test As String = Hex(goRead(arrLocation(121)))
        For Each k In aKeys
            If k.loc = "12102" Then MsgBox(Hex(arrLocation(121)) & ": " & test & ": " & checkLoc("12102").ToString)
        Next


        If False Then
            Dim outputXX As String = "Visited:"
            For i = 0 To aVisited.Length - 1
                outputXX = outputXX & vbCrLf & i.ToString & ": " & aVisited(i).ToString
            Next
            outputXX = outputXX & vbCrLf & vbCrLf
            For i = 0 To aExitMap.Length - 1
                For ii = 0 To aExitMap(i).Length - 1
                    outputXX = outputXX & vbCrLf & "aExitMap(" & i.ToString & ")(" & ii.ToString & "): " & aExitMap(i)(ii).ToString
                Next
            Next
            Clipboard.SetText(outputXX)
        End If


        If False Then
            Dim outputXX As String = String.Empty
            outputXX = "Adult:"
            For i = 0 To aReachA.Length - 1
                outputXX = outputXX & vbCrLf & i.ToString & ": " & aReachA(i).ToString
            Next
            outputXX = outputXX & vbCrLf & vbCrLf & "Young:"
            For i = 0 To aReachY.Length - 1
                outputXX = outputXX & vbCrLf & i.ToString & ": " & aReachY(i).ToString
            Next
            Clipboard.SetText(outputXX)
        End If

        If False Then
            Dim text2 As String = String.Empty
            For i = 0 To arrLocation.Length - 1
                text = text & "frmTrackerOfTime.arrLocation(" & i.ToString & ") = SAV(&H" & Hex(arrLocation(i) - &HEC8560) & ")" & vbCrLf
            Next
            Clipboard.SetText(text)
        End If
    End Sub
    Private Sub changeTheme(Optional theme As Byte = 0)
        Dim cBack As Color = Control.DefaultBackColor
        Dim cFore As Color = Control.DefaultForeColor
        Dim cpnlBack As Color = Control.DefaultBackColor
        Dim cpnlFore As Color = Control.DefaultForeColor
        Select Case theme
            Case 0  ' Default
                cBack = Color.WhiteSmoke
                cFore = Color.Black
            Case 1  ' Dark Mode
                cBack = Color.Black
                cFore = Color.White
            Case 2  ' Lavender 
                cBack = Color.Lavender
                cFore = Color.Purple
            Case 3  ' Midnight
                cBack = Color.Black
                cFore = Color.RoyalBlue
            Case 4  ' Hotdog Stand
                cBack = Color.Yellow
                cFore = Color.Red
            Case 5  ' The Hub
                cBack = Color.FromArgb(27, 27, 27)
                cFore = Color.FromArgb(255, 163, 26)
            Case Else
                Exit Sub
        End Select

        Me.BackColor = cBack
        Me.ForeColor = cFore

        Dim cbR1 As Integer = CInt((CInt(Me.BackColor.R) + CInt(Me.ForeColor.R)) / 2)
        Dim cbG1 As Integer = CInt((CInt(Me.BackColor.G) + CInt(Me.ForeColor.G)) / 2)
        Dim cbB1 As Integer = CInt((CInt(Me.BackColor.B) + CInt(Me.ForeColor.B)) / 2)
        cBlend = Color.FromArgb(cbR1, cbG1, cbB1)

        rtbOutputLeft.BackColor = cBack
        rtbOutputLeft.ForeColor = cFore
        rtbOutputRight.BackColor = cBack
        rtbOutputRight.ForeColor = cFore

        mnuOptions.BackColor = cBack
        mnuOptions.ForeColor = cFore
        custMenu.highlight = cBlend
        custMenu.backColour = cBack
        custMenu.foreColour = cFore

        My.Settings.setTheme = theme
        My.Settings.Save()

        For Each lbl In pnlSettings.Controls.OfType(Of Label)()
            lbl.Refresh()
        Next
        'updateSettingsPanel()
    End Sub
    Private Sub resizeForm()
        btnTest.Visible = False
        Button2.Visible = False

        Dim setHeight As Integer = 990
        If My.Settings.setShortForm Then
            pnlMain.VerticalScroll.Visible = True
            setHeight = 688
            pnlSettings.Left = 578 + 33
            pnlSettings.Width = 597 + 33
        Else
            pnlSettings.Left = 578
            pnlSettings.Width = 597
            pnlMain.VerticalScroll.Visible = False
            If My.Settings.setMap Then inc(setHeight, 6)
            If My.Settings.setExpand Then inc(setHeight, 39)
            If My.Settings.setMap And My.Settings.setExpand Then inc(setHeight, 18)
        End If
        pnlMain.Width = 564 + CByte(IIf(My.Settings.setShortForm, 33, 0))
        If My.Settings.setExpand Then
            pnlWorldMap.Height = 318
            rtbLines = 14
            rtbOutputLeft.Height = 196
            rtbOutputRight.Height = 196
            pnlER.Height = 516
            'pbxMap.Top = 49
        Else
            rtbLines = 11
            pnlWorldMap.Height = 300
            rtbOutputLeft.Height = 157
            rtbOutputRight.Height = 157
            pnlER.Height = 492
            'pbxMap.Top = 37
        End If
        pnlMain.Height = setHeight
        pnlSettings.VerticalScroll.Visible = pnlMain.VerticalScroll.Visible
        pnlSettings.Height = setHeight
        If My.Settings.setMap Then
            rtbOutputLeft.Top = pnlWorldMap.Top + pnlWorldMap.Height + 1
            rtbOutputRight.Top = rtbOutputLeft.Top
        Else
            rtbOutputLeft.Top = pnlDMCrater.Top + pnlDMCrater.Height
            rtbOutputRight.Top = rtbOutputLeft.Top
        End If
        pnlWorldMap.Visible = My.Settings.setMap

        Application.DoEvents()
        If showSetting Then
            Me.Width = pnlMain.Width + 627 + CByte(IIf(My.Settings.setShortForm, 33, 0))
        Else
            Me.Width = pnlMain.Width + 26
        End If
        Me.Height = pnlMain.Height + 52 '+ CByte(IIf(My.Settings.setMap, 0, 1))
        redrawOutputBoarder()
        If My.Settings.setShortForm Then ' pnlMain.VerticalScroll.Value >= pnlMain.VerticalScroll.Maximum - pnlMain.Height Then
            Me.CreateGraphics.DrawLine(New Pen(Me.ForeColor, 1), 6, Me.Height - 45, rtbOutputLeft.Width + 7, Me.Height - 45)
        Else
            Me.CreateGraphics.DrawLine(New Pen(Me.BackColor, 1), 6, Me.Height - 45, rtbOutputLeft.Width + 7, Me.Height - 45)
        End If
    End Sub
    Private Sub redrawOutputBoarder()
        Dim gfx As Graphics = Me.CreateGraphics
        gfx = pnlMain.CreateGraphics
        With rtbOutputLeft
            gfx.DrawRectangle(New Pen(Me.BackColor, 1), .Location.X - 1, .Location.Y - 1, .Width + 1, .Height + 20)
            gfx.DrawRectangle(New Pen(Me.ForeColor, 1), .Location.X - 1, .Location.Y - 1, .Width + 1, .Height + 1)
        End With
    End Sub



    Private Sub displayChecks(ByVal area As String, ByVal showChecked As Boolean, Optional setCounter As Boolean = True)
        'getWarps()
        Dim displayName As String = String.Empty
        Dim sOut As String = String.Empty

        ' Counter set to delay refreshing the output
        If setCounter Then rtbRefresh = 2

        Select Case area
            Case "KF"
                displayName = "Kokiri Forest"
            Case "LW"
                displayName = "Lost Woods"
            Case "SFM"
                displayName = "Sacred Forest Meadow"
            Case "HF"
                displayName = "Hyrule Field"
            Case "LLR"
                displayName = "Lon Lon Ranch"
            Case "MK"
                If My.Settings.setMap Then
                    displayName = "Market & Temple of Time"
                Else
                    displayName = "Market"
                End If
            Case "TT"
                If My.Settings.setMap Then
                    displayName = "Market & Temple of Time"
                Else
                    displayName = "Temple of Time"
                End If
            Case "HC"
                If My.Settings.setMap Then
                    displayName = "Hyrule Castle & Outside Ganon's Castle"
                Else
                    displayName = "Hyrule Castle"
                End If
            Case "KV"
                displayName = "Kakariko Village"
            Case "GY"
                displayName = "Graveyard"
            Case "DMT"
                displayName = "Death Mountain Trail"
            Case "DMC"
                displayName = "Death Mountain Crater"
            Case "GC"
                displayName = "Goron City"
            Case "ZR"
                displayName = "Zora's River"
            Case "ZD"
                displayName = "Zora's Domain"
            Case "ZF"
                displayName = "Zora's Fountain"
            Case "LH"
                displayName = "Lake Hylia"
            Case "GV"
                displayName = "Gerudo Valley"
            Case "GF"
                displayName = "Gerudo Fortress"
            Case "HW"
                displayName = "Haunted Wasteland"
            Case "DC"
                displayName = "Desert Colossus"
            Case "OGC"
                If My.Settings.setMap Then
                    displayName = "Hyrule Castle & Outside Ganon's Castle"
                Else
                    displayName = "Outside Ganon's Castle"
                End If
            Case "QGS"
                displayName = "Quest: Gold Skulltulas"
            Case "QM"
                displayName = "Quest: Masks"
            Case Else
                displayName = "Bad Area Code: " & area
        End Select

        ' If right-clicked to show checked, not that it is found checks
        If showChecked Then displayName = displayName & " (Found)"

        ' Create a list and fill it with checks
        emboldenList.Clear()
        Dim outputLines As New List(Of String)
        scanArea(outputLines, area, showChecked)

        Select Case area
            Case "HF"
                Do While outputLines.Count < rtbLines
                    outputLines.Add(String.Empty)
                Loop
                scanArea(outputLines, "QBPH", showChecked)
            Case "ZR"
                Do While outputLines.Count < rtbLines
                    outputLines.Add(String.Empty)
                Loop
                scanArea(outputLines, "QF", showChecked)
            Case "MK", "TT"
                If My.Settings.setMap Then
                    emboldenList.Clear()
                    outputLines.Clear()
                    scanArea(outputLines, "MK", showChecked)
                    scanArea(outputLines, "QM", showChecked)
                    Do While outputLines.Count < rtbLines
                        outputLines.Add(String.Empty)
                    Loop
                    scanArea(outputLines, "TT", showChecked)
                End If
            Case "HC", "OGC"
                If My.Settings.setMap Then
                    emboldenList.Clear()
                    outputLines.Clear()
                    scanArea(outputLines, "HC", showChecked)
                    Do While outputLines.Count < rtbLines
                        outputLines.Add(String.Empty)
                    Loop
                    scanArea(outputLines, "OGC", showChecked)
                End If
            Case "KV"
                If My.Settings.setMap Then
                    scanArea(outputLines, "QGS", showChecked)
                    Do While outputLines.Count > (rtbLines * 2)
                        outputLines.RemoveAt(outputLines.Count - 1)
                    Loop
                End If
        End Select

        Dim theyLookSoGodDamnLikeTheSameList As Boolean = False

        If outputLines.Count = lastOutput.Count And outputLines.Count > 0 Then
            theyLookSoGodDamnLikeTheSameList = True
            For i = 0 To outputLines.Count - 1
                If Not outputLines(i) = lastOutput(i) Then
                    theyLookSoGodDamnLikeTheSameList = False
                    Exit For
                End If
            Next
        End If

        If Not theyLookSoGodDamnLikeTheSameList Then
            Select Case area
                Case "HF"
                    rtbOutputRight.Text = "Big Poes:" & IIf(showChecked, " (Found)", "").ToString
                Case "ZR"
                    rtbOutputRight.Text = "Fabulous Five Froggish Tenors:" & IIf(showChecked, " (Found)", "").ToString
                Case Else
                    rtbOutputRight.Clear()
            End Select

            ' Clear out the output boxes and set the display name
            rtbOutputLeft.Text = displayName & ":"

            ' Output each line
            For Each line In outputLines
                rtbAddLine(line)
            Next
        End If
        lastOutput = outputLines
        Application.DoEvents()
        ' If logic setting is set, and the embolden list is not empty
        If My.Settings.setLogic And emboldenList.Count > 0 Then
            ' Run each line through the emboldening process
            For Each line In emboldenList
                embolden(line)
            Next
        End If
    End Sub
    Private Sub scanArea(ByRef lines As List(Of String), ByVal area As String, ByVal showChecked As Boolean)
        ' Prefix for adding indent and things like GS (Gold Skulltula) for other options, suffix for if it is forced
        Dim prefix As String = String.Empty
        Dim suffix As String = String.Empty
        ' Confirmation to add the check to the output list or not
        Dim addCheck As Boolean = False
        ' If in compact mode, float checks to a readable location
        Dim floatChecks As New List(Of String)
        Dim doFloat As Boolean = False
        For i = 0 To aKeys.Length - 1
            With aKeys(i)
                ' Stop if an empty key
                If .loc = String.Empty Then Exit For
                If .area = area Then
                    ' Reset and determine the prefix
                    prefix = "  "
                    Select Case True
                        Case .gs
                            prefix = "  GS: "
                        Case .cow
                            prefix = "  Cow: "
                        Case .scrub
                            prefix = "  Scrub: "
                        Case .shop
                            prefix = "  Shopsanity: "
                    End Select
                    ' So long as it is not checked and in good standing, add it to the output
                    If .scan = True Then
                        If .checked = showChecked Or (showChecked And .forced) Then
                            addCheck = False
                            If .gs Then
                                If My.Settings.setGSLoc >= 1 Then
                                    Select Case My.Settings.setSkulltula
                                        Case 0
                                            addCheck = False
                                        Case 1
                                            addCheck = True
                                        Case Else
                                            addCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                                    End Select
                                End If
                            ElseIf .cow Then
                                addCheck = My.Settings.setCow
                            ElseIf .scrub Then
                                addCheck = My.Settings.setScrub
                            ElseIf .shop Then
                                If My.Settings.setShop > 0 Then addCheck = True
                            Else
                                addCheck = True
                            End If
                            ' Remove the forced checks from the unchecked list
                            If Not showChecked And .forced Then addCheck = False
                            If My.Settings.setHideQuests And .area = "QBPH" Then addCheck = False

                            ' Reset and determing suffix
                            suffix = ""
                            If .forced Then suffix = " (Forced)"

                            ' Do not bolden the checked list. This will still bolden the forced items in the check list
                            doFloat = False
                            If Not .checked Then
                                ' If logic is on, bold the ones that are accessable
                                Dim logicResult As Byte = checkLogic(.logic, .zone)
                                If My.Settings.setLogic And Not logicResult = 0 Then
                                    Select Case logicResult
                                        Case 1
                                            suffix = " (Y) " & suffix
                                        Case 2
                                            suffix = " (A) " & suffix
                                    End Select
                                    emboldenList.Add(prefix & .name & suffix)
                                    doFloat = True
                                End If
                            End If
                            If addCheck Then
                                If lines.Count < (rtbLines * 2) Then
                                    lines.Add(prefix & .name & suffix & Chr(32))
                                ElseIf doFloat Then
                                    floatChecks.Add(prefix & .name & suffix & Chr(32))
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
        If Not floatChecks.Count = 0 Then
            Dim doRemove As Boolean = True
            For i = lines.Count - 1 To 0 Step -1
                doRemove = True
                For j = 0 To emboldenList.Count - 1
                    If lines(i) = emboldenList(j) & Chr(32) Then doRemove = False
                Next
                If doRemove Then lines.RemoveAt(i)
                If lines.Count + floatChecks.Count <= rtbLines * 2 Then Exit For
            Next
            For Each line In floatChecks
                lines.Add(line)
            Next
        End If
    End Sub
    Private Sub displayChecksDungeons(ByVal dungeon As Byte, ByVal showChecked As Boolean, Optional setCounter As Boolean = True)
        Dim displayName As String = String.Empty
        Dim sOut As String = String.Empty

        ' Counter set to delay refreshing the output
        If setCounter Then rtbRefresh = 2

        displayName = dungeonNumber2name(dungeon)

        ' Add 'MQ' to the display if it is a Master Quest dungeon
        If aMQ(dungeon) Then displayName = displayName & " MQ"

        ' If right-clicked to show checked, not that it is found checks
        If showChecked Then displayName = displayName & " (Found)"

        ' Create a list and fill it with checks
        Dim outputLines As New List(Of String)
        emboldenList.Clear()
        scanDungeon(outputLines, dungeon, showChecked)
        Dim theyLookSoGodDamnLikeTheSameList As Boolean = False

        If outputLines.Count = lastOutput.Count And outputLines.Count > 0 Then
            theyLookSoGodDamnLikeTheSameList = True
            For i = 0 To outputLines.Count - 1
                If Not outputLines(i) = lastOutput(i) Then
                    theyLookSoGodDamnLikeTheSameList = False
                    Exit For
                End If
            Next
        End If

        If Not theyLookSoGodDamnLikeTheSameList Then
            ' Clear out the output boxes and set the display name
            rtbOutputLeft.Text = displayName & ":"
            rtbOutputRight.Clear()

            ' Output each line
            For Each line In outputLines
                rtbAddLine(line)
            Next
            lastOutput = outputLines
        End If
        Application.DoEvents()
        ' If logic setting is set, and the embolden list is not empty
        If My.Settings.setLogic And emboldenList.Count > 0 Then
            ' Run each line through the emboldening process
            For Each line In emboldenList
                embolden(line)
            Next
        End If
    End Sub
    Private Sub scanDungeon(ByRef lines As List(Of String), ByVal dungeon As Byte, ByVal showChecked As Boolean)
        ' Prefix for adding indent and things like GS (Gold Skulltula) for other options, suffix for if it is forced
        Dim prefix As String = String.Empty
        Dim suffix As String = String.Empty
        ' Confirmation to add the check to the output list or not
        Dim addCheck As Boolean = False

        For i = 0 To aKeysDungeons(dungeon).Length - 1
            With aKeysDungeons(dungeon)(i)
                ' Stop if an empty key or event key
                If .loc = String.Empty Or .area = "EVENT" Then Exit For
                ' Reset and determine the prefix
                prefix = "  "
                Select Case True
                    Case .gs
                        prefix = "  GS: "
                    Case .cow
                        prefix = "  Cow: "
                    Case .scrub
                        prefix = "  Scrub: "
                    Case .shop
                        prefix = "  Shopsanity: "
                End Select

                ' So long as it is not checked and in good standing, add it to the output
                If .scan = True Then
                    If .checked = showChecked Or (showChecked And .forced) Then
                        addCheck = False
                        If .gs Then
                            If My.Settings.setGSLoc <= 1 Then
                                Select Case My.Settings.setSkulltula
                                    Case 0
                                        addCheck = False
                                    Case 1
                                        addCheck = True
                                    Case Else
                                        addCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                                End Select
                            End If
                        ElseIf .cow Then
                            addCheck = My.Settings.setCow
                        ElseIf .scrub Then
                            addCheck = My.Settings.setScrub
                        ElseIf .shop Then
                            If My.Settings.setShop > 0 Then addCheck = True
                        Else
                            addCheck = True
                        End If
                        ' Remove the forced checks from the unchecked list
                        If Not showChecked And .forced Then addCheck = False

                        ' Reset and determing suffix
                        suffix = ""
                        If .forced Then suffix = " (Forced)"

                        ' Do not bolden the checked list. This will still bolden the forced items in the check list
                        If Not .checked Then
                            ' If logic is on, bold the ones that are accessable
                            ' TESTLOGIC If My.Settings.setLogic And checkLogic(.logic, 99, .area) Then
                            ' TESTLOGIC emboldenList.Add(prefix & .name & suffix)
                            ' TESTLOGIC End If
                            Dim logicResult As Byte = checkLogic(.logic, .zone)
                            If My.Settings.setLogic And Not logicResult = 0 Then
                                Select Case logicResult
                                    Case 1
                                        suffix = " (Y) " & suffix
                                    Case 2
                                        suffix = " (A) " & suffix
                                End Select
                                emboldenList.Add(prefix & .name & suffix)
                            End If
                        End If

                        ' Output the check and note if it is forced
                        If addCheck Then lines.Add(prefix & .name & suffix & Chr(32))
                    End If
                End If
            End With
        Next
    End Sub

    Private Sub setupKeys()
        ' tK is short for thisKey
        Dim tK As Integer = 0

        makeKeysKF(tK)
        makeKeysLW(tK)
        makeKeysSFM(tK)
        makeKeysHF(tK)
        makeKeysLLR(tK)
        makeKeysMK(tK)
        makeKeysTT(tK)
        makeKeysHC(tK)
        makeKeysKV(tK)
        makeKeysGY(tK)
        makeKeysDMT(tK)
        makeKeysDMC(tK)
        makeKeysGC(tK)
        makeKeysZR(tK)
        makeKeysZD(tK)
        makeKeysZF(tK)
        makeKeysLH(tK)
        makeKeysGV(tK)
        makeKeysGF(tK)
        makeKeysHW(tK)
        makeKeysDC(tK)
        makeKeysOGC(tK)
        makeKeysQBPH(tK)
        makeKeysQF(tK)
        makeKeysQGS(tK)
        makeKeysQM(tK)
        makeKeysInventory(tK)
        'updateShoppes()
        'makeKeysShoppes(tK)
        changeScrubs()
        changeSalesmans()
    End Sub
    Private Sub getHighLows()
        ' Scans all the keys to determing what is the largest bit size needed to read from an address, and when we can stop reading bits
        ' This helps reduce scan time

        Dim bLoc As Byte = 0
        Dim bVal As Byte = 0
        Dim strLoc As String = String.Empty
        For Each key In aKeys
            With key
                If IsNumeric(.loc) Then
                    strLoc = .loc
                    fixHex(strLoc, 5)
                    bLoc = CByte(Mid(strLoc, 1, 3))
                    If bLoc <= CHECK_COUNT Then
                        bVal = CByte(Mid(strLoc, 4, 2))
                        If bVal > arrHigh(bLoc) Then arrHigh(bLoc) = bVal
                        If bVal < arrLow(bLoc) Then arrLow(bLoc) = bVal
                    End If
                End If
            End With
        Next
        For i As Byte = 0 To 11
            For Each key In aKeysDungeons(i)
                With key
                    If IsNumeric(.loc) Then
                        strLoc = .loc
                        fixHex(strLoc, 5)
                        bLoc = CByte(Mid(strLoc, 1, 3))
                        bVal = CByte(Mid(strLoc, 4, 2))
                        If bVal > arrHigh(bLoc) Then
                            arrHigh(bLoc) = bVal
                        End If
                        If bVal < arrLow(bLoc) Then arrLow(bLoc) = bVal
                    End If
                End With
            Next
        Next
    End Sub

    Private Sub reachAreas(ByVal area As Byte, ByVal asAdult As Boolean)

        ' Make sure the area is set to prevent branching from unreachable areas
        If asAdult Then
            If Not aReachA(area) Then Return
        Else
            If Not aReachY(area) Then Return
        End If

        ' Bridge the past and the future together
        If aReachY(12) Then aReachA(12) = True
        If aReachA(12) Then aReachY(12) = True

        Select Case area
            Case 0
                ' KF Main to LW Front, KF Deku Tree, LW Between Bridge
                addAreaExit(10, 2, asAdult) 'addArea(2, asAdult)
                If asAdult Then
                    addArea(1, asAdult)
                    addAreaExit(10, 1, asAdult) 'addArea(8, asAdult)
                Else
                    If checkLoc("6220") Or (item("kokiri sword") And item("deku shield")) Then addArea(1, asAdult)
                    If My.Settings.setOpenKF Or bSongWarps Or bSpawnWarps Or checkLoc("6223") Or iER Mod 2 = 1 Then addAreaExit(10, 1, asAdult) 'addArea(8, asAdult)
                End If
            Case 1
                ' KF Deku Tree to Deku Tree Lobby
                If Not asAdult Or iER > 1 Then addAreaExit(10, 0, asAdult) 'addArea(60, asAdult)
            Case 2
                ' LW Front to KF Main, GC Shortcut, LW Behind Mido, ZR Main
                addAreaExit(16, 1, asAdult) 'addArea(0, asAdult)
                addAreaExit(16, 3, asAdult) 'addArea(30, asAdult)
                If asAdult Then
                    If item("saria's song") Or checkLoc("6226") Then addArea(3, asAdult)
                    If item("dive") Or item("iron boots") Then addAreaExit(16, 2, asAdult) 'addArea(35, asAdult)
                Else
                    addArea(3, asAdult)
                    If item("dive") Then addAreaExit(16, 2, asAdult) 'addArea(35, asAdult)
                End If
            Case 3
                ' LW Behind Mido to KF Main, SFM Main, LW Front
                addAreaExit(16, 1, asAdult) 'addArea(0, asAdult)
                addAreaExit(16, 0, asAdult) 'addArea(5, asAdult)
                If asAdult Then
                    If item("saria's song") Or checkLoc("6226") Then addArea(2, asAdult)
                Else
                    addArea(2, asAdult)
                End If
            Case 4
                ' LW Bridge to HF, KF Between Bridge, LW Front
                addAreaExit(16, 5, asAdult) 'addArea(7, asAdult)
                If asAdult Then
                    addAreaExit(16, 4, asAdult)
                    If item("longshot") Then addArea(2, asAdult)
                Else
                    If My.Settings.setOpenKF Or bSongWarps Or bSpawnWarps Or checkLoc("6223") Or iER Mod 2 = 1 Then addAreaExit(16, 4, asAdult) 'addArea(0, asAdult)
                End If
            Case 5
                ' SFM Main to LW Behind Mido, SFM Temple Ledge
                addAreaExit(11, 1, asAdult) 'addArea(3, asAdult)
                If asAdult And item("hookshot") Then addArea(6, asAdult)
            Case 6
                ' SFM Temple Ledge to Forest Temple
                addAreaExit(11, 0, asAdult) 'addArea(88, asAdult)
            Case 7
                ' HF to LW Bridge, ZR Front, KV Main, MK Entrance, LLR, LH Main, GV Hyrule Side, Big Poe Hunt
                addAreaExit(6, 1, asAdult) 'addArea(4, asAdult)
                addAreaExit(6, 2, asAdult) 'addArea(34, asAdult)
                addAreaExit(6, 0, asAdult) 'addArea(15, asAdult)
                addAreaExit(6, 4, asAdult) 'addArea(197, asAdult)
                addAreaExit(6, 5, asAdult) 'addArea(9, asAdult)
                addAreaExit(6, 6, asAdult) 'addArea(42, asAdult)
                addAreaExit(6, 3, asAdult) 'addArea(44, asAdult)
                If asAdult And item("epona") And item("bow") And item("bottle") Then addArea(59, asAdult)
            Case 8
                ' LW Between Bridge
                addArea(4, asAdult)
            Case 9
                'LLR to HF
                addAreaExit(24, 0, asAdult) 'addArea(7, asAdult)
            Case 10
                ' MK to MK Entrance, ToT Front, OGC, HC, MK Back Alley (left) and (right)
                If asAdult Then
                    addAreaExit(3, 1, asAdult) 'addArea(197, asAdult)
                    addAreaExit(3, 2, asAdult) 'addArea(11, asAdult)
                    addAreaExit(3, 0, asAdult) 'addArea(51, asAdult)
                Else
                    addAreaExit(2, 1, asAdult) 'addArea(197, asAdult)
                    addAreaExit(2, 2, asAdult) 'addArea(11, asAdult)
                    addAreaExit(2, 0, asAdult) 'addArea(13, asAdult)
                    addAreaExit(2, 4, asAdult) 'addArea(198, asAdult)
                    addAreaExit(2, 3, asAdult) 'addArea(199, asAdult)
                End If
            Case 11
                ' Outside ToT to MK, ToT Front
                addAreaExit(4, 1, asAdult) 'addArea(10, asAdult)
                addAreaExit(4, 0, asAdult) 'addArea(200, asAdult)
            Case 12
                ' ToT Behind Door to ToT Front, ToT Behind Door Both Ages
                If checkLoc("6427") Then addArea(200, asAdult)
                addArea(12, Not asAdult)
            Case 13
                ' HC to MK, GFF
                addAreaExit(20, 0, asAdult) 'addArea(10, asAdult)
                If Not asAdult And canExplode() Then addArea(54, asAdult)
            Case 14
                ' KV Windmill to KV Main
                addArea(15, asAdult)
            Case 15
                ' KV Main to HF, GY, KF Behind Gate, KV Rooftops, Bottom of the Well
                addAreaExit(7, 2, asAdult) 'addArea(7, asAdult)
                addAreaExit(7, 3, asAdult) 'addArea(18, asAdult)
                If asAdult Then
                    addArea(17, asAdult)
                    If item("hookshot") Or My.Settings.setHoverTricks Then addArea(16, asAdult)
                    If iER > 1 And (checkLoc("C02") Or item("song of storms")) Then addAreaExit(7, 0, asAdult)
                Else
                    If allItems.Contains("y3") Or checkLoc("C05") Then addArea(17, asAdult)
                    If aReachY(15) And (checkLoc("C02") Or item("song of storms")) Then addAreaExit(7, 0, asAdult) 'addArea(174, asAdult)
                End If
            Case 16
                ' KV Rooftops to KV Main
                addArea(15, asAdult)
            Case 17
                ' KV Behind Gate to KV Main, DMT Lower
                addArea(15, asAdult)
                addAreaExit(7, 1, asAdult) 'addArea(20, asAdult)
            Case 18
                ' GY Main to KV Main, KV Windmill
                addAreaExit(8, 1, asAdult) 'addArea(15, asAdult)
                If asAdult And item("song of time") Then addArea(14, asAdult)
            Case 19
                ' GY Upper to GY Main, Shadow Temple
                addArea(18, asAdult)
                If item("din's fire") Then addAreaExit(8, 0, asAdult) 'addArea(150, asAdult)
            Case 20
                ' DMT Lower to KV Behind Gate, GC Main, Dodongo's Cavern Lobby, DMT Upper
                addAreaExit(21, 2, asAdult) 'addArea(17, asAdult)
                addAreaExit(21, 1, asAdult) 'addArea(29, asAdult)
                If asAdult Then
                    If canBreakRocks() Or item("lift") Then addAreaExit(21, 0, asAdult) 'addArea(69, asAdult)
                    If canBreakRocks() Or (canMagicBean("B2") Or canHoverTricks()) Then addArea(21, asAdult)
                Else
                    If canExplode() Or item("lift") Then addAreaExit(21, 0, asAdult) 'addArea(69, asAdult)
                    If canExplode() Then addArea(21, asAdult)
                End If
            Case 21
                ' DMT Upper to DMT Lower, DMC Upper Local, KV Rooftops
                addArea(20, asAdult)
                addAreaExit(21, 3, asAdult) 'addArea(22, asAdult)
                If Not asAdult Then addArea(16, asAdult) ' Hoot hoot
            Case 22
                ' DMC Upper to DMT Upper, DMC Ladder, DMC Central Nearby
                addAreaExit(22, 2, asAdult) 'addArea(21, asAdult)
                addArea(23, asAdult)
                If asAdult And item("goron tunic") And item("longshot") Then addArea(26, asAdult)
            Case 23
                ' DMC Ladder to DMT Upper, DMC Upper, DMC Lower Nearby
                If asAdult Then
                    addAreaExit(22, 2, asAdult) 'addArea(21, asAdult)
                    If item("goron tunic") Then addArea(22, asAdult)
                    If item("hover boots") Then addArea(24, asAdult)
                End If
            Case 24
                ' DMC Lower Nearby to GC Darunias, DMC Lower Local
                addAreaExit(22, 1, asAdult) 'addArea(31, asAdult)
                If asAdult And item("goron tunic") Then addArea(25, asAdult)
            Case 25
                ' DMC Lower Local to DMC Ladder, DMC Lower Nearby, DMC Central Nearby, DMC Fire Temple Entrance
                addArea(23, asAdult)
                addArea(24, asAdult)
                If asAdult Then
                    If item("hover boots") Or item("hookshot") Then addArea(26, asAdult)
                    If (item("hover boots") Or item("hookshot")) And canFewerGoron() Then addArea(28, asAdult)
                End If
            Case 26
                ' DMC Central Nearby to DMC Central Local
                If asAdult And item("goron tunic") Then addArea(27, asAdult)
            Case 27
                ' DMC Central Local to DMC Central Nearby, DMC Local Nearby, DMT Upper, DMC Upper, DMC Fire Temple Entrance
                addArea(26, asAdult)
                If asAdult Then
                    If item("hover boots") Or item("hookshot") Then addArea(24, asAdult)
                    If canMagicBean("B3") Then
                        addArea(24, asAdult)
                        addAreaExit(22, 0, asAdult)
                        If item("goron tunic") Then addArea(22, asAdult)
                    End If
                    If canFewerGoron() Then addArea(28, asAdult)
                Else
                    If iER > 1 Then addArea(28, asAdult)
                End If
            Case 28
                ' DMC Fire Temple Entrance to DMC Fire Temple
                addAreaExit(22, 0, asAdult) 'addArea(108, asAdult)
            Case 29
                ' GC Main to DMT Lower, GC Shortcut, GC Darunia, GC Shoppe, GC Grotto Platform
                addAreaExit(23, 1, asAdult) 'addArea(20, asAdult)
                If asAdult Then
                    If canBreakRocks() Or canBurnAdult() Or item("lift") Or checkLoc("11508") Or checkLoc("11511") Or checkLoc("11512") Then addArea(30, asAdult)
                    If checkLoc("7025") Then
                        addArea(31, asAdult)
                        addArea(32, asAdult)
                    End If
                    If item("song of time") Or (item("hookshot") And (item("goron tunic") Or item("nayru's love"))) Then addArea(33, asAdult)
                Else
                    If canBreakRocks() Or item("din's fire") Or item("lift") Or checkLoc("11508") Or checkLoc("11511") Or checkLoc("11512") Then addArea(30, asAdult)
                    If item("zelda's lullaby") Or checkLoc("15527") Then addArea(31, asAdult)
                    If checkLoc("15501") Or canExplode() Or item("lift") Or item("din's fire") Or ((item("zelda's lullaby") Or checkLoc("15527") Or checkLoc("15528")) And item("deku stick")) Then addArea(32, asAdult)
                End If
            Case 30
                ' GC Shortcut to LW Front, GC Main
                addAreaExit(23, 2, asAdult) 'addArea(2, asAdult)
                If asAdult Then
                    If canBreakRocks() Or item("din's fire") Or checkLoc("11508") Or checkLoc("11511") Or checkLoc("11512") Then addArea(29, asAdult)
                Else
                    If canExplode() Or item("din's fire") Or checkLoc("11508") Or checkLoc("11511") Or checkLoc("11512") Then addArea(29, asAdult)
                End If
            Case 31
                ' GC Darunia to GC Main, DMC Lower Local
                addArea(29, asAdult)
                If asAdult Then addAreaExit(23, 0, asAdult) 'addArea(25, asAdult)
            Case 32
                ' GC Shoppe to GC Main
                addArea(29, asAdult)
            Case 33
                ' GC Grotto Platform to GC Main
                addArea(29, asAdult)
            Case 34
                ' ZR Front to HF, ZR Main
                addAreaExit(9, 0, asAdult) 'addArea(7, asAdult)
                If asAdult Then
                    addArea(35, asAdult)
                Else
                    If canExplode() Or checkLoc("11603") Or checkLoc("11608") Or checkLoc("11609") Then addArea(35, asAdult)
                End If
            Case 35
                ' ZR Main to ZR Front, ZR Behind Waterfall, LW Front, Frogs
                addArea(34, asAdult)
                If asAdult Then
                    If item("zelda's lullaby") Or canHoverTricks() Then addArea(36, asAdult)
                    If item("dive") Or item("iron boots") Then addAreaExit(9, 2, asAdult) 'addArea(2, asAdult)
                Else
                    If item("zelda's lullaby") Or My.Settings.setCucco Then addArea(36, asAdult)
                    If item("dive") Then addAreaExit(9, 2, asAdult) 'addArea(2, asAdult)
                    If item("ocarina") Then addArea(58, asAdult)
                End If
            Case 36
                'ZR Behind Waterfall to ZR Main and ZD Main
                addArea(35, asAdult)
                addAreaExit(9, 1, asAdult) 'addArea(37, asAdult)
            Case 37
                ' ZD Main to ZR Behind Waterfall, ZR Shoppe, ZR Behind King, LH
                addAreaExit(13, 1, asAdult) 'addArea(36, asAdult)
                If asAdult Then
                    If item("blue fire") Then addArea(39, asAdult)
                    If checkLoc("6303") Or My.Settings.setOpenZF Then addArea(38, asAdult)
                Else
                    addArea(39, asAdult)
                    If checkLoc("6303") Or allItems.Contains("v") Then addArea(38, asAdult)
                    If item("dive") Then addAreaExit(13, 2, asAdult) 'addArea(42, asAdult)
                End If
            Case 38
                ' ZD Behind King to ZF Main, ZD Main
                addAreaExit(13, 0, asAdult) 'addArea(40, asAdult)
                If asAdult Then
                    If checkLoc("6303") Or My.Settings.setOpenZF Then addArea(37, asAdult)
                Else
                    If checkLoc("6303") Then addArea(37, asAdult)
                End If
            Case 39
                ' ZD Shoppe to ZD Main
                addArea(37, asAdult)
            Case 40
                ' ZF Main to ZD Behind King, ZF Ice Cavern Ledge, Jabu-Jabu's Belly
                addAreaExit(14, 2, asAdult) 'addArea(38, asAdult)
                If asAdult Then
                    addArea(41, asAdult)
                Else
                    If item("bottle") Then addAreaExit(14, 0, asAdult) 'addArea(79, asAdult)
                End If
            Case 41
                ' ZF Ice Cavern Ledge to ZF Main, Ice Cavern
                addArea(40, asAdult)
                If asAdult Then addAreaExit(14, 1, asAdult) 'addArea(178, asAdult)
            Case 42
                ' LH Main to HF, LH Fishing Ledge, ZD Main, Water Temple
                addAreaExit(12, 1, asAdult) 'addArea(7, asAdult)
                If asAdult Then
                    If item("scarecrow") Or canMagicBean("202") Or canChangeLake() Then addArea(43, asAdult)
                    If item("iron boots") And (checkLoc("231") Or allItems.Contains("k")) Then addAreaExit(12, 0, asAdult) 'addArea(119, asAdult)
                    If item("dive", 2) And (checkLoc("231") Or allItems.Contains("l")) Then addAreaExit(12, 0, asAdult) 'addArea(119, asAdult)
                Else
                    'addArea(7, asAdult)   ' Hoot hoot!
                    addArea(43, asAdult)
                    If item("dive") Then addAreaExit(12, 2, asAdult) 'addArea(37, asAdult)
                End If
            Case 43
                ' LH Fishing Ledge to LH Main
                addArea(42, asAdult)
            Case 44
                ' GV Hyrule Side to HF, GV Upper Stream, GV Gerudo Side, GV Crate Ledge
                addAreaExit(15, 1, asAdult) 'addArea(7, asAdult)
                addArea(56, asAdult)
                If asAdult Then
                    If item("epona") Or item("longshot") Or checkLoc("CARD") Then addArea(45, asAdult)
                    If item("longshot") Then addArea(57, asAdult)
                Else
                    addArea(57, asAdult)
                End If
            Case 45
                ' GV Gerudo Side to GF Main, GV Upper Stream, GV Hyrule Side, GV Crate Ledge
                addAreaExit(15, 2, asAdult) 'addArea(46, asAdult)
                addArea(56, asAdult)
                If asAdult Then
                    If item("epona") Or item("longshot") Or checkLoc("CARD") Then addArea(44, asAdult)
                    If canHoverTricks() Then addArea(57, asAdult)
                Else
                    addArea(44, asAdult)
                End If
            Case 46
                ' GF Main to GV Gerudo Side, GF Behind Gate, Gerudo Training Ground
                addAreaExit(18, 1, asAdult) 'addArea(45, asAdult)
                If asAdult Then
                    If item("membership card") Then
                        addArea(47, asAdult)
                        addAreaExit(18, 0, asAdult) 'addArea(179, asAdult)
                    End If
                Else
                    If checkLoc("C00") Then addArea(47, asAdult)
                End If
            Case 47
                ' GF Behind Gate to HW Gerudo Side, GF Main
                addAreaExit(18, 2, asAdult) 'addArea(48, asAdult)
                If asAdult Then
                    addArea(46, asAdult)
                Else
                    If iER Mod 2 = 1 And checkLoc("C00") Then addArea(46, asAdult)
                End If
            Case 48
                ' HW Gerudo Side to GF Behind Gate, HW Main
                addAreaExit(19, 1, asAdult) 'addArea(47, asAdult)
                If asAdult Then
                    If My.Settings.setWastelandBackstep Or item("hover boots") Or item("longshot") Then addArea(55, asAdult)
                Else
                    If My.Settings.setWastelandBackstep Then addArea(55, asAdult)
                End If
            Case 49
                ' HW Colossus Side to DC, HW Main
                addAreaExit(19, 0, asAdult) 'addArea(50, asAdult)
                If My.Settings.setWastelandReverse Then addArea(55, asAdult)
            Case 50
                ' DC to HW Colossus Side, Spirit Temple
                addAreaExit(17, 1, asAdult) 'addArea(49, asAdult)
                addAreaExit(17, 0, asAdult) 'addArea(132, asAdult)
            Case 51
                ' OGC to MK, GFF, Ganon's Castle
                addAreaExit(25, 1, asAdult) 'addArea(10, asAdult)
                If asAdult And item("lift", 3) Then addArea(53, asAdult)
                If canEnterGanonsCastle() Then addArea(193, asAdult)
            Case 52
                ' Schrodinger's Castle: Adult to OGC, Young to HC
                If asAdult Then
                    addArea(51, asAdult)
                Else
                    addArea(13, asAdult)
                End If
            Case 53, 54
                ' OGC Fairy Fountain and HC Fairy Fountain
                addArea(52, asAdult)
            Case 55
                ' HW Main to HW Colossus Side, HW Gerudo Side
                If My.Settings.setWastelandLensless Or item("lens of truth") Then addArea(49, asAdult)
                If asAdult Then
                    If My.Settings.setWastelandBackstep Or item("hover boots") Or item("longshot") Then addArea(48, asAdult)
                Else
                    If My.Settings.setWastelandBackstep Then addArea(48, asAdult)
                End If
            Case 56
                ' GV Upper Stream to LH
                addAreaExit(15, 0, asAdult) 'addArea(42, asAdult)
            Case 57
                ' GV Crate Ledge to LH
                addAreaExit(15, 0, asAdult) 'addArea(42, asAdult)
            Case 60
                'addAreaExit(26, 0, asAdult)
                If Not aMQ(0) Then
                    ' 60	DT Lobby
                    ' 61	DT Slinghsot Room
                    ' 62	DT Basement Backroom
                    ' 63	DT Boss Room

                    ' DT: Lobby to Slingshot Room, Basement Backroom, Boss Room

                    If iER < 2 Then
                        If Not asAdult Then
                            If item("deku shield") Then addArea(61, asAdult)
                            If (item("slingshot") And canBurnYoung()) Or My.Settings.setDekuB1Skip Or checkLoc("10316") Then
                                addArea(62, asAdult)
                                addArea(63, asAdult)
                            End If
                        End If
                    Else
                        If (aReachY(60) And item("deku shield")) Or (aReachA(60) And item("hylian shield")) Then addArea(61, asAdult)
                        If (((aReachY(60) And canBurnYoung()) Or (aReachA(60) And (canBurnAdult() Or item("bow")))) And
                                ((aReachY(60) And item("slingshot")) Or (aReachA(60) And item("bow")))) Or
                            (Not asAdult And (My.Settings.setDekuB1Skip Or aReachA(60) Or checkLoc("10316"))) Then addArea(62, asAdult)
                        ' TODO: False is placeholder for "(logic_deku_b1_webs_with_bow and can_use(Bow)))"
                        If (((aReachY(60) And canBurnYoung()) Or (aReachA(60) And canBurnAdult())) Or
                                False) And
                            (Not asAdult And (My.Settings.setDekuB1Skip Or aReachA(60) Or checkLoc("10316"))) Then addArea(63, asAdult)
                    End If
                Else
                    ' 60	DT MQ Lobby
                    ' 64	DT MQ Compass Room
                    ' 65	DT MQ Basement Water Room Front
                    ' 66	DT MQ Basement Water Room Back
                    ' 67	DT MQ Back Room
                    ' 68	DT MQ Basement Ledge

                    ' DT MQ: Lobby to Compass Room, Basement Water Room Front, Basement Ledge
                    If iER < 2 Then
                        If Not asAdult Then
                            If item("slingshot") And canBurnYoung() Then
                                addArea(64, asAdult)
                                addArea(65, asAdult)
                            End If
                            If My.Settings.setDekuB1Skip Or checkLoc("10316") Then addArea(68, asAdult)
                        End If
                    Else
                        If ((aReachY(60) And item("slingshot")) Or (aReachA(60) And item("bow"))) And
                            ((aReachY(60) And canBurnYoung()) Or (aReachA(60) And (canBurnAdult() Or item("bow")))) Then addArea(64, asAdult)
                        If ((aReachY(60) And item("slingshot")) Or (aReachA(60) And item("bow"))) And ((aReachY(60) And canBurnYoung()) Or (aReachA(60) And canBurnAdult())) Then addArea(65, asAdult)
                        If My.Settings.setDekuB1Skip Or aReachA(60) Or checkLoc("10316") Then addArea(68, asAdult)
                    End If
                End If
            Case 61 To 63
                ' DT: Slingshot Room, Basement Backroom, Boss Room to Lobby
                addArea(60, asAdult)
            Case 64
                ' DT MQ: Compass Room to Lobby
                addArea(60, asAdult)
            Case 65
                ' DT MQ: Basement Water Rooom Front to Lobby, Basement Water Rooom Front
                addArea(60, asAdult)
                If iER > 2 Then
                    If Not asAdult And item("deku shield") Or item("hylian shield") Then addArea(66, asAdult)
                Else
                    ' TODO: False is placeholder for "logic_deku_mq_log"
                    If False Or (Not asAdult And (item("deku shield") Or item("hylian shield"))) Or
                        (asAdult And (item("longshot") Or (item("hookshot") And item("iron boots")))) Then addArea(66, asAdult)
                End If
            Case 66
                ' DT MQ: Basement Water Rooom Back to Basement Water Rooom Front, Basement Back Room
                addArea(65, asAdult)
                If iER < 2 Then
                    If Not asAdult And (item("deku stick") Or item("din's fire")) And (item("kokiri sword") Or canProjectile(0) Or (item("deku nuts") And item("deku stick"))) Then addArea(67, asAdult)
                Else
                    If ((aReachY(66) And item("deku stick")) Or item("din's fire") Or
                        (aReachA(65) And item("fire arrows"))) And
                    (aReachA(66) Or item("kokiri sword") Or (aReachY(67) And canProjectile(0)) Or (item("deku nuts") And item("deku stick"))) Then addArea(67, asAdult)
                End If
            Case 67
                ' DT MQ: Basement Back Room to Basement Ledge, Basement Water Room Back
                If iER < 2 Then
                    If Not asAdult Then
                        addArea(68, asAdult)
                        If item("kokiri sword") Or canProjectile(0) Or (item("deku nuts") And item("deku stick")) Then addArea(66, asAdult)
                    End If
                Else
                    If Not asAdult Then addArea(68, asAdult)
                    If item("kokiri sword") Or ((aReachY(67) And canProjectile(0)) Or (aReachA(67) And canProjectile(2))) Or (item("deku nuts") And item("deku stick")) Then addArea(66, asAdult)
                End If
            Case 68
                ' DT MQ: Basement Ledge to Lobby, Basement Back Room
                addArea(60, asAdult)
                If Not asAdult Then addArea(67, asAdult)
            Case 69
                'addAreaExit(27, 0, asAdult)
                If Not aMQ(1) Then
                    ' 69	DC Lobby
                    ' 70	DC Staircase
                    ' 71	DC Climb
                    ' 72	DC Far Bridge
                    ' 73	DC Boss Area

                    ' DC: Lobby to Boss Area, Staircase Room
                    If (checkLoc("10410") And canExplode()) Or checkLoc("10426") Then addArea(73, asAdult)
                    If asAdult Then
                        addArea(70, asAdult)
                    Else
                        If (canExplode() Or item("lift")) And (item("deku stick") Or (item("din's fire") And (item("slingshot") Or item("bombs") Or item("kokiri sword")))) Then addArea(70, asAdult)
                    End If
                Else
                    ' 69	DC MQ Lobby
                    ' 74	DC MQ Elevator
                    ' 75	DC MQ Lower Right Side
                    ' 76	DC MQ Bomb Bag Area
                    ' 77	DC MQ Boss Area
                    ' 78	DC MQ Boss Room

                    ' DC MQ: Lobby to Boss Area, Elevator, Lower Right Side, Bomb Bag Area
                    If (checkLoc("10410") And canExplode()) Or checkLoc("10426") Then addArea(78, asAdult)
                    If asAdult Then
                        If canBreakRocks() Or item("lift") Then addArea(74, asAdult)
                        If canBreakRocks() Then addArea(75, asAdult)
                        addArea(76, asAdult)
                    Else
                        If canExplode() Or item("lift") Then addArea(74, asAdult)
                        If canExplode() Then addArea(75, asAdult)
                    End If
                End If
            Case 70
                ' DC: Staircase Room to Lobby, Climb
                addArea(69, asAdult)
                If asAdult Then
                    If canExplode() Or item("lift") Or item("din's fire") Or (My.Settings.setDCStaircase And item("bow")) Then addArea(71, asAdult)
                Else
                    If canExplode() Or item("lift") Or item("din's fire") Then addArea(71, asAdult)
                End If
            Case 71
                ' DC: Climb to Lobby, Far Bridge
                addArea(69, asAdult)
                If asAdult Then
                    If item("hover boots") Or item("longshot") Or My.Settings.setDCSpikeJump Or ((canBreakRocks() Or item("lift")) And item("bow")) Then addArea(72, asAdult)
                Else
                    If (canBreakRocks() Or item("lift")) And (item("slingshot") Or (My.Settings.setDCSlingshotSkips And (item("deku stick") Or item("bomb") Or item("kokiri sword")))) Then addArea(72, asAdult)
                End If
            Case 72
                ' DC: Far Bridge to Lobby, Boss Area
                addArea(69, asAdult)
                If canExplode() Then addArea(73, asAdult)
            Case 73
                ' DC: Boss Area to Lobby
                addArea(69, asAdult)
            Case 74
                ' DC MQ: Elevator to Lower Right Side, Boss Area
                If asAdult Then
                    If item("din's fire") And item("lift") Then addArea(75, asAdult)
                    If canExplode() Or (item("lift") And My.Settings.setDCLightEyes And (item("din's fire") Or My.Settings.setDCSpikeJump Or item("hammer") Or item("hover boots") Or item("hookshot"))) Then addArea(77, asAdult)
                Else
                    If (item("deku stick") Or item("din's fire")) And item("lift") Then addArea(75, asAdult)
                    If canExplode() Or (item("lift") And My.Settings.setDCLightEyes And (item("deku stick") Or item("din's fire"))) Then addArea(77, asAdult)
                End If
            Case 75
                ' DC MQ: Lower Right Side to Bomb Bag Area
                If Not asAdult And (item("lift") Or item("din's fire") Or canExplode()) And item("slingshot") Then addArea(76, asAdult)
            Case 76
                ' DC MQ: Bomb Bag Area to Lower Right Side
                addArea(75, asAdult)
            Case 77
                ' DC MQ: Boss Area to Boss Room
                If asAdult Then
                    addArea(78, asAdult)
                Else
                    If canExplode() Or item("din's fire") Or (item("deku stick") Or ((item("deku nuts") Or item("boomerang")) And (item("kokiri sword") Or item("slingshot")))) Then addArea(78, asAdult)
                End If
            Case 79
                'addAreaExit(28, 0, asAdult)
                If Not aMQ(2) Then
                    ' 79	JB Beginning
                    ' 80	JB Main
                    ' 81	JB Depths
                    ' 82	JB Boss Area

                    ' JB: Beginning to Main
                    If iER < 2 Then
                        If Not asAdult And canProjectile(0) Then addArea(80, asAdult)
                    Else
                        If asAdult Then
                            If canProjectile(2) Then addArea(80, asAdult)
                        Else
                            If canProjectile(0) Then addArea(80, asAdult)
                        End If
                    End If
                Else
                    ' 79	JB MQ Beginning
                    ' 83	JB MQ Elevator Room
                    ' 84	JB MQ Main
                    ' 85	JB MQ Depths
                    ' 86	JB MQ Past Big Octo
                    ' 87	JB MQ Boss Area

                    ' JB MQ: Beginning to Main
                    If Not asAdult And item("slingshot") Then addArea(84, asAdult)
                End If
            Case 80
                ' JB: Main to Beginning, Boss Area, Depths
                addArea(79, asAdult)
                If checkLoc("10529") Then addArea(82, asAdult)
                If asAdult Then
                    ' TODO: False is placeholder for "logic_jabu_boss_gs_adult"
                    If False And item("hover boots") Then addArea(82, asAdult)
                Else
                    If item("boomerang") Then addArea(81, asAdult)
                End If
            Case 81
                ' JB: Depths to Main, Boss Area
                addArea(80, asAdult)
                If Not asAdult And (item("deku stick") Or item("kokiri sword")) Then addArea(82, asAdult)
            Case 82
                ' JB: Boss Area to Main
                addArea(80, asAdult)
            Case 83
                ' JB MQ: Elevator to Beginning, Boss Area, Main
                addArea(79, asAdult)
                If checkLoc("10529") Then addArea(87, asAdult)
                If Not asAdult Then addArea(84, asAdult)
            Case 84
                ' JB MQ: Main to Elevator, Depths
                addArea(83, asAdult)
                If Not asAdult And canExplode() And item("boomerang") And item("slingshot") Then addArea(85, asAdult)
            Case 85
                ' JB MQ: Depths to Main, Past Big Octo
                addArea(84, asAdult)
                If Not asAdult And (item("deku stick") Or (item("din's fire") And item("kokiri sword"))) Then addArea(86, asAdult)
            Case 86
                ' JB MQ: Past Big Octo to Main, Boss Area
                addArea(84, asAdult)
                addArea(87, asAdult)
            Case 87
                ' JB MQ: Boss Area to Main
                addArea(84, asAdult)
            Case 88
                'addAreaExit(29, 0, asAdult)
                If Not aMQ(3) Then
                    ' 88	FoT Lobby
                    ' 89	FoT NW Outdoors
                    ' 90	FoT NE OutDoors
                    ' 91	FoT Outdoors High Balconies
                    ' 92	FoT Falling Room
                    ' 93	FoT Block Push Room
                    ' 94	FoT Straightened Hall
                    ' 95	FoT Outside Upper Ledge
                    ' 96	FoT Bow Region
                    ' 97	FoT Boss Region

                    ' FoT Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("lift", 1) And item("bow") And ((item("hover boots") And item("hookshot")) Or Not My.Settings.setFoTFrame) Then
                            canDungeon(3) = True
                        Else
                            canDungeon(3) = False
                        End If
                    End If

                    ' FoT: Lobby to Block Push Room, NW Outdoors, NE Outdoors, Boss Region
                    If dungeonKeyCounter(3, "01") Then addArea(93, asAdult)
                    If asAdult Then
                        If item("song of time") Then addArea(89, asAdult)
                        If item("bow") Then addArea(90, asAdult)
                        If canDungeon(3) Or checkLoc("10628") Or (checkLoc("10629") And checkLoc("10630") And checkLoc("10631") And item("bow")) Then addArea(97, asAdult)
                    Else
                        addArea(89, asAdult)
                        If item("slingshot") Then addArea(90, asAdult)
                        If checkLoc("10628") Then addArea(97, asAdult)
                    End If
                Else
                    '  88	FoT MQ Lobby
                    '  98	FoT MQ Central Area
                    '  99	FoT MQ After Block Puzzle
                    ' 100	FoT MQ Outdoor Ledge
                    ' 101	FoT MQ NW Outdoors
                    ' 102	FoT MQ NE Outdoors
                    ' 103	FoT MQ Outdoors Top Ledges
                    ' 104	FoT MQ NE Outdoors Ledge
                    ' 105	FoT MQ Bow Region
                    ' 106	FoT MQ Falling Room
                    ' 107	FoT MQ Boss Region

                    ' FoT MQ Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("bow") And item("song of time") And ((item("lift") And Not My.Settings.setFoTMQPuzzle) Or (My.Settings.setFoTMQPuzzle And item("bombchu") And item("hookshot"))) And
                            ((item("hover boots") And (item("hookshot") Or Not My.Settings.setFoTBackdoor)) Or Not My.Settings.setFoTMQTwisted) Then
                            canDungeon(3) = True
                        Else
                            canDungeon(3) = False
                        End If
                    End If

                    ' FoT MQ: Lobby to Central Area
                    If asAdult Then
                        If dungeonKeyCounter(3, "06") Then addArea(98, asAdult)
                    Else
                        If dungeonKeyCounter(3, "06") And (canAttackYoung(True) Or item("deku nuts")) Then addArea(98, asAdult)
                    End If
                End If
            Case 89
                ' FoT: NW Outdoors to NE Outdoors, Outdoors High Balconies
                If item("dive", 2) Then addArea(90, asAdult)
                If asAdult Then
                    addArea(91, asAdult)
                Else
                    If canExplode() Or ((item("boomerang") Or item("deku nuts") Or item("deku shield")) And (item("deku stick") Or item("kokiri sword") Or item("slingshot"))) Then addArea(91, asAdult)
                End If
            Case 90
                ' FoT: NE Outdoors to Lobby, Outdoors High Balconies, NW Outdoors
                addArea(88, asAdult)
                If asAdult Then
                    If item("longshot") Or (My.Settings.setFoTVines And item("hookshot")) Then addArea(91, asAdult)
                    If item("iron boots") Or item("dive", 2) Then addArea(89, asAdult)
                Else
                    If item("dive", 2) Then addArea(89, asAdult)
                End If
            Case 91
                ' FoT: Outdoors High Balconies to NW Outdoors, NE Outdoors, Falling Room
                addArea(89, asAdult)
                addArea(90, asAdult)
                If asAdult And My.Settings.setFoTLedge And item("hover boots") And item("scarecrow") Then addArea(92, asAdult)
            Case 92
                ' FoT: Falling room to NE Outdoors
                addArea(90, asAdult)
            Case 93
                ' FoT: Block Push Room to Outside Upper Ledge, Bow Region, Straightened Hall
                If asAdult Then
                    If item("hover boots") Or (My.Settings.setFoTBackdoor And item("lift")) Then addArea(95, asAdult)
                    If item("lift") And dungeonKeyCounter(3, "010203") Then addArea(96, asAdult)
                    If item("lift") And dungeonKeyCounter(3, "0102") And item("bow") Then addArea(94, asAdult)
                End If
            Case 94
                ' FoT: Straightened Hallway to Outside Upper Ledge
                addArea(95, asAdult)
            Case 95
                ' FoT: Outside Upper Ledge to NW Outdoors
                addArea(89, asAdult)
            Case 96
                ' FoT: Bow Region to Falling Room
                If asAdult Then
                    If dungeonKeyCounter(3, "0001020304") And (item("bow") Or item("din's fire")) Then addArea(92, asAdult)
                Else
                    If dungeonKeyCounter(3, "0001020304") And item("din's fire") Then addArea(92, asAdult)
                End If
            Case 98
                ' FoT MQ: Central Area to NW Outdoors, NE Outdoors, Boss Region, After Block Puzzle, Outdoor Ledge
                If asAdult Then
                    If item("bow") Then
                        addArea(101, asAdult)
                        addArea(102, asAdult)
                    End If
                    If canDungeon(3) Or checkLoc("10628") Or (checkLoc("10629") And checkLoc("10630") And checkLoc("10631") And item("bow")) Then addArea(107, asAdult)
                    If item("lift") Or (My.Settings.setFoTMQPuzzle And item("bombchu") And item("hookshot")) Then addArea(99, asAdult)
                    If My.Settings.setFoTMQTwisted And item("hover boots") Then addArea(100, asAdult)
                Else
                    If item("slingshot") Then
                        addArea(101, asAdult)
                        addArea(102, asAdult)
                    End If
                    If checkLoc("10628") Then addArea(107, asAdult)
                End If
            Case 99
                ' FoT MQ: After Block Puzzle to Bow Region, NW Outdoors, Outdoor Ledge
                If dungeonKeyCounter(3, "060201") Then addArea(105, asAdult)
                If dungeonKeyCounter(3, "0602") Then addArea(101, asAdult)
                If dungeonKeyCounter(3, "0602") Or (My.Settings.setFoTMQTwisted And (item("hookshot") Or My.Settings.setFoTBackdoor)) Then addArea(100, asAdult)
            Case 100
                ' FoT MQ: Outdoor Ledge to NW Outdoors
                addArea(101, asAdult)
            Case 101
                ' FoT MQ: NW Outdoors to NE Outdoors, Outdoors Top Ledges
                If asAdult Then
                    If item("iron boots") Or item("longshot") Or (item("dive", 2) Or (My.Settings.setFoTMQWell And item("hookshot"))) Then addArea(102, asAdult)
                    If item("fire arrow") Then addArea(103, asAdult)
                End If
            Case 102
                ' FoT MQ: NE Outdoors to Outdoors Top Ledges, NE Outdoors Ledge
                If asAdult Then
                    If item("hookshot") Or (item("longshot") Or item("hover boots") Or item("song of time") Or My.Settings.setFoTVines) Then addArea(103, asAdult)
                    If item("longshot") Then addArea(104, asAdult)
                End If
            Case 103
                ' FoT MQ: Outdoors Top Ledges to NE Outdoors, NE Outdoors Ledge
                addArea(102, asAdult)
                If asAdult And My.Settings.setFoTLedge And item("hover boots") Then addArea(104, asAdult)
            Case 104
                ' FoT MQ: NE Outdoors Ledge to NE Outdoors, Falling Room
                addArea(102, asAdult)
                If item("song of time") Then addArea(106, asAdult)
            Case 105
                ' FoT MQ: Bow Region to Falling Room
                If dungeonKeyCounter(3, "0602010004") And (item("bow") Or item("din's fire")) Then addArea(106, asAdult)
            Case 106
                ' FoT MQ: Falling Room to NE Outdoors Ledge
                addArea(104, asAdult)
            Case 108
                'addAreaExit(30, 0, asAdult)
                If Not aMQ(4) Then
                    ' 108	FiT Lower
                    ' 109	FiT Big Lava Room
                    ' 110	FiT Middle
                    ' 111	FiT Upper
                    ' 112	FiT Boss Loop

                    ' FiT Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("goron tunic") And item("hammer") And item("bow") And canExplode() And (item("lift") Or My.Settings.setFiTClimb) And
                            (item("scarecrow") Or (My.Settings.setFiTScarecrow And item("longshot"))) And (item("hover boots") Or Not My.Settings.setFiTMaze) Then
                            canDungeon(4) = True
                        Else
                            canDungeon(4) = False
                        End If
                    End If

                    ' FiT: Lower to Big Lava Room, Boss Loop
                    If dungeonKeyCounter(4, "29") And canFewerGoron() Then addArea(109, asAdult)
                    If asAdult And item("hammer") And (Not My.Settings.setSmallKeys = 1 Or dungeonKeyCounter(4, "23")) Then addArea(112, asAdult)
                Else
                    ' 108	FiT MQ Lower
                    ' 113	FiT MQ Lower Locked Door
                    ' 114	FiT MQ Big Lava Room
                    ' 115	FiT MQ Lower Maze
                    ' 116	FiT MQ Upper Maze
                    ' 117	FiT MQ Upper
                    ' 118	FiT MQ Boss Room

                    ' FiT MQ Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("goron tunic") And item("hammer") And item("hover boots") And item("bow") And item("hookshot") And canExplode() And canBurnAdult() Then
                            canDungeon(4) = True
                        Else
                            canDungeon(4) = False
                        End If
                    End If

                    ' FiT MQ: Lower to Boss Room, Big Lava Room, Lower Locked Door
                    If asAdult Then
                        If item("goron tunic") And item("hammer") And (canBurnAdult() And (item("hover boots") Or checkLoc("10710"))) Then addArea(118, asAdult)
                        If canFewerGoron() And item("hammer") Then addArea(114, asAdult)
                        If dungeonKeyCounter(4, "23") Then addArea(113, asAdult)
                    Else
                        If dungeonKeyCounter(4, "23") And item("kokiri sword") Then addArea(113, asAdult)
                    End If
                End If
            Case 109
                ' FiT: Big Lava Room to Lower, Middle
                addArea(108, asAdult)
                If asAdult And item("goron tunic") And dungeonKeyCounter(4, "293024") And (item("lift") Or My.Settings.setFiTClimb) And
                    (canExplode() Or item("bow") Or item("hookshot")) Then addArea(110, asAdult)
            Case 110
                ' FiT: Middle to Upper
                If asAdult And dungeonKeyCounter(4, "29302427263125") Or (dungeonKeyCounter(4, "293024272631") And (item("hover boots") Or item("hammer") Or My.Settings.setFiTMaze)) Then addArea(111, asAdult)
            Case 114
                ' Fit MQ: Big Lava Room to Lower Maze
                If asAdult And item("goron tunic") And dungeonKeyCounter(4, "30") And (canBurnAdult() Or (My.Settings.setFiTMQClimb And item("hover boots"))) Then addArea(115, asAdult)
            Case 115
                ' Fit MQ: Lower Maze to Upper Maze
                If asAdult And (((canExplode() Or My.Settings.setFiTRustedSwitches) And item("hookshot")) Or (My.Settings.setFiTMQMazeHovers And item("hover boots")) Or My.Settings.setFiTMQMazeJump) _
                    Then addArea(116, asAdult)
            Case 116
                ' Fit MQ: Upper Maze to Upper
                If asAdult And dungeonKeyCounter(4, "3026") And ((item("bow") And item("hookshot")) Or item("fire arrows")) Then addArea(117, asAdult)
            Case 117
                ' Fit MQ: Upper to Boss Room
                ' This is not normally there, but this will prevent failed checks
                If asAdult And item("goron tunic") And item("hammer") Then addArea(118, asAdult)
            Case 119
                'addAreaExit(31, 0, asAdult)
                If Not aMQ(5) Then
                    ' 119	WaT Lobby
                    ' 120	WaT Highest Water Level
                    ' 121	WaT Dive
                    ' 122	WaT North Basement
                    ' 123	WaT Cracked Wall
                    ' 124	WaT Dragon Statue
                    ' 125	WaT Middle Water Level
                    ' 126	WaT Falling Platform Room
                    ' 127	WaT Dark Link Region

                    ' WaT Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("longshot") And item("zelda's lullaby") And item("song of time") And item("lift") And item("bow") And item("iron boots") And item("zora tunic") And canExplode() Then
                            canDungeon(5) = True
                        Else
                            canDungeon(5) = False
                        End If

                    End If

                    ' WaT: Lobby to Highest Water Level, Dive
                    If asAdult Then
                        If item("hookshot") Or item("bow") Or (canBurnAdult() And canProjectile(2)) Then addArea(120, asAdult)
                        If canFewerZora() And ((My.Settings.setWaTTorch And item("longshot")) Or item("iron boots")) Then addArea(121, asAdult)
                    Else
                        If canBurnYoung() And canProjectile(0) Then addArea(120, asAdult)
                    End If
                Else
                    ' 119	WaT MQ Lobby
                    ' 128	WaT MQ Dive
                    ' 129	WaT MQ Lowered Water Levels
                    ' 130	WaT MQ Dark Link Region
                    ' 131	WaT MQ Basement Grated Areas

                    ' WaT MQ Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("longshot") And item("din's fire") And item("zora tunic") And item("song of time") And item("zelda's lullaby") And item("iron boots") And (item("hover boots") Or item("scarecrow")) Then
                            canDungeon(5) = True
                        Else
                            canDungeon(5) = False
                        End If
                    End If

                    ' WaT MQ: Lobby to Dive, Dark Link Region
                    If asAdult Then
                        If canFewerZora() And item("iron boots") Then addArea(128, asAdult)
                        If dungeonKeyCounter(5, "21") And item("longshot") Then addArea(130, asAdult)
                    End If
                End If
            Case 120
                ' WaT: Highest Water Level to Falling Platform Room
                If dungeonKeyCounter(5, "05") Then addArea(126, asAdult)
            Case 121
                ' WaT: Dive to Middle Water Level, Cracked Wall, North Basement, Dragon Statue

                ' If adult Link can make it here, and young Link can make it into the Water Temple, then adult Link can lower the water to let young Link down here.
                If aReachY(119) Then addArea(121, False)

                If asAdult Then
                    If (item("bow") Or item("din's fire") Or (dungeonKeyCounter(5, "06") And item("hookshot"))) And item("zelda's lullaby") Then addArea(125, asAdult)
                    If item("zelda's lullaby") And (item("hookshot") Or item("hover boots")) And
                        (My.Settings.setWaTCrackNothing Or (My.Settings.setWaTCrackHovers And item("hover boots"))) Then addArea(123, asAdult)
                    If dungeonKeyCounter(5, "01") And (item("longshot") Or (My.Settings.setWaTBKR And item("hover boots"))) And (item("iron boots") Or item("zelda's lullaby")) Then addArea(122, asAdult)
                    If item("zelda's lullaby") And item("lift") And
                        (item("iron boots") And item("hookshot")) Or (My.Settings.setWaTDragonDive And (item("bombchu") Or item("bow") Or item("hookshot")) And (item("dive") Or item("iron boots"))) _
                        Then addArea(124, asAdult)
                Else
                    If (item("din's fire") Or item("deku stick")) And item("zelda's lullaby") Then addArea(125, asAdult)
                End If
            Case 125
                ' WaT: Middle Water Level to Cracked Wall
                addArea(123, asAdult)
            Case 126
                ' WaT: Falling Platform Room to Dark Link Region
                If asAdult And dungeonKeyCounter(5, "0502") And item("hookshot") Then addArea(127, asAdult)
            Case 127
                ' WaT: Dark Link Region to Dragon Statue
                If asAdult And canFewerZora() And item("song of time") And item("bow") And (item("iron boots") Or My.Settings.setWaTDragonDive Or My.Settings.setWaTDrgnSwitch) Then addArea(124, asAdult)
            Case 128
                ' WaT MQ: Dive to Lowered Water Levels
                If item("zelda's lullaby") Then addArea(129, asAdult)
            Case 130
                ' WaT MQ: Dark Link Region to Basement Grated Areas
                If asAdult And canFewerZora() And item("din's fire") And item("iron boots") Then addArea(131, asAdult)
            Case 132
                'addAreaExit(32, 0, asAdult)
                If Not aMQ(6) Then
                    ' 132	SpT Lobby
                    ' 133	SpT Child
                    ' 134	SpT Child Before Locked Door
                    ' 135	SpT Child Climb
                    ' 136	SpT Early Adult
                    ' 137	SpT Central Chamber
                    ' 138	SpT Outdoor Hands
                    ' 139	SpT Beyond Central Locked Door
                    ' 140	SpT Beyond Final Locked Door
                    ' 141	SpT Boss Platform

                    ' SpT Keys in Dungeon
                    If aReachY(132) And aReachA(132) And My.Settings.setSmallKeys = 0 Then
                        If item("lift", 2) And canExplode() And item("hookshot") And item("zelda's lullaby") And item("hover boots") And item("mirror shield") And (My.Settings.setSpTLensless Or item("lens of truth")) And
                            canProjectile(1) And (item("boomerang") Or item("slingshot")) And (item("din's fire") Or item("fire arrows")) Then
                            canDungeon(6) = True
                        Else
                            canDungeon(6) = False
                        End If

                    End If

                    ' SpT: Lobby to Early Adult, Central Chamber, Child 
                    If asAdult Then
                        If item("lift", 2) Then addArea(136, asAdult)
                        If canSpiritShortcut() Then addArea(137, asAdult)
                    Else
                        addArea(133, asAdult)
                    End If
                Else
                    ' 132	SpT MQ Lobby
                    ' 142	SpT MQ Child
                    ' 143	SpT MQ Adult
                    ' 144	SpT MQ Shared
                    ' 145	SpT MQ Lower Adult
                    ' 146	SpT MQ Boss Area
                    ' 147	SpT MQ Boss Platform
                    ' 148	SpT MQ Shield Hand
                    ' 149	SpT MQ Silver Gauntlets Hand

                    ' SpT MQ Keys in Dungeon
                    If aReachY(132) And aReachA(132) And My.Settings.setSmallKeys = 0 Then
                        If item("longshot") And item("bombchus") And item("slingshot") And item("din's fire") And item("zelda's lullaby") And item("song of time") And item("hammer") And item("mirror shield") And item("bow") And
                            (My.Settings.setSpTLensless Or item("lens of truth")) And (item("fire arrows") Or My.Settings.setSpTMQLowAdult) And
                            (item("deku sticks") Or item("kokiri sword") Or item("bombs")) Then
                            canDungeon(6) = True
                        Else
                            canDungeon(6) = False
                        End If

                    End If

                    ' SpT MQ: Lobby to Child, Adult
                    If asAdult Then
                        If item("longshot") And (item("lift", 2) Or checkLoc("10909")) And (item("bombchu") Or checkLoc("10912")) Then addArea(143, asAdult)
                    Else
                        addArea(142, asAdult)
                    End If
                End If
            Case 133
                ' SpT: Child to Before Locked Door
                addArea(134, asAdult)
            Case 134
                ' SpT: Child Before Locked Door to Child Climb
                If Not asAdult And dungeonKeyCounter(6, "30") Then addArea(135, asAdult)
            Case 135
                ' SpT: Child Climb to Central Chamber
                If canExplode() Then addArea(137, asAdult)
            Case 136
                ' SpT: Early Adult to Central Chamber
                If dungeonKeyCounter(6, "13") Then addArea(137, asAdult)
            Case 137
                ' SpT: Central Chamber to Child Climb, Outdoor Hands, Beyond Central Locked Door, Boss Platform, Early Adult
                addArea(135, asAdult)
                addArea(138, asAdult)
                If asAdult Then
                    If dungeonKeyCounter(6, "1327") Or (canSpiritShortcut() And dungeonKeyCounter(6, "27") And item("hookshot")) Then addArea(139, asAdult)
                    If item("longshot") And checkLoc("10923") And checkLoc("10904") Then addArea(141, asAdult)
                    If canSpiritShortcut() And item("hookshot") Then addArea(136, asAdult)
                End If
            Case 138
                ' SpT: Outdoor Hands to DC
                addArea(50, asAdult)
            Case 139
                ' SpT: Beyond Central Locked Door to Beyond Final Locked Door
                If asAdult And (dungeonKeyCounter(6, "132728") Or (canSpiritShortcut() And dungeonKeyCounter(6, "2728"))) And
                    (My.Settings.setSpTWall Or item("longshot") Or item("bombchu") Or
                     ((item("bombs") Or item("deku nuts") Or item("din's fire")) And (item("bow") Or item("hookshot") Or item("hammer")))) Then addArea(140, asAdult)
            Case 140
                ' SpT: Beyond Final Locked Door to Boss Platform
                If asAdult And item("mirror shield") And canExplode() Then addArea(141, asAdult)
            Case 142
                ' SpT MQ: Child to Shared
                If item("bombchu") And dungeonKeyCounter(6, "3018") Then addArea(144, asAdult)
            Case 143
                ' SpT MQ: Adult to Shared, Shield Hand, Lower Adult, Boss Area, Boss Platform
                addArea(144, asAdult)
                If dungeonKeyCounter(6, "27") And item("song of time") And (My.Settings.setSpTLensless Or item("lens of truth")) And (canExplode() Or item("deku nuts")) Then addArea(148, asAdult)
                If asAdult Then
                    If item("mirror shield") And (item("fire arrows") Or (My.Settings.setSpTMQLowAdult And item("din's fire") And item("bow"))) Then addArea(145, asAdult)
                    If dungeonKeyCounter(6, "2728") And item("zelda's lullaby") And item("hammer") Then addArea(146, asAdult)
                    If checkLoc("10923") And checkLoc("10904") Then addArea(147, asAdult)
                End If
            Case 144
                ' SpT MQ: Shared to Silver Gauntlets Hand
                If asAdult Then
                    If dungeonKeyCounter(6, "21") Then addArea(149, asAdult)
                Else
                    If dungeonKeyCounter(6, "301821") And (item("song of time") Or My.Settings.setSpTMQSunRoom) Then addArea(149, asAdult)
                End If
            Case 146
                ' SpT MQ: Boss Area to Boss Platform
                If asAdult And item("mirror shield") Then addArea(147, asAdult)
            Case 149
                ' SpT MQ: Silver Gauntlets Hand to DC
                addArea(50, asAdult)
            Case 150
                'addAreaExit(33, 0, asAdult)
                If Not aMQ(7) Then
                    ' 150	ShT Entryway
                    ' 151	ShT Beginning
                    ' 152	ShT First Beamos
                    ' 153	ShT Huge Pit
                    ' 154	ShT Invisible Spikes
                    ' 155	ShT Wind Tunnel
                    ' 156	ShT After Wind
                    ' 157	ShT Boat
                    ' 158	ShT Beyond Boat
                    ' 159	ShT Boss Room

                    ' ShT Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("hover boots") And item("hookshot") And item("zelda's lullaby") And item("din's fire") And canExplode() And (item("lift") Or My.Settings.setShTUmbrella) And
                            (item("lens of truth") Or (My.Settings.setShTLensless And My.Settings.setShTPlatform)) And
                            ((item("bow") Or item("scarecrow", 2)) And (item("bombchus") Or Not My.Settings.setShTStatue)) Then
                            canDungeon(7) = True
                        Else
                            canDungeon(7) = False
                        End If
                    End If

                    ' ShT: Entry to Beginning
                    If asAdult And (My.Settings.setShTLensless Or item("lens of truth")) And (item("hover boots") Or item("hookshot")) Then addArea(151, asAdult)
                Else
                    ' 150	ShT MQ Entryway
                    ' 160	ShT MQ Beginning
                    ' 161	ShT MQ Dead Hand Area
                    ' 162	ShT MQ First Beamos
                    ' 163	ShT MQ Upper Huge Pit
                    ' 164	ShT MQ Invisible Blades
                    ' 165	ShT MQ Lower Huge Pit
                    ' 166	ShT MQ Falling Spikes
                    ' 167	ShT MQ Invisible Spikes
                    ' 168	ShT MQ Wind Tunnel 
                    ' 169	ShT MQ After Wind
                    ' 170	ShT MQ Boat
                    ' 171	ShT MQ Beyond Boat
                    ' 172	ShT MQ Invisible Maze
                    ' 173	ShT MQ Near Boss

                    ' ShT MQ Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("bow") And item("hover boots") And item("longshot") And item("song of time") And item("zelda's lullaby") And canExplode() And (item("lift") Or My.Settings.setShTUmbrella) And
                            (canBurnAdult() Or My.Settings.setShTMQPit) And (item("lens of truth") Or (My.Settings.setShTLensless And My.Settings.setShTPlatform)) Then
                            canDungeon(7) = True
                        Else
                            canDungeon(7) = False
                        End If
                    End If

                    ' ShT MQ: Entry to Beginning
                    If asAdult And (My.Settings.setShTLensless Or item("lens of truth")) And (item("hover boots") Or item("hookshot")) Then addArea(160, asAdult)
                End If
            Case 151
                ' ShT: Beginning to Entry Way, First Beamos
                addArea(150, asAdult)
                If asAdult And item("hover boots") Then addArea(152, asAdult)
            Case 152
                ' ShT: First Beamos to Huge Pit, Boat
                If canExplode() And dungeonKeyCounter(7, "22") Then addArea(153, asAdult)
                If checkLoc("11029") Then addArea(157, asAdult)
            Case 153
                ' ShT: Huge Pit to Invisible Spikes
                If dungeonKeyCounter(7, "2223") And (My.Settings.setShTPlatform Or item("lens of truth")) Then addArea(154, asAdult)
            Case 154
                ' ShT: Invisible Spikes to Wind Tunnel, Huge Pit
                If asAdult And item("hookshot") And dungeonKeyCounter(7, "222324") Then addArea(155, asAdult)
                If checkLoc("11029") And dungeonKeyCounter(7, "212423") And (My.Settings.setShTPlatform Or item("lens of truth")) Then addArea(153, asAdult)
            Case 155
                ' ShT: Wind Tunnel to After Wind, Invisible Spikes
                addArea(156, asAdult)
                If checkLoc("11029") And dungeonKeyCounter(7, "2124") Then addArea(154, asAdult)
            Case 156
                ' ShT: After Wind to Wind Tunnel, Boat
                addArea(155, asAdult)
                If dungeonKeyCounter(7, "22232421") Then addArea(157, asAdult)
            Case 157
                ' ShT: Boat to Beyond Boat, After Wind
                If item("zelda's lullaby") Then addArea(158, asAdult)
                If checkLoc("11029") And dungeonKeyCounter(7, "21") Then addArea(156, asAdult)
            Case 158
                ' ShT: Beyond Boat to Boss Room
                If dungeonKeyCounter(7, "2223242125") And (item("bow") Or item("scarecrow", 2) Or checkLoc("11016") Or (My.Settings.setShTStatue And item("bombchu"))) Then addArea(159, asAdult)
            Case 160
                ' ShT MQ: Beginning to Dead Hand Area, First Beamos
                If canExplode() And dungeonKeyCounter(7, "25") Then addArea(161, asAdult)
                If asAdult And item("fire arrows") Or item("hover boots") Or (My.Settings.setShTMQGap And item("longshot")) Then addArea(162, asAdult)
            Case 162
                ' ShT MQ: First Beamos to Upper Huge Pit, Boat
                If canExplode() And dungeonKeyCounter(7, "22") Then addArea(163, asAdult)
                If checkLoc("11029") Then addArea(169, asAdult)
            Case 163
                ' ShT MQ: Upper Huge Pit to Invisible Blades, Lower Huge Pit
                If item("song of time") Then addArea(164, asAdult)
                If (asAdult And canBurnAdult()) Or My.Settings.setShTMQPit Then addArea(165, asAdult)
            Case 165
                ' ShT MQ: Lower Huge Pit to Falling Spikes, Upper Huge Pit,Invisible Spikes
                addArea(166, asAdult)
                If asAdult Then
                    If item("longshot") Then addArea(163, asAdult)
                    If item("hover boots") And dungeonKeyCounter(7, "2223") And (My.Settings.setShTPlatform Or item("lens of truth")) Then addArea(167, asAdult)
                End If
            Case 167
                ' ShT MQ: Invisible Spikes to Wind Tunnel, Lower Huge Pit
                If asAdult Then
                    If item("hookshot") And dungeonKeyCounter(7, "222324") Then addArea(168, asAdult)
                    If (My.Settings.setShTPlatform Or item("lens of truth")) And item("hover boots") And checkLoc("11029") And dungeonKeyCounter(7, "212423") Then addArea(165, asAdult)
                End If
            Case 168
                ' ShT MQ: Wind Tunnel to After Wind, Invisible Spikes
                addArea(169, asAdult)
                If asAdult And item("hookshot") And checkLoc("11029") And dungeonKeyCounter(7, "2124") Then addArea(167, asAdult)
            Case 169
                ' ShT MQ: After Wind to Boat, Wind Tunnel
                If dungeonKeyCounter(7, "22232421") Then addArea(170, asAdult)
                If asAdult And item("hover boots") Then addArea(168, asAdult)
            Case 170
                ' ShT MQ: Boat to Beyond Boat, After Wind
                If item("zelda's lullaby") Then addArea(171, asAdult)
                If checkLoc("11029") And dungeonKeyCounter(7, "21") Then addArea(169, asAdult)
            Case 171
                ' ShT MQ: Beyond Boat to Invisible Maze, Near Boss
                If asAdult Then
                    If item("bow") And item("song of time") And item("longshot") Then addArea(172, asAdult)
                    If (item("bow") Or (My.Settings.setShTStatue And item("bombchu")) Or checkLoc("11016")) And item("hover boots") Then addArea(173, asAdult)
                End If
            Case 174
                'addAreaExit(34, 0, asAdult)
                If Not aMQ(8) Then
                    ' 174	BotW Entrance
                    ' 175	BotW Main Area
                    ' 176	BotW Main Area with Lens of Truth Check

                    ' BotW Keys in Dungeon
                    If Not asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("zelda's lullaby") And canExplode() And (item("lens of truth") Or My.Settings.setBotWLensless) And (item("deku stick") Or item("din's fire")) And
                            (item("kokiri sword") Or (item("deku stick") And My.Settings.setBotWDeadHand)) Then
                            canDungeon(8) = True
                        Else
                            canDungeon(8) = False
                        End If
                    End If

                    ' BotW: Entrance to Main Area
                    If Not asAdult And canAttackYoung(True) Or item("deku nuts") Then addArea(175, asAdult)
                Else
                    ' 174	BotW Entrance and Perimeter
                    ' 177	BotW Middle

                    ' BotW MQ Keys in Dungeon
                    If Not asAdult And My.Settings.setSmallKeys = 0 Then
                        If canExplode() And (item("lens of truth") Or My.Settings.setBotWLensless) And (item("kokiri sword") Or (item("deku stick") And My.Settings.setBotWDeadHand)) And
                            (item("zelda's lullaby") Or My.Settings.setBotWMQPits) Then
                            canDungeon(8) = True
                        Else
                            canDungeon(8) = False
                        End If
                    End If

                    ' BotW MQ: Entrance and Perimeter to Middle
                    If Not asAdult And item("zelda's lullaby") Or (My.Settings.setBotWMQPits And canExplode()) Then addArea(177, asAdult)
                End If
            Case 175
                ' BotW: Main Area Main Area with Lens of Truth Check
                If My.Settings.setBotWLensless Or item("lens of truth") Then addArea(176, asAdult)
            Case 178
                'addAreaExit(35, 0, asAdult)

                If Not aMQ(9) Then
                    ' 178   IC Beginning
                    ' 201   IC Main

                    ' IC: Beginning to Main
                    If asAdult Then
                        addArea(201, asAdult)
                    Else
                        If canExplode() Or item("din's fire") Or item("deku stick") Then addArea(201, asAdult)
                    End If
                Else
                    ' 178   IC Beginning
                    ' 202   IC Map Room
                    ' 203   IC Compass Room
                    ' 204   IC Iron Boots Region

                    If item("blue fire") Then addArea(204, asAdult)
                    If asAdult Then
                        addArea(202, asAdult)
                        If item("blue fire") Then addArea(203, asAdult)
                    Else
                        If item("din's fire") Or (canExplode() And (item("deku stick") Or item("slingshot") Or item("kokiri sword"))) Then addArea(202, asAdult)
                    End If
                End If
            Case 179
                'addAreaExit(36, 0, asAdult)
                If Not aMQ(10) Then
                    ' 179	GTG Lobby and Central Maze
                    ' 180	GTG Central Maze Right
                    ' 181	GTG Lava Room
                    ' 182	GTG Hammer Room
                    ' 183	GTG Eye Statue Lower
                    ' 184	GTG Eye Statue Upper
                    ' 185	GTG Heavy Block Room
                    ' 186	GTG Like Like Room

                    ' GTG Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("bow") And item("hammer") And item("hookshot") And item("iron boots") And item("song of time") And item("lift", 2) And canExplode() And canFewerZora() And
                            (item("lens of truth") Or My.Settings.setGTGLensless) Then
                            canDungeon(10) = True
                        Else
                            canDungeon(10) = False
                        End If
                    End If

                    ' GTG: Lobby and Central Maze to Central Maze Right, Lava Room, Heavy Block Room
                    If dungeonKeyCounter(10, "1003") Then addArea(180, asAdult)
                    If asAdult Then
                        If canExplode() Then addArea(181, asAdult)
                        If item("hookshot") Then addArea(185, asAdult)
                    Else
                        If canExplode() And item("kokiri sword") Then addArea(181, asAdult)
                    End If
                Else
                    ' 179	GTG MQ Lobby
                    ' 187	GTG MQ Right Side
                    ' 188	GTG MQ Underwater
                    ' 189	GTG MQ Left Side
                    ' 190	GTG MQ Stalfos Room
                    ' 191	GTG MQ Back Areas
                    ' 192	GTG MQ Central Maze Right

                    ' GTG MQ Keys in Dungeon
                    If asAdult And My.Settings.setSmallKeys = 0 Then
                        If item("bottle") And item("bow") And item("hammer") And item("hover boots") And item("iron boots") And item("longshot") And item("song of time") And
                            item("lift", 2) And canBurnAdult() And canFewerZora() And (item("lens of truth") Or My.Settings.setGTGLensless) Then
                            canDungeon(10) = True
                        Else
                            canDungeon(10) = False
                        End If
                    End If

                    ' GTG MQ: Lobby to Left Side, Right Side
                    If asAdult Then
                        If canBurnAdult() Then addArea(189, asAdult)
                        If item("bow") Then addArea(187, asAdult)
                    Else
                        If canBurnYoung() Then addArea(189, asAdult)
                        If item("slingshot") Then addArea(187, asAdult)
                    End If
                End If
            Case 180
                ' GTG: Central Maze Right to Lava Room, Hammer Room
                addArea(181, asAdult)
                If asAdult And item("hookshot") Then addArea(182, asAdult)
            Case 181
                ' GTG: Lava Room to Central Maze Right, Hammer Room
                If asAdult Then
                    If item("song of time") Then addArea(180, asAdult)
                    If item("longshot") Or (item("hookshot") And item("hover boots")) Then addArea(182, asAdult)
                Else
                    addArea(180, asAdult)
                End If
            Case 182
                ' GTG: Hammer Room to Lava Room, Eye Statue Lower
                addArea(181, asAdult)
                If asAdult And (item("hammer") Or item("bow")) Then addArea(183, asAdult)
            Case 183
                ' GTG: Eye Statue Lower to Hammer Room
                addArea(182, asAdult)
            Case 184
                ' GTG: Eye Statue Upper to Eye Statue Lower
                addArea(183, asAdult)
            Case 185
                ' GTG: Heavy Block Room to Eye Statue Upper, Like Like Room
                If asAdult Then
                    If (item("lens of truth") Or My.Settings.setGTGLensless) And item("hookshot") Then addArea(184, asAdult)
                    If item("lift", 2) And (item("lens of truth") Or My.Settings.setGTGLensless) And item("hookshot") Then addArea(186, asAdult)
                End If
            Case 187
                ' GTG MQ: Right Side to Underwater
                If asAdult And (item("bow") Or item("longshot")) And item("hover boots") Then addArea(188, asAdult)
            Case 189
                ' GTG MQ: Left Side to Stalfos Room
                If asAdult And item("longshot") Then addArea(190, asAdult)
            Case 190
                ' GTG MQ: Stalfos Room to Back Areas
                If asAdult And (item("lens of truth") Or My.Settings.setGTGLensless) And item("blue fire") And item("song of time") Then addArea(191, asAdult)
            Case 191
                ' GTG MQ: Back Areas to Right Side, Central Maze Right
                If asAdult Then
                    If item("longshot") Then addArea(187, asAdult)
                    If item("hammer") Then addArea(192, asAdult)
                End If
            Case 192
                ' GTG MQ: Central Maze Right to Right Side, Underwater
                If asAdult Then
                    If item("hookshot") Then addArea(187, asAdult)
                    If item("longshot") Or (item("hookshot") And item("bow")) Then addArea(188, asAdult)
                End If
            Case 193
                ' Normal and MQ share same structure
                ' 193	Ganon's Castle
                ' 194	Ganon's Castle Deku Room
                ' 195	Ganon's Castle Light Trial
                ' 196	Ganon's Tower

                ' Ganon's Castle to Deku Room, Light Trial, Tower
                If item("lens of truth") Or My.Settings.setIGCLensless Then addArea(194, asAdult)
                If asAdult And item("lift", 3) Then addArea(195, asAdult)
                If (checkLoc("6611") And checkLoc("6612") And checkLoc("6613") And checkLoc("6614") And checkLoc("6615") And checkLoc("6629")) Or checkLoc("6719") Then addArea(196, asAdult)
            Case 197
                ' MK Entrance to MK and HF
                addAreaExit(0, 1, asAdult) 'addArea(10, asAdult)
                addAreaExit(0, 0, asAdult) 'addArea(7, asAdult)
            Case 198
                ' MK Back Alley (left) to MK, MK Back Alley (right)
                addAreaExit(1, 0, asAdult) 'addArea(10, asAdult)
                addArea(199, asAdult)
            Case 199
                ' MK Back Alley (right) to MK, MK Back Alley (left)
                addAreaExit(1, 1, asAdult) 'addArea(10, asAdult)
                addArea(198, asAdult)
            Case 200
                ' ToT Front to Outside ToT, ToT Behind Door
                addAreaExit(5, 0, asAdult) 'addArea(11, asAdult)
                If item("song of time") Or checkLoc("6427") Then addArea(12, asAdult)
        End Select
    End Sub
    Private Sub addArea(ByVal area As Byte, ByVal asAdult As Boolean)
        ' Depending on if as an adult or not, focus on the correct array
        Dim addArray() As Boolean
        If asAdult Then
            addArray = aReachA
        Else
            addArray = aReachY
        End If
        ' If already true, return
        If addArray(area) Then Return
        ' If not already true, make true and test that new location
        addArray(area) = True
        reachAreas(area, asAdult)
    End Sub
    Private Sub addAreaExit(ByVal area As Byte, ByVal ext As Byte, ByVal asAdult As Boolean)
        ' Depending on if as an adult or not, focus on the correct array
        Dim realArea As Byte = aExitMap(area)(ext)
        addArea(realArea, asAdult)

    End Sub

    Private Function dungeonKeyCounter(ByVal dungeon As Byte, ByVal theseDoors As String) As Boolean
        ' Works to count dungeon keys and unlocked doors
        dungeonKeyCounter = False
        Dim countDoors As Byte = 0

        ' If set to remove small keys, always true. This setting is only for OOTR, as AP handles bit flags better
        If My.Settings.setSmallKeys = 2 Or canDungeon(dungeon) Then Return True

        ' Use the dungeon to set variabled to refocus on dungeon keys array and checkloc for door
        Dim dunKeys As Byte = 0
        Dim dunArea As String = String.Empty
        Select Case dungeon
            Case 3
                ' Forest Temple
                dunKeys = 0
                dunArea = "106"
            Case 4
                ' Fire Temple
                dunKeys = 1
                dunArea = "107"
            Case 5
                ' Water Temple
                dunKeys = 2
                dunArea = "108"
            Case 6
                ' Spirit Temple
                dunKeys = 3
                dunArea = "109"
            Case 7
                ' Shadow Temple
                dunKeys = 4
                dunArea = "110"
            Case 8
                ' Bottom of the Well
                dunKeys = 5
                dunArea = "117"
            Case 10
                ' Gerudo Training Ground
                dunKeys = 6
                dunArea = "112"
            Case 11
                ' Gerudo Training Ground
                dunKeys = 7
                dunArea = "113"
            Case Else
                Return False
        End Select

        ' Make sure doors are at least 1 and even to be safe
        If theseDoors.Length = 0 Or theseDoors.Length Mod 2 = 1 Then Return False

        Dim needKeys As Byte = CByte(theseDoors.Length / 2)
        ' If keys are already enough, return false and skip from having to do the door checks
        If aDungeonKeys(dunKeys) >= needKeys Then Return True

        ' Step through the selected doors, used as 2-digit codes
        For i = 1 To theseDoors.Length Step 2
            ' If the door is checked (opened) then add it to be counted as a key
            If checkLoc(dunArea & Mid(theseDoors, i, 2)) Then incB(countDoors)
        Next

        ' Check if both open doors and keys are enough
        If aDungeonKeys(dunKeys) + countDoors >= needKeys Then Return True
    End Function

    Private Function item(ByVal name As String, Optional rank As Byte = 1) As Boolean
        item = False
        Dim count As Integer = 0
        Select Case LCase(name)
            Case "kokiri sword"
                If canYoung And checkLoc("7416") Then Return True
            Case "deku shield"
                If canYoung And checkLoc("7420") Then Return True
            Case "hylian shield"
                If checkLoc("7421") Then Return True
            Case "mirror shield"
                If canAdult And checkLoc("7422") Then Return True
            Case "goron tunic"
                If canAdult And checkLoc("7425") Then Return True
            Case "zora tunic"
                If canAdult And checkLoc("7426") Then Return True
            Case "iron boots"
                If canAdult And checkLoc("7429") Then Return True
            Case "hover boots"
                If canAdult And checkLoc("7430") Then Return True
            Case "blue fire"
                If allItems.Contains("u") And (allItems.Contains("w") Or aReachA(190) Or aReachA(193) Or aReachA(201) Or aReachY(202) Or aReachA(202)) Then Return True
            Case "bomb", "bombs"
                If allItems.Contains("c") Then Return True
            Case "bombchu", "bombchus"
                If allItems.Contains("j") Then Return True
            Case "boomerang"
                If canYoung And allItems.Contains("o") Then Return True
            Case "bottle"
                If allItems.Contains("u") Then Return True
            Case "bow"
                If canAdult And allItems.Contains("d") Then Return True
            Case "deku nuts"
                If allItems.Contains("b") Then Return True
            Case "deku stick", "deku sticks"
                If canYoung And allItems.Contains("a") Then Return True
            Case "din's fire"
                If canMagic And allItems.Contains("f") Then Return True
            Case "fire arrow", "fire arrows"
                If canAdult And allItems.Contains("d") And allItems.Contains("e") And canMagic Then Return True
            Case "megaton hammer", "hammer"
                If canAdult And allItems.Contains("r") Then Return True
            Case "hookshot"
                If canAdult And allItems.Contains("k") Then Return True
            Case "lens of truth", "lens"
                If canMagic And allItems.Contains("p") Then Return True
            Case "longshot"
                If canAdult And allItems.Contains("l") Then Return True
            Case "membership card"
                If checkLoc("7722") Then Return True
            Case "nayru's love"
                If canMagic And allItems.Contains("t") Then Return True
            Case "ocarina"
                If allItems.Contains("h") Then Return True
            Case "slingshot"
                If canYoung And allItems.Contains("g") Then Return True
            Case "zelda's lullaby"
                If allItems.Contains("h") And aQuestItems(12) Then Return True
            Case "epona's song"
                If allItems.Contains("h") And aQuestItems(13) Then Return True
            Case "saria's song"
                If allItems.Contains("h") And aQuestItems(14) Then Return True
            Case "sun's song"
                If allItems.Contains("h") And aQuestItems(15) Then Return True
            Case "song of time"
                If allItems.Contains("h") And aQuestItems(16) Then Return True
            Case "song of storms"
                If allItems.Contains("h") And aQuestItems(17) Then Return True
            Case "minuet of forest"
                If allItems.Contains("h") And aQuestItems(6) And Not bSongWarps Then Return True
            Case "bolero of fire"
                If allItems.Contains("h") And aQuestItems(7) And Not bSongWarps Then Return True
            Case "serenade of water"
                If allItems.Contains("h") And aQuestItems(8) And Not bSongWarps Then Return True
            Case "requiem of spirit"
                If allItems.Contains("h") And aQuestItems(9) And Not bSongWarps Then Return True
            Case "nocturne of shadow"
                If allItems.Contains("h") And aQuestItems(10) And Not bSongWarps Then Return True
            Case "prelude of light"
                If allItems.Contains("h") And aQuestItems(11) And Not bSongWarps Then Return True
            Case "scarecrow"
                If canAdult And allItems.Contains("h") And checkLoc("6512") Then
                    If rank = 1 And allItems.Contains("k") Then
                        Return True
                    ElseIf rank = 2 And allItems.Contains("l") Then
                        Return True
                    End If
                End If
            Case "dive"
                If checkLoc("7609") Then inc(count)
                If checkLoc("7610") Then inc(count, 2)
                If count >= rank Then Return True
            Case "lift"
                If checkLoc("7606") Then inc(count)
                If checkLoc("7607") Then inc(count, 2)
                If count >= rank Then Return True
            Case "wallet"
                If checkLoc("7612") Then inc(count)
                If checkLoc("7613") Then inc(count, 2)
                If count >= rank Then Return True
            Case "epona"
                If canAdult And allItems.Contains("h") And checkLoc("7713") And checkLoc("6208") Then Return True
            Case Else
                MsgBox("Invalid item: " & name)
        End Select
    End Function
    Private Function canAttackYoung(ByVal countBoomerang As Boolean) As Boolean
        canAttackYoung = False
        ' First make sure you can be young Link
        If Not canYoung Then Return False
        ' Check for Kokiri Sword, Deku Sticks, Fairy Slingshot, can explode, or Din's Fire
        If item("kokiri sword") Or item("deku stick") Or item("slingshot") Or canExplode() Or item("din's fire") Then Return True
        ' If counting Boomerang and you have it
        If countBoomerang And item("boomerang") Then Return True
    End Function
    Private Function canBombGrotto() As Boolean
        ' Check to see if you can open a non-storms grotto
        canBombGrotto = False
        ' If stone of agony is in logic but player does not have it, then return false
        If My.Settings.setStoneOfAgony And Not checkLoc("7721") Then Return False
        ' Really the only difference between Rock and Bomb Grotto is the Stone of Agony check
        If canBreakRocks() Then Return True
    End Function
    Private Function canBreakRocks() As Boolean
        ' Check to see if you can break rocks
        canBreakRocks = False
        ' Check if you can explode things
        If canExplode() Then Return True
        ' Check for Megaton Hammer
        If item("hammer") Then Return True
    End Function
    Private Function canBurnAdult() As Boolean
        canBurnAdult = False
        ' Check for Din's Fire access
        If item("din's fire") Then Return True
        ' Check for Fire Arrow access
        If item("fire arrows") Then Return True
    End Function
    Private Function canBurnYoung() As Boolean
        canBurnYoung = False
        ' Check for Din's Fire access
        If item("din's fire") Then Return True
        ' If sticks are allowed, check they can be used
        If item("deku stick") Then Return True
    End Function
    Private Function canChangeLake() As Boolean
        canChangeLake = False
        If checkLoc("C04") Or checkLoc("1531") Then Return True
    End Function
    Private Function canSpiritShortcut() As Boolean
        canSpiritShortcut = False
        If checkLoc("10909") And checkLoc("10911") And checkLoc("10912") Then Return True
    End Function
    Private Function canDeku() As Boolean
        ' A check if you have anything that can stun a deku scrub
        canDeku = False

        ' Adult Link always has Master Sword and can stun a Deku
        '        If canAdult Then Return True

        ' Check for Kokiri Sword, Deku Shield, or Din's Fire
        If item("kokiri sword") Or item("deku shield") Or item("din's fire") Then Return True
        ' Check for deku sticks, nuts, bombs, slingshot, or boomerang
        If allItems.Contains("a") Then Return True
        If allItems.Contains("b") Then Return True
        If allItems.Contains("c") Then Return True
        If allItems.Contains("g") Then Return True
        If allItems.Contains("o") Then Return True
        ' Check for bombchus and that they are in logic
        If allItems.Contains("j") And My.Settings.setBombchus Then Return True
    End Function
    Private Function canEnterGanonsCastle() As Boolean
        canEnterGanonsCastle = False
        ' OGC access
        If Not aReachA(51) Then Return False
        ' If blocked by spoiler guard
        'If My.Settings.setHideSpoiler And Not checkLoc("6429") Then Return False
        ' Check if the rainbow bridge is there
        If checkLoc("6429") Then Return True
        ' Will be counting various checks
        Dim countChecks As Byte = 0
        Select Case rainbowBridge(0)
            Case 0
                ' Always Open, should never trigger here, the checkloc on 6429 should have caught this
                Return True
            Case 1
                ' Medallions
                For i = 7700 To 7705
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                If countChecks >= rainbowBridge(1) Then Return True
            Case 2
                ' Dungeons (technically combining Stones and Medallions)
                ' Medallions first
                For i = 7700 To 7705
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                ' Spiritual Stones second
                For i = 7718 To 7720
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                If countChecks >= rainbowBridge(1) Then Return True
            Case 3
                ' Spiritual Stones
                For i = 7718 To 7720
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                If countChecks >= rainbowBridge(1) Then Return True
            Case 4
                ' Vanilla: Spirit Medallion, Shadow Medallion, and Light Arrows
                If checkLoc("7703") And checkLoc("7704") And allItems.Contains("s") Then Return True
            Case 5
                ' Gold Skulltulas
                If goldSkulltulas >= rainbowBridge(1) Then Return True
        End Select
    End Function
    Private Function canExplode() As Boolean
        ' Check to see if you can explode things
        canExplode = False
        ' Check for bombs
        If allItems.Contains("c") Then Return True
        ' Check for bombchus and that they are in logic
        If allItems.Contains("j") And My.Settings.setBombchus Then Return True
    End Function
    Private Function canLens() As Boolean
        ' Checks for Lens of Truth requirement
        canLens = False
        ' If not selected as in logic, return true
        If Not My.Settings.setLensOfTruth Then Return True
        ' If it is in logic, check that player has Lens of Truth and magic
        If item("lens of truth") Then Return True
    End Function
    Private Function canLACS() As Boolean
        ' Checks if you have what it takes to get Light Arrow Cutscene
        canLACS = False
        ' Check for ToT access as adult Link
        If Not aReachA(11) Then Return False
        ' If hiding information, treat like vanilla
        If My.Settings.setHideSpoiler Then
            If checkLoc("7703") And checkLoc("7704") Then Return True
            Return False
        End If

        ' Determine what is needed for LACS
        If aAddresses(1) = 0 Then
            ' No address entered, treat as default
            If checkLoc("7703") And checkLoc("7704") Then Return True
            Return False
        End If

        Dim conditionLACS As Byte
        Dim countLACS As Byte
        Dim countChecks As Byte = 0
        If isSoH Then   ' only vanilla LACS currently
            conditionLACS = 0
            countLACS = 0
        Else
            conditionLACS = CByte(goRead(aAddresses(1), 1))
            countLACS = CByte(goRead(aAddresses(2), 1))
        End If
        Select Case conditionLACS
            Case 0
                ' Vanilla: Spirit Medallion and Shadow Medallion
                If checkLoc("7703") And checkLoc("7704") Then Return True
            Case 1
                ' Medallions
                For i = 7700 To 7705
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                If countChecks >= countLACS Then Return True
            Case 2
                'Dungeons
                ' Medallions first
                For i = 7700 To 7705
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                ' Spiritual Stones second
                For i = 7718 To 7720
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                If countChecks >= countLACS Then Return True
            Case 3
                'Spiritual Stones
                For i = 7718 To 7720
                    If checkLoc(i.ToString) Then incB(countChecks)
                Next
                If countChecks >= countLACS Then Return True
            Case 4
                ' Gold Skulltula Tokens
                If goldSkulltulas >= countLACS Then Return True
        End Select
    End Function
    Private Function canFewerGoron() As Boolean
        canFewerGoron = False
        If My.Settings.setFewerTunics Or (canAdult And checkLoc("7425")) Then Return True
    End Function
    Private Function canFewerZora() As Boolean
        canFewerZora = False
        If My.Settings.setFewerTunics Or (canAdult And checkLoc("7426")) Then Return True
    End Function
    Private Function canGetBugs() As Boolean
        ' Checks if you can reach bugs
        canGetBugs = False
        ' Auto False if you do not have a bottle (does not count Ruto's Letter)
        If Not allItems.Contains("u") Then Return False

        ' KV Main, GY Main, and DC as both ages, or HC, DMT Upper, and GV Hyrule Side as young Link, with just a bottle
        If aReachA(15) Or aReachY(15) Or aReachA(18) Or aReachY(18) Or aReachA(50) Or aReachY(50) Or aReachY(13) Or aReachY(19) Or aReachY(44) Then Return True

        ' GC Main as both ages, break rocks or silver guantlets
        If aReachA(29) Or aReachY(29) Then
            If canBreakRocks() Or item("lift", 2) Then Return True
        End If

        ' Storms Grottos with bugs
        If canStormsGrotto() Then
            ' KF Main or DMT Lower as both ages
            ' Adult Link automatically passes
            If aReachA(0) Or aReachA(20) Then Return True
            ' Young Link needs sticks, boomerang, explosions, or the Kokiri Sword
            If aReachY(0) Or aReachY(20) Then
                If item("deku stick") Or item("boomerang") Or canExplode() Or item("kokiri sword") Then Return True
            End If
        End If

        ' LW Front, HF, ZR Main, or LH as young Link
        If aReachY(2) Or aReachY(7) Or aReachY(35) Or aReachY(42) Then
            ' Need sticks, boomerang, explosions, or the Kokiri Sword
            If item("deku stick") Or item("boomerang") Or canExplode() Or item("kokiri sword") Then Return True
        End If
    End Function
    Private Function canHoverTricks() As Boolean
        canHoverTricks = False
        If checkLoc("7430") And My.Settings.setHoverTricks Then Return True
    End Function
    Private Function canProjectile(Optional onlyAge As Byte = 1) As Boolean
        canProjectile = False
        If canExplode() Then Return True
        If onlyAge >= 1 And item("bow") Or item("hookshot") Then Return True
        If onlyAge <= 1 And item("boomerang") Or item("slingshot") Then Return True
    End Function
    Private Function canStormsGrotto() As Boolean
        ' Check to see if you can open a storms grotto
        canStormsGrotto = False
        ' If stone of agony is in logic but player does not have it, then return false
        If My.Settings.setStoneOfAgony And Not checkLoc("7721") Then Return False
        ' Check if player has an ocarina and that they have song of storms
        If allItems.Contains("h") And checkLoc("7717") Then Return True
    End Function
    Private Function canTimeChange() As Boolean
        canTimeChange = False
        ' Check Sun's Song
        If item("sun's song") Then Return True

        ' Going to only check 3 of the areas because the others all depend on Hyrule Field
        If aReachA(7) Or aReachY(7) Or aReachA(20) Or aReachY(20) Or aReachA(50) Or aReachY(50) Then Return True
        ' HF, DMT, DC || HC, ZR, LH, GV
    End Function
    Private Function canMagicBean(ByVal bean As String) As Boolean
        canMagicBean = False
        ' First check if the Magic Bean has been planted
        If checkLoc(bean) Then
            'MsgBox("checkloc: " & bean)
            Return True
        End If
        ' Next make sure young Link can reach the soil
        Select Case bean
            Case "B0"
                If Not aReachY(2) Then Return False
            Case "B1"
                If Not aReachY(50) Then Return False
            Case "B2"
                If Not (aReachY(20) And (item("lift") Or canExplode())) Then Return False
            Case "B3"
                If Not aReachY(27) Then Return False
            Case "B4"
                If Not aReachY(18) Then Return False
            Case "202"
                If Not aReachY(42) Then Return False
        End Select
        ' Check if you have a Magic Bean in your inventory
        If magicBeans > 0 Then
            'MsgBox("Magic Beans: " & magicBeans.ToString)
            Return True
        End If
        ' If young Link can reach ZR Main, and the salesman is not shuffled, you can buy beans
        If aReachY(35) And Not My.Settings.setBeans Then
            'MsgBox("aReachY(35): " & aReachY(35).ToString)
            Return True
        End If
    End Function
    Private Function inlogicBombchus() As Boolean
        inlogicBombchus = False
        ' Check for bombchu logic setting
        If My.Settings.setBombchus Then
            ' If bombchus are in logic, check for bombchus
            If allItems.Contains("j") Then Return True
        Else
            ' If bombchus are not in logic, check for bomb bag 1 or bomb bag 2
            If checkLoc("7603") Or checkLoc("7604") Then Return True
        End If
    End Function

    Private Function checkLogic(ByVal logicKey As String, ByVal zone As Byte) As Byte ' Boolean
        ' Checks the logic key to see if it is available. Start with false
        checkLogic = 0 ' TESTLOGIC False
        Dim canDoThis As Boolean = True

        ' Check we can reach the overworld zone, or dungeon area
        If aReachA(zone) Or aReachY(zone) Then canDoThis = True
        ' If ER, then check the visited array
        If Not iER = 0 Then
            Dim iVisited As Byte = zone2map(zone)
            If iVisited = 255 Then
                canDoThis = False
            Else
                If aVisited(iVisited) = False Then canDoThis = False
            End If
        End If

        ' Do not convert dungeons or quests
        convertLogic(logicKey)


        ' If either of those fail, return false
        If Not canDoThis Then Return 0 ' TESTLOGIC False
        canDoThis = False
        ' If there is no logic key, then item is always accessable, return true
        ' TESTLOGIC If logicKey = String.Empty Then Return True

        ' H is for Hover Boot tricks
        If logicKey.Contains("H") And canHoverTricks() Then logicKey = logicKey.Replace("H", "")
        ' I is for checking you can damage as young Link
        If logicKey.Contains("I") And canAttackYoung(False) Then logicKey = logicKey.Replace("I", "")
        ' J is for checking you can attack as young Link
        If logicKey.Contains("J") And canAttackYoung(True) Then logicKey = logicKey.Replace("J", "")
        ' M is for Magic access
        If logicKey.Contains("M") And canMagic Then logicKey = logicKey.Replace("M", "")
        ' N is for Night/Day cycling
        If logicKey.Contains("N") And canTimeChange() Then logicKey = logicKey.Replace("N", "")
        ' S is for checking if you can stun a deku scrub
        If logicKey.Contains("S") And canDeku() Then logicKey = logicKey.Replace("S", "")
        ' T is for checking if you need the lens of truth
        If logicKey.Contains("T") And canLens() Then logicKey = logicKey.Replace("T", "")
        ' U is for Saria's Gift, either opened Forest setting or cleared Deku Tree
        If logicKey.Contains("U") And My.Settings.setOpenKF Or checkLoc("6223") Or bSpawnWarps Or bSongWarps Then logicKey = logicKey.Replace("U", "")
        ' x is for checking if you can explode things
        If logicKey.Contains("x") And canExplode() Then logicKey = logicKey.Replace("x", "")
        ' X is for checking if Bombchus are in logic. If they are, test them, if not, test bomb bag
        If logicKey.Contains("X") And inlogicBombchus() Then logicKey = logicKey.Replace("X", "")
        ' Y is for young Link access
        If logicKey.Contains("Y") Then
            'If zone = 99 Then
            'If canYoung Then logicKey = logicKey.Replace("Y", "")
            'Else
            If aReachY(zone) Then logicKey = logicKey.Replace("Y", "")
            'End If
        End If
        ' Z is for adult Link access
        If logicKey.Contains("Z") Then
            'If zone = 99 Then
            'If canAdult Then logicKey = logicKey.Replace("Z", "")
            'Else
            If aReachA(zone) Then logicKey = logicKey.Replace("Z", "")
            'End If
        End If

        Dim tempString As String = String.Empty

        ' Step through the a-w for the basic inventory check
        For i = 97 To 119
            tempString = Chr(i)
            ' If it exists in both need and have, remove it from the need list
            Select Case tempString
                Case "a", "g", "o"
                    ' Young only item checks
                    If logicKey.Contains(tempString) And allItems.Contains(tempString) And canYoung Then logicKey = logicKey.Replace(tempString, "")
                Case "d", "l", "k", "r"
                    ' Adult only item checks
                    If logicKey.Contains(tempString) And allItems.Contains(tempString) And canAdult Then logicKey = logicKey.Replace(tempString, "")
                Case "e", "f", "m", "n", "p", "s", "t"
                    ' Magic checks
                    If logicKey.Contains(tempString) And allItems.Contains(tempString) And canMagic Then logicKey = logicKey.Replace(tempString, "")
                Case "q"
                    ' Magic Bean Checks
                    If logicKey.Contains(tempString) And allItems.Contains(tempString) And magicBeans > 0 And canYoung Then logicKey = logicKey.Replace(tempString, "")
                Case Else
                    ' Normal checks
                    If logicKey.Contains(tempString) And allItems.Contains(tempString) Then logicKey = logicKey.Replace(tempString, "")
            End Select
        Next

        ' Midway check to see if we can leave this mess of a function early
        ' TESTLOGIC If isLogicGood(logicKey) Then Return True

        ' Check for various keys
        For i = 1 To logicKey.Length
            Select Case Mid(logicKey, i, 1)
                Case "G"
                    ' G# = Test Types:  00 = Break Rock         01 = Break Rock with SoA    02 = Storms Grotto      03 = SoA Check          09 = Fewer Goron Tunic      0A = Fewer Zora Tunic       0D = Can Get Bugs
                    '                   14 = SpT Shortcut       15 = SpT Lensless           16 = Can Change Lake    17 = Blue Fire Access 
                    '
                    '                   20 = Can Projectile Y   21 = Can Projectile A       22 = Can LACS           23 = Burn Young         24 = Burn Adult
                    '
                    '                   30 = Magic Beans LW     31 = Magic Beans DC         32 = Magic Beans DMT    33 = Magic Beans DMC    34 = Magic Beans GY         35 = Magic Beans LH
                    '                   
                    '                   A# = Logics
                    '
                    '                   A0 = DC Jump        A1 = FoT Ledge          A2 = FoT Doorframe      A3 = FiT Rusted Switches        A4 = FiT Flame Maze Skip        A5 = SpT Young Bombchu          A6 = SpT MQ Sun Block without Song
                    '                   A7 = ShT Umbrella   A8 = BotW Dead Hand     A9 - MQ Spirit Trial Rusted Switches                    AA = BotW MQ Pits               AB = Ganon's Castle Lensess     AC = FiT Scarecrow
                    '                   AD = GTG Lensless
                    '
                    '                   B# = Boss Key: # Dungeon
                    '                   C0 = Fire Temple Keys: 3        C1 = Fire Temple Keys: 4    C2 = Fire Temple Keys: 5    C3 = Fire Temple MQ Keys: 3     C4 = Fire Temple MQ Keys: 4     C5 = Water Temple Pillar
                    '                   C6 = Water Temple BK Chest      C7 = Water Temple MQ All    C8 = Sp. Temple Young All   C9 = Sp. Temple Young Last Key  CA = Sp. Temple MQ Young Door 1 CB = Sp Temple MQ Statue Left Door 1
                    '                   CC = Sp. Temple MQ 9 Thrones    CD = Sp. Temple MQ Symphony CE = Sh. Temple MQ Spikes   CF = BotW Right                 D0 = BotW Centre Left           D1 = BotW Centre Right
                    '                   D2 = BotW MQ Left Side          D3 = BotW MQ Right Side     D4 = GTG Hidden Ceiling     D5 = GTG 1st Chest              D6 = GTG 2nd Chest              D7 = GTG 3rd Chest
                    '                   D8 = GTG Final Chest            D9 = GTG MQ First Door      DA = GTG MQ All Doors       DB = GC Light Trail First Door  DC = Forest Temple MQ Keys: 2   DD = Fire Temple Keys: 6
                    '
                    '                   Removed
                    '                   03 = Ascending Death Mountain   04 = Crater Entrance Check      05 = Half Crater Climb  06 = Full Crater Climb  07 = LW to ZR           08 = Can Buy Beans      
                    '                   0B = Entrance Check GV          0C = Reach Structure
                    '                   0E = Young ZD   0F = Adult ZD   10 = Young DC                   11 = Adult DC           12 = Young SpT Centre   13 = Adult SpT Centre
                    '                   15 = Young ZF   16 = Adult ZF   18 = Can Reach Bean Salesman    19 = KV Entrance        20 = Adult SFM (Never actually put it in...)

                    tempString = Mid(logicKey, i + 1, 2)
                    canDoThis = False
                    ' Depending on the grotto type, test it to different grotto tests
                    Select Case tempString
                        Case "00"
                            ' Needs hammer, bomb, or bombchus if they are in logic
                            canDoThis = canBreakRocks()
                        Case "01"
                            ' Same thing as canBreakRocks, but checking for stone of agony if in logic
                            canDoThis = canBombGrotto()
                        Case "02"
                            ' Needs an ocarina, song of storms, and stone of agony if in logic
                            canDoThis = canStormsGrotto()
                        Case "03"
                            If Not My.Settings.setStoneOfAgony Or checkLoc("7721") Then canDoThis = True
                        Case "09"
                            canDoThis = canFewerGoron()
                        Case "0A"
                            canDoThis = canFewerZora()
                        Case "0D"
                            canDoThis = canGetBugs()
                        Case "14"
                            canDoThis = canSpiritShortcut()
                        Case "15"
                            If My.Settings.setSpTLensless Or item("lens of truth") Then canDoThis = True
                        Case "16"
                            canDoThis = canChangeLake()
                        Case "17"
                            canDoThis = item("blue fire")
                        Case "20"
                            canDoThis = canProjectile(0)
                        Case "21"
                            canDoThis = canProjectile(2)
                        Case "22"
                            canDoThis = canLACS()
                        Case "23"
                            canDoThis = canBurnYoung()
                        Case "24"
                            canDoThis = canBurnAdult()
                        Case "30"
                            canDoThis = canMagicBean("B0")
                        Case "31"
                            canDoThis = canMagicBean("B1")
                        Case "32"
                            canDoThis = canMagicBean("B2")
                        Case "33"
                            canDoThis = canMagicBean("B3")
                        Case "34"
                            canDoThis = canMagicBean("B4")
                        Case "35"
                            canDoThis = canMagicBean("202")
                        Case "A0"
                            canDoThis = My.Settings.setDCSpikeJump
                        Case "A1"
                            canDoThis = My.Settings.setFoTLedge
                        Case "A2"
                            canDoThis = My.Settings.setFoTFrame
                        Case "A3"
                            canDoThis = My.Settings.setFiTRustedSwitches
                        Case "A4"
                            canDoThis = My.Settings.setFiTMaze
                        Case "A5"
                            canDoThis = My.Settings.setSpTBombchu
                        Case "A6"
                            canDoThis = My.Settings.setSpTMQSunRoom
                        Case "A7"
                            canDoThis = My.Settings.setShTUmbrella
                        Case "A8"
                            canDoThis = My.Settings.setBotWDeadHand
                        Case "A9"
                            canDoThis = My.Settings.setIGCMQRustedSwitches
                        Case "AA"
                            canDoThis = My.Settings.setBotWMQPits
                        Case "AB"
                            If My.Settings.setIGCLensless Or item("lens of truth") Then canDoThis = True
                        Case "AC"
                            If item("scarecrow") Or (My.Settings.setFiTScarecrow And item("longshot")) Then canDoThis = True
                        Case "AD"
                            If My.Settings.setGTGLensless Or item("lens of truth") Then canDoThis = True
                        Case "B0", "B1", "B2", "B3", "B4"
                            canDoThis = aBossKeys(CByte(Mid(tempString, 2)))
                            If Not canDoThis Then
                                ' If the key was not found, check each boss door
                                Select Case Mid(tempString, 2)
                                    Case "0"    ' Forest
                                        If checkLoc("10620") Then canDoThis = True
                                    Case "1"    ' Fire
                                        If checkLoc("10720") Then canDoThis = True
                                    Case "2"    ' Water
                                        If checkLoc("10820") Then canDoThis = True
                                    Case "3"    ' Spirit
                                        If checkLoc("10920") Then canDoThis = True
                                    Case "4"    ' Shadow
                                        If checkLoc("11020") Then canDoThis = True
                                    Case Else
                                        MsgBox("G" & tempString)
                                End Select
                            End If
                        Case "C0"
                            canDoThis = dungeonKeyCounter(4, "293024")
                        Case "C1"
                            canDoThis = dungeonKeyCounter(4, "29302427")
                        Case "C2"
                            canDoThis = dungeonKeyCounter(4, "2930242726")
                        Case "C3"
                            canDoThis = dungeonKeyCounter(4, "302624")
                        Case "C4"
                            canDoThis = dungeonKeyCounter(4, "30262427")
                        Case "C5"
                            canDoThis = dungeonKeyCounter(5, "06")
                        Case "C6"
                            canDoThis = dungeonKeyCounter(5, "0109")
                        Case "C7"
                            canDoThis = dungeonKeyCounter(5, "2102")
                        Case "C8"
                            canDoThis = dungeonKeyCounter(6, "3021")
                        Case "C9"
                            canDoThis = dungeonKeyCounter(6, "21")
                        Case "CA"
                            canDoThis = dungeonKeyCounter(6, "30")
                        Case "CB"
                            canDoThis = dungeonKeyCounter(6, "27")
                        Case "CC"
                            canDoThis = dungeonKeyCounter(6, "272801")
                        Case "CD"
                            canDoThis = dungeonKeyCounter(6, "03")
                        Case "CE"
                            canDoThis = dungeonKeyCounter(7, "27")
                        Case "CF"
                            canDoThis = dungeonKeyCounter(8, "27")
                        Case "D0"
                            canDoThis = dungeonKeyCounter(8, "28")
                        Case "D1"
                            canDoThis = dungeonKeyCounter(8, "29")
                        Case "D2"
                            canDoThis = dungeonKeyCounter(8, "20")
                        Case "D3"
                            canDoThis = dungeonKeyCounter(8, "21")
                        Case "D4"
                            canDoThis = dungeonKeyCounter(10, "01")
                        Case "D5"
                            canDoThis = dungeonKeyCounter(10, "0109")
                        Case "D6"
                            canDoThis = dungeonKeyCounter(10, "01090405")
                        Case "D7"
                            canDoThis = dungeonKeyCounter(10, "0109040506")
                        Case "D8"
                            canDoThis = dungeonKeyCounter(10, "01090405060723")
                        Case "D9"
                            canDoThis = dungeonKeyCounter(10, "29")
                        Case "DA"
                            canDoThis = dungeonKeyCounter(10, "292023")
                        Case "DB"
                            canDoThis = dungeonKeyCounter(11, "30")
                        Case "DC"
                            canDoThis = dungeonKeyCounter(3, "0602")
                        Case "DD"
                            canDoThis = dungeonKeyCounter(4, "293024272631")

                        Case Else
                            MsgBox(logicKey & ": G" & tempString)
                    End Select
                    If canDoThis Then
                        logicKey = logicKey.Replace("G" & tempString, "")
                        ' Step i back a step to not miss from the removed text
                        dec(i)
                    Else
                        inc(i, 2)
                    End If
                Case "K"
                    ' K# = Gold Skulltulas x10
                    tempString = Mid(logicKey, i + 1, 1)
                    ' Compare the number of gold skulltulas to the need times 10
                    If goldSkulltulas >= CByte("&H" & tempString) * 10 Then
                        logicKey = logicKey.Replace("K" & tempString, "")
                        ' Step i back a step to not miss from the removed text
                        dec(i)
                    Else
                        inc(i)
                    End If
                Case "L"
                    ' L###### is for checking a key by loc value
                    tempString = Mid(logicKey, i + 1, 5)
                    ' Since a key's loc can be 3-5 characters long, Ls are used to fill the space, remove them from the check request
                    Dim tempValue As Integer = CInt(Val(tempString.Replace("L", "")))
                    canDoThis = False
                    Select Case tempValue
                        Case 7706 To 7711
                            If Not bSongWarps Then
                                If checkLoc(tempValue.ToString) Then canDoThis = True
                            End If
                        Case Else
                            If checkLoc(tempValue.ToString) Then canDoThis = True
                    End Select
                    If canDoThis Then
                        logicKey = logicKey.Replace("L" & tempString, "")
                        ' Step i back a step to not miss from the removed text
                        dec(i)
                    Else
                        inc(i, 5)
                    End If
                Case "Q"
                    ' Q### = Check DAC #|## = Dungeon|Area
                    'tempString = Mid(logicKey, i + 1, 3)
                    'If arrDAC(CByte("&H" & (Mid(tempString, 1, 1))))(CByte("&H" & (Mid(tempString, 2)))) Then
                    'logicKey = logicKey.Replace("Q" & tempString, "")
                    'i = 1 - 1
                    'Else
                    'i = i + 3
                    'End If
                    ' Q#### = Check aReach #|### = Age (0 = Y, 1 = A) | aReach 
                    canDoThis = False
                    tempString = Mid(logicKey, i + 1, 4)
                    Select Case Mid(tempString, 1, 1)
                        Case "0"
                            If aReachY(CInt(Mid(tempString, 2))) Then canDoThis = True
                        Case "1"
                            If aReachA(CInt(Mid(tempString, 2))) Then canDoThis = True
                    End Select
                    If canDoThis Then
                        logicKey = logicKey.Replace("Q" & tempString, "")
                        dec(i)
                    Else
                        inc(i, 4)
                    End If

                Case "V"
                    ' V## = Value of tests
                    tempString = Mid(logicKey, i + 1, 2)

                    ' Decern which item this is for
                    Dim tempItem As String = String.Empty
                    Select Case Mid(tempString, 1, 1)
                        Case "0"
                            tempItem = "lift"
                        Case "1"
                            tempItem = "dive"
                        Case "2"
                            tempItem = "wallet"
                    End Select

                    ' Compare item check and back up i
                    If item(tempItem, CByte(Mid(tempString, 2))) Then
                        logicKey = logicKey.Replace("V" & tempString, "")
                        dec(i)
                    Else
                        inc(i, 2)
                    End If
                Case "y", "z"
                    ' y# = Young Link Quest Item
                    ' z# = Adult Link Quest Item
                    tempString = Mid(logicKey, i, 2)
                    If allItems.Contains(tempString) Then
                        logicKey = logicKey.Replace(tempString, "")
                        dec(i)
                    Else
                        inc(i)
                    End If
            End Select
        Next
        ' After all the removals, check the key
        ' TESTLOGIC Return isLogicGood(logicKey)

        If logicKey.Contains("|") Then
            checkLogic = isLogicGoodx(logicKey)
        Else
            If isLogicGood(logicKey) Then checkLogic = 3
        End If

    End Function
    Private Function isLogicGood(ByVal logicString As String) As Boolean
        ' If logic key is empty, contains a '..', or starts/end with '.' then all or some logic must have been cleared, return true
        isLogicGood = False

        If logicString = String.Empty Then Return True
        If logicString.Contains("..") Or Mid(logicString, 1, 1) = "." Or Mid(logicString, logicString.Length, 1) = "." Then Return True
    End Function
    Private Sub convertLogic(ByRef logic As String)
        If logic = String.Empty Then
            ' Make any blank one fully opened
            logic = "Y|Z"
        Else
            ' Make into 2 split arrays
            Dim sideY() As String = logic.Split(CChar("."))
            Dim sideA() As String = logic.Split(CChar("."))

            Dim newLogicY As String = String.Empty
            Dim newLogicA As String = String.Empty
            Dim added As String = String.Empty
            ' Check young Link side first
            For i = 0 To sideY.Length - 1

                ' If no age identifier, add "Y" to it
                If Not sideY(i).Contains("Y") And Not sideY(i).Contains("Z") Then sideY(i) = "Y" & sideY(i)
                ' If just adult Link check, remove it
                If sideY(i).Contains("Z") And Not sideY(i).Contains("Y") Then sideY(i) = String.Empty
                ' Just young Link checks
                If Not (sideY(i).Contains("Z") Or sideY(i).Contains("Q1")) Then
                    If sideY(i).Contains("d") Then sideY(i) = String.Empty
                    If sideY(i).Contains("l") Then sideY(i) = String.Empty
                    If sideY(i).Contains("k") Then sideY(i) = String.Empty
                    If sideY(i).Contains("r") Then sideY(i) = String.Empty
                    If sideY(i).Contains("G17") Then sideY(i) = String.Empty
                    If sideY(i).Contains("G21") Then sideY(i) = String.Empty
                    If sideY(i).Contains("G24") Then sideY(i) = String.Empty
                    ' G00 is for breaking rocks, which is technically explosions and hammer. Change to just explosion for young Link checks
                    If sideY(i).Contains("G00") Then sideY(i) = sideY(i).Replace("G00", "x")
                    ' Same as G00 but with Stone of Agony check. Turn into a stand-alone SoA check and then explosions
                    If sideY(i).Contains("G01") Then sideY(i) = sideY(i).Replace("G01", "G03x")
                End If
                ' Rebuild the logic side for young Link
                If Not sideY(i) = String.Empty Then
                    newLogicY = newLogicY & added & sideY(i)
                    added = "."
                End If
            Next
            ' Reset the added "."
            added = String.Empty
            ' Now do adult Link side
            For i = 0 To sideA.Length - 1
                ' If no age identifier, add "Z" to it
                If Not sideA(i).Contains("Y") And Not sideA(i).Contains("Z") Then sideA(i) = "Z" & sideA(i)
                ' If just young Link check, remove it
                If sideA(i).Contains("Y") And Not sideY(i).Contains("Z") Then sideA(i) = String.Empty
                ' Just adult Link checks
                If Not sideA(i).Contains("Y") Then
                    If sideA(i).Contains("a") Then sideA(i) = String.Empty
                    If sideA(i).Contains("g") Then sideA(i) = String.Empty
                    If sideA(i).Contains("o") Then sideA(i) = String.Empty
                    If sideA(i).Contains("I") Then sideA(i) = String.Empty
                    If sideA(i).Contains("J") Then sideA(i) = String.Empty
                    If sideA(i).Contains("G20") Then sideA(i) = String.Empty
                    If sideA(i).Contains("G23") Then sideA(i) = String.Empty
                    ' "S" is for young Link being able to stun a deku scrub. No need to check it for adult Link who will always have the Master Sword
                    If sideA(i).Contains("S") Then sideA(i) = sideA(i).Replace("S", "")
                End If

                ' Rebuild the logic side for young Link
                If Not sideA(i) = String.Empty Then
                    newLogicA = newLogicA & added & sideA(i)
                    added = "."
                End If
            Next

            ' If either side is empty, change it to a "W" for impossible logic
            If newLogicY = String.Empty Then newLogicY = "W"
            If newLogicA = String.Empty Then newLogicA = "W"
            ' Make it a full single logic string
            logic = newLogicY & "|" & newLogicA
        End If

        'If tempLogic.Contains("Y") And Not tempLogic.Contains("Z") Then logic = logic & "|W"
        'If tempLogic.Contains("Z") And Not tempLogic.Contains("Y") Then logic = "W|" & logic



    End Sub

    Private Function isLogicGoodx(ByVal logicString As String) As Byte
        isLogicGoodx = 0
        Dim testLogic As String = String.Empty
        For i As Byte = 0 To 1
            testLogic = Split(logicString, "|")(i)
            If testLogic = String.Empty Then
                isLogicGoodx = CByte(isLogicGoodx + i + 1)
            Else
                If testLogic.Contains("..") Or Mid(testLogic, 1, 1) = "." Or Mid(testLogic, testLogic.Length, 1) = "." Then isLogicGoodx = CByte(isLogicGoodx + i + 1)
            End If
        Next
        Return isLogicGoodx
    End Function

    ' Make the keys for each area
    Private Sub makeKeysKF(ByRef tK As Integer)
        ' Set up keys for Kokiri 
        ' 00 KF Main
        ' 01 KF Deku Tree
        ' 08 KF Between Bridges

        With aKeys(tK)
            .loc = "4500"
            .area = "KF"
            .zone = 0
            .name = "Mido's House Upper Left Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4501"
            .area = "KF"
            .zone = 0
            .name = "Mido's House Upper Right Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4502"
            .area = "KF"
            .zone = 0
            .name = "Mido's House Lower Left Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4503"
            .area = "KF"
            .zone = 0
            .name = "Mido's House Lower Right Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "5100"
            .area = "KF"
            .zone = 0
            .name = "Kokiri Sword Chest"
            .logic = "Y"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4612"
            .area = "KF"
            .zone = 0
            .name = "Storms Grotto Chest"
            .logic = "G02"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8100"
            .area = "KF"
            .zone = 0
            .name = "Soft Soil Near Shoppe"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8101"
            .area = "KF"
            .zone = 0
            .name = "On Know-it-All Brother's House (N)"
            .gs = True
            .logic = "YJN"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8102"
            .area = "KF"
            .zone = 0
            .name = "Above the House of Twins (N)"
            .gs = True
            .logic = "ZNk"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "10024"
            .area = "KF"
            .zone = 0
            .name = "Link's House"
            .cow = True
            .logic = "ZLL6214hLL7713"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6220"
            .area = "EVENT"
            .name = "Mido Moved From Deku Path"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6223"
            .area = "EVENT"
            .name = "Kid Moved From Exit to MF"
        End With
        inc(tK)
    End Sub
    Private Sub makeKeysLW(ByRef tK As Integer)
        ' Set up keys for the Lost Woods
        ' 02 LW Front
        ' 03 LW Behind Mido
        ' 04 LW Bridge

        With aKeys(tK)
            .loc = "7202"
            .area = "LW"
            .zone = 2
            .name = "Deku Scrub Near Bridge"
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4620"
            .area = "LW"
            .zone = 2
            .name = "Near Shortcuts Grotto Chest"
            .logic = "G00"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "7203"
            .area = "LW"
            .zone = 3
            .name = "Deku Scrub in Grotto"
            .logic = "YG00S.ZG00"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6717"
            .area = "LW"
            .zone = 8
            .name = "Gift from Saria"
            .logic = "YU.Z"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6813"
            .area = "LW"
            .zone = 2
            .name = "Target in Woods"
            .logic = "Yg"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6807"
            .area = "LW"
            .zone = 2
            .name = "Ocarina Memory Game"
            .logic = "Yh"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6806"
            .area = "LW"
            .zone = 2
            .name = "Skull Kid"
            .logic = "YhLL7714"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = locSwap(13)
            .area = "LW"
            .zone = 3
            .name = "Deku Theatre Skull Mask"
            .logic = "YLL6908.Yy5"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6815"
            .area = "LW"
            .zone = 3
            .name = "Deku Theatre Mask of Truth"
            .logic = "YLL6911.YyB"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8108"
            .area = "LW"
            .zone = 2
            .name = "Soft Soil Near Bridge"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8109"
            .area = "LW"
            .zone = 3
            .name = "Soft Soil Near Theatre"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8110"
            .area = "LW"
            .zone = 3
            .name = "Above Theatre (N)"
            .gs = True
            .logic = "ZNG30"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9810"
            .area = "LW"
            .zone = 2
            .name = "Near Bridge"
            .scrub = True
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9801"
            .area = "LW"
            .zone = 3
            .name = "Outside Theatre Right"
            .scrub = True
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9802"
            .area = "LW"
            .zone = 3
            .name = "Outside Theatre Left"
            .scrub = True
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9311"
            .area = "LW"
            .zone = 3
            .name = "Grotto Left"
            .scrub = True
            .logic = "YG00S.ZG00"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9304"
            .area = "LW"
            .zone = 3
            .name = "Grotto Right"
            .scrub = True
            .logic = "YG00S.ZG00"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "B0"
            .area = "EVENT"
            .name = "Beans Planted Near Theatre"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6226"
            .area = "EVENT"
            .name = "Play Saria's Song for Mido"
        End With
        inc(tK)
    End Sub
    Private Sub makeKeysSFM(ByRef tk As Integer)
        ' Set up keys for the Sacred Forest Meadow
        ' 05 SFM Main
        ' 06 SFM Temple Ledge

        With aKeys(tk)
            .loc = "4617"
            .area = "SFM"
            .zone = 5
            .name = "Wolfos Grotto Chest"
            .logic = "ZG01.YG03xa.YG03xf.YG03xg.YG03xLL7416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(6)
            .area = "SFM"
            .zone = 5
            .name = "Song from Saria"
            .logic = "YLL6416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6400"
            .area = "SFM"
            .zone = 5
            .name = "Song from Sheik"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8111"
            .area = "SFM"
            .zone = 5
            .name = "On East Wall (N)"
            .gs = True
            .logic = "ZNk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9009"
            .area = "SFM"
            .zone = 5
            .name = "Storms Grotto Front"
            .scrub = True
            .logic = "YG02S.G02Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9008"
            .area = "SFM"
            .zone = 5
            .name = "Storms Grotto Back"
            .scrub = True
            .logic = "YG02S.G02Z"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysHF(ByRef tk As Integer)
        ' Set up keys for Hyrule Field
        ' 07 HF

        With aKeys(tk)
            .loc = "4600"
            .area = "HF"
            .zone = 7
            .name = "Near Market Grotto Chest"
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4602"
            .area = "HF"
            .zone = 7
            .name = "Southeast Grotto Chest"
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4603"
            .area = "HF"
            .zone = 7
            .name = "Open Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6827"
            .area = "HF"
            .zone = 7
            .name = "Deku Scrub Inside Fence Grotto"
            .logic = "YG03xS.G01Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1901"
            .area = "HF"
            .zone = 7
            .name = "Tektite Grotto Piece of Heart"
            .logic = "ZG01LL7429.G01LL7610"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "103"
            .area = "HF"
            .zone = 7
            .name = "Get Ocarina of Time"
            .logic = "YLL7718LL7719LL7720"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(9)
            .area = "HF"
            .zone = 7
            .name = "Song from Ocarina of Time"
            .logic = "YLL7718LL7719LL7720"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8016"
            .area = "HF"
            .zone = 7
            .name = "Grotto Near Gerudo Valley"
            .gs = True
            .logic = "YG03xfo.ZrG24k"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8017"
            .area = "HF"
            .zone = 7
            .name = "Grotto Near Kakariko Village"
            .gs = True
            .logic = "YG01o.ZG01k"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1925"
            .area = "HF"
            .zone = 7
            .name = "Grotto Near Gerudo Valley"
            .cow = True
            .logic = "YG03xfhLL7713.ZrG24hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8803"
            .area = "HF"
            .zone = 7
            .name = "Inside Fence Grotto"
            .scrub = True
            .logic = "YG03xS.G01Z"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysLLR(ByRef tk As Integer)
        ' Set up keys for Lon Lon Ranch
        ' 09 LLR

        With aKeys(tk)
            .loc = "2101"
            .area = "LLR"
            .zone = 9
            .name = "Tower Piece of Heart"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6818"
            .area = "LLR"
            .zone = 9
            .name = "Talon's Chickens"
            .logic = "YLL6204h"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(7)
            .area = "LLR"
            .zone = 9
            .name = "Song from Malon"
            .logic = "YLL6204h"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6208"
            .area = "LLR"
            .zone = 9
            .name = "Get Epona"
            .logic = "ZhLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8024"
            .area = "LLR"
            .zone = 9
            .name = "Wall Near Tower (N)"
            .gs = True
            .logic = "YNo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8025"
            .area = "LLR"
            .zone = 9
            .name = "Behind Corral (N)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8026"
            .area = "LLR"
            .zone = 9
            .name = "On House Window (N)"
            .gs = True
            .logic = "YNo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8027"
            .area = "LLR"
            .zone = 9
            .name = "Tree Near House"
            .gs = True
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "10124"
            .area = "LLR"
            .zone = 9
            .name = "Stables Left"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "10125"
            .area = "LLR"
            .zone = 9
            .name = "Stables Right"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2125"
            .area = "LLR"
            .zone = 9
            .name = "Tower Left"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2124"
            .area = "LLR"
            .zone = 9
            .name = "Tower Right"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6214"
            .area = "LLR"
            .zone = 9
            .name = "Win Cow from Malon"
            .cow = True
            .logic = "ZhLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9601"
            .area = "LLR"
            .zone = 9
            .name = "Open Grotto Left"
            .scrub = True
            .logic = "YS"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9604"
            .area = "LLR"
            .zone = 9
            .name = "Open Grotto Centre"
            .scrub = True
            .logic = "YS"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9606"
            .area = "LLR"
            .zone = 9
            .name = "Open Grotto Right"
            .scrub = True
            .logic = "YS"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysMK(ByRef tk As Integer)
        ' Set up keys for the Market
        ' 10 ML
        With aKeys(tk)
            .loc = "6829"
            .area = "MK"
            .zone = 10
            .name = "Shooting Gallery Reward"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6801"
            .area = "MK"
            .zone = 10
            .name = "Bombchu Bowling First Prize"
            .logic = "YX"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6802"
            .area = "MK"
            .zone = 10
            .name = "Bombchu Bowling Second Prize"
            .logic = "YX"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7201"
            .area = "MK"
            .zone = 10
            .name = "Lost Dog (N)"
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6811"
            .area = "MK"
            .zone = 10
            .name = "Treasure Chest Game"
            .logic = "YNT"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8119"
            .area = "MK"
            .zone = 10
            .name = "Crate in Guard House"
            .gs = True
            .logic = "Y"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysTT(ByRef tk As Integer)
        ' Set up keys for the Temple of Time
        ' 11 ToT Front
        ' 12 ToT Behind Door

        With aKeys(tk)
            .loc = "6405"
            .area = "TT"
            .zone = 11
            .name = "Song from Sheik"
            .logic = "ZLL7700"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(5)
            .area = "TT"
            .zone = 11
            .name = "Light Arrows Cutscene"
            .logic = "ZG22"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6427"
            .area = "EVENT"
            .name = "Opened Temple of Time"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysHC(ByRef tk As Integer)
        ' Set up keys for Hyrule Castle
        ' 13 HC
        ' 54 HC Great Fairy Fountain

        With aKeys(tk)
            .loc = "6202"
            .area = "HC"
            .zone = 13
            .name = "Malon's Egg"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6416"
            .area = "HC"
            .zone = 13
            .name = "Zelda's Letter"
            .logic = "Yy1.Yy2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6409"
            .area = "HC"
            .zone = 13
            .name = "Song from Impa"
            .logic = "Yy1.Yy2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(11)
            .area = "HC"
            .zone = 54
            .name = "HC Great Fairy Fountain"
            .logic = "hLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8117"
            .area = "HC"
            .zone = 13
            .name = "Inside Storms Grotto"
            .gs = True
            .logic = "YG02o"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8118"
            .area = "HC"
            .zone = 13
            .name = "Tree Near Entrance"
            .gs = True
            .logic = "YJ"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6204"
            .area = "EVENT"
            .name = "Woke Talon"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysKV(ByRef tk As Integer)
        ' Set up keys for Kakariko Village
        ' 14 KV Windmill
        ' 15 KV Main
        ' 16 KV Rooftops
        ' 17 KV Behind Gate

        With aKeys(tk)
            .loc = "4610"
            .area = "KV"
            .zone = 15
            .name = "Redead Grotto Chest"
            .logic = "YG03xa.YG03xf.YG03xLL7416.G01Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4608"
            .area = "KV"
            .zone = 15
            .name = "Open Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6828"
            .area = "KV"
            .zone = 15
            .name = "Anju's Chickens"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6928"
            .area = "KV"
            .zone = 15
            .name = "Talk to Anju"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6805"
            .area = "KV"
            .zone = 15
            .name = "Talk to Man on Roof"
            .logic = "Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1801"
            .area = "KV"
            .zone = 15
            .name = "Impa's House Piece of Heart"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(14)
            .area = "KV"
            .zone = 15
            .name = "Shooting Gallery"
            .logic = "Zd"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2001"
            .area = "KV"
            .zone = 15
            .name = "Windmill Piece of Heart"
            .logic = "Yo.ZhLL7716"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6411"
            .area = "KV"
            .zone = 15
            .name = "Song from Windmill"
            .logic = "Zh"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6404"
            .area = "KV"
            .zone = 15
            .name = "Song from Shiek"
            .logic = "ZLL7700LL7701LL7702"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8205"
            .area = "KV"
            .zone = 15
            .name = "Tree Near Entrance (N)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8203"
            .area = "KV"
            .zone = 15
            .name = "Construction Site (N)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8204"
            .area = "KV"
            .zone = 15
            .name = "On Skulltula House (N)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8201"
            .area = "KV"
            .zone = 15
            .name = "House Near Death Mountain (N)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8202"
            .area = "KV"
            .zone = 15
            .name = "Watchtower Ladder (N)"
            .gs = True
            .logic = "YNg.YNj"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8206"
            .area = "KV"
            .zone = 16
            .name = "Above Impa's House (N)"
            .gs = True
            .logic = "ZN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1824"
            .area = "KV"
            .zone = 15
            .name = "Inside Impa's House"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "C02"
            .area = "EVENT"
            .name = "Drained Kakariko Well"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGY(ByRef tk As Integer)
        ' Set up keys for the Graveyard
        ' 18 GY Main
        ' 19 GY Upper

        With aKeys(tk)
            .loc = locSwap(2)
            .area = "GY"
            .zone = 18
            .name = "Dampe's Gravedigging Tour (N)"
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4800"
            .area = "GY"
            .zone = 18
            .name = "Shield Grave Chest"
            .logic = "YN.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2204"
            .area = "GY"
            .zone = 18
            .name = "Ledge Piece of Heart"
            .logic = "ZG34.Zl"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4700"
            .area = "GY"
            .zone = 18
            .name = "Heart Piece Grave Chest"
            .logic = "YNhLL7715.ZhLL7715"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4900"
            .area = "GY"
            .zone = 18
            .name = "Royal Family's Tomb Chest"
            .logic = "YhLL7712f.ZhLL7712G24"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(8)
            .area = "GY"
            .zone = 18
            .name = "Song from Royal Family's Tomb"
            .logic = "YJhLL7712.ZhLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5000"
            .area = "GY"
            .zone = 18
            .name = "Hookshot Chest"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2007"
            .area = "GY"
            .zone = 18
            .name = "Dampe's Race Piece of Heart"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8200"
            .area = "GY"
            .zone = 18
            .name = "Soft Soil"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8207"
            .area = "GY"
            .zone = 18
            .name = "On South Wall (N)"
            .gs = True
            .logic = "YNo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "B4"
            .area = "EVENT"
            .name = "Beans Planted in Graveyard"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDMT(ByRef tk As Integer)
        ' Set up keys for Death Mountain Trail
        ' 20 DMT Lower
        ' 21 DMT Upper

        With aKeys(tk)
            .loc = "5801"
            .area = "DMT"
            .zone = 20
            .name = "Blast Wall Chest"
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2830"
            .area = "DMT"
            .zone = 20
            .name = "Piece of Heart Above Cavern"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(1)
            .area = "DMT"
            .zone = 21
            .name = "Great Fairy Fountain"
            .logic = "G00hLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4623"
            .area = "DMT"
            .zone = 20
            .name = "Storms Grotto Chest"
            .logic = "G02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6008"
            .area = "DMT"
            .zone = 21
            .name = "Help Biggoron"
            .logic = "Zz9.ZzA.ZzB"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8125"
            .area = "DMT"
            .zone = 20
            .name = "Soft Soil Near Dodongo's Cavern"
            .gs = True
            .logic = "YxG0D.YJG0DV01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8126"
            .area = "DMT"
            .zone = 20
            .name = "Blast Wall Near Village"
            .gs = True
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8127"
            .area = "DMT"
            .zone = 20
            .name = "Rock Near Bomb Flower (N)"
            .gs = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8128"
            .area = "DMT"
            .zone = 21
            .name = "Rock Near Climb Wall (N)"
            .gs = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1924"
            .area = "DMT"
            .zone = 21
            .name = "In Grotto"
            .cow = True
            .logic = "G00hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "B2"
            .area = "EVENT"
            .name = "DMT Planted Beans"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDMC(ByRef tk As Integer)
        ' Set up keys for Death Mountain Crater
        ' 22 DMC Upper
        ' 23 DMC Ladder
        ' 24 DMC Lower Nearby
        ' 25 DMC Lower Local
        ' 26 DMC Central Nearby
        ' 27 DMC Central Local
        ' 28 DMC Fire Temple Entrance

        With aKeys(tk)
            .loc = "4626"
            .area = "DMC"
            .zone = 22
            .name = "Upper Grotto Chest"
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2902"
            .area = "DMC"
            .zone = 22
            .name = "Wall Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2908"
            .area = "DMC"
            .zone = 26
            .name = "Volcano Piece of Heart"
            .logic = "ZG33.ZH"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(0)
            .area = "DMC"
            .zone = 24
            .name = "Great Fairy Fountain"
            .logic = "ZrhLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6401"
            .area = "DMC"
            .zone = 26
            .name = "Song from Sheik"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8131"
            .area = "DMC"
            .zone = 22
            .name = "Crate at Entrance"
            .gs = True
            .logic = "YJ"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8124"
            .area = "DMC"
            .zone = 27
            .name = "Soft Soil Near Warp"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "C01"
            .area = "DMC"
            .zone = 23
            .name = "Near Ladder"
            .scrub = True
            .logic = "YS"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9401"
            .area = "DMC"
            .zone = 26
            .name = "Grotto Left"
            .scrub = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9404"
            .area = "DMC"
            .zone = 26
            .name = "Grotto Centre"
            .scrub = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9406"
            .area = "DMC"
            .zone = 26
            .name = "Grotto Right"
            .scrub = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "B3"
            .area = "EVENT"
            .name = "DMC Planted Beans"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGC(ByRef tk As Integer)
        ' Set up keys for Goron City
        ' 29 GC Main
        ' 30 GC Shortcut
        ' 31 GC Darunia
        ' 32 GC Shoppe
        ' 33 GC Grotto Platform

        With aKeys(tk)
            .loc = "5900"
            .area = "GC"
            .zone = 29
            .name = "Maze Left Chest"
            .logic = "Zr.ZV02.ZHx"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5902"
            .area = "GC"
            .zone = 29
            .name = "Maze Centre Chest"
            .logic = "G00.ZV02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5901"
            .area = "GC"
            .zone = 29
            .name = "Maze Right Chest"
            .logic = "G00.ZV02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7014"
            .area = "GC"
            .zone = 29
            .name = "Rolling Goron"
            .logic = "Yx"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7025"
            .area = "GC"
            .zone = 29
            .name = "Rolling Goron"
            .logic = "Zx.Zd.ZV01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "3031"
            .area = "GC"
            .zone = 29
            .name = "Pot Piece of Heart"
            .logic = "Yfx.YfV01.YhLL7712x.YhLL7712V01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(4)
            .area = "GC"
            .zone = 31
            .name = "Darunia's Joy"
            .logic = "YhLL7714"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "3001"
            .area = "GC"
            .zone = 29
            .name = "Medigoron"
            .logic = "ZV21G00.ZV21V01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8130"
            .area = "GC"
            .zone = 29
            .name = "Boulder Maze"
            .gs = True
            .logic = "Yx"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8129"
            .area = "GC"
            .zone = 29
            .name = "Rope Platform"
            .gs = True
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9501"
            .area = "GC"
            .zone = 33
            .name = "Open Grotto Left"
            .scrub = True
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9504"
            .area = "GC"
            .zone = 33
            .name = "Open Grotto Centre"
            .scrub = True
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9506"
            .area = "GC"
            .zone = 33
            .name = "Open Grotto Right"
            .scrub = True
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11501"
            .area = "EVENT"
            .name = "Young Shoppe Opened"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11508"
            .area = "EVENT"
            .name = "Shortcut Rock 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11511"
            .area = "EVENT"
            .name = "Shortcut Rock 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11512"
            .area = "EVENT"
            .name = "Shortcut Rock 3"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11527"
            .area = "EVENT"
            .name = "Darunia's Door Opened"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11528"
            .area = "EVENT"
            .name = "Torches Lit"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZR(ByRef tk As Integer)
        ' Set up keys for Zora's River
        ' 34 ZR Front
        ' 35 ZR Main
        ' 36 ZR Behind Waterfall

        With aKeys(tk)
            .loc = "2304"
            .area = "ZR"
            .zone = 35
            .name = "Near Open Grotto Piece of Heart"
            .logic = "Y.ZLL7430"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4609"
            .area = "ZR"
            .zone = 35
            .name = "Open Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2311"
            .area = "ZR"
            .zone = 35
            .name = "Near Domain Piece of Heart"
            .logic = "Y.ZLL7430"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2301"
            .area = "ZR"
            .zone = 35
            .name = "Magic Bean Salesman"
            .logic = "Y"
            .scan = False
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8209"
            .area = "ZR"
            .zone = 34
            .name = "Tree Near Entrance"
            .gs = True
            .logic = "YJ"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8208"
            .area = "ZR"
            .zone = 35
            .name = "On Ladder Near Zora's Domain (N)"
            .gs = True
            .logic = "YJ"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8212"
            .area = "ZR"
            .zone = 35
            .name = "Wall Near Open Grotto (N)"
            .gs = True
            .logic = "Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8211"
            .area = "ZR"
            .zone = 35
            .name = "Wall Above Bridge (N)"
            .gs = True
            .logic = "Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8909"
            .area = "ZR"
            .zone = 35
            .name = "Storms Grotto Front"
            .scrub = True
            .logic = "YG02S.G02Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8908"
            .area = "ZR"
            .zone = 35
            .name = "Storms Grotto Back"
            .scrub = True
            .logic = "YG02S.G02Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11603"
            .area = "EVENT"
            .name = "Front Rock 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11608"
            .area = "EVENT"
            .name = "Front Rock 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11609"
            .area = "EVENT"
            .name = "Front Rock 3"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZD(ByRef tk As Integer)
        ' Set up keys for Zora's Domain
        ' 37 ZD Main
        ' 38 ZD Behind King
        ' 39 ZD Shoppe

        With aKeys(tk)
            .loc = "5300"
            .area = "ZD"
            .zone = 37
            .name = "Torches Chest"
            .logic = "Ya"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6308"
            .area = "ZD"
            .zone = 37
            .name = "Diving Minigame"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7109"
            .area = "ZD"
            .zone = 37
            .name = "Zora King"
            .logic = "ZG17"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8214"
            .area = "ZD"
            .zone = 37
            .name = "Frozen Waterfall Top (N)"
            .gs = True
            .logic = "ZNk.ZNd.ZNM"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6303"
            .area = "EVENT"
            .name = "Delivered Ruto's Letter"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZF(ByRef tk As Integer)
        ' Set up keys for Zora's Fountain
        ' 40 ZF Main
        ' 41 ZF Ice Cavern Ledge

        With aKeys(tk)
            .loc = locSwap(10)
            .area = "ZF"
            .zone = 40
            .name = "Great Fairy Fountain"
            .logic = "xhLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2501"
            .area = "ZF"
            .zone = 40
            .name = "Iceberg Piece of Heart"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2520"
            .area = "ZF"
            .zone = 40
            .name = "Bottom Piece of Heart"
            .logic = "ZG0ALL7429"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8210"
            .area = "ZF"
            .zone = 40
            .name = "Lower West Wall (N)"
            .gs = True
            .logic = "YNo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8215"
            .area = "ZF"
            .zone = 40
            .name = "Southeast Corner Tree"
            .gs = True
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8213"
            .area = "ZF"
            .zone = 40
            .name = "Hidden Cave (N)"
            .gs = True
            .logic = "ZV02G00Nk"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysLH(ByRef tk As Integer)
        ' 42 LH Main
        ' 43 LH Fishing Ledge

        ' Set up keys for Lake Hylia
        With aKeys(tk)
            .loc = "7316"
            .area = "LH"
            .zone = 42
            .name = "Scarecrow Bonooru"
            .logic = "Yh"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6512"
            .area = "LH"
            .zone = 42
            .name = "Scarecrow Pierre"
            .logic = "ZhLL7316"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "211"
            .area = "LH"
            .zone = 42
            .name = "Underwater Item"
            .logic = "YV11"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6110"
            .area = "LH"
            .zone = 43
            .name = "Big Fish First"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6111"
            .area = "LH"
            .zone = 43
            .name = "Big Fish Second"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2430"
            .area = "LH"
            .zone = 42
            .name = "Lab Tower Piece of Heart"
            .logic = "YZG35.ZkhLL6512"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6800"
            .area = "LH"
            .zone = 42
            .name = "Lab Dive"
            .logic = "ZV12"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = locSwap(3)
            .area = "LH"
            .zone = 42
            .name = "Shoot the Sun"
            .logic = "ZdlhLL6512.ZdG16"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8216"
            .area = "LH"
            .zone = 42
            .name = "Soft Soil Near Lab"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8218"
            .area = "LH"
            .zone = 42
            .name = "Behind Lab (N)"
            .gs = True
            .logic = "Yo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8217"
            .area = "LH"
            .zone = 42
            .name = "On Small Island Pillar (N)"
            .gs = True
            .logic = "YJ"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8219"
            .area = "LH"
            .zone = 42
            .name = "Crate in Lab Pool"
            .gs = True
            .logic = "ZLL7429k"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8220"
            .area = "LH"
            .zone = 42
            .name = "On Tree Near Warp (N)"
            .gs = True
            .logic = "Zl"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9101"
            .area = "LH"
            .zone = 42
            .name = "Grave Grotto Left"
            .scrub = True
            .logic = "YS.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9104"
            .area = "LH"
            .zone = 42
            .name = "Grave Grotto Centre"
            .scrub = True
            .logic = "YS.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9106"
            .area = "LH"
            .zone = 42
            .name = "Grave Grotto Right"
            .scrub = True
            .logic = "YS.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "231"
            .area = "EVENT"
            .name = "Open Water Temple"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "C04"
            .area = "EVENT"
            .name = "Lake Restored"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "202"
            .area = "EVENT"
            .name = "Lake Hylia Bean Planted"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGV(ByRef tk As Integer)
        ' Set up keys for Gerudo Valley
        ' 44 GV Hyrule Side
        ' 45 GV Gerudo Side
        ' 56 GV Upper Stream
        ' 57 GV Crate Ledge

        With aKeys(tk)
            .loc = "2602"
            .area = "GV"
            .zone = 57
            .name = "Crate Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2601"
            .area = "GV"
            .zone = 56
            .name = "Waterfall Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5400"
            .area = "GV"
            .zone = 45
            .name = "Chest Behind Rocks"
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8225"
            .area = "GV"
            .zone = 44
            .name = "Small Bridge (N)"
            .gs = True
            .logic = "Yo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8224"
            .area = "GV"
            .zone = 56
            .name = "Soft Soil Near Cow"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8227"
            .area = "GV"
            .zone = 45
            .name = "Behind Tents (N)"
            .gs = True
            .logic = "ZNk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8226"
            .area = "GV"
            .zone = 45
            .name = "Pillar Near Tents (N)"
            .gs = True
            .logic = "ZNk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2624"
            .area = "GV"
            .zone = 56
            .name = "Near Soft Soil"
            .cow = True
            .logic = "YhLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9209"
            .area = "GV"
            .name = "Storms Grotto Front"
            .zone = 45
            .scrub = True
            .logic = "ZG02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9208"
            .area = "GV"
            .zone = 45
            .name = "Storms Grotto Back"
            .scrub = True
            .logic = "ZG02"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGF(ByRef tk As Integer)
        ' Set up keys for Gerudo's Fortress
        ' 46 GF Main
        ' 47 GF Behind Gate

        With aKeys(tk)
            .loc = "5600"
            .area = "GF"
            .zone = 46
            .name = "Chest on Top"
            .logic = "Zl.ZLL7430.ZkhLL6512"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "CARD"
            .area = "GF"
            .zone = 46
            .name = "Membership Card"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7200"
            .area = "GF"
            .zone = 46
            .name = "Archery 1000 points"
            .logic = "ZhLL7713LL6208LL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6831"
            .area = "GF"
            .zone = 46
            .name = "Archery 1500 points"
            .logic = "ZhLL7713LL6208LL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6500"
            .area = "GF"
            .zone = 46
            .name = "Carpenter with 1 Torch"
            .scan = False
            .logic = "Z.LL7416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6502"
            .area = "GF"
            .zone = 46
            .name = "Carpenter with 2 Torches"
            .scan = False
            .logic = "Z.LL7416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6503"
            .area = "GF"
            .zone = 46
            .name = "Carpenter with 3 Torches"
            .scan = False
            .logic = "ZLL7722.Zd.Zk.ZLL7430.LL7416LL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6501"
            .area = "GF"
            .zone = 46
            .name = "Carpenter with 4 Torches"
            .scan = False
            .logic = "Z.LL7416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8301"
            .area = "GF"
            .zone = 46
            .name = "Top Wall Near Scarecrow (N)"
            .gs = True
            .logic = "ZNd.ZNk.ZNLL7430.ZNLL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8300"
            .area = "GF"
            .zone = 46
            .name = "Far Archery Target (N)"
            .gs = True
            .logic = "ZNLL7722k"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "C00"
            .area = "EVENT"
            .name = "Opened Haunted Wasteland Gate"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysHW(ByRef tk As Integer)
        ' Set up keys for the Haunted Wasteland
        ' 48 HW Gerudo Side
        ' 49 HW Colossus Side
        ' 55 HW Main

        With aKeys(tk)
            .loc = "5700"
            .area = "HW"
            .zone = 55
            .name = "Structure Chest"
            .logic = "Yf.ZG24"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8309"
            .area = "HW"
            .zone = 55
            .name = "Structure Basement"
            .gs = True
            .logic = "Yo.Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "11401"
            .area = "HW"
            .zone = 55
            .name = "Carpet Salesman"
            .logic = "YV21a.YV21LL7416.ZV21"
            .scan = False
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDC(ByRef tk As Integer)
        ' Set up keys for the Desert Colossus
        ' 50 DC

        With aKeys(tk)
            .loc = locSwap(12)
            .area = "DC"
            .zone = 50
            .name = "Great Fairy Fountain"
            .logic = "xhLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6628"
            .area = "DC"
            .zone = 50
            .name = "Song from Shiek"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2713"
            .area = "DC"
            .zone = 50
            .name = "Piece of Heart"
            .logic = "ZG31"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8308"
            .area = "DC"
            .zone = 50
            .name = "Soft Soil By Temple Entrance"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8310"
            .area = "DC"
            .zone = 50
            .name = "Bean Plant Ride Hill (N)"
            .gs = True
            .logic = "ZG31"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8311"
            .area = "DC"
            .zone = 50
            .name = "Southern Edge On Tree (N)"
            .gs = True
            .logic = "Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9709"
            .area = "DC"
            .zone = 50
            .name = "Grotto Front"
            .scrub = True
            .logic = "ZV02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9708"
            .area = "DC"
            .zone = 50
            .name = "Grotto Back"
            .scrub = True
            .logic = "ZV02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "B1"
            .area = "EVENT"
            .name = "Desert Colossus Bean Planted"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysOGC(ByRef tk As Integer)
        ' Set up keys for Outside Ganon's Castle
        ' 51 OGC
        ' 53 OGC Great Fairy Fountain

        With aKeys(tk)
            .loc = "008"
            .area = "OGC"
            .zone = 53
            .name = "OGC Great Fairy Fountain"
            .logic = "hLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8116"
            .area = "OGC"
            .zone = 51
            .name = "On Pillar"
            .gs = True
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6429"
            .area = "EVENT"
            .name = "Rainbow Bridge"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQBPH(ByRef tk As Integer)
        ' Set up keys for Quest: Big Poe Hunt
        ' 57 Big Poe Hunt
        With aKeys(tk)
            .loc = "124"
            .area = "QBPH"
            .zone = 59
            .name = "#1: Near Castle Gate"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "123"
            .area = "QBPH"
            .zone = 59
            .name = "#2: Near Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "122"
            .area = "QBPH"
            .zone = 59
            .name = "#3: West of Castle"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "130"
            .area = "QBPH"
            .zone = 59
            .name = "#4: Between Gerudo Valley and Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "131"
            .area = "QBPH"
            .zone = 59
            .name = "#5: Near Gerudo Valley"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "128"
            .area = "QBPH"
            .zone = 59
            .name = "#6: Southeast Field Near Path"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "129"
            .area = "QBPH"
            .zone = 59
            .name = "#7: Southeast Field Near Grotto"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "127"
            .area = "QBPH"
            .zone = 59
            .name = "#8: Between Kokiri Forest and Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "126"
            .area = "QBPH"
            .zone = 59
            .name = "#9: Wall East of Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "125"
            .area = "QBPH"
            .zone = 59
            .name = "#10: Near Kakariko Village"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQF(ByRef tk As Integer)
        ' Set up keys for Quest: Frogs
        ' 58 Frogs

        With aKeys(tk)
            .loc = "6706"
            .area = "QF"
            .zone = 58
            .name = "Play Song of Storms"
            .logic = "LL7717"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6705"
            .area = "QF"
            .zone = 58
            .name = "Play Song of Time"
            .logic = "LL7716"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6704"
            .area = "QF"
            .zone = 58
            .name = "Play Saria's Song"
            .logic = "LL7714"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6703"
            .area = "QF"
            .zone = 58
            .name = "Play Sun's Song"
            .logic = "LL7715"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6702"
            .area = "QF"
            .zone = 58
            .name = "Play Epona's Song"
            .logic = "LL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6701"
            .area = "QF"
            .zone = 58
            .name = "Play Zelda's Lullaby"
            .logic = "LL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6700"
            .area = "QF"
            .zone = 58
            .name = "Ocarina Game"
            .logic = "LL7712LL7713LL7714LL7715LL7716LL7717"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQGS(ByRef tk As Integer)
        ' Set up keys for Quest: Gold Skulltula Rewards
        ' 15 KV Main
        With aKeys(tk)
            .loc = "6710"
            .area = "QGS"
            .zone = 15
            .name = "10 Gold Skulltulas"
            .logic = "K1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6711"
            .area = "QGS"
            .zone = 15
            .name = "20 Gold Skulltulas"
            .logic = "K2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6712"
            .area = "QGS"
            .zone = 15
            .name = "30 Gold Skulltulas"
            .logic = "K3"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6713"
            .area = "QGS"
            .zone = 15
            .name = "40 Gold Skulltulas"
            .logic = "K4"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6714"
            .area = "QGS"
            .zone = 15
            .name = "50 Gold Skulltulas"
            .logic = "K5"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQM(ByRef tk As Integer)
        ' Set up keys for the Quest: Masks
        With aKeys(tk)
            .loc = "C05"
            .area = "QM"
            .zone = 15
            .name = "Deliver Zelda's Letter to Unlock Mask Shoppe"
            .logic = "Yy3"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6908"
            .area = "QM"
            .zone = 10
            .name = "Sell Keaton Mask"
            .logic = "YLLLC05Q0015"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6909"
            .area = "QM"
            .zone = 10
            .name = "Sell Skull Mask"
            .logic = "YLL6908hLL7714Q0002"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6910"
            .area = "QM"
            .zone = 10
            .name = "Sell Spooky Mask"
            .logic = "YLL6909Q0018"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6911"
            .area = "QM"
            .zone = 10
            .name = "Sell Bunny Hood"
            .logic = "YLL6910LL6625Q0007"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysInventory(ByRef tk As Integer)
        ' Set up keys for the inventory
        With aKeys(tk)
            .loc = "7416"
            .area = "INV"
            .name = "Kokiri Sword"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7417"
            .area = "INV"
            .name = "Master Sword"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7418"
            .area = "INV"
            .name = "Biggoron's Sword"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7419"
            .area = "INV"
            .name = "Broken Knife"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7420"
            .area = "INV"
            .name = "Deku Shield"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7421"
            .area = "INV"
            .name = "Hylian Shield"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7422"
            .area = "INV"
            .name = "Mirror Shield"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7424"
            .area = "INV"
            .name = "Kokiri Tunic"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7425"
            .area = "INV"
            .name = "Goron Tunic"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7426"
            .area = "INV"
            .name = "Zora Tunic"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7428"
            .area = "INV"
            .name = "Kokiri Boots"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7429"
            .area = "INV"
            .name = "Iron Boots"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7430"
            .area = "INV"
            .name = "Hover Boots"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7508"
            .area = "INV"
            .name = "Biggoron or Knife"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7600"
            .area = "INV"
            .name = "Quiver 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7601"
            .area = "INV"
            .name = "Quiver 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7603"
            .area = "INV"
            .name = "Bomb Bag 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7604"
            .area = "INV"
            .name = "Bomb Bag 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7606"
            .area = "INV"
            .name = "Gauntlet 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7607"
            .area = "INV"
            .name = "Gauntlet 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7609"
            .area = "INV"
            .name = "Scale 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7610"
            .area = "INV"
            .name = "Scale 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7612"
            .area = "INV"
            .name = "Wallet 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7613"
            .area = "INV"
            .name = "Wallet 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7614"
            .area = "INV"
            .name = "Bullet Bag 1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7615"
            .area = "INV"
            .name = "Bullet Bag 2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7700"
            .area = "INV"
            .name = "Forest Medallion"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7701"
            .area = "INV"
            .name = "Fire Medallion"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7702"
            .area = "INV"
            .name = "Water Medallion"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7703"
            .area = "INV"
            .name = "Spirit Medallion"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7704"
            .area = "INV"
            .name = "Shadow Medallion"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7705"
            .area = "INV"
            .name = "Light Medallion"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7706"
            .area = "INV"
            .name = "Minuet of Forest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7707"
            .area = "INV"
            .name = "Bolero of Fire"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7708"
            .area = "INV"
            .name = "Serenade of Water"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7709"
            .area = "INV"
            .name = "Requiem of Spirit"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7710"
            .area = "INV"
            .name = "Nocturne of Shadow"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7711"
            .area = "INV"
            .name = "Prelude of Light"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7712"
            .area = "INV"
            .name = "Zelda's Lullaby"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7713"
            .area = "INV"
            .name = "Epona's Song"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7714"
            .area = "INV"
            .name = "Saria's Song"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7715"
            .area = "INV"
            .name = "Sun's Song"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7716"
            .area = "INV"
            .name = "Song of Time"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7717"
            .area = "INV"
            .name = "Song of Storms"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7718"
            .area = "INV"
            .name = "Kokiri Emerald"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7719"
            .area = "INV"
            .name = "Goron Ruby"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7720"
            .area = "INV"
            .name = "Zora Sapphire"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7721"
            .area = "INV"
            .name = "Stone of Agony"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7722"
            .area = "INV"
            .name = "Gerudo's Membership Card"
        End With
        inc(tk)
    End Sub

    Private Sub updateShoppes()
        Select Case aAddresses(6)
            Case 1
                generateShoppeKeys(4)
                ' For Archipelago, once we have a full shopsanity key set, disable the unneccessary ones
                For i = keyCount - 31 To keyCount
                    With aKeys(i)
                        If .name.Contains(": Upper-Right") Then .scan = CBool(IIf(My.Settings.setShop >= 4, True, False))
                        If .name.Contains(": Upper-Left") Then .scan = CBool(IIf(My.Settings.setShop >= 3, True, False))
                        If .name.Contains(": Lower-Right") Then .scan = CBool(IIf(My.Settings.setShop >= 2, True, False))
                    End With
                Next
            Case Else
                generateShoppeKeys(My.Settings.setShop)
        End Select
    End Sub
    Private Sub generateShoppeKeys(Optional items As Byte = 5)
        ' Failsafe for none to be full and anything over full to limit itself
        If items > 4 Then items = 4

        ' Set starting key for shopsanity. Being the last 32 keys, -31 (need to count the one it is on)
        Dim startKey = keyCount - 31
        Dim workKey As String = String.Empty
        'Dim totalItems As Byte = My.Settings.setShop
        Dim curShop As Integer = 0
        'Dim shopArea() As String = {"MK", "MK", "ZD", "KV", "MK", "GC", "KF", "KV"}
        'Dim shopZone() As Byte = {10, 10, 39, 15, 10, 32, 0, 15}
        'Dim shopName() As String = {"Potion Shop", "Bombchu Shop", "Shop", "Potion Shop", "Bazaar", "Shop", "Shop", "Bazaar"}
        'Dim shopLogic() As String = {"Y", "Y", "", "Z", "Y", "", "", "Z"}
        Dim shopArea() As String = {"KF", "KV", "MK", "GC", "ZD", "KV", "MK", "MK"}
        Dim shopZone() As Byte = {0, 15, 10, 32, 39, 15, 10, 10}
        Dim shopName() As String = {"Shop", "Bazaar", "Bazaar", "Shop", "Shop", "Potion Shop", "Potion Shop", "Bombchu Shop"}
        Dim shopLogic() As String = {"", "Z", "Y", "", "", "Z", "Y", "Y"}
        Dim curShelf As Integer = 0
        'Dim shelfLoc() As String = {"Lower-Left", "Lower-Right", "Upper-Left", "Upper-Right"}
        'Dim shelfLoc() As String = {"Lower-Right", "Upper-Right", "Lower-Left", "Upper-Left"}
        Dim shelfLoc() As String = {"Lower-Left", "Lower-Right", "Upper-Left", "Upper-Right"}

        Select Case items
            Case 1
                shelfLoc = {"Lower-Left", String.Empty, String.Empty, String.Empty}
            Case 2, 3
                shelfLoc = {"Lower-Right", "Lower-Left", "Upper-Left", String.Empty}
            Case Else
                shelfLoc = {"Lower-Right", "Upper-Right", "Lower-Left", "Upper-Left"}
        End Select

        ' Step through the last 32 keys, which will have to be kept the Shopsanity keys
        For i = 0 To 31

            ' Start with resetting the key
            aKeys(startKey + i) = New keyCheck

            If i < (items * 8) Then
                With aKeys(startKey + i)
                    ' = 24 - (math.floor(i /8)*4)
                    workKey = ((3 - Math.Floor(i / 8)) * 8 + (i Mod 8)).ToString
                    fixHex(workKey, 2)
                    .loc = "99" & workKey
                    .area = shopArea(curShop)
                    .zone = shopZone(curShop)
                    .name = shopName(curShop) & ": " & shelfLoc(curShelf)
                    .logic = shopLogic(curShop)
                    .shop = True
                    .scan = True
                    inc(curShelf)
                    'If curShelf >= totalItems Then
                    If curShelf >= items Then
                        inc(curShop)
                        curShelf = 0
                    End If
                End With
            End If
        Next
        updateLabels()
    End Sub

    Private Sub updateMQs()
        ' Set up the dungeons based on if they are Master Quest versions or not

        ' Step through the aMQ array, where the current Master Quest settings are set
        For i = 0 To 11
            ' If the new read is different than the old read (old is what is already read and loaded), update the dungeon array
            If Not aMQ(i) = aMQOld(i) Then
                Select Case i
                    Case 0
                        makeKeysDekuTree(aMQ(i))
                    Case 1
                        makeKeysDodongosCavern(aMQ(i))
                    Case 2
                        makeKeysJabuJabusBelly(aMQ(i))
                    Case 3
                        makeKeysForestTemple(aMQ(i))
                    Case 4
                        makeKeysFireTemple(aMQ(i))
                    Case 5
                        makeKeysWaterTemple(aMQ(i))
                    Case 6
                        makeKeysSpiritTemple(aMQ(i))
                    Case 7
                        makeKeysShadowTemple(aMQ(i))
                    Case 8
                        makeKeysBottomOfTheWell(aMQ(i))
                    Case 9
                        makeKeysIceCavern(aMQ(i))
                    Case 10
                        makeKeysGerudoTrainingGround(aMQ(i))
                    Case 11
                        makeKeysGanonsCastle(aMQ(i))
                End Select
            End If
        Next

        ' Update the High/Lows, used for saving time when scanning
        getHighLows()
    End Sub
    Private Sub makeKeysDekuTree(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(0) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(0)(CInt(IIf(isMQ, 13, 11)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(0).Length - 1
            aKeysDungeons(0)(i) = New keyCheck
            aKeysDungeons(0)(i).scan = True
        Next

        ' Start off with keys for both versions of the Deku Tree
        If Not isMQ Then
            ' The non-Master Quest keys for the Deku Tree
            '   0: Lobby
            '   1: Slingshot Room
            '   2: Basement Backroom
            '   3: Boss Room

            With aKeysDungeons(0)(0)
                .loc = "3103"
                .area = "DT0"
                .zone = 60
                .name = "Map Chest"
            End With
            With aKeysDungeons(0)(1)
                .loc = "3101"
                .area = "DT1"
                .zone = 61
                .name = "Slingshot Chest"
            End With
            With aKeysDungeons(0)(2)
                .loc = "3105"
                .area = "DT1"
                .zone = 61
                .name = "Slingshot Room Side Chest"
            End With
            With aKeysDungeons(0)(3)
                .loc = "3102"
                .area = "DT0"
                .zone = 60
                .name = "Compass Chest"
            End With
            With aKeysDungeons(0)(4)
                .loc = "3106"
                .area = "DT0"
                .zone = 60
                .name = "Compass Room Side Chest"
            End With
            With aKeysDungeons(0)(5)
                .loc = "3104"
                .area = "DT0"
                .zone = 60
                .name = "Basement Chest"
                .logic = "YJ.Yb.Z"
                '.logic = "YJ.b"
            End With
            With aKeysDungeons(0)(6)
                .loc = "1031"
                .area = "DT3"
                .zone = 63
                .name = "Queen Gohma"
                .logic = "YLL7420LL7416.YLL7420a.ZLL7421"
                '.logic = "YLL7420LL7416.YLL7420a"
            End With
            With aKeysDungeons(0)(7)
                .loc = "7803"
                .area = "DT0"
                .zone = 60
                .name = "Compass Room"
                .gs = True
                .logic = "YJ.Z"
                '.logic = "YJ"
            End With
            With aKeysDungeons(0)(8)
                .loc = "7802"
                .area = "DT0"
                .zone = 60
                .name = "Basement Vines"
                .gs = True
                .logic = "YG20.Yf.ZG21.Zf"
                '.logic = "f.x.Yg.Yo"
            End With
            With aKeysDungeons(0)(9)
                .loc = "7801"
                .area = "DT0"
                .zone = 60
                .name = "Basement Gate"
                .gs = True
                .logic = "YJ.Z"
                '.logic = "YJ"
            End With
            With aKeysDungeons(0)(10)
                .loc = "7800"
                .area = "DT2"
                .zone = 62
                .name = "Basement Back Room"
                .gs = True
                .logic = "YG23xo.ZG24xk.Zdxk.ZG24rk.Zdrk"
                '.logic = "YG23ox"
            End With
            tK = 11
        Else
            ' The Master Quest keys for the Deku Tree
            '   0: MQ Lobby
            '   1: MQ Compass Room
            '   2: MQ Basement Water Room Front
            '   3: MQ Basement Water Room Back
            '   4: MQ Basement Back Room
            '   5: MQ Basement Ledge

            With aKeysDungeons(0)(0)
                .loc = "3103"
                .area = "DT0"
                .zone = 60
                .name = "Map Chest"
            End With
            With aKeysDungeons(0)(1)
                .loc = "3102"
                .area = "DT0"
                .zone = 60
                .name = "Slingshot Chest"
                .logic = "YJ"
            End With
            With aKeysDungeons(0)(2)
                .loc = "3106"
                .area = "DT0"
                .zone = 60
                .name = "Slingshot Back Room Chest"
                .logic = "YG23"
            End With
            With aKeysDungeons(0)(3)
                .loc = "3101"
                .area = "DT1"
                .zone = 64
                .name = "Compass Chest"
            End With
            With aKeysDungeons(0)(4)
                .loc = "3104"
                .area = "DT0"
                .zone = 60
                .name = "Basement Chest"
                .logic = "YG23"
            End With
            With aKeysDungeons(0)(5)
                .loc = "3105"
                .area = "DT2"
                .zone = 65
                .name = "Before Spinning Log Chest"
            End With
            With aKeysDungeons(0)(6)
                .loc = "3100"
                .area = "DT3"
                .zone = 66
                .name = "After Spinning Log Chest"
                .logic = "hLL7716"
            End With
            With aKeysDungeons(0)(7)
                .loc = "1031"
                .area = "DT5"
                .zone = 68
                .name = "Queen Gohma"
                .logic = "YG23LL7420"
            End With
            With aKeysDungeons(0)(8)
                .loc = "7801"
                .area = "DT0"
                .zone = 60
                .name = "Lobby"
                .gs = True
                .logic = "YJ"
            End With
            With aKeysDungeons(0)(9)
                .loc = "7803"
                .area = "DT1"
                .zone = 64
                .name = "Compass Room"
                .gs = True
                .logic = "Yoj.YochLL7716"
            End With
            With aKeysDungeons(0)(10)
                .loc = "7802"
                .area = "DT4"
                .zone = 67
                .name = "Basement Graves Room"
                .gs = True
                .logic = "YohLL7716"
            End With
            With aKeysDungeons(0)(11)
                .loc = "7800"
                .area = "DT4"
                .zone = 67
                .name = "Basement Back Room"
                .gs = True
                .logic = "YG23o"
            End With
            With aKeysDungeons(0)(12)
                .loc = "8405"
                .area = "DT5"
                .zone = 68
                .name = "Basement"
                .scrub = True
                .logic = "YS"
            End With
            tK = 13
        End If
        With aKeysDungeons(0)(tK)
            .loc = "10316"
            .area = "EVENT"
            .name = "Block in Basement"
        End With
    End Sub
    Private Sub makeKeysDodongosCavern(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(1) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(1)(CInt(IIf(isMQ, 18, 17)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(1).Length - 1
            aKeysDungeons(1)(i) = New keyCheck
            aKeysDungeons(1)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys to Dodongo's Cavern
            ' 0: Dodongo's Cavern Beginning
            ' 1: Dodongo's Cavern Lobby
            ' 2: Dodongo's Cavern Staircase Room
            ' 3: Dodongo's Cavern Climb
            ' 4: Dodongo's Cavern Far Bridge
            ' 5: Dodongo's Cavern Boss Area

            With aKeysDungeons(1)(0)
                .loc = "3208"
                .area = "DDC1"
                .zone = 69
                .name = "Map Chest"
                .logic = "x.V01"
            End With
            With aKeysDungeons(1)(1)
                .loc = "3205"
                .area = "DDC2"
                .zone = 70
                .name = "Compass Chest"
                .logic = "x.V01"
            End With
            With aKeysDungeons(1)(2)
                .loc = "3206"
                .area = "DDC3"
                .zone = 71
                .name = "Bomb Flower Platform Chest"
            End With
            With aKeysDungeons(1)(3)
                .loc = "3204"
                .area = "DDC4"
                .zone = 72
                .name = "Bomb Bag Chest"
            End With
            With aKeysDungeons(1)(4)
                .loc = "3210"
                .area = "DDC4"
                .zone = 72
                .name = "End of Bridge Chest"
                .logic = "G00"
            End With
            With aKeysDungeons(1)(5)
                .loc = "4400"
                .area = "DDC5"
                .zone = 73
                .name = "Boss Room Chest"
            End With
            With aKeysDungeons(1)(6)
                .loc = "1131"
                .area = "DDC5"
                .zone = 73
                .name = "King Dodongo"
                .logic = "Zc.Yac.YLL7416c.ZV01.YaV01.YLL7416V01"
            End With
            With aKeysDungeons(1)(7)
                .loc = "7812"
                .area = "DDC1"
                .zone = 69
                .name = "Side Room Near Lower Lizalfos"
                .gs = True
                .logic = "Z.Yx.YV01LL7416.YV01a.YV01g.YV01o"
            End With
            With aKeysDungeons(1)(8)
                .loc = "7809"
                .area = "DDC1"
                .zone = 69
                .name = "Scarecrow"
                .gs = True
                .logic = "ZkhLL6512.Zl"
            End With
            With aKeysDungeons(1)(9)
                .loc = "7810"
                .area = "DDC4"
                .zone = 72
                .name = "Alcove Above Stairs"
                .gs = True
                .logic = "Zk.Yo"
            End With
            With aKeysDungeons(1)(10)
                .loc = "7808"
                .area = "DDC3"
                .zone = 71
                .name = "Vines Above Stairs"
                .gs = True
            End With
            With aKeysDungeons(1)(11)
                .loc = "7811"
                .area = "DDC5"
                .zone = 73
                .name = "Back Room"
                .gs = True
                .logic = "G00"
            End With
            With aKeysDungeons(1)(12)
                .loc = "8505"
                .area = "DDC1"
                .zone = 69
                .name = "Lobby"
                .scrub = True
                .logic = "Z.YS"
            End With
            With aKeysDungeons(1)(13)
                .loc = "8502"
                .area = "DDC1"
                .zone = 69
                .name = "Side Room Near Dodongos"
                .scrub = True
                .logic = "ZG00.ZV01.G00LL7416.YV01LL7416.YG00a.YV01a.G00c.V01c.YG00g.YV01g"
            End With
            With aKeysDungeons(1)(14)
                .loc = "8501"
                .area = "DDC3"
                .zone = 71
                .name = "Near Bomb Bag Left"
                .scrub = True
                .logic = "G00"
            End With
            With aKeysDungeons(1)(15)
                .loc = "8504"
                .area = "DDC3"
                .zone = 71
                .name = "Near Bomb Bag Right"
                .scrub = True
                .logic = "G00"
            End With
            tK = 16
        Else
            ' The Master Quest keys to Dodongo's Cavern
            ' 0: MW Dodongo's Cavern Beginning
            ' 1: MQ Dodongo's Cavern Lobby
            ' 2: MQ Dodongo's Cavern Elevator
            ' 3: MQ Dodongo's Cavern Lower Right Side
            ' 4: MQ Dodongo's Cavern Bomb Bag Area
            ' 5: MQ Dodongo's Cavern Boss Area
            ' 6: MQ Dodongo's Cavern Boss Room
            With aKeysDungeons(1)(0)
                .loc = "3200"
                .area = "DDC1"
                .zone = 69
                .name = "Map Chest"
                .logic = "G00.V01"
            End With
            With aKeysDungeons(1)(1)
                .loc = "3204"
                .area = "DDC4"
                .zone = 76
                .name = "Bomb Bag Chest"
            End With
            With aKeysDungeons(1)(2)
                .loc = "3203"
                .area = "DDC2"
                .zone = 74
                .name = "Torch Puzzle Room Chest"
                .logic = "G00.Ya.f.ZA0.ZLL7430.Zk"
            End With
            With aKeysDungeons(1)(3)
                .loc = "3202"
                .area = "DDC2"
                .zone = 74
                .name = "Larvae Room Chest"
                .logic = "Ya.f.ZG24"
            End With
            With aKeysDungeons(1)(4)
                .loc = "3205"
                .area = "DDC2"
                .zone = 74
                .name = "Compass Chest"
                .logic = "Z.YJ.Yb"
            End With
            With aKeysDungeons(1)(5)
                .loc = "3201"
                .area = "DDC5"
                .zone = 77
                .name = "Under Grave Chest"
            End With
            With aKeysDungeons(1)(6)
                .loc = "4400"
                .area = "DDC6"
                .zone = 78
                .name = "Boss Room Chest"
            End With
            With aKeysDungeons(1)(7)
                .loc = "1131"
                .area = "DDC6"
                .zone = 78
                .name = "King Dodongo"
                '.logic = "rc.rV01.xa.xLL7416.Zx"
                .logic = "Zrc.ZrV01.Yca.YxV01a.YcLL7416.YxV01LL7416"
            End With
            With aKeysDungeons(1)(8)
                .loc = "7809"
                .area = "DDC4"
                .zone = 76
                .name = "Scrub Room"
                .gs = True
                .logic = "Zdl.ZV01l.Zfl.Zxl.Ydo.YV01o.Yfo.Yxo"
            End With
            With aKeysDungeons(1)(9)
                .loc = "7812"
                .area = "DDC2"
                .zone = 74
                .name = "Larvae Room"
                .gs = True
                .logic = "Ya.f.ZG24"
            End With
            With aKeysDungeons(1)(10)
                .loc = "7810"
                .area = "DDC2"
                .zone = 74
                .name = "Lizalfos Room"
                .gs = True
                .logic = "G00"
            End With
            With aKeysDungeons(1)(11)
                .loc = "7811"
                .area = "DDC2"
                .zone = 74
                .name = "Song of Time Block Room"
                .gs = True
                .logic = "JhLL7716.ZhLL7716"
            End With
            With aKeysDungeons(1)(12)
                .loc = "7808"
                .area = "DDC5"
                .zone = 77
                .name = "Back Area"
                .gs = True
                .logic = "Z.Yx.Yo.Yf.YabLL7416.YaoLL7416.Yabg.Yaog"
            End With
            With aKeysDungeons(1)(13)
                .loc = "8504"
                .area = "DDC1"
                .zone = 69
                .name = "Lobby Front"
                .scrub = True
                .logic = "Z.YS"
            End With
            With aKeysDungeons(1)(14)
                .loc = "8502"
                .area = "DDC1"
                .zone = 69
                .name = "Lobby Rear"
                .scrub = True
                .logic = "Z.YS"
            End With
            With aKeysDungeons(1)(15)
                .loc = "8508"
                .area = "DDC3"
                .zone = 75
                .name = "Side Room Near Lower Lizalfos"
                .scrub = True
                .logic = "G00Z.G00J.V01Z.V01J"
            End With
            With aKeysDungeons(1)(16)
                .loc = "8505"
                .area = "DDC2"
                .zone = 74
                .name = "Staircase Room"
                .scrub = True
                .logic = "Z.YS"
            End With
            tK = 17
        End If

        With aKeysDungeons(1)(tK)
            .loc = "10410"
            .area = "EVENT"
            .name = "Big Raising Platform"
        End With
        incB(tK)
        With aKeysDungeons(1)(tK)
            .loc = "10426"
            .area = "EVENT"
            .name = "Eyes Lit Mouth Open"
        End With
    End Sub
    Private Sub makeKeysJabuJabusBelly(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(2) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(2)(CInt(IIf(isMQ, 17, 9)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(2).Length - 1
            aKeysDungeons(2)(i) = New keyCheck
            aKeysDungeons(2)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for Jabu-Jabu's Belly
            ' 0: Jabu-Jabu's Belly Beginning
            ' 1: Jabu-Jabu's Belly Main
            ' 2: Jabu-Jabu's Belly Depths
            ' 3: Jabu-Jabu's Belly Boss Area

            With aKeysDungeons(2)(0)
                .loc = "3301"
                .area = "JB1"
                .zone = 80
                .name = "Boomerang Chest"
            End With
            With aKeysDungeons(2)(1)
                .loc = "3302"
                .area = "JB2"
                .zone = 81
                .name = "Map Chest"
            End With
            With aKeysDungeons(2)(2)
                .loc = "3304"
                .area = "JB2"
                .zone = 81
                .name = "Compass Chest"
            End With
            With aKeysDungeons(2)(3)
                .loc = "1231"
                .area = "JB3"
                .zone = 82
                .name = "Barinade"
                .logic = "YLL7416o.Yao"
            End With
            With aKeysDungeons(2)(4)
                .loc = "7819"
                .area = "JB1"
                .zone = 80
                .name = "Water Switch Room"
                .gs = True
            End With
            With aKeysDungeons(2)(5)
                .loc = "7817"
                .area = "JB1"
                .zone = 80
                .name = "Lobby Basement Upper"
                .gs = True
                .logic = "Yo.Zk"
            End With
            With aKeysDungeons(2)(6)
                .loc = "7816"
                .area = "JB1"
                .zone = 80
                .name = "Lobby Basement Lower"
                .gs = True
                .logic = "Yo.Zk"
            End With
            With aKeysDungeons(2)(7)
                .loc = "7818"
                .area = "JB3"
                .zone = 82
                .name = "Near Boss"
                .gs = True
            End With
            With aKeysDungeons(2)(8)
                .loc = "8601"
                .area = "JB1"
                .zone = 80
                .name = "Through Water Passage"
                .scrub = True
                .logic = "Y.ZV11.ZLL4729"
            End With
            tK = 9
        Else
            ' The Master Quest keys for Jabu-Jabu's Belly
            ' 0: MQ Jabu-Jabu's Belly Beginning
            ' 1: MQ Jabu-Jabu's Belly Elevator Room
            ' 2: MQ Jabu-Jabu's Belly Main
            ' 3: MQ Jabu-Jabu's Belly Depth
            ' 4: MQ Jabu-Jabu's Belly Past Big Octo
            ' 5: MQ Jabu-Jabu's Belly Boss Area

            With aKeysDungeons(2)(0)
                .loc = "3303"
                .area = "JB0"
                .zone = 79
                .name = "Map Chest"
                .logic = "Yx.Zx.Zr"
            End With
            With aKeysDungeons(2)(1)
                .loc = "3305"
                .area = "JB0"
                .zone = 79
                .name = "First Room Side Chest"
                .logic = "Yg"
            End With
            With aKeysDungeons(2)(2)
                .loc = "3302"
                .area = "JB1"
                .zone = 83
                .name = "Second Room Lower Chest"
            End With
            With aKeysDungeons(2)(3)
                .loc = "3300"
                .area = "JB1"
                .zone = 83
                .name = "Compass Chest"
                .logic = "Yg.Yj"
            End With
            With aKeysDungeons(2)(4)
                .loc = "3304"
                .area = "JB2"
                .zone = 84
                .name = "Basement Near Vines Chest"
                .logic = "Yg"
            End With
            With aKeysDungeons(2)(5)
                .loc = "3308"
                .area = "JB2"
                .zone = 84
                .name = "Basement Near Switches Chest"
                .logic = "Yg"
            End With
            With aKeysDungeons(2)(6)
                .loc = "3306"
                .area = "JB2"
                .zone = 84
                .name = "Boomerang Chest"
                .logic = "YLL7416.Ya.c.Yg"
            End With
            With aKeysDungeons(2)(7)
                .loc = "3301"
                .area = "JB2"
                .zone = 84
                .name = "Boomerang Room Small Chest"
            End With
            With aKeysDungeons(2)(8)
                .loc = "3309"
                .area = "JB3"
                .zone = 85
                .name = "Falling Like Like Room"
            End With
            With aKeysDungeons(2)(9)
                .loc = "3307"
                .area = "JB5"
                .zone = 87
                .name = "Second Room Upper Chest"
                .logic = "Yg"
            End With
            With aKeysDungeons(2)(10)
                .loc = "3310"
                .area = "JB5"
                .zone = 87
                .name = "Near Boss Chest"
                .logic = "Yg"
            End With
            With aKeysDungeons(2)(11)
                .loc = "1231"
                .area = "JB5"
                .zone = 87
                .name = "Barinade"
                .logic = "YoLL7416.Yoa"
            End With
            With aKeysDungeons(2)(12)
                .loc = "7816"
                .area = "JB2"
                .zone = 84
                .name = "Boomerang Chest Room"
                .gs = True
                .logic = "YhLL7716J"
            End With
            With aKeysDungeons(2)(13)
                .loc = "7818"
                .area = "JB3"
                .zone = 85
                .name = "Tailpasaran Room"
                .gs = True
                .logic = "Ya.f"
            End With
            With aKeysDungeons(2)(14)
                .loc = "7819"
                .area = "JB3"
                .zone = 85
                .name = "Invisible Enemies Room"
                .gs = True
                .logic = "YTgo"
            End With
            With aKeysDungeons(2)(15)
                .loc = "7817"
                .area = "JB5"
                .zone = 87
                .name = "Near Boss"
                .gs = True
                .logic = "Yo"
            End With
            With aKeysDungeons(2)(16)
                .loc = "10224"
                .area = "JB4"
                .zone = 86
                .name = "Red Jelly Room"
                .cow = True
                .logic = "hLL7713"
            End With
            tK = 17
        End If

        With aKeysDungeons(2)(tK)
            .loc = "10529"
            .area = "EVENT"
            .name = "Platform to Boss Area"
        End With
    End Sub
    Private Sub makeKeysForestTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(3) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(3)(28)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(3).Length - 1
            aKeysDungeons(3)(i) = New keyCheck
            aKeysDungeons(3)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for Forest Temple
            ' 0: Forest Temple Lobby
            ' 1: Forest Temple NW Outdoors
            ' 2: Forest Temple NE Outdoors
            ' 3: Forest Temple Outdoors High Balconies
            ' 4: Forest Temple Falling Room
            ' 5: Forest Temple Block Push Room
            ' 6: Forest Temple Straightened Hall
            ' 7: Forest Temple Outside Upper Ledge
            ' 8: Forest Temple Bow Region
            ' 9: Forest Temple Boss Region

            With aKeysDungeons(3)(0)
                .loc = "3403"
                .area = "FOT0"
                .zone = 88
                .name = "First Room Chest"
            End With
            With aKeysDungeons(3)(1)
                .loc = "3400"
                .area = "FOT0"
                .zone = 88
                .name = "First Stalfos Chest"
                .logic = "YLL7416.Z"
            End With
            With aKeysDungeons(3)(2)
                .loc = "3405"
                .area = "FOT2"
                .zone = 90
                .name = "Raised Island Courtyard Chest"
                .logic = "Zk.Q1092.ZGA1LL7430Q1091"
            End With
            With aKeysDungeons(3)(3)
                .loc = "3401"
                .area = "FOT3"
                .zone = 91
                .name = "Map Chest"
            End With
            With aKeysDungeons(3)(4)
                .loc = "3409"
                .area = "FOT3"
                .zone = 91
                .name = "Well Chest"
            End With
            With aKeysDungeons(3)(5)
                .loc = "3404"
                .area = "FOT5"
                .zone = 93
                .name = "Eye Switch Chest"
                .logic = "YV01g.ZV01d"
            End With
            With aKeysDungeons(3)(6)
                .loc = "3414"
                .area = "FOT6"
                .zone = 94
                .name = "Boss Key Chest"
            End With
            With aKeysDungeons(3)(7)
                .loc = "3402"
                .area = "FOT7"
                .zone = 95
                .name = "Floormaster Chest"
            End With
            With aKeysDungeons(3)(8)
                .loc = "3413"
                .area = "FOT8"
                .zone = 96
                .name = "Red Poe Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(3)(9)
                .loc = "3412"
                .area = "FOT8"
                .zone = 96
                .name = "Bow Chest"
            End With
            With aKeysDungeons(3)(10)
                .loc = "3415"
                .area = "FOT8"
                .zone = 96
                .name = "Blue Poe Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(3)(11)
                .loc = "3407"
                .area = "FOT4"
                .zone = 92
                .name = "Falling Ceiling Room Chest"
            End With
            With aKeysDungeons(3)(12)
                .loc = "3411"
                .area = "FOT9"
                .zone = 97
                .name = "Basement Chest"
                .logic = "YJ.Z.b"
            End With
            With aKeysDungeons(3)(13)
                .loc = "1331"
                .area = "FOT9"
                .zone = 97
                .name = "Phantom Ganon"
                .logic = "YLL7416.Yg.ZGB0d.ZGB0k"
            End With
            With aKeysDungeons(3)(14)
                .loc = "7825"
                .area = "FOT0"
                .zone = 88
                .name = "First Room"
                .gs = True
                .logic = "Yg.Yo.Zc.Zd.Zk.f.j"
                '.logic = "Zc.Zd.f.j.Zk"
            End With
            With aKeysDungeons(3)(15)
                .loc = "7827"
                .area = "FOT0"
                .zone = 88
                .name = "Lobby"
                .gs = True
                .logic = "Yo.Zk"
            End With
            With aKeysDungeons(3)(16)
                .loc = "7826"
                .area = "FOT1"
                .zone = 89
                .name = "Level Island Courtyard"
                .gs = True
                .logic = "Zl.ZQ1095k"
            End With
            With aKeysDungeons(3)(17)
                .loc = "7824"
                .area = "FOT2"
                .zone = 90
                .name = "Raised Island Courtyard"
                .gs = True
                .logic = "Zk.ZQ1092d.Q1092f.Q1092x"
            End With
            With aKeysDungeons(3)(18)
                .loc = "7828"
                .area = "FOT9"
                .zone = 97
                .name = "Basement"
                .gs = True
                .logic = "Yo.Zk"
            End With
            With aKeysDungeons(3)(19)
                .loc = "10600"
                .area = "EVENT"
                .name = "Door 1"
            End With
            With aKeysDungeons(3)(20)
                .loc = "10601"
                .area = "EVENT"
                .name = "Door 2"
            End With
            With aKeysDungeons(3)(21)
                .loc = "10602"
                .area = "EVENT"
                .name = "Door 3"
            End With
            With aKeysDungeons(3)(22)
                .loc = "10603"
                .area = "EVENT"
                .name = "Door 4"
            End With
            With aKeysDungeons(3)(23)
                .loc = "10604"
                .area = "EVENT"
                .name = "Door 5"
            End With
            tK = 24
        Else
            ' The Master Quest keys for Forest Temple
            ' 0: MQ Forest Temple Lobby
            ' 1: MQ Forest Temple Central Area
            ' 2: MQ Forest Temple After Block Puzzle
            ' 3: MQ Forest Temple Outdoor Ledge
            ' 4: MQ Forest Temple NW Outdoors
            ' 5: MQ Forest Temple NE Outdoors
            ' 6: MQ Forest Temple Outdoors Top Ledges
            ' 7: MQ Forest Temple NW Outdoor Ledge
            ' 8: MQ Forest Temple Bow Region
            ' 9: MQ Forest Temple Falling Room
            ' A: MQ Forest Temple Boss Region

            With aKeysDungeons(3)(0)
                .loc = "3403"
                .area = "FOT0"
                .zone = 88
                .name = "First Room Chest"
                .logic = "YLL7416.Ya.Yb.Yc.Yf.Yg.Yo.Z"
            End With
            With aKeysDungeons(3)(1)
                .loc = "3400"
                .area = "FOT1"
                .zone = 98
                .name = "Wolfos Chest"
                .logic = "YLL7416.Ya.Yf.Yg.ZhLL7716"
            End With
            With aKeysDungeons(3)(2)
                .loc = "3409"
                .area = "FOT5"
                .zone = 102
                .name = "Well Chest"
                .logic = "Yg.Zd"
            End With
            With aKeysDungeons(3)(3)
                .loc = "3401"
                .area = "FOT7"
                .zone = 104
                .name = "Raised Island Courtyard Lower Chest"
            End With
            With aKeysDungeons(3)(4)
                .loc = "3405"
                .area = "FOT6"
                .zone = 103
                .name = "Raised Island Courtyard Upper Chest"
            End With
            With aKeysDungeons(3)(5)
                .loc = "3414"
                .area = "FOT2"
                .zone = 99
                .name = "Boss Key Chest"
                .logic = "GDC"
            End With
            With aKeysDungeons(3)(6)
                .loc = "3402"
                .area = "FOT3"
                .zone = 100
                .name = "Redead Chest"
            End With
            With aKeysDungeons(3)(7)
                .loc = "3413"
                .area = "FOT8"
                .zone = 105
                .name = "Map Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(3)(8)
                .loc = "3412"
                .area = "FOT8"
                .zone = 105
                .name = "Bow Chest"
            End With
            With aKeysDungeons(3)(9)
                .loc = "3415"
                .area = "FOT8"
                .zone = 105
                .name = "Compass Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(3)(10)
                .loc = "3406"
                .area = "FOT9"
                .zone = 106
                .name = "Falling Ceiling Room Chest"
            End With
            With aKeysDungeons(3)(11)
                .loc = "3411"
                .area = "FOTA"
                .zone = 107
                .name = "Basement Chest"
            End With
            With aKeysDungeons(3)(12)
                .loc = "1331"
                .area = "FOTA"
                .zone = 107
                .name = "Phantom Ganon"
                .logic = "YLL7416.Yg.ZGB0d.ZGB0k"
            End With
            With aKeysDungeons(3)(13)
                .loc = "7825"
                .area = "FOT0"
                .zone = 88
                .name = "First Hallway"
                .gs = True
                .logic = "Yo.Zk"
            End With
            With aKeysDungeons(3)(14)
                .loc = "7824"
                .area = "FOT5"
                .zone = 102
                .name = "Raised Island Courtyard"
                .gs = True
                .logic = "Yo.Zk.ZdeLL7716.ZdeLL7430GA2"
                '.logic = "Zk.ZdeLL7716.ZdeLL7430GA2"
            End With
            With aKeysDungeons(3)(15)
                .loc = "7826"
                .area = "FOT4"
                .zone = 101
                .name = "Level Island Courtyard"
                .gs = True
            End With
            With aKeysDungeons(3)(16)
                .loc = "7827"
                .area = "FOT5"
                .zone = 102
                .name = "Well"
                .gs = True
                .logic = "Yg.ZLL7429k.Zd"
            End With
            With aKeysDungeons(3)(17)
                .loc = "7828"
                .area = "FOT1"
                .zone = 98
                .name = "Block Push Room"
                .gs = True
                .logic = "YLL7416.Z"
            End With
            With aKeysDungeons(3)(18)
                .loc = "10600"
                .area = "EVENT"
                .name = "Door: Beth"
            End With
            With aKeysDungeons(3)(19)
                .loc = "10601"
                .area = "EVENT"
                .name = "Door: First Twisted Hallway"
            End With
            With aKeysDungeons(3)(20)
                .loc = "10602"
                .area = "EVENT"
                .name = "Door: Push Block Room"
            End With
            With aKeysDungeons(3)(21)
                .loc = "10603"
                .area = "EVENT"
                .name = "Door: Falling Ceiling Room"
            End With
            With aKeysDungeons(3)(22)
                .loc = "10604"
                .area = "EVENT"
                .name = "Door: Second Twisted Hallway"
            End With
            With aKeysDungeons(3)(23)
                .loc = "10606"
                .area = "EVENT"
                .name = "Door: Lobby"
            End With
            tK = 24
        End If

        With aKeysDungeons(3)(tK)
            .loc = "10620"
            .area = "EVENT"
            .name = "Boss Door Unlocked"
        End With
        incB(tK)
        With aKeysDungeons(3)(tK)
            .loc = "10628"
            .area = "EVENT"
            .name = "Meg"
        End With
        incB(tK)
        With aKeysDungeons(3)(tK)
            .loc = "10629"
            .area = "EVENT"
            .name = "Joelle"
        End With
        incB(tK)
        With aKeysDungeons(3)(tK)
            .loc = "10630"
            .area = "EVENT"
            .name = "Beth"
        End With
        incB(tK)
        With aKeysDungeons(3)(tK)
            .loc = "10631"
            .area = "EVENT"
            .name = "Amy"
        End With
    End Sub
    Private Sub makeKeysFireTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(4) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(4)(CInt(IIf(isMQ, 24, 29)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(4).Length - 1
            aKeysDungeons(4)(i) = New keyCheck
            aKeysDungeons(4)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Fire Temple
            ' 0: Fire Temple Lower
            ' 1: Fire Temple Big Lava Room
            ' 2: Fire Temple Middle
            ' 3: Fire Temple Upper
            ' 4: Fire Temple Lobby Back Room

            With aKeysDungeons(4)(0)
                .loc = "3501"
                .area = "FIT0"
                .zone = 108
                .name = "Near Boss Chest"
                .logic = "G09"
            End With
            With aKeysDungeons(4)(1)
                .loc = "3500"
                .area = "FIT4"
                .zone = 112
                .name = "Flare Dancer Chest"
            End With
            With aKeysDungeons(4)(2)
                .loc = "3512"
                .area = "FIT4"
                .zone = 112
                .name = "Boss Key Chest"
            End With
            With aKeysDungeons(4)(3)
                .loc = "3504"
                .area = "FIT1"
                .zone = 109
                .name = "Big Lava Room Lower Open Door Chest"
            End With
            With aKeysDungeons(4)(4)
                .loc = "3502"
                .area = "FIT1"
                .zone = 109
                .name = "Big Lava Room Blocked Door Chest"
                .logic = "Zx"
            End With
            With aKeysDungeons(4)(5)
                .loc = "3503"
                .area = "FIT2"
                .zone = 110
                .name = "Boulder Maze Lower Chest"
            End With
            With aKeysDungeons(4)(6)
                .loc = "3508"
                .area = "FIT2"
                .zone = 110
                .name = "Boulder Maze Side Room Chest"
            End With
            With aKeysDungeons(4)(7)
                .loc = "3510"
                .area = "FIT2"
                .zone = 110
                .name = "Map Chest"
                .logic = "GC2.ZGC1d"
            End With
            With aKeysDungeons(4)(8)
                .loc = "3506"
                .area = "FIT2"
                .zone = 110
                .name = "Boulder Maze Upper Chest"
                .logic = "GC2"
            End With
            With aKeysDungeons(4)(9)
                .loc = "3511"
                .area = "FIT2"
                .zone = 110
                .name = "Boulder Maze Shortcut Chest"
                .logic = "GC2x"
            End With
            With aKeysDungeons(4)(10)
                .loc = "3513"
                .area = "FIT2"
                .zone = 110
                .name = "Scarecrow Chest"
                .logic = "ZGC2GAC"
            End With
            With aKeysDungeons(4)(11)
                .loc = "3507"
                .area = "FIT2"
                .zone = 110
                .name = "Compass Chest"
                .logic = "GDD"
            End With
            With aKeysDungeons(4)(12)
                .loc = "3509"
                .area = "FIT3"
                .zone = 111
                .name = "Highest Goron Chest"
                .logic = "ZrhLL7716.ZrGA3LL7430.ZrGA3x"
            End With
            With aKeysDungeons(4)(13)
                .loc = "3505"
                .area = "FIT3"
                .zone = 111
                .name = "Megaton Hammer Chest"
                .logic = "x"
            End With
            With aKeysDungeons(4)(14)
                .loc = "1431"
                .area = "FIT0"
                .zone = 108
                .name = "Volvagia"
                .logic = "ZGB1LL7425rLL7430.ZGB1LL7425rL10718.ZGB1LL7425rQ1111hLL7716.ZGB1LL7425rQ1111x"
            End With
            With aKeysDungeons(4)(15)
                .loc = "7901"
                .area = "FIT4"
                .zone = 112
                .name = "Boss Key Loop"
                .gs = True
            End With
            With aKeysDungeons(4)(16)
                .loc = "7900"
                .area = "FIT1"
                .zone = 109
                .name = "Song of Time Room"
                .gs = True
                .logic = "ZhLL7716"
            End With
            With aKeysDungeons(4)(17)
                .loc = "7902"
                .area = "FIT2"
                .zone = 110
                .name = "Boulder Maze"
                .gs = True
                .logic = "GC0x"
            End With
            With aKeysDungeons(4)(18)
                .loc = "7904"
                .area = "FIT2"
                .zone = 110
                .name = "Scarecrow Climb"
                .gs = True
                .logic = "ZGC2GAC"
            End With
            With aKeysDungeons(4)(19)
                .loc = "7903"
                .area = "FIT2"
                .zone = 110
                .name = "Scarecrow Top"
                .gs = True
                .logic = "ZGC2GAC"
            End With
            With aKeysDungeons(4)(20)
                .loc = "10729"
                .area = "EVENT"
                .name = "Door 1: Entrance to BLR"
            End With
            With aKeysDungeons(4)(21)
                .loc = "10730"
                .area = "EVENT"
                .name = "Door 2: BLR to Elevator"
            End With
            With aKeysDungeons(4)(22)
                .loc = "10724"
                .area = "EVENT"
                .name = "Door 3: Elevator Exit"
            End With
            With aKeysDungeons(4)(23)
                .loc = "10727"
                .area = "EVENT"
                .name = "Door 4: Maze Lower Exit"
            End With
            With aKeysDungeons(4)(24)
                .loc = "10726"
                .area = "EVENT"
                .name = "Door 5: Above BLR"
            End With
            With aKeysDungeons(4)(25)
                .loc = "10731"
                .area = "EVENT"
                .name = "Door 6: Flame Trap Room"
            End With
            With aKeysDungeons(4)(26)
                .loc = "10725"
                .area = "EVENT"
                .name = "Door 7: Flame Walls Room"
            End With
            With aKeysDungeons(4)(27)
                .loc = "10723"
                .area = "EVENT"
                .name = "Keysanity Door in Lobby to Back Room"
            End With
            With aKeysDungeons(4)(28)
                .loc = "10718"
                .area = "EVENT"
                .name = "Block to Reach Boss Door"
            End With
            tK = 29
        Else
            ' The Master Quest keys for the Fire Temple
            ' 0: MQ Fire Temple Lower
            ' 1: MQ Fire Temple Locked Door
            ' 2: MQ Fire Temple Big Lava Room
            ' 3: MQ Fire Temple Lower Maze
            ' 4: MQ Fire Temple Upper Maze
            ' 5: MQ Fire Temple Upper
            ' 6: MQ Fire Temple Boss Room

            With aKeysDungeons(4)(0)
                .loc = "3507"
                .area = "FIT0"
                .zone = 108
                .name = "Near Boss Chest"
                .logic = "Yf.Zf.ZLL7430G24.Zkde"
            End With
            With aKeysDungeons(4)(1)
                .loc = "3502"
                .area = "FIT0"
                .zone = 108
                .name = "Map Room Side Chest"
                .logic = "YLL7416.Ya.Yc.Yf.Yg.Z"
            End With
            With aKeysDungeons(4)(2)
                .loc = "3500"
                .area = "FIT1"
                .zone = 108
                .name = "Megaton Hammer Chest"
                .logic = "Zk.Zr.Zx"
            End With
            With aKeysDungeons(4)(3)
                .loc = "3512"
                .area = "FIT1"
                .zone = 108
                .name = "Map Chest"
                .logic = "Zr"
            End With
            With aKeysDungeons(4)(4)
                .loc = "3501"
                .area = "FIT2"
                .zone = 114
                .name = "Big Lava Room Blocked Door Chest"
                .logic = "ZkxG24"
            End With
            With aKeysDungeons(4)(5)
                .loc = "3504"
                .area = "FIT2"
                .zone = 114
                .name = "Boss Key Chest"
                .logic = "ZdkG24"
            End With
            With aKeysDungeons(4)(6)
                .loc = "3503"
                .area = "FIT3"
                .zone = 115
                .name = "Lizalfos Maze Lower Chest"
            End With
            With aKeysDungeons(4)(7)
                .loc = "3508"
                .area = "FIT3"
                .zone = 115
                .name = "Lizalfos Maze Side Room Chest"
                .logic = "ZQ1116x"
            End With
            With aKeysDungeons(4)(8)
                .loc = "3506"
                .area = "FIT4"
                .zone = 116
                .name = "Lizalfos Maze Upper Chest"
            End With
            With aKeysDungeons(4)(9)
                .loc = "3511"
                .area = "FIT4"
                .zone = 116
                .name = "Compass Chest"
                .logic = "x"
            End With
            With aKeysDungeons(4)(10)
                .loc = "328"
                .area = "FIT5"
                .zone = 117
                .name = "Under Platform Freestanding Key"
                .logic = "Zk.GA4"
            End With
            With aKeysDungeons(4)(11)
                .loc = "3505"
                .area = "FIT5"
                .zone = 117
                .name = "Chest on Fire"
                .logic = "ZGC3k.GC3GA4"
            End With
            With aKeysDungeons(4)(12)
                .loc = "1431"
                .area = "FIT6"
                .zone = 118
                .name = "Volvagia"
                .logic = "ZGB1"
            End With
            With aKeysDungeons(4)(13)
                .loc = "7900"
                .area = "FIT2"
                .zone = 114
                .name = "Big Lava Room Open Door"
                .gs = True
            End With
            With aKeysDungeons(4)(14)
                .loc = "7902"
                .area = "FIT4"
                .zone = 116
                .name = "Skulltula On Fire"
                .gs = True
                .logic = "Zl.ZhxkLL7716.ZGA3khLL7716"
            End With
            With aKeysDungeons(4)(15)
                .loc = "7903"
                .area = "FIT5"
                .zone = 117
                .name = "Fire Wall Maze Centre"
                .gs = True
                .logic = "x"
            End With
            With aKeysDungeons(4)(16)
                .loc = "7904"
                .area = "FIT5"
                .zone = 117
                .name = "Fire Wall Maze Side Room"
                .gs = True
                .logic = "GA4.ZLL7430.hLL7716"
            End With
            With aKeysDungeons(4)(17)
                .loc = "7901"
                .area = "FIT5"
                .zone = 117
                .name = "Above Fire Wall Maze"
                .gs = True
                .logic = "ZGC4k"
            End With
            With aKeysDungeons(4)(18)
                .loc = "10723"
                .area = "EVENT"
                .name = "Door to Boss Loop"
            End With
            With aKeysDungeons(4)(19)
                .loc = "10730"
                .area = "EVENT"
                .name = "Door to Exit BLR"
            End With
            With aKeysDungeons(4)(20)
                .loc = "10726"
                .area = "EVENT"
                .name = "Door to Exit Maze"
            End With
            With aKeysDungeons(4)(21)
                .loc = "10724"
                .area = "EVENT"
                .name = "Door after Flare Dancer"
            End With
            With aKeysDungeons(4)(22)
                .loc = "10727"
                .area = "EVENT"
                .name = "Door after Top Area"
            End With
            With aKeysDungeons(4)(23)
                .loc = "10710"
                .area = "EVENT"
                .name = "Block to Reach Boss Door"
            End With
            tK = 24
        End If

        With aKeysDungeons(4)(tK)
            .loc = "10720"
            .area = "EVENT"
            .name = "Boss Door Unlocked"
        End With
    End Sub
    Private Sub makeKeysWaterTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old arra
        aMQOld(5) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(5)(CInt(IIf(isMQ, 14, 21)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(5).Length - 1
            aKeysDungeons(5)(i) = New keyCheck
            aKeysDungeons(5)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Water Temple
            ' 0: Water Temple Lobby
            ' 1: Water Temple Highest Water Level
            ' 2: Water Temple Dive
            ' 3: Water Temple North Basement
            ' 4: Water Temple Cracked Wall
            ' 5: Water Temple Dragon Statue
            ' 6: Water Temple Middle Water Level
            ' 7: Water Temple Falling Platform Room
            ' 8: Water Temple Dark Link Region

            With aKeysDungeons(5)(0)
                .loc = "3602"
                .area = "WAT2"
                .zone = 121
                .name = "Map Chest"
                .logic = "Q1120"
            End With
            With aKeysDungeons(5)(1)
                .loc = "3600"
                .area = "WAT4"
                .zone = 123
                .name = "Cracked Wall Chest"
                .logic = "x"
            End With
            With aKeysDungeons(5)(2)
                .loc = "3601"
                .area = "WAT2"
                .zone = 121
                .name = "Torches Chest"
                .logic = "YLL7416MahLL7712.YfhLL7712.ZdhLL7712.ZfhLL7712"
                ' .logic = "ZdehLL7712.fhLL7712"
            End With
            With aKeysDungeons(5)(3)
                .loc = "3609"
                .area = "WAT2"
                .zone = 121
                .name = "Compass Chest"
                .logic = "ZhLL7716k.ZLL7429k"
            End With
            With aKeysDungeons(5)(4)
                .loc = "3608"
                .area = "WAT2"
                .zone = 121
                .name = "Central Bow Target Chest"
                .logic = "ZV01hLL7712dLL7730.ZV01hLL7712dl"
            End With
            With aKeysDungeons(5)(5)
                .loc = "3606"
                .area = "WAT6"
                .zone = 125
                .name = "Central Pillar Chest"
                .logic = "ZLL7429LL7426kGC5.ZLL7429LL7426kd.ZLL7429LL7426kf"
            End With
            With aKeysDungeons(5)(6)
                .loc = "3607"
                .area = "WAT8"
                .zone = 127
                .name = "Longshot Chest"
            End With
            With aKeysDungeons(5)(7)
                .loc = "3603"
                .area = "WAT8"
                .zone = 127
                .name = "River Chest"
                .logic = "ZhLL7716d"
            End With
            With aKeysDungeons(5)(8)
                .loc = "3610"
                .area = "WAT5"
                .zone = 124
                .name = "Dragon Chest"
            End With
            With aKeysDungeons(5)(9)
                .loc = "3605"
                .area = "WAT3"
                .zone = 122
                .name = "Boss Key Chest"
                .logic = "ZGC6LL7429LL7430.ZGC6LL7429xV01"
            End With
            With aKeysDungeons(5)(10)
                .loc = "1531"
                .area = "WAT1"
                .zone = 120
                .name = "Morpha"
                .logic = "ZGB2l"
            End With
            With aKeysDungeons(5)(11)
                .loc = "7908"
                .area = "WAT2"
                .zone = 121
                .name = "Behind Gate"
                .gs = True
                .logic = "ZkxhLL7712LL7429.ZLL7430xhLL7712LL7429.ZkxhLL7712V11.ZLL7430xhLL7712V11"
            End With
            With aKeysDungeons(5)(12)
                .loc = "7911"
                .area = "WAT3"
                .zone = 122
                .name = "Near Boss Key Chest"
                .gs = True
            End With
            With aKeysDungeons(5)(13)
                .loc = "7910"
                .area = "WAT2"
                .zone = 121
                .name = "Central Pillar"
                .gs = True
                .logic = "ZhLL7712lGC5.ZhLL7712ld.ZhLL7712lf"
            End With
            With aKeysDungeons(5)(14)
                .loc = "7909"
                .area = "WAT7"
                .zone = 126
                .name = "Falling Platform Room"
                .gs = True
                .logic = "Zl"
            End With
            With aKeysDungeons(5)(15)
                .loc = "7912"
                .area = "WAT8"
                .zone = 127
                .name = "River"
                .gs = True
                .logic = "ZhLL7716LL7429G0A"
            End With
            With aKeysDungeons(5)(16)
                .loc = "10806"
                .area = "EVENT"
                .name = "Door Central Pillar"
            End With
            With aKeysDungeons(5)(17)
                .loc = "10805"
                .area = "EVENT"
                .name = "Door Highest Level"
            End With
            With aKeysDungeons(5)(18)
                .loc = "10802"
                .area = "EVENT"
                .name = "Door Sliding Platforms"
            End With
            With aKeysDungeons(5)(19)
                .loc = "10801"
                .area = "EVENT"
                .name = "Door Basement Whirlpools"
            End With
            With aKeysDungeons(5)(20)
                .loc = "10809"
                .area = "EVENT"
                .name = "Door Three Geysers"
            End With
            tK = 21
        Else
            ' The Master Quest keys for the Water Temple
            ' 0: MQ Water Temple Lobby
            ' 1: MQ Water Temple Dive
            ' 2: MQ Water Temple Lowered Water Levels
            ' 3: MQ Water Temple Dark Link Region
            ' 4: MQ Water Temple Basement Gated Areas

            With aKeysDungeons(5)(0)
                .loc = "3602"
                .area = "WAT1"
                .zone = 128
                .name = "Map Chest"
                .logic = "ZkG24"
            End With
            With aKeysDungeons(5)(1)
                .loc = "3600"
                .area = "WAT2"
                .zone = 129
                .name = "Longshot Chest"
                .logic = "Zk"
            End With
            With aKeysDungeons(5)(2)
                .loc = "3601"
                .area = "WAT2"
                .zone = 129
                .name = "Compass Chest"
                .logic = "Zd.Zf"
            End With
            With aKeysDungeons(5)(3)
                .loc = "3606"
                .area = "WAT1"
                .zone = 128
                .name = "Central Pillar Chest"
                .logic = "ZLL7426kfhLL7716"
            End With
            With aKeysDungeons(5)(4)
                .loc = "3605"
                .area = "WAT3"
                .zone = 130
                .name = "Boss Key Chest"
                .logic = "ZG0AfV11.ZG0AfLL7429"
            End With
            With aKeysDungeons(5)(5)
                .loc = "401"
                .area = "WAT4"
                .zone = 131
                .name = "Stalfos Room Freestanding Key"
                .logic = "ZLL7430.ZhLL6512k"
            End With
            With aKeysDungeons(5)(6)
                .loc = "1531"
                .area = "WAT0"
                .zone = 128
                .name = "Morpha"
                .logic = "ZGB2l"
            End With
            With aKeysDungeons(5)(7)
                .loc = "7908"
                .area = "WAT2"
                .zone = 129
                .name = "Lizalfos Hallway"
                .gs = True
                .logic = "f"
            End With
            With aKeysDungeons(5)(8)
                .loc = "7909"
                .area = "WAT3"
                .zone = 130
                .name = "River"
                .gs = True
            End With
            With aKeysDungeons(5)(9)
                .loc = "7910"
                .area = "WAT2"
                .zone = 129
                .name = "Before Upper Water Switch"
                .gs = True
                .logic = "Zl"
            End With
            With aKeysDungeons(5)(10)
                .loc = "7911"
                .area = "WAT4"
                .zone = 131
                .name = "Freestanding Key Area"
                .gs = True
                .logic = "ZGC7LL7430.ZGC7hLL6512k"
            End With
            With aKeysDungeons(5)(11)
                .loc = "7912"
                .area = "WAT4"
                .zone = 131
                .name = "Triple Wall Torch"
                .gs = True
                .logic = "ZdeLL7430.ZdehLL6512k"
            End With
            With aKeysDungeons(5)(12)
                .loc = "10821"
                .area = "EVENT"
                .name = "Door to Dark Link"
            End With
            With aKeysDungeons(5)(13)
                .loc = "10802"
                .area = "EVENT"
                .name = "Door in Basement"
            End With
            tK = 14
        End If

        With aKeysDungeons(5)(tK)
            .loc = "10820"
            .area = "EVENT"
            .name = "Boss Door Unlocked"
        End With
    End Sub
    Private Sub makeKeysSpiritTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(6) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0


        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(6)(CInt(IIf(isMQ, 39, 35)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(6).Length - 1
            aKeysDungeons(6)(i) = New keyCheck
            aKeysDungeons(6)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the  Spirit Temple
            ' 0: Spirit Temple Lobby
            ' 1: Child Spirit Temple
            ' 2: Child Spirit Before Locked Door
            ' 3: Child Spirit Temple Climb
            ' 4: Early Adult Spirit Temple
            ' 5: Spirit Temple Central Chamber
            ' 6: Spirit Temple Outdoor Hands
            ' 7: Spirit Temple Beyond Central Locked Door
            ' 8: Spirit Temple Beyond Final Locked Door
            ' 9: Spirit Temple Boss Platform

            With aKeysDungeons(6)(0)
                .loc = "3708"
                .area = "SPT1"
                .zone = 133
                .name = "Child Bridge Chest"
                .logic = "Yga.Ygb.Ygo.Ygx.Yoa.Yox.YoLL7416.YjGA5a.YjGA5x.YjGA5bLL7416"
            End With
            With aKeysDungeons(6)(1)
                .loc = "3700"
                .area = "SPT1"
                .zone = 133
                .name = "Child Early Torches Chest"
                .logic = "Yga.Ygbf.Ygof.Ygxf.Yoa.Yoxf.YoLL7416f.YjGA5a.YjGA5xf.YjGA5bLL7416f"
            End With
            With aKeysDungeons(6)(2)
                .loc = "3706"
                .area = "SPT3"
                .zone = 135
                .name = "Child Climb North Chest"
                .logic = "YG20.ZG21"
            End With
            With aKeysDungeons(6)(3)
                .loc = "3712"
                .area = "SPT3"
                .zone = 135
                .name = "Child Climb East Chest"
                .logic = "YG20.ZG21"
            End With
            With aKeysDungeons(6)(4)
                .loc = "3701"
                .area = "SPT5"
                .zone = 137
                .name = "Sun Block Room Chest"
                .logic = "Ya.Zde.f"
            End With
            With aKeysDungeons(6)(5)
                .loc = "3703"
                .area = "SPT5"
                .zone = 137
                .name = "Map Chest"
                .logic = "f.Zde"
            End With
            With aKeysDungeons(6)(6)
                .loc = "5511"
                .area = "SPT6"
                .zone = 138
                .name = "Silver Gauntlets Chest"
                .logic = "YGC8.ZGC9"
            End With
            With aKeysDungeons(6)(7)
                .loc = "3704"
                .area = "SPT4"
                .zone = 136
                .name = "Compass Chest"
                .logic = "ZkhLL7712"
            End With
            With aKeysDungeons(6)(8)
                .loc = "3707"
                .area = "SPT4"
                .zone = 136
                .name = "Early Adult Right Chest"
                .logic = "Zd.j.Zk"
            End With
            With aKeysDungeons(6)(9)
                .loc = "3713"
                .area = "SPT5"
                .zone = 137
                .name = "First Mirror Left Chest"
                .logic = "ZV02.ZG14k"
            End With
            With aKeysDungeons(6)(10)
                .loc = "3714"
                .area = "SPT5"
                .zone = 137
                .name = "First Mirror Right Chest"
                .logic = "ZV02.ZG14k"
            End With
            With aKeysDungeons(6)(11)
                .loc = "3702"
                .area = "SPT5"
                .zone = 137
                .name = "Statue Room Hand Chest"
                '.logic = "ZQ1141hLL7712.ZLL7430hLL7712"
                .logic = "ZG14khLL7712.ZV02hLL7712"
            End With
            With aKeysDungeons(6)(12)
                .loc = "3715"
                .area = "SPT5"
                .zone = 137
                .name = "Statue Room Northeast Chest"
                '.logic = "ZQ1141hLL7712k.ZhLL7712LL7430"
                .logic = "ZG14khLL7712.ZV02hLL7712LL7430"
            End With
            With aKeysDungeons(6)(13)
                .loc = "3705"
                .area = "SPT7"
                .zone = 139
                .name = "Near Four Armos Chest"
                .logic = "ZLL7422x"
            End With
            With aKeysDungeons(6)(14)
                .loc = "3721"
                .area = "SPT7"
                .zone = 139
                .name = "Hallway Left Invisible Chest"
                .logic = "G15x"
            End With
            With aKeysDungeons(6)(15)
                .loc = "3720"
                .area = "SPT7"
                .zone = 139
                .name = "Hallway Right Invisible Chest"
                .logic = "G15x"
            End With
            With aKeysDungeons(6)(16)
                .loc = "5509"
                .area = "SPT7"
                .zone = 139
                .name = "Mirror Shield Chest"
                .logic = "x"
            End With
            With aKeysDungeons(6)(17)
                .loc = "3710"
                .area = "SPT8"
                .zone = 140
                .name = "Boss Key Chest"
                .logic = "ZhLL7712dk"
            End With
            With aKeysDungeons(6)(18)
                .loc = "3718"
                .area = "SPT8"
                .zone = 140
                .name = "Topmost Chest"
                .logic = "ZLL7422"
            End With
            With aKeysDungeons(6)(19)
                .loc = "1631"
                .area = "SPT9"
                .zone = 141
                .name = "Twinrova"
                .logic = "ZGB3LL7422k"
            End With
            With aKeysDungeons(6)(20)
                .loc = "7920"
                .area = "SPT1"
                .zone = 133
                .name = "Metal Fence"
                .gs = True
                .logic = "Yga.Ygb.Ygo.Ygx.Yoa.Yox.YoLL7416.YjGA5a.YjGA5x.YjGA5bLL7416"
            End With
            With aKeysDungeons(6)(21)
                .loc = "7919"
                .area = "SPT3"
                .zone = 135
                .name = "Sun on Floor Room"
                .gs = True
                .logic = "f.YG20.Z"
            End With
            With aKeysDungeons(6)(22)
                .loc = "7916"
                .area = "SPT5"
                .zone = 137
                .name = "Hall After Sun Block Room"
                .gs = True
                .logic = "Yo.Zk"
            End With
            With aKeysDungeons(6)(23)
                .loc = "7918"
                .area = "SPT5"
                .zone = 137
                .name = "Lobby"
                .gs = True
                .logic = "ZLL7430.ZkhLL6512"
            End With
            With aKeysDungeons(6)(24)
                .loc = "7917"
                .area = "SPT4"
                .zone = 136
                .name = "Boulder Room"
                .gs = True
                .logic = "ZhLL7716d.hLL7716j.ZhLL7716k"
            End With
            With aKeysDungeons(6)(25)
                .loc = "10930"
                .area = "EVENT"
                .name = "Door Young First Door"
            End With
            With aKeysDungeons(6)(26)
                .loc = "10921"
                .area = "EVENT"
                .name = "Door Statue Right (Young 2)"
            End With
            With aKeysDungeons(6)(27)
                .loc = "10913"
                .area = "EVENT"
                .name = "Door Adult First Door"
            End With
            With aKeysDungeons(6)(28)
                .loc = "10927"
                .area = "EVENT"
                .name = "Door Statue Left (Adult 2)"
            End With
            With aKeysDungeons(6)(29)
                .loc = "10928"
                .area = "EVENT"
                .name = "Door Mummy Room (Adult 3)"
            End With
            With aKeysDungeons(6)(30)
                .loc = "10911"
                .area = "EVENT"
                .name = "Shortcut Part 3"
            End With
            tK = 31
        Else
            ' The Master Quest keys for the Spirit Temple
            ' 0: MQ Spirit Temple Lobby
            ' 1: MQ Child Spirit Temple
            ' 2: MQ Adult Spirit Temple
            ' 3: MQ Spirit Temple Shared
            ' 4: MQ Lower Adult Spirit Temple
            ' 5: MQ Spirit Temple Boss Area
            ' 6: MQ Spirit Temple Boss Platform
            ' 7: MQ Mirror Shield Hand
            ' 8: MQ Silver Gauntlets Hand

            With aKeysDungeons(6)(0)
                .loc = "3726"
                .area = "SPT0"
                .zone = 132
                .name = "Entrance Front Left Chest"
            End With
            With aKeysDungeons(6)(1)
                .loc = "3730"
                .area = "SPT0"
                .zone = 132
                .name = "Entrance Back Left Chest"
                .logic = "Yxg.ZG00d"
            End With
            With aKeysDungeons(6)(2)
                .loc = "3731"
                .area = "SPT0"
                .zone = 132
                .name = "Entrance Back Right Chest"
                .logic = "j.Zd.Zk.Yg.Yo"
            End With
            With aKeysDungeons(6)(3)
                .loc = "3708"
                .area = "SPT1"
                .zone = 142
                .name = "Map Room Enemy Chest"
                .logic = "Yajgf.YLL7416jgf"
            End With
            With aKeysDungeons(6)(4)
                .loc = "3700"
                .area = "SPT1"
                .zone = 142
                .name = "Map Chest"
                .logic = "Ya.YLL7416.c"
            End With
            With aKeysDungeons(6)(5)
                .loc = "3729"
                .area = "SPT1"
                .zone = 142
                .name = "Child Hammer Switch Chest"
                .logic = "YQ1143GCAr"
            End With
            With aKeysDungeons(6)(6)
                .loc = "3728"
                .area = "SPT3"
                .zone = 144
                .name = "Silver Block Hallway Chest"
                .logic = "YQ1144de.Yg"
            End With
            With aKeysDungeons(6)(7)
                .loc = "3703"
                .area = "SPT3"
                .zone = 144
                .name = "Compass Chest"
                .logic = "Yg.Zd"
            End With
            With aKeysDungeons(6)(8)
                .loc = "3701"
                .area = "SPT3"
                .zone = 144
                .name = "Sun Block Room Chest"
                .logic = "Z.YhLL7716.YGA6"
            End With
            With aKeysDungeons(6)(9)
                .loc = "5511"
                .area = "SPT8"
                .zone = 149
                .name = "Silver Gauntlets Chest"
            End With
            With aKeysDungeons(6)(10)
                .loc = "3706"
                .area = "SPT1"
                .zone = 142
                .name = "Child Climb North Chest"
                .logic = "x"
            End With
            With aKeysDungeons(6)(11)
                .loc = "3712"
                .area = "SPT2"
                .zone = 143
                .name = "Child Climb South Chest"
                .logic = "Zx"
            End With
            With aKeysDungeons(6)(12)
                .loc = "3715"
                .area = "SPT2"
                .zone = 143
                .name = "Statue Room Lullaby Chest"
                .logic = "hLL7712"
            End With
            With aKeysDungeons(6)(13)
                .loc = "3702"
                .area = "SPT2"
                .zone = 143
                .name = "Statue Room Invisible Chest"
                .logic = "G15"
            End With
            With aKeysDungeons(6)(14)
                .loc = "3704"
                .area = "SPT4"
                .zone = 145
                .name = "Leever Room Chest"
            End With
            With aKeysDungeons(6)(15)
                .loc = "3707"
                .area = "SPT4"
                .zone = 145
                .name = "Symphony Room Chest"
                .logic = "ZGCDrhLL7712LL7713LL7715LL7716LL7717"
            End With
            With aKeysDungeons(6)(16)
                .loc = "3727"
                .area = "SPT4"
                .zone = 145
                .name = "Entrance Front Right Chest"
                .logic = "Zr"
            End With
            With aKeysDungeons(6)(17)
                .loc = "3725"
                .area = "SPT2"
                .zone = 143
                .name = "Beamos Room Chest"
                .logic = "GCBx"
            End With
            With aKeysDungeons(6)(18)
                .loc = "3724"
                .area = "SPT2"
                .zone = 143
                .name = "Chest Switch Chest"
                .logic = "GCBhLL7716x.GCBhLL7716b"
            End With
            With aKeysDungeons(6)(19)
                .loc = "3705"
                .area = "SPT2"
                .zone = 143
                .name = "Boss Key Chest"
                .logic = "GCBhLL7716LL7422x.GCBhLL7716LL7422b"
            End With
            With aKeysDungeons(6)(20)
                .loc = "5509"
                .area = "SPT7"
                .zone = 148
                .name = "Mirror Shield Chest"
            End With
            With aKeysDungeons(6)(21)
                .loc = "3718"
                .area = "SPT5"
                .zone = 146
                .name = "Mirror Puzzle Invisible Chest"
                .logic = "G15"
            End With
            With aKeysDungeons(6)(22)
                .loc = "1631"
                .area = "SPT6"
                .zone = 147
                .name = "Twinrova"
                .logic = "ZGB3LL7422"
            End With
            With aKeysDungeons(6)(23)
                .loc = "7916"
                .area = "SPT3"
                .zone = 144
                .name = "Sun Block Room"
                .gs = True
                .logic = "Z"
            End With
            With aKeysDungeons(6)(24)
                .loc = "7917"
                .area = "SPT4"
                .zone = 145
                .name = "Leever Room"
                .gs = True
            End With
            With aKeysDungeons(6)(25)
                .loc = "7919"
                .area = "SPT4"
                .zone = 145
                .name = "Symphony Room"
                .gs = True
                .logic = "ZGCDrhLL7712LL7713LL7715LL7716LL7717"
            End With
            With aKeysDungeons(6)(26)
                .loc = "7918"
                .area = "SPT2"
                .zone = 143
                .name = "Nine Thrones Room West"
                .gs = True
                .logic = "GCC"
            End With
            With aKeysDungeons(6)(27)
                .loc = "7920"
                .area = "SPT2"
                .zone = 143
                .name = "Nine Thrones Room North"
                .gs = True
                .logic = "GCC"
            End With
            With aKeysDungeons(6)(28)
                .loc = "10930"
                .area = "EVENT"
                .name = "Door Young Side 1"
            End With
            With aKeysDungeons(6)(29)
                .loc = "10918"
                .area = "EVENT"
                .name = "Door Young Side 2"
            End With
            With aKeysDungeons(6)(30)
                .loc = "10903"
                .area = "EVENT"
                .name = "Door Adult Side"
            End With
            With aKeysDungeons(6)(31)
                .loc = "10921"
                .area = "EVENT"
                .name = "Door Statue Silver's Side"
            End With
            With aKeysDungeons(6)(32)
                .loc = "10927"
                .area = "EVENT"
                .name = "Door Statue Mirror's Side 1"
            End With
            With aKeysDungeons(6)(33)
                .loc = "10928"
                .area = "EVENT"
                .name = "Door Statue Mirror's Side 2"
            End With
            With aKeysDungeons(6)(34)
                .loc = "10901"
                .area = "EVENT"
                .name = "Door Statue Mirror's Side 3"
            End With
            tK = 35
        End If

        With aKeysDungeons(6)(tK)
            .loc = "10909"
            .area = "EVENT"
            .name = "Shortcut Part 1"
        End With
        incB(tK)
        With aKeysDungeons(6)(tK)
            .loc = "10912"
            .area = "EVENT"
            .name = "Shortcut Part 2"
        End With
        incB(tK)
        With aKeysDungeons(6)(tK)
            .loc = "10923"
            .area = "EVENT"
            .name = "Boss Platform Lowered"
        End With
        incB(tK)
        With aKeysDungeons(6)(tK)
            .loc = "10904"
            .area = "EVENT"
            .name = "Statue Face Crumbled"
        End With
        incB(tK)
        With aKeysDungeons(6)(tK)
            .loc = "10920"
            .area = "EVENT"
            .name = "Boss Door Unlocked"
        End With
    End Sub
    Private Sub makeKeysShadowTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(7) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(7)(CInt(IIf(isMQ, 34, 30)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(7).Length - 1
            aKeysDungeons(7)(i) = New keyCheck
            aKeysDungeons(7)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Shadow Temple
            ' 0: Shadow Temple Entryway
            ' 1: Shadow Temple Beginning
            ' 2: Shadow Temple First Beamos
            ' 3: Shadow Temple Huge Pit
            ' 4: Shadow Temple Invisible Spikes
            ' 5: Shadow Temple Wind Tunnel
            ' 6: Shadow Temple After Wind
            ' 7: Shadow Temple Boat
            ' 8: Shadow Temple Beyond Boat
            ' 9: Shadow Temple Boss Room

            With aKeysDungeons(7)(0)
                .loc = "3801"
                .area = "SHT1"
                .zone = 151
                .name = "Map Chest"
            End With
            With aKeysDungeons(7)(1)
                .loc = "3807"
                .area = "SHT1"
                .zone = 151
                .name = "Hover Boots Chest"
            End With
            With aKeysDungeons(7)(2)
                .loc = "3803"
                .area = "SHT2"
                .zone = 152
                .name = "Compass Chest"
            End With
            With aKeysDungeons(7)(3)
                .loc = "3802"
                .area = "SHT2"
                .zone = 152
                .name = "Early Silver Rupee Chest"
            End With
            With aKeysDungeons(7)(4)
                .loc = "3812"
                .area = "SHT3"
                .zone = 153
                .name = "Invisible Blades Visible Chest"
            End With
            With aKeysDungeons(7)(5)
                .loc = "3822"
                .area = "SHT3"
                .zone = 153
                .name = "Invisible Blades Invisible Chest"
            End With
            With aKeysDungeons(7)(6)
                .loc = "3805"
                .area = "SHT3"
                .zone = 153
                .name = "Falling Spikes Lower Chest"
            End With
            With aKeysDungeons(7)(7)
                .loc = "3804"
                .area = "SHT3"
                .zone = 153
                .name = "Falling Spikes Switch Chest"
                .logic = "GA7.V01"
            End With
            With aKeysDungeons(7)(8)
                .loc = "3806"
                .area = "SHT3"
                .zone = 153
                .name = "Falling Spikes Upper Chest"
                .logic = "GA7.V01"
            End With
            With aKeysDungeons(7)(9)
                .loc = "3809"
                .area = "SHT4"
                .zone = 154
                .name = "Invisible Spikes Chest"
            End With
            With aKeysDungeons(7)(10)
                .loc = "501"
                .area = "SHT4"
                .zone = 154
                .name = "Skull Pot Room Freestanding Key"
                .logic = "Zkx.ZkV01"
            End With
            With aKeysDungeons(7)(11)
                .loc = "3821"
                .area = "SHT5"
                .zone = 155
                .name = "Wind Hint Chest"
            End With
            With aKeysDungeons(7)(12)
                .loc = "3808"
                .area = "SHT6"
                .zone = 156
                .name = "After Wind Enemy Chest"
            End With
            With aKeysDungeons(7)(13)
                .loc = "3820"
                .area = "SHT6"
                .zone = 156
                .name = "After Wind Hidden Chest"
                .logic = "x"
            End With
            With aKeysDungeons(7)(14)
                .loc = "3810"
                .area = "SHT8"
                .zone = 158
                .name = "Spike Walls Left Chest"
                .logic = "f"
            End With
            With aKeysDungeons(7)(15)
                .loc = "3811"
                .area = "SHT8"
                .zone = 158
                .name = "Boss Key Chest"
                .logic = "f"
            End With
            With aKeysDungeons(7)(16)
                .loc = "3813"
                .area = "SHT8"
                .zone = 158
                .name = "Invisible Floormaster Chest"
            End With
            With aKeysDungeons(7)(17)
                .loc = "1731"
                .area = "SHT9"
                .zone = 159
                .name = "Bongo Bongo"
                .logic = "ZGB4d.ZGB4k"
            End With
            With aKeysDungeons(7)(18)
                .loc = "7927"
                .area = "SHT3"
                .zone = 153
                .name = "Like Like Room"
                .gs = True
            End With
            With aKeysDungeons(7)(19)
                .loc = "7925"
                .area = "SHT3"
                .zone = 153
                .name = "Falling Spikes Room"
                .gs = True
                .logic = "Zk"
            End With
            With aKeysDungeons(7)(20)
                .loc = "7924"
                .area = "SHT4"
                .zone = 154
                .name = "Single Giant Pot"
                .gs = True
                .logic = "Zk"
            End With
            With aKeysDungeons(7)(21)
                .loc = "7928"
                .area = "SHT7"
                .zone = 157
                .name = "Near Ship"
                .gs = True
                .logic = "Zl"
            End With
            With aKeysDungeons(7)(22)
                .loc = "7926"
                .area = "SHT8"
                .zone = 158
                .name = "Triple Giant Pot"
                .gs = True
            End With
            With aKeysDungeons(7)(23)
                .loc = "11022"
                .area = "EVENT"
                .name = "Door #1: Beamos to Huge Pit"
            End With
            With aKeysDungeons(7)(24)
                .loc = "11023"
                .area = "EVENT"
                .name = "Door #2: Huge Pit to Invisible Spikes"
            End With
            With aKeysDungeons(7)(25)
                .loc = "11024"
                .area = "EVENT"
                .name = "Door #3: Invisible Spikes to Wind Tunnel"
            End With
            With aKeysDungeons(7)(26)
                .loc = "11021"
                .area = "EVENT"
                .name = "Door #4: After Wind to Boat"
            End With
            With aKeysDungeons(7)(27)
                .loc = "11025"
                .area = "EVENT"
                .name = "Door #5: Beyond Boat to Boss Room"
            End With
            tK = 28
        Else
            ' The Master Quest keys for the Shadow Temple
            ' 0: MQ Shadow Temple Entryway
            ' 1: MQ Shadow Temple Beginning
            ' 2: MQ Shadow Temple Dead Hand Area
            ' 3: MQ Shadow Temple First Beamos
            ' 4: MQ Shadow Temple Upper Huge Pit
            ' 5: MQ Shadow Temple Invisible Blades
            ' 6: MQ Shadow Temple Lower Huge Pit
            ' 7: MQ Shadow Temple Falling Spikes
            ' 8: MQ Shadow Temple Invisible Spikes
            ' 9: MQ Shadow Temple Wind Tunnel 
            ' A: MQ Shadow Temple After Wind
            ' B: MQ Shadow Temple Boat
            ' C: MQ Shadow Temple Beyond Boat
            ' D: MQ Shadow Temple Invisible Maze
            ' E: MQ Shadow Temple Near Boss

            With aKeysDungeons(7)(0)
                .loc = "3801"
                .area = "SHT2"
                .zone = 161
                .name = "Compass Chest"
            End With
            With aKeysDungeons(7)(1)
                .loc = "3807"
                .area = "SHT2"
                .zone = 161
                .name = "Hover Boots Chest"
                .logic = "ZhLL7716d"
            End With
            With aKeysDungeons(7)(2)
                .loc = "3802"
                .area = "SHT3"
                .zone = 162
                .name = "Map Chest"
            End With
            With aKeysDungeons(7)(3)
                .loc = "3803"
                .area = "SHT3"
                .zone = 162
                .name = "Early Gibdos Chest"
            End With
            With aKeysDungeons(7)(4)
                .loc = "3814"
                .area = "SHT3"
                .zone = 162
                .name = "Near Ship Invisible Chest"
            End With
            With aKeysDungeons(7)(5)
                .loc = "3812"
                .area = "SHT5"
                .zone = 164
                .name = "Invisible Blades Visible Chest"
            End With
            With aKeysDungeons(7)(6)
                .loc = "3822"
                .area = "SHT5"
                .zone = 164
                .name = "Invisible Blades Invisible Chest"
            End With
            With aKeysDungeons(7)(7)
                .loc = "3815"
                .area = "SHT7"
                .zone = 166
                .name = "Beamos Silver Rupee Chest"
                .logic = "Zl"
            End With
            With aKeysDungeons(7)(8)
                .loc = "3805"
                .area = "SHT7"
                .zone = 166
                .name = "Falling Spikes Lower Chest"
            End With
            With aKeysDungeons(7)(9)
                .loc = "3806"
                .area = "SHT7"
                .zone = 166
                .name = "Falling Spikes Upper Chest"
                .logic = "GA7.V01"
            End With
            With aKeysDungeons(7)(10)
                .loc = "3804"
                .area = "SHT7"
                .zone = 166
                .name = "Falling Spikes Switch Chest"
                .logic = "GA7.V01"
            End With
            With aKeysDungeons(7)(11)
                .loc = "3809"
                .area = "SHT8"
                .zone = 167
                .name = "Invisible Spikes Chest"
            End With
            With aKeysDungeons(7)(12)
                .loc = "3816"
                .area = "SHT8"
                .zone = 167
                .name = "Stalfos Room Chest"
                .logic = "Zk"
            End With
            With aKeysDungeons(7)(13)
                .loc = "3821"
                .area = "SHT9"
                .zone = 168
                .name = "Wind Hint Chest"
            End With
            With aKeysDungeons(7)(14)
                .loc = "3808"
                .area = "SHTA"
                .zone = 169
                .name = "After Wind Enemy Chest"
            End With
            With aKeysDungeons(7)(15)
                .loc = "3820"
                .area = "SHTA"
                .zone = 169
                .name = "After Wind Hidden Chest"
                .logic = "x"
            End With
            With aKeysDungeons(7)(16)
                .loc = "3810"
                .area = "SHTD"
                .zone = 172
                .name = "Spike Walls Left Chest"
                .logic = "ZfGCE"
            End With
            With aKeysDungeons(7)(17)
                .loc = "3811"
                .area = "SHTD"
                .zone = 172
                .name = "Boss Key Chest"
                .logic = "ZfGCE"
            End With
            With aKeysDungeons(7)(18)
                .loc = "3813"
                .area = "SHTD"
                .zone = 172
                .name = "Bomb Flower Chest"
            End With
            With aKeysDungeons(7)(19)
                .loc = "506"
                .area = "SHTD"
                .zone = 172
                .name = "Skull Pots Room Freestanding Key"
            End With
            With aKeysDungeons(7)(20)
                .loc = "1731"
                .area = "SHTE"
                .zone = 173
                .name = "Bongo Bongo"
                .logic = "ZGB4d.ZGB4k"
            End With
            With aKeysDungeons(7)(21)
                .loc = "7925"
                .area = "SHT7"
                .zone = 166
                .name = "Falling Spikes Room"
                .gs = True
                .logic = "Zk"
            End With
            With aKeysDungeons(7)(22)
                .loc = "7924"
                .area = "SHT9"
                .zone = 168
                .name = "Wind Hint Room"
                .gs = True
            End With
            With aKeysDungeons(7)(23)
                .loc = "7927"
                .area = "SHTA"
                .zone = 169
                .name = "After Wind"
                .gs = True
                .logic = "x"
            End With
            With aKeysDungeons(7)(24)
                .loc = "7928"
                .area = "SHTC"
                .zone = 171
                .name = "After Ship"
                .gs = True
            End With
            With aKeysDungeons(7)(25)
                .loc = "7926"
                .area = "SHTE"
                .zone = 173
                .name = "Near Boss"
                .gs = True
                .logic = "ZG21.f"
            End With
            With aKeysDungeons(7)(26)
                .loc = "11025"
                .area = "EVENT"
                .name = "Door to Dead Hands"
            End With
            With aKeysDungeons(7)(27)
                .loc = "11022"
                .area = "EVENT"
                .name = "Door #1: Beamos to Huge Pit"
            End With
            With aKeysDungeons(7)(28)
                .loc = "11023"
                .area = "EVENT"
                .name = "Door #2: Huge Pit to Invisible Spikes"
            End With
            With aKeysDungeons(7)(29)
                .loc = "11024"
                .area = "EVENT"
                .name = "Door #3: Invisible Spikes to Wind Tunnel"
            End With
            With aKeysDungeons(7)(30)
                .loc = "11021"
                .area = "EVENT"
                .name = "Door #4: After Wind to Boat"
            End With
            With aKeysDungeons(7)(31)
                .loc = "11027"
                .area = "EVENT"
                .name = "Door to Spike Walls"
            End With
            tK = 32
        End If

        With aKeysDungeons(7)(tK)
            .loc = "11029"
            .area = "EVENT"
            .name = "Block at Boat"
        End With
        incB(tK)
        With aKeysDungeons(7)(tK)
            .loc = "11016"
            .area = "EVENT"
            .name = "Pillar Knocked Down"
        End With
        incB(tK)
        With aKeysDungeons(7)(tK)
            .loc = "11020"
            .area = "EVENT"
            .name = "Boss Door Unlocked"
        End With
    End Sub
    Private Sub makeKeysBottomOfTheWell(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(8) = isMQ

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(8)(CInt(IIf(isMQ, 9, 19)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(8).Length - 1
            aKeysDungeons(8)(i) = New keyCheck
            aKeysDungeons(8)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Bottom of the Well
            ' 1: Bottom of the Well Main Area
            ' 2: Bottom of the Well Main Area with Lens of Truth checks

            With aKeysDungeons(8)(0)
                .loc = "3908"
                .area = "BW2"
                .zone = 176
                .name = "Front Left Fake Wall Chest"
                .logic = "Y"
            End With
            With aKeysDungeons(8)(1)
                .loc = "3902"
                .area = "BW1"
                .zone = 175
                .name = "Front Centre Bombable Chest"
                .logic = "Yx"
            End With
            With aKeysDungeons(8)(2)
                .loc = "3905"
                .area = "BW2"
                .zone = 176
                .name = "Right Bottom Fake Wall Chest"
                .logic = "Y"
            End With
            With aKeysDungeons(8)(3)
                .loc = "3901"
                .area = "BW2"
                .zone = 176
                .name = "Compass Chest"
                .logic = "Y"
            End With
            With aKeysDungeons(8)(4)
                .loc = "3914"
                .area = "BW2"
                .zone = 176
                .name = "Centre Skulltula Chest"
                .logic = "Y"
            End With
            With aKeysDungeons(8)(5)
                .loc = "3904"
                .area = "BW2"
                .zone = 176
                .name = "Back Left Bombable Chest"
                .logic = "Yx"
            End With
            With aKeysDungeons(8)(6)
                .loc = "601"
                .area = "BW1"
                .zone = 175
                .name = "Coffin Freestanding Key"
                .logic = "Ya.Yf"
            End With
            With aKeysDungeons(8)(7)
                .loc = "3903"
                .area = "BW1"
                .zone = 175
                .name = "Lens of Truth Chest"
                .logic = "YhLL7712LL7416.YhLL7712GA8a"
            End With
            With aKeysDungeons(8)(8)
                .loc = "3920"
                .area = "BW2"
                .zone = 176
                .name = "Invisible Chest"
                .logic = "YhLL7712"
            End With
            With aKeysDungeons(8)(9)
                .loc = "3916"
                .area = "BW1"
                .zone = 175
                .name = "Underwater Front Chest"
                .logic = "YhLL7712"
            End With
            With aKeysDungeons(8)(10)
                .loc = "3909"
                .area = "BW1"
                .zone = 175
                .name = "Underwater Left Chest"
                .logic = "YhLL7712"
            End With
            With aKeysDungeons(8)(11)
                .loc = "3907"
                .area = "BW1"
                .zone = 175
                .name = "Map Chest"
                .logic = "Yx.YQ0176V01.YfV01"
            End With
            With aKeysDungeons(8)(12)
                .loc = "3910"
                .area = "BW2"
                .zone = 176
                .name = "Fire Keese Chest"
                .logic = "YGCF"
            End With
            With aKeysDungeons(8)(13)
                .loc = "3912"
                .area = "BW2"
                .zone = 176
                .name = "Like Like Chest"
                .logic = "YGCF"
            End With
            With aKeysDungeons(8)(14)
                .loc = "8000"
                .area = "BW2"
                .zone = 176
                .name = "Like Like Cage"
                .gs = True
                .logic = "YGCFo"
            End With
            With aKeysDungeons(8)(15)
                .loc = "8002"
                .area = "BW2"
                .zone = 176
                .name = "West Inner Room"
                .gs = True
                .logic = "YGD0o"
            End With
            With aKeysDungeons(8)(16)
                .loc = "8001"
                .area = "BW2"
                .zone = 176
                .name = "East Inner Room"
                .gs = True
                .logic = "YGD1o"
            End With
            With aKeysDungeons(8)(17)
                .loc = "11727"
                .area = "EVENT"
                .name = "Door Fire Keese and Like Like Cage"
            End With
            With aKeysDungeons(8)(18)
                .loc = "11728"
                .area = "EVENT"
                .name = "Door West GS"
            End With
            With aKeysDungeons(8)(19)
                .loc = "11729"
                .area = "EVENT"
                .name = "Door East GS"
            End With
        Else
            ' The Master Quest keys for the Bottom of the Well
            ' 0: Bottom of the Well and Perimeter
            ' 1: Bottom of the Well Middle

            With aKeysDungeons(8)(0)
                .loc = "3903"
                .area = "BW1"
                .zone = 177
                .name = "Map Chest"
                .logic = "Y"
            End With
            With aKeysDungeons(8)(1)
                .loc = "601"
                .area = "BW1"
                .zone = 177
                .name = "East Inner Room Freestanding Key"
                .logic = "Y"
            End With
            With aKeysDungeons(8)(2)
                .loc = "3902"
                .area = "BW0"
                .zone = 174
                .name = "Compass Chest"
                .logic = "YLL7416.YGA8a"
            End With
            With aKeysDungeons(8)(3)
                .loc = "602"
                .area = "BW0"
                .zone = 174
                .name = "Dead Hand Freestanding Key"
                .logic = "Yx"
            End With
            With aKeysDungeons(8)(4)
                .loc = "3901"
                .area = "BW1"
                .zone = 177
                .name = "Lens of Truth Chest"
                .logic = "YGD3x"
            End With
            With aKeysDungeons(8)(5)
                .loc = "8000"
                .area = "BW0"
                .zone = 174
                .name = "Basement"
                .gs = True
                .logic = "YJ"
            End With
            With aKeysDungeons(8)(6)
                .loc = "8001"
                .area = "BW1"
                .zone = 177
                .name = "West Inner Room"
                .gs = True
                .logic = "Yx"
            End With
            With aKeysDungeons(8)(7)
                .loc = "8002"
                .area = "BW0"
                .zone = 174
                .name = "Coffin Room"
                .gs = True
                .logic = "YJGD2"
            End With
            With aKeysDungeons(8)(8)
                .loc = "11720"
                .area = "EVENT"
                .name = "Door Coffin Room"
            End With
            With aKeysDungeons(8)(9)
                .loc = "11721"
                .area = "EVENT"
                .name = "Door First Switch"
            End With
        End If
    End Sub
    Private Sub makeKeysIceCavern(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(9) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys, nice that they are the same for both Master Quest and not
        ReDim aKeysDungeons(9)(7)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(9).Length - 1
            aKeysDungeons(9)(i) = New keyCheck
            aKeysDungeons(9)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Ice Cavern
            With aKeysDungeons(9)(0)
                .loc = "4000"
                .area = "IC"
                .zone = 201
                .name = "Map Chest"
                .logic = "Zu"
            End With
            With aKeysDungeons(9)(1)
                .loc = "4001"
                .area = "IC"
                .zone = 201
                .name = "Compass Chest"
                .logic = "G17"
            End With
            With aKeysDungeons(9)(2)
                .loc = "701"
                .area = "IC"
                .zone = 201
                .name = "Piece of Heart"
                .logic = "G17"
            End With
            With aKeysDungeons(9)(3)
                .loc = "4002"
                .area = "IC"
                .zone = 201
                .name = "Iron Boots Chest"
                .logic = "YLL7416G17.YaG17.YfG17.YgG17.Zu"
            End With
            With aKeysDungeons(9)(4)
                .loc = "6402"
                .area = "IC"
                .zone = 201
                .name = "Song from Sheik"
                .logic = "YLL7416G17.YaG17.YfG17.YgG17.Zu"
            End With
            With aKeysDungeons(9)(5)
                .loc = "8009"
                .area = "IC"
                .zone = 201
                .name = "Spinning Scythe Room"
                .gs = True
                .logic = "Yo.Zk"
            End With
            With aKeysDungeons(9)(6)
                .loc = "8010"
                .area = "IC"
                .zone = 201
                .name = "Heart Piece Room"
                .gs = True
                .logic = "YoG17.ZG17k"
            End With
            With aKeysDungeons(9)(7)
                .loc = "8008"
                .area = "IC"
                .zone = 201
                .name = "Push Block Room"
                .gs = True
                .logic = "YoG17.ZG17k"
            End With
        Else
            ' 178   IC Beginning
            ' 202   IC Map Room
            ' 203   IC Compass Room
            ' 204   IC Iron Boots Region

            ' The Master Quest keys for the Ice Carvern
            With aKeysDungeons(9)(0)
                .loc = "4001"
                .area = "IC"
                .zone = 202
                .name = "Map Chest"
                .logic = "YLL7416u.Yau.YG20u.Zu"
            End With
            With aKeysDungeons(9)(1)
                .loc = "4000"
                .area = "IC"
                .zone = 203
                .name = "Compass Chest"
            End With
            With aKeysDungeons(9)(2)
                .loc = "701"
                .area = "IC"
                .zone = 203
                .name = "Piece of Heart"
                .logic = "x"
            End With
            With aKeysDungeons(9)(3)
                .loc = "4002"
                .area = "IC"
                .zone = 204
                .name = "Iron Boots Chest"
                .logic = "Z"
            End With
            With aKeysDungeons(9)(4)
                .loc = "6402"
                .area = "IC"
                .zone = 204
                .name = "Song from Sheik"
                .logic = "Z"
            End With
            With aKeysDungeons(9)(5)
                .loc = "8009"
                .area = "IC"
                .zone = 203
                .name = "Red Ice"
                .gs = True
                .logic = "hLL7716"
            End With
            With aKeysDungeons(9)(6)
                .loc = "8010"
                .area = "IC"
                .zone = 204
                .name = "Ice Block"
                .gs = True
                .logic = "YJ.Z"
            End With
            With aKeysDungeons(9)(7)
                .loc = "8008"
                .area = "IC"
                .zone = 204
                .name = "Scarecrow"
                .gs = True
                .logic = "ZhLL6512k.ZLL7730l"
            End With
        End If
    End Sub
    Private Sub makeKeysGerudoTrainingGround(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(10) = isMQ

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(10)(CInt(IIf(isMQ, 19, 30)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(10).Length - 1
            aKeysDungeons(10)(i) = New keyCheck
            aKeysDungeons(10)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Gerudo Training Ground
            ' 0: Gerudo Training Ground Lobby and Central Maze
            ' 1: Gerudo Training Ground Central Maze Right
            ' 2: Gerudo Training Ground Lava Room
            ' 3: Gerudo Training Ground Hammer Room
            ' 4: Gerudo Training Ground Eye Statue Lower
            ' 5: Gerudo Training Ground Eye Statue Upper
            ' 6: Gerudo Training Ground Heavy Block Room
            ' 7: Gerudo Training Ground Like Like Room

            With aKeysDungeons(10)(0)
                .loc = "4219"
                .area = "GTG0"
                .zone = 179
                .name = "Lobby Left Chest"
                .logic = "Yg.Zd"
            End With
            With aKeysDungeons(10)(1)
                .loc = "4207"
                .area = "GTG0"
                .zone = 179
                .name = "Lobby Right Chest"
                .logic = "Yg.Zd"
            End With
            With aKeysDungeons(10)(2)
                .loc = "4200"
                .area = "GTG0"
                .zone = 179
                .name = "Stalfos Chest"
                .logic = "YLL7416.Z"
            End With
            With aKeysDungeons(10)(3)
                .loc = "4201"
                .area = "GTG0"
                .zone = 179
                .name = "Beamos Chest"
                .logic = "YLL7416x.Zx"
            End With
            With aKeysDungeons(10)(4)
                .loc = "4217"
                .area = "GTG6"
                .zone = 185
                .name = "Before Heavy Block Chest"
            End With
            With aKeysDungeons(10)(5)
                .loc = "4215"
                .area = "GTG7"
                .zone = 186
                .name = "Heavy Block First Chest"
            End With
            With aKeysDungeons(10)(6)
                .loc = "4214"
                .area = "GTG7"
                .zone = 186
                .name = "Heavy Block Second Chest"
            End With
            With aKeysDungeons(10)(7)
                .loc = "4220"
                .area = "GTG7"
                .zone = 186
                .name = "Heaby Block Third Chest"
            End With
            With aKeysDungeons(10)(8)
                .loc = "4202"
                .area = "GTG7"
                .zone = 186
                .name = "Heavy Block Fourth Chest"
            End With
            With aKeysDungeons(10)(9)
                .loc = "4203"
                .area = "GTG4"
                .zone = 183
                .name = "Eye Statue Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(10)(10)
                .loc = "4204"
                .area = "GTG5"
                .zone = 184
                .name = "Near Scarecrow Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(10)(11)
                .loc = "4218"
                .area = "GTG3"
                .zone = 182
                .name = "Hammer Room Clear Chest"
            End With
            With aKeysDungeons(10)(12)
                .loc = "4216"
                .area = "GTG3"
                .zone = 182
                .name = "Hammer Room Switch Chest"
                .logic = "Zr"
            End With
            With aKeysDungeons(10)(13)
                .loc = "4205"
                .area = "GTG1"
                .zone = 180
                .name = "Maze Right Central Chest"
            End With
            With aKeysDungeons(10)(14)
                .loc = "4208"
                .area = "GTG1"
                .zone = 180
                .name = "Maze Right Side Chest"
            End With
            With aKeysDungeons(10)(15)
                .loc = "801"
                .area = "GTG1"
                .zone = 180
                .name = "Lava Room Freestanding Key"
            End With
            With aKeysDungeons(10)(16)
                .loc = "4213"
                .area = "GTG2"
                .zone = 181
                .name = "Underwater Silver Rupee Chest"
                .logic = "ZkhLL7716LL7429G0A"
            End With
            With aKeysDungeons(10)(17)
                .loc = "4211"
                .area = "GTG0"
                .zone = 179
                .name = "Hidden Ceiling Chest"
                .logic = "GD4GAD"
            End With
            With aKeysDungeons(10)(18)
                .loc = "4206"
                .area = "GTG0"
                .zone = 179
                .name = "Maze Path First Chest"
                .logic = "GD5"
            End With
            With aKeysDungeons(10)(19)
                .loc = "4210"
                .area = "GTG0"
                .zone = 179
                .name = "Maze Path Second Chest"
                .logic = "GD6"
            End With
            With aKeysDungeons(10)(20)
                .loc = "4209"
                .area = "GTG0"
                .zone = 179
                .name = "Maze Path Third Chest"
                .logic = "GD7"
            End With
            With aKeysDungeons(10)(21)
                .loc = "4212"
                .area = "GTG0"
                .zone = 179
                .name = "Maze Path Final Chest"
                .logic = "GD8"
            End With
            With aKeysDungeons(10)(22)
                .loc = "11210"
                .area = "EVENT"
                .name = "Door Right 1"
            End With
            With aKeysDungeons(10)(23)
                .loc = "11203"
                .area = "EVENT"
                .name = "Door Right 2"
            End With
            With aKeysDungeons(10)(24)
                .loc = "11201"
                .area = "EVENT"
                .name = "Door Left 1"
            End With
            With aKeysDungeons(10)(25)
                .loc = "11209"
                .area = "EVENT"
                .name = "Door Left 2"
            End With
            With aKeysDungeons(10)(26)
                .loc = "11204"
                .area = "EVENT"
                .name = "Door Left 3"
            End With
            With aKeysDungeons(10)(27)
                .loc = "11205"
                .area = "EVENT"
                .name = "Door Left 4"
            End With
            With aKeysDungeons(10)(28)
                .loc = "11206"
                .area = "EVENT"
                .name = "Door Left 5"
            End With
            With aKeysDungeons(10)(29)
                .loc = "11207"
                .area = "EVENT"
                .name = "Door Left 6"
            End With
            With aKeysDungeons(10)(30)
                .loc = "11223"
                .area = "EVENT"
                .name = "Door Left 7"
            End With
        Else
            ' The Master Quest keys for the Gerudo Training Ground
            ' 0: MQ Gerudo Training Ground Lobby
            ' 1: MQ Gerudo Training Ground Right Side
            ' 2: MQ Gerudo Training Ground Underwater
            ' 3: MQ Gerudo Training Ground Left Side
            ' 4: MQ Gerudo Training Ground Stalfos Room
            ' 5: MQ Gerudo Training Ground Back Areas
            ' 6: MQ Gerudo Training Ground Central Maze Right

            With aKeysDungeons(10)(0)
                .loc = "4219"
                .area = "GTG0"
                .zone = 179
                .name = "Lobby Left Chest"
            End With
            With aKeysDungeons(10)(1)
                .loc = "4207"
                .area = "GTG0"
                .zone = 179
                .name = "Lobby Right Chest"
            End With
            With aKeysDungeons(10)(2)
                .loc = "4211"
                .area = "GTG0"
                .zone = 179
                .name = "Hidden Ceiling Chest"
                .logic = "GAD"
            End With
            With aKeysDungeons(10)(3)
                .loc = "4206"
                .area = "GTG0"
                .zone = 179
                .name = "Maze Path First Chest"
            End With
            With aKeysDungeons(10)(4)
                .loc = "4210"
                .area = "GTG0"
                .zone = 179
                .name = "Maze Path Second Chest"
            End With
            With aKeysDungeons(10)(5)
                .loc = "4209"
                .area = "GTG0"
                .zone = 179
                .name = "Maze Path Third Chest"
                .logic = "GD9"
            End With
            With aKeysDungeons(10)(6)
                .loc = "4201"
                .area = "GTG1"
                .zone = 187
                .name = "Dinolfos Chest"
                .logic = "Z"
            End With
            With aKeysDungeons(10)(7)
                .loc = "4213"
                .area = "GTG2"
                .zone = 188
                .name = "Underwater Silver Rupee Chest"
                .logic = "ZG24LL7429G0A"
            End With
            With aKeysDungeons(10)(8)
                .loc = "4200"
                .area = "GTG3"
                .zone = 189
                .name = "First Iron Knuckle Chest"
                .logic = "YLL7416.Yx.Z"
            End With
            With aKeysDungeons(10)(9)
                .loc = "4217"
                .area = "GTG4"
                .zone = 190
                .name = "Before Heavy Block Chest"
            End With
            With aKeysDungeons(10)(10)
                .loc = "4202"
                .area = "GTG4"
                .zone = 190
                .name = "Heavy Block Chest"
                .logic = "ZV02"
            End With
            With aKeysDungeons(10)(11)
                .loc = "4203"
                .area = "GTG5"
                .zone = 191
                .name = "Eye Statue Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(10)(12)
                .loc = "4218"
                .area = "GTG5"
                .zone = 191
                .name = "Second Iron Knuckle Chest"
            End With
            With aKeysDungeons(10)(13)
                .loc = "4214"
                .area = "GTG5"
                .zone = 191
                .name = "Flame Circle Chest"
                .logic = "Zk.Zd.x"
            End With
            With aKeysDungeons(10)(14)
                .loc = "4205"
                .area = "GTG6"
                .zone = 192
                .name = "Maze Right Central Chest"
            End With
            With aKeysDungeons(10)(15)
                .loc = "4208"
                .area = "GTG6"
                .zone = 192
                .name = "Maze Right Side Chest"
            End With
            With aKeysDungeons(10)(16)
                .loc = "4204"
                .area = "GTG6"
                .zone = 192
                .name = "Ice Arrows Chest"
                .logic = "GDA"
            End With
            With aKeysDungeons(10)(17)
                .loc = "11229"
                .area = "EVENT"
                .name = "Door 1"
            End With
            With aKeysDungeons(10)(18)
                .loc = "11220"
                .area = "EVENT"
                .name = "Door 2"
            End With
            With aKeysDungeons(10)(19)
                .loc = "11223"
                .area = "EVENT"
                .name = "Door 3"
            End With
        End If
    End Sub
    Private Sub makeKeysGanonsCastle(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(11) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Byte = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(11)(CInt(IIf(isMQ, 26, 28)))

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(11).Length - 1
            aKeysDungeons(11)(i) = New keyCheck
            aKeysDungeons(11)(i).scan = True
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for Ganon's Castle
            With aKeysDungeons(11)(0)
                .loc = "4309"
                .area = "IGC0"
                .zone = 193
                .name = "Forest Trial Chest"
            End With
            With aKeysDungeons(11)(1)
                .loc = "4307"
                .area = "IGC0"
                .zone = 193
                .name = "Water Trial Left Chest"
            End With
            With aKeysDungeons(11)(2)
                .loc = "4306"
                .area = "IGC0"
                .zone = 193
                .name = "Water Trial Right Chest"
            End With
            With aKeysDungeons(11)(3)
                .loc = "4308"
                .area = "IGC0"
                .zone = 193
                .name = "Shadow Trial Front Chest"
                .logic = "Zde.Zk.ZLL7430.hLL7716"
            End With
            With aKeysDungeons(11)(4)
                .loc = "4305"
                .area = "IGC0"
                .zone = 193
                .name = "Shadow Trial Golden Gauntlets Chest"
                .logic = "Zde.ZlLL7439.Zlf"
            End With
            With aKeysDungeons(11)(5)
                .loc = "4318"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial Crystal Switch Chest"
                .logic = "Zk"
            End With
            With aKeysDungeons(11)(6)
                .loc = "4320"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial Invisible Chest"
                .logic = "ZjkGAB"
            End With
            With aKeysDungeons(11)(7)
                .loc = "4312"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial First Left Chest"
            End With
            With aKeysDungeons(11)(8)
                .loc = "4311"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial Second Left Chest"
            End With
            With aKeysDungeons(11)(9)
                .loc = "4313"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial Third Left Chest"
            End With
            With aKeysDungeons(11)(10)
                .loc = "4314"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial First Right Chest"
            End With
            With aKeysDungeons(11)(11)
                .loc = "4310"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial Second Right Chest"
            End With
            With aKeysDungeons(11)(12)
                .loc = "4315"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial Third Right Chest"
            End With
            With aKeysDungeons(11)(13)
                .loc = "4316"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial Invisible Enemies Chest"
                .logic = "GAB"
            End With
            With aKeysDungeons(11)(14)
                .loc = "4317"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial Lullaby Chest"
                .logic = "GDBhLL7712"
            End With
            With aKeysDungeons(11)(15)
                .loc = "4111"
                .area = "IGC3"
                .zone = 196
                .name = "Ganon’s Tower Boss Key Chest"
            End With
            With aKeysDungeons(11)(16)
                .loc = "8709"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(17)
                .loc = "8706"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Centre Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(18)
                .loc = "8704"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Centre Right"
                .scrub = True
            End With
            With aKeysDungeons(11)(19)
                .loc = "8708"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Right"
                .scrub = True
            End With
            With aKeysDungeons(11)(20)
                .loc = "11330"
                .area = "EVENT"
                .name = "Door: Light Trial 1"
            End With
            tK = 21
        Else
            ' The Master Quest keys for Ganon's Castle
            With aKeysDungeons(11)(0)
                .loc = "4302"
                .area = "IGC0"
                .zone = 193
                .name = "Forest Trial Eye Switch Chest"
                .logic = "Zd"
            End With
            With aKeysDungeons(11)(1)
                .loc = "4303"
                .area = "IGC0"
                .zone = 193
                .name = "Forest Trial Frozen Eye Switch Chest"
                .logic = "G24"
            End With
            With aKeysDungeons(11)(2)
                .loc = "901"
                .area = "IGC0"
                .zone = 193
                .name = "Forest Trial Stalfos Room Freestanding Key"
                .logic = "Zk"
            End With
            With aKeysDungeons(11)(3)
                .loc = "4301"
                .area = "IGC0"
                .zone = 193
                .name = "Water Trial Chest"
                .logic = "u"
            End With
            With aKeysDungeons(11)(4)
                .loc = "4300"
                .area = "IGC0"
                .zone = 193
                .name = "Shadow Trial Bomb Flower Chest"
                .logic = "Zdk.ZdLL7430.ZLL7430GABx.ZLL7430GABV01.ZLL7430GABf"
            End With
            With aKeysDungeons(11)(5)
                .loc = "4305"
                .area = "IGC0"
                .zone = 193
                .name = "Shadow Trial Eye Switch Chest"
                .logic = "ZdGABLL7430.ZdGABkG24"
            End With
            With aKeysDungeons(11)(6)
                .loc = "4310"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial First Chest"
                .logic = "Zdr.ZGA9r"
            End With
            With aKeysDungeons(11)(7)
                .loc = "4320"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial Invisible Chest"
                .logic = "ZdrjGAB.ZGA9rjGAB"
            End With
            With aKeysDungeons(11)(8)
                .loc = "4309"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial Sun Front Left Chest"
                .logic = "ZdejrLL7422"
            End With
            With aKeysDungeons(11)(9)
                .loc = "4308"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial Sun Back Left Chest"
                .logic = "ZdejrLL7422"
            End With
            With aKeysDungeons(11)(10)
                .loc = "4307"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial Sun Back Right Chest"
                .logic = "ZdejrLL7422"
            End With
            With aKeysDungeons(11)(11)
                .loc = "4306"
                .area = "IGC0"
                .zone = 193
                .name = "Spirit Trial Golden Gauntlets Chest"
                .logic = "ZdejrLL7422"
            End With
            With aKeysDungeons(11)(12)
                .loc = "4304"
                .area = "IGC2"
                .zone = 195
                .name = "Light Trial Lullaby Chest"
                .logic = "hLL7712"
            End With
            With aKeysDungeons(11)(13)
                .loc = "4111"
                .area = "IGC3"
                .zone = 196
                .name = "Ganon’s Tower Boss Key Chest"
            End With
            With aKeysDungeons(11)(14)
                .loc = "8709"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(15)
                .loc = "8706"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Centre Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(16)
                .loc = "8704"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Centre Right"
                .scrub = True
            End With
            With aKeysDungeons(11)(17)
                .loc = "8708"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Right"
                .scrub = True
            End With
            With aKeysDungeons(11)(18)
                .loc = "8701"
                .area = "IGC1"
                .zone = 194
                .name = "Deku Room Front Right"
                .scrub = True
            End With
            tK = 19
        End If

        With aKeysDungeons(11)(tK)
            .loc = "11304"
            .area = "EVENT"
            .name = "Light Trial Unblocked"
        End With
        incB(tK)
        With aKeysDungeons(11)(tK)
            .loc = "6611"
            .area = "EVENT"
            .name = "Forest Trial Finished"
        End With
        incB(tK)
        With aKeysDungeons(11)(tK)
            .loc = "6614"
            .area = "EVENT"
            .name = "Fire Trial Finished"
        End With
        incB(tK)
        With aKeysDungeons(11)(tK)
            .loc = "6612"
            .area = "EVENT"
            .name = "Water Trial Finished"
        End With
        incB(tK)
        With aKeysDungeons(11)(tK)
            .loc = "6629"
            .area = "EVENT"
            .name = "Spirit Trial Finished"
        End With
        incB(tK)
        With aKeysDungeons(11)(tK)
            .loc = "6613"
            .area = "EVENT"
            .name = "Shadow Trial Finished"
        End With
        incB(tK)
        With aKeysDungeons(11)(tK)
            .loc = "6615"
            .area = "EVENT"
            .name = "Light Trial Finished"
        End With
        incB(tK)
        With aKeysDungeons(11)(tK)
            .loc = "6719"
            .area = "EVENT"
            .name = "Barrier Lowered"
        End With
    End Sub

    ' Handle all the clicks to output checks
    Private Sub lblKokiriForest_MouseClick(sender As Object, e As MouseEventArgs) Handles lblKokiriForest.MouseClick
        displayChecks("KF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlKokiriForest_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlKokiriForest.MouseClick
        displayChecks("KF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblLostWoods_MouseClick(sender As Object, e As MouseEventArgs) Handles lblLostWoods.MouseClick
        displayChecks("LW", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlLostWoods_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlLostWoods.MouseClick
        displayChecks("LW", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblSacredForestMeadow_MouseClick(sender As Object, e As MouseEventArgs) Handles lblSacredForestMeadow.MouseClick
        displayChecks("SFM", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlSacredForestMeadow_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlSacredForestMeadow.MouseClick
        displayChecks("SFM", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblHyruleField_MouseClick(sender As Object, e As MouseEventArgs) Handles lblHyruleField.MouseClick
        displayChecks("HF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlHyruleField_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlHyruleField.MouseClick
        displayChecks("HF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblLonLonRanch_MouseClick(sender As Object, e As MouseEventArgs) Handles lblLonLonRanch.MouseClick
        displayChecks("LLR", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlLonLonRanch_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlLonLonRanch.MouseClick
        displayChecks("LLR", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblMarket_MouseClick(sender As Object, e As MouseEventArgs) Handles lblMarket.MouseClick
        displayChecks("MK", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlMarket_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlMarket.MouseClick
        displayChecks("MK", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblTempleOfTime_MouseClick(sender As Object, e As MouseEventArgs) Handles lblTempleOfTime.MouseClick
        displayChecks("TT", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlTempleOfTime_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlTempleOfTime.MouseClick
        displayChecks("TT", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblHyruleCastle_MouseClick(sender As Object, e As MouseEventArgs) Handles lblHyruleCastle.MouseClick
        displayChecks("HC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlHyruleCastle_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlHyruleCastle.MouseClick
        displayChecks("HC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblKakarikoVillage_MouseClick(sender As Object, e As MouseEventArgs) Handles lblKakarikoVillage.MouseClick
        displayChecks("KV", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlKakarikoVillage_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlKakarikoVillage.MouseClick
        displayChecks("KV", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblGraveyard_MouseClick(sender As Object, e As MouseEventArgs) Handles lblGraveyard.MouseClick
        displayChecks("GY", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlGraveyard_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlGraveyard.MouseClick
        displayChecks("GY", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblDMTrail_MouseClick(sender As Object, e As MouseEventArgs) Handles lblDMTrail.MouseClick
        displayChecks("DMT", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlDMTrail_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlDMTrail.MouseClick
        displayChecks("DMT", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblDMCrater_MouseClick(sender As Object, e As MouseEventArgs) Handles lblDMCrater.MouseClick
        displayChecks("DMC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlDMCrater_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlDMCrater.MouseClick
        displayChecks("DMC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblGoronCity_MouseClick(sender As Object, e As MouseEventArgs) Handles lblGoronCity.MouseClick
        displayChecks("GC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlGoronCity_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlGoronCity.MouseClick
        displayChecks("GC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblZorasRiver_MouseClick(sender As Object, e As MouseEventArgs) Handles lblZorasRiver.MouseClick
        displayChecks("ZR", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlZorasRiver_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlZorasRiver.MouseClick
        displayChecks("ZR", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblZorasDomain_MouseClick(sender As Object, e As MouseEventArgs) Handles lblZorasDomain.MouseClick
        displayChecks("ZD", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlZorasDomain_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlZorasDomain.MouseClick
        displayChecks("ZD", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblZorasFountain_MouseClick(sender As Object, e As MouseEventArgs) Handles lblZorasFountain.MouseClick
        displayChecks("ZF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlZorasFountain_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlZorasFountain.MouseClick
        displayChecks("ZF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblLakeHylia_MouseClick(sender As Object, e As MouseEventArgs) Handles lblLakeHylia.MouseClick
        displayChecks("LH", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlLakeHylia_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlLakeHylia.MouseClick
        displayChecks("LH", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblGerudoValley_MouseClick(sender As Object, e As MouseEventArgs) Handles lblGerudoValley.MouseClick
        displayChecks("GV", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlGerudoValley_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlGerudoValley.MouseClick
        displayChecks("GV", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblGerudoFortress_MouseClick(sender As Object, e As MouseEventArgs) Handles lblGerudoFortress.MouseClick
        displayChecks("GF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlGerudoFortress_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlGerudoFortress.MouseClick
        displayChecks("GF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblHauntedWasteland_MouseClick(sender As Object, e As MouseEventArgs) Handles lblHauntedWasteland.MouseClick
        displayChecks("HW", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlHauntedWasteland_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlHauntedWasteland.MouseClick
        displayChecks("HW", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblDesertColossus_MouseClick(sender As Object, e As MouseEventArgs) Handles lblDesertColossus.MouseClick
        displayChecks("DC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlDesertColossus_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlDesertColossus.MouseClick
        displayChecks("DC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblOutsideGanonsCastle_MouseClick(sender As Object, e As MouseEventArgs) Handles lblOutsideGanonsCastle.MouseClick
        displayChecks("OGC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlOutsideGanonsCastle_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlOutsideGanonsCastle.MouseClick
        displayChecks("OGC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblDekuTree_MouseClick(sender As Object, e As MouseEventArgs) Handles lblDekuTree.MouseClick
        displayChecksDungeons(0, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlDekuTree_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlDekuTree.MouseClick
        displayChecksDungeons(0, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblDodongosCavern_MouseClick(sender As Object, e As MouseEventArgs) Handles lblDodongosCavern.MouseClick
        displayChecksDungeons(1, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlDodongosCavern_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlDodongosCavern.MouseClick
        displayChecksDungeons(1, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblJabuJabusBelly_MouseClick(sender As Object, e As MouseEventArgs) Handles lblJabuJabusBelly.MouseClick
        displayChecksDungeons(2, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlJabuJabusBelly_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlJabuJabusBelly.MouseClick
        displayChecksDungeons(2, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblForestTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles lblForestTemple.MouseClick
        displayChecksDungeons(3, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlForestTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlForestTemple.MouseClick
        displayChecksDungeons(3, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblFireTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles lblFireTemple.MouseClick
        displayChecksDungeons(4, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlFireTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlFireTemple.MouseClick
        displayChecksDungeons(4, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblWaterTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles lblWaterTemple.MouseClick
        displayChecksDungeons(5, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlWaterTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlWaterTemple.MouseClick
        displayChecksDungeons(5, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblSpiritTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles lblSpiritTemple.MouseClick
        displayChecksDungeons(6, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlSpiritTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlSpiritTemple.MouseClick
        displayChecksDungeons(6, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblShadowTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles lblShadowTemple.MouseClick
        displayChecksDungeons(7, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlShadowTemple_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlShadowTemple.MouseClick
        displayChecksDungeons(7, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblBottomOfTheWell_MouseClick(sender As Object, e As MouseEventArgs) Handles lblBottomOfTheWell.MouseClick
        displayChecksDungeons(8, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlBottomOfTheWell_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlBottomOfTheWell.MouseClick
        displayChecksDungeons(8, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblIceCavern_MouseClick(sender As Object, e As MouseEventArgs) Handles lblIceCavern.MouseClick
        displayChecksDungeons(9, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlIceCavern_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlIceCavern.MouseClick
        displayChecksDungeons(9, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblGerudoTrainingGround_MouseClick(sender As Object, e As MouseEventArgs) Handles lblGerudoTrainingGround.MouseClick
        displayChecksDungeons(10, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlGerudoTrainingGround_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlGerudoTrainingGround.MouseClick
        displayChecksDungeons(10, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblGanonsCastle_MouseClick(sender As Object, e As MouseEventArgs) Handles lblGanonsCastle.MouseClick
        displayChecksDungeons(11, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlGanonsCastle_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlGanonsCastle.MouseClick
        displayChecksDungeons(11, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblQuestGoldSkulltulas_MouseClick(sender As Object, e As MouseEventArgs) Handles lblQuestGoldSkulltulas.MouseClick
        displayChecks("QGS", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlQuestGoldSkulltulas_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlQuestGoldSkulltulas.MouseClick
        displayChecks("QGS", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblQuestMasks_MouseClick(sender As Object, e As MouseEventArgs) Handles lblQuestMasks.MouseClick
        displayChecks("QM", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlQuestMasks_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlQuestMasks.MouseClick
        displayChecks("QM", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub

    Private Sub dec(ByRef value As Integer, Optional ByVal amount As Byte = 1)
        ' Small sub for shorthand decrements
        value = value - amount
    End Sub
    Private Sub decB(ByRef value As Byte, Optional ByVal amount As Byte = 1)
        ' Small sub for shorthand decrements
        If value > 0 Then value = value - amount
    End Sub
    Private Sub inc(ByRef value As Integer, Optional ByVal amount As Integer = 1)
        ' Small sub for shorthand increments
        value = value + amount
    End Sub
    Private Sub incB(ByRef value As Byte, Optional ByVal amount As Byte = 1)
        ' Small sub for shorthand increments
        value = value + amount
    End Sub
    Private Sub fixHex(ByRef hex As String, Optional digits As Byte = 8)
        ' Small sub for fixing hex digits
        While hex.Length < digits
            hex = "0" & hex
        End While
    End Sub

    Private Sub checkCarpenters()
        ' Checks that all 4 carpenters have been saved
        Dim allFound As Boolean = False

        ' Check for the 4 carpenter 
        If checkLoc("6500") And checkLoc("6501") And checkLoc("6502") And checkLoc("6503") Then allFound = True

        ' Find the key for the Membership Card and make it the allFound
        For Each key In aKeys
            With key
                Select Case .loc
                    Case "CARD"
                        .checked = allFound
                End Select
            End With
        Next
    End Sub
		Private Sub checkEquipment()
        ' Pair the pictureboxes up with the checkboxes to make them visible or not
        pbxKokiriSword.Visible = aEquipment(16)
        pbxMasterSword.Visible = aEquipment(17)
        pbxDekuShield.Visible = aEquipment(20)
        pbxHylianShield.Visible = aEquipment(21)
        pbxMirrorShield.Visible = aEquipment(22)
        pbxKokiriTunic.Visible = aEquipment(24)
        pbxGoronTunic.Visible = aEquipment(25)
        pbxZoraTunic.Visible = aEquipment(26)
        pbxKokiriBoots.Visible = aEquipment(28)
        pbxIronBoots.Visible = aEquipment(29)
        pbxHoverBoots.Visible = aEquipment(30)

        ' Run the check to see if you have the Biggoron's Sword or the Knife
        biggoronSwordCheck()
    End Sub
    Private Sub biggoronSwordCheck()
        '0 = none, 1 = Biggoron's Sword, 2 = Broken Knife
        Dim who As Byte = 0
        If aEquipment(18) Then
            who = 1
            If knifeCheck Then
                who = 1
            ElseIf aEquipment(19) Then
                who = 2
            End If
        Else
            If aEquipment(19) Then
                who = 2
            End If
        End If
        Select Case who
            Case 1
                pbxBiggoronsSword.Visible = True
                pbxBrokenKnife.Visible = False
            Case 2
                pbxBiggoronsSword.Visible = False
                pbxBrokenKnife.Visible = True
            Case Else
                pbxBiggoronsSword.Visible = False
                pbxBrokenKnife.Visible = False
        End Select
    End Sub
    Private Sub checkUpgrades()
        Dim upgrades() As Integer = {0, 0, 0, 0, 0, 0, 0, 0}

        ' Quivers
        If aUpgrades(0) Then upgrades(0) = upgrades(0) + 1
        If aUpgrades(1) Then upgrades(0) = upgrades(0) + 2

        ' Bomb Bags
        If aUpgrades(3) Then upgrades(1) = upgrades(1) + 1
        If aUpgrades(4) Then upgrades(1) = upgrades(1) + 2

        ' Gauntlets
        If aUpgrades(6) Then upgrades(2) = upgrades(2) + 1
        If aUpgrades(7) Then upgrades(2) = upgrades(2) + 2

        ' Scales
        If aUpgrades(9) Then upgrades(3) = upgrades(3) + 1
        If aUpgrades(10) Then upgrades(3) = upgrades(3) + 2

        ' Wallets
        If aUpgrades(12) Then upgrades(4) = upgrades(4) + 1
        If aUpgrades(13) Then upgrades(4) = upgrades(4) + 2

        ' Bullet Bag
        If aUpgrades(14) Then upgrades(5) = upgrades(5) + 1
        If aUpgrades(15) Then upgrades(5) = upgrades(5) + 2

        ' Sticks
        If aUpgrades(17) Then upgrades(6) = upgrades(6) + 1
        If aUpgrades(18) Then upgrades(6) = upgrades(6) + 2

        ' Nuts
        If aUpgrades(20) Then upgrades(7) = upgrades(7) + 1
        If aUpgrades(21) Then upgrades(7) = upgrades(7) + 2

        With pbxQuiver
            Select Case upgrades(0)
                Case 1
                    .Image = My.Resources.upgradeQuiver1
                    If Not .Visible Then .Visible = True
                Case 2
                    .Image = My.Resources.upgradeQuiver2
                    If Not .Visible Then .Visible = True
                Case 3
                    .Image = My.Resources.upgradeQuiver3
                    If Not .Visible Then .Visible = True
                Case Else
                    If .Visible Then .Visible = False
            End Select
        End With

        With pbxBombBag
            Select Case upgrades(1)
                Case 1
                    .Image = My.Resources.upgradeBombBag1
                    If Not .Visible Then .Visible = True
                Case 2
                    .Image = My.Resources.upgradeBombBag2
                    If Not .Visible Then .Visible = True
                Case 3
                    .Image = My.Resources.upgradeBombBag3
                    If Not .Visible Then .Visible = True
                Case Else
                    If .Visible Then .Visible = False
            End Select
        End With

        With pbxGauntlet
            Select Case upgrades(2)
                Case 1
                    .Image = My.Resources.upgradeGoronsBracelet
                    If Not .Visible Then .Visible = True
                Case 2
                    .Image = My.Resources.upgradeSilverGauntlets
                    If Not .Visible Then .Visible = True
                Case 3
                    .Image = My.Resources.upgradeGoldenGauntlets
                    If Not .Visible Then .Visible = True
                Case Else
                    If .Visible Then .Visible = False
            End Select
        End With

        With pbxScale
            Select Case upgrades(3)
                Case 1
                    .Image = My.Resources.upgradeSilverScale
                    If Not .Visible Then .Visible = True
                Case 2
                    .Image = My.Resources.upgradeGoldenScale
                    If Not .Visible Then .Visible = True
                Case Else
                    If .Visible Then .Visible = False
            End Select
        End With

        With pbxWallet
            Select Case upgrades(4)
                Case 1
                    .Image = My.Resources.upgradeWallet1
                    If Not .Visible Then .Visible = True
                Case 2
                    .Image = My.Resources.upgradeWallet2
                    If Not .Visible Then .Visible = True
                Case 3
                    .Image = My.Resources.upgradeWallet3
                    If Not .Visible Then .Visible = True
                Case Else
                    If .Visible Then .Visible = False
            End Select
        End With

        With pbxBulletBag
            Select Case upgrades(5)
                Case 1
                    .Image = My.Resources.upgradeBulletBag1
                    If Not .Visible Then .Visible = True
                Case 2
                    .Image = My.Resources.upgradeBulletBag2
                    If Not .Visible Then .Visible = True
                Case 3
                    .Image = My.Resources.upgradeBulletBag3
                    If Not .Visible Then .Visible = True
                Case Else
                    If .Visible Then .Visible = False
            End Select
        End With
    End Sub
    Private Sub updateQuestItems()
        For i As Byte = 0 To 22
            With aoQuestItems(i)
                If aQuestItems(i) Then
                    .Image = aoQuestItemImages(i)
                Else
                    .Image = aoQuestItemImagesEmpty(i)
                End If
            End With

            ' Store for use when drawing text over each quest reward
            aQuestRewardsCollected(i) = aQuestItems(i)

            ' Update the added dungeon text on each medallion and stone
            Select Case i
                Case 0 To 11, 18 To 20
                    changeAddedText(i, , True)
            End Select
        Next
        For i As Byte = 0 To 7
            Select Case i
                Case 0, 1
                    If bSpawnWarps Then displaySpawns(i)
                Case Else
                    If bSongWarps Then displayWarps(i)
            End Select
        Next
    End Sub
    Private Sub attachToSoH()
        emulator = String.Empty
        If Not IS_64BIT Then Exit Sub
        ' Look for soh.exe
        Dim target As Process = Nothing
        Try
            target = Process.GetProcessesByName("soh")(0)
        Catch ex As Exception
            Return
        End Try
        ' If found, grab base address
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "soh.exe" Then
                romAddrStart64 = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next
        ' Finally, attach to soh.exe
        SetProcessName("soh")
        emulator = "soh"

        isSoH = True

        ' Check if we are changing from non-SoH to SoH, or just reconnecting to SoH again
        'For Each key In aKeys
        'Select Case key.loc

        'Case "016"      ' Redirect DMC Great Fairy
        'key.loc = "12202"
        'Case "024"      ' Redirect DMT Great Fairy
        'key.loc = "12201"
        'Case "2208"     ' Redirect Dampe's Gravedigging
        'key.loc = "2231"
        'Case "5200"     ' Redirect Shoot the Sun
        'key.loc = "5231"
        'Case "6306"     ' Redirect Darunia's Joy
        'key.loc = "5630"
        'Case "6720"     ' Light Arrows Cutscene Reward
        'key.loc = "12330"
        'Case "6407"     ' Redirect Song from Saria
        'key.loc = "11931"
        'Case "6408"     ' Redirect Song from Malon
        'key.loc = "11831"
        'Case "6410"     ' Redirect Sun's Song (you need to check the message twice!!!)
        'key.loc = "4931"
        'Case "6625"     ' Redirect Song From Ocarina of Time
        'key.loc = "12431"
        'Case "6808"     ' Redirect GF Zora
        'key.loc = "12101"
        'Case "6809"     ' Redirect GF Castle (Young)
        'key.loc = "12102"
        'Case "6810"     ' Redirect GF Desert
        'key.loc = "12103"
        'Case "6814"     ' Redirect Deku Theatre Skull Mask
        'key.loc = "4631"
        'Case "6830"     ' Shooting Gallery Adult (Child worked, todo: test child again)
        'key.loc = "12031"



        '   Case "6828"     ' Redirect Anju's Chickens
        '   key.loc = "7728"
        'End Select
        'Next

        If wasSoH = False Then
            soh.sohSetup(romAddrStart64)
            redirectChecks(False)
        End If
        wasSoH = True
    End Sub

    Private Sub redirectChecks(Optional ByVal regularRando As Boolean = True)
        ' Handles what the locs will be when making and checking keys
        If regularRando Then
            locSwap(0) = "016"      ' DMC Great Fairy
            locSwap(1) = "024"      ' DMT Great Fairy
            locSwap(2) = "2208"     ' Dampe's Gravedigging
            locSwap(3) = "5200"     ' Shoot the Sun
            locSwap(4) = "6306"     ' Darunia's Joy
            locSwap(5) = "6720"     ' Light Arrows Cutscene Reward
            locSwap(6) = "6407"     ' Song from Saria
            locSwap(7) = "6408"     ' Song from Malon
            locSwap(8) = "6410"     ' Sun's Song
            locSwap(9) = "6625"     ' Song From Ocarina of Time
            locSwap(10) = "6808"    ' GF Zora
            locSwap(11) = "6809"    ' GF Castle
            locSwap(12) = "6810"    ' GF Desert
            locSwap(13) = "6814"    ' Deku Theatre Skull Mask
            locSwap(14) = "6830"    ' Shooting Gallery Adult
        Else
            locSwap(0) = "12202"    ' DMC Great Fairy
            locSwap(1) = "12201"    ' DMT Great Fairy
            locSwap(2) = "2231"     ' Dampe's Gravedigging
            locSwap(3) = "5231"     ' Shoot the Sun
            locSwap(4) = "5630"     ' Darunia's Joy
            locSwap(5) = "12330"    ' Light Arrows Cutscene Reward
            locSwap(6) = "11931"    ' Song from Saria
            locSwap(7) = "11831"    ' Song from Malon
            locSwap(8) = "4931"     ' Sun's Song
            locSwap(9) = "12431"    ' Song From Ocarina of Time
            locSwap(10) = "12101"   ' GF Zora
            locSwap(11) = "12102"   ' GF Castle
            locSwap(12) = "12103"   ' GF Desert
            locSwap(13) = "4631"    ' Deku Theatre Skull Mask
            locSwap(14) = "12031"   ' Shooting Gallery Adult
        End If

        ' Clear all the keys
        For i = 0 To aKeys.Length - 1
            aKeys(i) = New keyCheck
            aKeys(i).scan = True
        Next

        setupKeys()
        ' Will need to fix the shoppes afterwards
        If regularRando Then updateShoppes()
        getHighLows()
    End Sub

    Private Sub getSoHRandoSettings()
        Dim offset As Integer = SAV(&H13E8)
        Dim varEnum As Byte = 0
        Dim varVal As Byte = 0

        Dim firstEnum As Byte = CByte(goRead(offset, 1))
        If lastFirstEnum = firstEnum Then Exit Sub
        lastFirstEnum = firstEnum

        ReDim aRandoSet(50)
        For i = 0 To aRandoSet.Length - 1
            aRandoSet(i) = 0
        Next

        For i = 0 To aRandoSet.Length - 1
            varEnum = CByte(goRead(offset + (i * 8), 1))
            If varEnum = 0 Then Exit For
            varVal = CByte(goRead(offset + (i * 8) + 4, 1))
            aRandoSet(varEnum) = varVal
        Next
    End Sub

    Private Sub attachToBizHawk()
        emulator = String.Empty
        If Not IS_64BIT Then Exit Sub
        Dim target As Process = Nothing
        Try
            target = Process.GetProcessesByName("emuhawk")(0)
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                Return
            End If
        End Try
        Dim addressDLL As Int64 = 0
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next
        If addressDLL = 0 Then Exit Sub
        romAddrStart64 = addressDLL + &H658E0
        SetProcessName("emuhawk")
        emulator = "emuhawk"
    End Sub
    Private Sub attachToRMG()
        ' This should already be empty in order to reach this point, but never hurts to make sure
        emulator = String.Empty
        ' If not 64bit, do not even bother
        If IS_64BIT = False Then Exit Sub
        ' Prepare new target process
        Dim target As Process = Nothing

        Try
            ' Try to attach to application
            target = Process.GetProcessesByName("rmg")(0)
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                ' This is the expected error if process was not found, just return
                Return
            Else
                ' Any other error, output error message to textbox
                rtbOutputLeft.Text = "Attachment Problem: " & ex.Message & vbCrLf
                Return
            End If
        End Try

        ' Prepare new address variable
        Dim addressDLL As Int64 = 0

        ' Step through all modules to find mupen64plus.dll's base address
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        ' If 0, aka did not find mupen64plus.dll, then exit
        If addressDLL = 0 Then Exit Sub

        ' Add location of variable to base address
        addressDLL = addressDLL + &H29C15D8

        ' Attach to process and set it as the current emulator
        SetProcessName("rmg")
        emulator = "rmg"

        ' Read the first half of the address
        Dim readR15 As Integer = ReadMemory(Of Integer)(addressDLL)
        ' Convert to hex
        Dim hexR15 As String = Hex(readR15)

        ' Make sure length is 8 digit for any dropped 0's
        While hexR15.Length < 8
            hexR15 = "0" & hexR15
        End While

        ' Read the second half of the address
        readR15 = ReadMemory(Of Integer)(addressDLL + 4)
        ' Convert to hex and attach to first half
        hexR15 = Hex(readR15) & hexR15

        ' Set it + 0x8000000 as starting address, done as 4's since 8 invokes negative value
        romAddrStart64 = CLng("&H" & hexR15) + &H40000000 + &H40000000
    End Sub
    Private Sub attachToM64P(Optional attempt As Byte = 0)
        ' This should already be empty in order to reach this point, but never hurts to make sure
        emulator = String.Empty
        ' If not 64bit, do not even bother
        If IS_64BIT = False Then Exit Sub
        ' Prepare new target process
        Dim target As Process = Nothing

        Try
            ' Try to attach to application
            target = Process.GetProcessesByName("mupen64plus-gui")(0)
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                ' This is the expected error if process was not found, just return
                Return
            Else
                ' Any other error, output error message to textbox
                rtbOutputLeft.Text = "Attachment Problem: " & ex.Message & vbCrLf
                Return
            End If
        End Try

        ' Prepare new address variable
        Dim addressDLL As Int64 = 0
        Dim attemptOffset As Int64 = 0
        Dim attemptAdded As Int64 = 0

        ' Step through all modules to find mupen64plus.dll's base address
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        Select Case attempt
            Case 0
                ' Builds October 27, 2021 to March 31, 2022
                attemptOffset = &H29C95D8
                attemptAdded = 2147483648
            Case 1
                ' Builds July 13, 2021 to October 11, 2021
                attemptOffset = &HCA6B8
            Case 2
                ' Builds October 27, 2021 to June, 2, 2022

                ' I hate this sooo much. Rather than the confusing "find the address that has the address", it is:
                ' "Find the address that, that has an address (with an increase of 0x140), that has the address we need"
                attemptOffset = &H1177888
                attemptAdded = 0
                'attemptOffset = attemptOffset + &H140
                Dim readRCX As Integer = ReadMemory(Of Integer)(addressDLL + attemptOffset)

                If Not readRCX = 0 Then
                    inc(readRCX, &H140)
                    Dim hexRCX As String = Hex(readRCX)
                    fixHex(hexRCX)
                    readRCX = ReadMemory(Of Integer)(addressDLL + attemptOffset + 4)
                    hexRCX = Hex(readRCX) & hexRCX
                    attemptOffset = CLng("&H" & hexRCX) - addressDLL
                End If
            Case Else
                Return
        End Select

        ' Check if mupen64plus.dll was found
        Dim nextAttempt As Boolean = True
        If Not addressDLL = 0 Then
            ' Add location of variable to base address
            addressDLL = addressDLL + attemptOffset
            'rtbOutput.AppendText(Hex(addressDLL) & vbCrLf)
            ' Attach to process and set it as the current emulator
            SetProcessName("mupen64plus-gui")
            emulator = "mupen64plus-gui"
            ' Read the first half of the address
            Dim readR15 As Integer = ReadMemory(Of Integer)(addressDLL)
            ' Convert to hex
            Dim hexR15 As String = Hex(readR15)
            If Not hexR15 = "0" Then
                ' Make sure length is 8 digit for any dropped 0's
                fixHex(hexR15)
                ' Read the second half of the address
                readR15 = ReadMemory(Of Integer)(addressDLL + 4)
                ' Convert to hex and attach to first half
                hexR15 = Hex(readR15) & hexR15
                romAddrStart64 = CLng("&H" & hexR15) + attemptAdded
                If ReadMemory(Of Integer)(romAddrStart64 + &H11A5EC) = 1514490948 Then nextAttempt = False
            End If
        End If
        If nextAttempt Then attachToM64P(CByte(attempt + 1))
    End Sub
    Private Sub attachToModLoader64()
        ' This sub, and thus support for ML64, is credited to subenji, who also added the memory function to attach by process ID -- 2022.06.17

        ' This should already be empty in order to reach this point, but never hurts to make sure
        emulator = String.Empty
        ' If not 64bit, do not even bother
        If IS_64BIT = False Then Exit Sub
        ' Prepare new target process
        Dim target As Process = Nothing

        Try
            ' Try to attach to application
            ' target = Process.GetProcessesByName("modloader64-gui")(0)
            ''
            ' ModLoader64 runs 5 or 6 processes with only one of them actually being the emulator window, so we need to check the emulation is running
            ''
            Dim processes As Process() = Process.GetProcessesByName("modloader64-gui")
            If processes.Length = 0 Then
                ' If process was not found, just return
                Return
            End If
            For Each p As Process In processes
                For Each pModule As ProcessModule In p.Modules
                    If LCase(pModule.ModuleName) = "mupen64plus.dll" Then
                        ''
                        ' As the process is run several times and I need to specify by PID rather than name, I need to attach this once by hand
                        ''
                        If Not OpenProcessHandleById(p.Id) Then
                            rtbOutputLeft.Text = "Attachment Problem: Could not open process handle: " & p.Id & vbCrLf
                            Return
                        End If

                        target = p
                        ' Prepare new address variable
                        Dim addressDLL As Int64 = 0
                        Dim attemptOffset As Int64 = 0
                        Dim attemptAdded As Int64 = 0
                        Dim positiveHit As Boolean = False
                        ''
                        ' These pointers to the Base ROM location were all listed as static, I felt it best just to add them all. The first one worked every time in testing.
                        ' I found adding 0x80000000 was never necessary.
                        ''
                        For attempt = 0 To 5
                            Select Case attempt
                                Case 0
                                    attemptOffset = &H116ECF8
                                Case 1
                                    attemptOffset = &H12EED10
                                Case 2
                                    attemptOffset = &H6A6F0
                                Case 3
                                    attemptOffset = &H6C400
                                Case 4
                                    attemptOffset = &H6C460
                                Case 5
                                    attemptOffset = &H6C538
                                Case Else
                                    Return
                            End Select

                            ''
                            ' As the emulator is the same mupen64plus module, the rest of the code is the same as the existing attachToM64P function.
                            ''
                            ' Step through all modules to find mupen64plus.dll's base address
                            For Each mo As ProcessModule In target.Modules
                                If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                                    addressDLL = mo.BaseAddress.ToInt64
                                    Exit For
                                End If
                            Next
                            ' Check if mupen64plus.dll was found
                            If Not addressDLL = 0 Then
                                ' Add location of variable to base address
                                addressDLL = addressDLL + attemptOffset
                                ' Set it as the current emulator
                                emulator = "modloader64-gui"
                                ' Read the first half of the address
                                Dim readR15 As Integer = ReadMemory(Of Integer)(addressDLL)
                                ' Convert to hex
                                Dim hexR15 As String = Hex(readR15)
                                If Not hexR15 = "0" Then
                                    ' Make sure length is 8 digit for any dropped 0's
                                    fixHex(hexR15)
                                    ' Read the second half of the address
                                    readR15 = ReadMemory(Of Integer)(addressDLL + 4)
                                    ' Convert to hex and attach to first half
                                    hexR15 = Hex(readR15) & hexR15
                                    romAddrStart64 = CLng("&H" & hexR15) + attemptAdded
                                    If ReadMemory(Of Integer)(romAddrStart64 + &H11A5EC) = 1514490948 Then
                                        ' Note as a positive hit and exit the loop
                                        positiveHit = True
                                        Exit For
                                    End If

                                End If
                            End If
                        Next
                        ' If a positive hit, finish up. If not, reset the target to advance to the next process
                        If positiveHit Then
                            Exit For
                        Else
                            target = Nothing
                        End If
                    End If
                Next
                If target IsNot Nothing Then
                    Exit For
                End If
            Next
            ''
            ' I added a line to handle the case that the launcher had started but emulation hadn't yet. Printing here isn't really necessary.
            ''
            If target Is Nothing Then
                rtbOutputLeft.Text = "Found modloader64 but couldn't find a running emulator." & vbCrLf
                Return
            End If
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                ' This is the expected error if process was not found, just return
                Return
            Else
                ' Any other error, output error message to textbox
                rtbOutputLeft.Text = "Attachment Problem: " & ex.Message & vbCrLf
                Return
            End If
        End Try
    End Sub

    Private Sub attachToModLoader642()
        ' This sub, and thus support for ML64, is credited to subenji, who also added the memory function to attach by process ID -- 2022.06.17

        ' This should already be empty in order to reach this point, but never hurts to make sure
        emulator = String.Empty
        ' If not 64bit, do not even bother
        If IS_64BIT = False Then Exit Sub
        ' Prepare new target process
        Dim target As Process = Nothing

        Try
            ' Try to attach to application
            ' target = Process.GetProcessesByName("modloader64-gui")(0)
            ''
            ' ModLoader64 runs 5 or 6 processes with only one of them actually being the emulator window, so we need to check the emulation is running
            ''
            Dim processes As Process() = Process.GetProcessesByName("modloader64-gui")
            If processes.Length = 0 Then
                ' If process was not found, just return
                Return
            End If
            For Each p As Process In processes
                For Each pModule As ProcessModule In p.Modules
                    If LCase(pModule.ModuleName) = "mupen64plus.dll" Then
                        ''
                        ' As the process is run several times and I need to specify by PID rather than name, I need to attach this once by hand
                        ''
                        If Not OpenProcessHandleById(p.Id) Then
                            rtbOutputLeft.Text = "Attachment Problem: Could not open process handle: " & p.Id & vbCrLf
                            Return
                        End If
                        target = p
                        Exit For
                    End If
                Next
                If target IsNot Nothing Then
                    Exit For
                End If
            Next
            ''
            ' I added a line to handle the case that the launcher had started but emulation hadn't yet. Printing here isn't really necessary.
            ''
            If target Is Nothing Then
                rtbOutputLeft.Text = "Found modloader64 but couldn't find a running emulator." & vbCrLf
                Return
            End If
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                ' This is the expected error if process was not found, just return
                Return
            Else
                ' Any other error, output error message to textbox
                rtbOutputLeft.Text = "Attachment Problem: " & ex.Message & vbCrLf
                Return
            End If
        End Try

        ' Prepare new address variable
        Dim addressDLL As Int64 = 0
        Dim attemptOffset As Int64 = 0
        Dim attemptAdded As Int64 = 0

        ''
        ' These pointers to the Base ROM location were all listed as static, I felt it best just to add them all. The first one worked every time in testing.
        ' I found adding 0x80000000 was never necessary.
        ''
        For attempt = 0 To 5
            Select Case attempt
                Case 0
                    attemptOffset = &H116ECF8
                Case 1
                    attemptOffset = &H12EED10
                Case 2
                    attemptOffset = &H6A6F0
                Case 3
                    attemptOffset = &H6C400
                Case 4
                    attemptOffset = &H6C460
                Case 5
                    attemptOffset = &H6C538
                Case Else
                    Return
            End Select

            ''
            ' As the emulator is the same mupen64plus module, the rest of the code is the same as the existing attachToM64P function.
            ''
            ' Step through all modules to find mupen64plus.dll's base address
            For Each mo As ProcessModule In target.Modules
                If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                    addressDLL = mo.BaseAddress.ToInt64
                    Exit For
                End If
            Next
            ' Check if mupen64plus.dll was found
            If Not addressDLL = 0 Then
                ' Add location of variable to base address
                addressDLL = addressDLL + attemptOffset
                ' Set it as the current emulator
                emulator = "modloader64-gui"
                ' Read the first half of the address
                Dim readR15 As Integer = ReadMemory(Of Integer)(addressDLL)
                ' Convert to hex
                Dim hexR15 As String = Hex(readR15)
                If Not hexR15 = "0" Then
                    ' Make sure length is 8 digit for any dropped 0's
                    fixHex(hexR15)
                    ' Read the second half of the address
                    readR15 = ReadMemory(Of Integer)(addressDLL + 4)
                    ' Convert to hex and attach to first half
                    hexR15 = Hex(readR15) & hexR15
                    romAddrStart64 = CLng("&H" & hexR15) + attemptAdded
                    If ReadMemory(Of Integer)(romAddrStart64 + &H11A5EC) = 1514490948 Then Exit For
                End If
            End If
        Next
    End Sub

    Private Sub attachToRetroArch()
        ' This should already be empty in order to reach this point, but never hurts to make sure
        emulator = String.Empty
        ' If not 64bit, do not even bother
        If IS_64BIT = False Then Exit Sub
        ' Prepare new target process
        Dim target As Process = Nothing

        Try
            ' Try to attach to application
            target = Process.GetProcessesByName("retroarch")(0)
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                ' This is the expected error if process was not found, just return
                Return
            Else
                ' Any other error, output error message to textbox
                rtbOutputLeft.Text = "Attachment Problem: " & ex.Message & vbCrLf
                Return
            End If
        End Try

        ' Prepare new address variable
        Dim addressDLL As Int64 = 0

        ' Step through all modules to find RetroArch's mupen64plus.dll's base address
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus_next_libretro.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        ' RetroArch will be coded for two cores, though no idea why someone would use the second. If it found the mupen64plus, go ahead with it
        If addressDLL <> 0 Then
            ' Add location of variable to base address
            addressDLL = addressDLL + &H8E795E0

            ' Attach to process and set it as the current emulator
            SetProcessName("retroarch")
            emulator = "retroarch - mupen64plus"

            ' Read the first half of the address
            Dim readR15 As Integer = ReadMemory(Of Integer)(addressDLL)
            ' Convert to hex
            Dim hexR15 As String = Hex(readR15)

            ' Make sure length is 8 digit for any dropped 0's
            While hexR15.Length < 8
                hexR15 = "0" & hexR15
            End While

            ' Read the second half of the address
            readR15 = ReadMemory(Of Integer)(addressDLL + 4)
            ' Convert to hex and attach to first half
            hexR15 = Hex(readR15) & hexR15

            ' Set it + 0x8000000 as starting address, done as 4's since 8 invokes negative value
            romAddrStart64 = CLng("&H" & hexR15) + &H40000000 + &H40000000
        Else
            ' Check for RetroArch's parallel core, again, for whatever reason but best to cover extra bases
            For Each mo As ProcessModule In target.Modules
                If LCase(mo.ModuleName) = "parallel_n64_libretro.dll" Then
                    addressDLL = mo.BaseAddress.ToInt64
                    Exit For
                End If
            Next
            If addressDLL = 0 Then Exit Sub

            ' Prepare new address variable
            Dim attemptOffset As Int64 = 0

            For i = 0 To 1
                Select Case i
                    Case 0
                        ' 1.9.0 - 1.10.2
                        attemptOffset = &H845000
                    Case 1
                        ' 1.10.3+
                        attemptOffset = &H844000
                    Case Else
                        Return
                End Select
                romAddrStart64 = addressDLL + attemptOffset
                SetProcessName("retroarch")
                If ReadMemory(Of Integer)(romAddrStart64 + &H11A5EC) = 1514490948 Then Exit For
            Next

            emulator = "retroarch - parallel"
        End If

    End Sub
    Private Sub attachToProject64()
        emulator = String.Empty
        If IS_64BIT Then Exit Sub

        p = Nothing
        If Process.GetProcessesByName("project64").Count > 0 Then
            p = Process.GetProcessesByName("project64")(0)
        Else
            Exit Sub
        End If

        For i = 0 To 1
            Select Case i
                Case 0
                    romAddrStart = &HDFE40000
                Case 1
                    romAddrStart = &HDFFB0000
                Case Else
                    Exit Sub
            End Select
            If quickRead32(romAddrStart + &H11A5EC, "project64", False) = 1514490948 Then
                emulator = "project64"
                Exit For
            End If
        Next
    End Sub
    Private Sub attachToM64PY()
        emulator = String.Empty
        If IS_64BIT Then Exit Sub
        ' Dim target As Process = Nothing
        Try
            ' Try to attach to application
            p = Process.GetProcessesByName("m64py")(0)
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                ' This is the expected error if process was not found, just return
                Return
            Else
                ' Any other error, output error message to textbox
                rtbOutputLeft.Text = "Attachment Problem: " & ex.Message & vbCrLf
                Return
            End If
        End Try

        ' Prepare new address variable
        Dim addressDLL As Int64 = 0

        For Each mo As ProcessModule In p.Modules
            If LCase(mo.ModuleName) = "mupen64plus-audio-sdl.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        If addressDLL <> 0 Then
            SetProcessName("m64py")
            romAddrStart = ReadMemory(Of Integer)(addressDLL + &H172060)
            If ReadMemory(Of Integer)(romAddrStart + &H11A5EC) = 1514490948 Then emulator = "m64py"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Select Case emulator
            Case String.Empty
                Exit Sub
            Case "variousX64"
                WriteMemory(Of Integer)(romAddrStart64 + &H11A644, &H10203)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A648, &H4050608)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A64C, &H90B0C0D)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A650, &HE0F1011)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A654, &H12131818)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A658, &H1418372B)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A65C, &H1E28281B)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A660, &H2B00)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A664, &H2B000000)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A668, &HA0A)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A66C, &H77770000)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A670, &H36E4DB)
                WriteMemory(Of Integer)(romAddrStart64 + &H11A674, &H7FFFFF)
            Case Else
                quickWrite(romAddrStart + &H11A644, &H10203, emulator)
                quickWrite(romAddrStart + &H11A648, &H4050608, emulator)
                quickWrite(romAddrStart + &H11A64C, &H90B0C0D, emulator)
                quickWrite(romAddrStart + &H11A650, &HE0F1011, emulator)
                quickWrite(romAddrStart + &H11A654, &H12131818, emulator)
                quickWrite(romAddrStart + &H11A658, &H1818372B, emulator)
                quickWrite(romAddrStart + &H11A65C, &H1E28281B, emulator)
                quickWrite(romAddrStart + &H11A660, &H2B00, emulator)
                quickWrite(romAddrStart + &H11A664, &H2B000000, emulator)
                quickWrite(romAddrStart + &H11A668, &HA0A, emulator)
                quickWrite(romAddrStart + &H11A66C, &H77770000, emulator)
                quickWrite(romAddrStart + &H11A670, &H36E4DB, emulator)
                quickWrite(romAddrStart + &H11A674, &H7FFFFF, emulator)
        End Select
    End Sub

    Private Sub clearItems()
        For Each pbx In pnlItems.Controls.OfType(Of PictureBox)()
            pbx.Visible = False
        Next
        For Each pbx In pnlEquips.Controls.OfType(Of PictureBox)()
            pbx.Visible = False
        Next

        For i As Byte = 0 To 22
            With aoQuestItems(i)
                If Not .Visible Then .Visible = True
                .Image = aoQuestItemImagesEmpty(i)
            End With
            aQuestRewardsCollected(i) = False
            aQuestRewardsText(i) = 0
            Select Case i
                Case 0 To 11, 18 To 20
                    changeAddedText(i, , True)
            End Select
        Next

        For Each pbx In pnlDungeonItems.Controls.OfType(Of PictureBox)()
            pbx.Visible = False
        Next
        pbxPoH.Visible = True
    End Sub
    Private Sub updateItems()
        If emulator = String.Empty Then Exit Sub
        If isLoadedGame() = False Then Exit Sub
        Dim items As String = String.Empty
        Dim quantity As String = String.Empty
        Dim temp As String = String.Empty
        Dim cTemp As Integer = 0

        Dim addrItems As Integer = CInt(IIf(isSoH, &HEC85D8, &H11A644))

        For i = addrItems To addrItems + 20 Step 4
            temp = Hex(goRead(i))
            fixHex(temp)
            items = items & temp
            temp = Hex(goRead(i + 24))
            fixHex(temp)
            quantity = quantity & temp
        Next

        If isSoH Then
            endianFlip(items)
            endianFlip(quantity)
            If Mid(items, 1, 8) = "00000000" Then
                lastFirstEnum = 255
                Exit Sub
            End If
        End If
        ' Storing items to an easy-to-scan string for logic detection
        allItems = String.Empty

        For i = 0 To 23 ' 1 To items.Length Step 2
            temp = Mid(items, (i * 2) + 1, 2)
            cTemp = CInt("&H" & temp)
            With aoInventory(i)
                Select Case cTemp
                    Case 0 To 55
                        If Not .Visible Then .Visible = True
                    Case Else
                        .Visible = False
                End Select
                Select Case cTemp
                    Case 0
                        .Image = My.Resources.dekuStick
                        If Not Mid(quantity, (i * 2) + 1, 2) = "00" Then
                            allItems = allItems & "a"
                        End If
                    Case 1
                        .Image = My.Resources.dekuNut
                        If Not Mid(quantity, (i * 2) + 1, 2) = "00" Then
                            allItems = allItems & "b"
                        End If
                    Case 2
                        .Image = My.Resources.bombs
                        If Not Mid(quantity, (i * 2) + 1, 2) = "00" Then
                            allItems = allItems & "c"
                        End If
                    Case 3
                        .Image = My.Resources.bow
                        If Not Mid(quantity, (i * 2) + 1, 2) = "00" Then
                            allItems = allItems & "d"
                        End If
                    Case 4
                        .Image = My.Resources.fireArrow
                        allItems = allItems & "e"
                    Case 5
                        .Image = My.Resources.dinsFire
                        allItems = allItems & "f"
                    Case 6
                        .Image = My.Resources.fairySlingshot
                        If Not Mid(quantity, (i * 2) + 1, 2) = "00" Then
                            allItems = allItems & "g"
                        End If
                    Case 7
                        .Image = My.Resources.fairyOcarina
                        allItems = allItems & "h"
                    Case 8
                        .Image = My.Resources.ocarinaOfTime
                        allItems = allItems & "hi"
                    Case 9
                        .Image = My.Resources.bombchus
                        If Not Mid(quantity, (i * 2) + 1, 2) = "00" Then
                            allItems = allItems & "j"
                        End If
                    Case 10
                        .Image = My.Resources.hookshot
                        allItems = allItems & "k"
                    Case 11
                        .Image = My.Resources.longshot
                        allItems = allItems & "kl"
                    Case 12
                        .Image = My.Resources.iceArrows
                        allItems = allItems & "m"
                    Case 13
                        .Image = My.Resources.faroresWind
                        allItems = allItems & "n"
                    Case 14
                        .Image = My.Resources.boomerang
                        allItems = allItems & "o"
                    Case 15
                        .Image = My.Resources.lensOfTruth
                        allItems = allItems & "p"
                    Case 16
                        .Image = My.Resources.magicBeans
                        allItems = allItems & "q"
                    Case 17
                        .Image = My.Resources.megatonHammer
                        allItems = allItems & "r"
                    Case 18
                        .Image = My.Resources.lightArrows
                        allItems = allItems & "s"
                    Case 19
                        .Image = My.Resources.nayrusLove
                        allItems = allItems & "t"
                    Case 20
                        .Image = My.Resources.bottleEmpty
                        allItems = allItems & "u"
                    Case 21
                        .Image = My.Resources.bottleRedPotion
                        allItems = allItems & "u"
                    Case 22
                        .Image = My.Resources.bottleGreenPotion
                        allItems = allItems & "u"
                    Case 23
                        .Image = My.Resources.bottleBluePotion
                        allItems = allItems & "u"
                    Case 24
                        .Image = My.Resources.bottleBottledFairy
                        allItems = allItems & "u"
                    Case 25
                        .Image = My.Resources.bottleFish
                        allItems = allItems & "u"
                    Case 26
                        .Image = My.Resources.bottleLonLonMilk
                        allItems = allItems & "u"
                    Case 27
                        .Image = My.Resources.bottleLetter
                        allItems = allItems & "v"
                    Case 28
                        .Image = My.Resources.bottleBlueFire
                        allItems = allItems & "uw"
                    Case 29
                        .Image = My.Resources.bottleBug
                        allItems = allItems & "u"
                    Case 30
                        .Image = My.Resources.bottleBigPoe
                        allItems = allItems & "u"
                    Case 31
                        .Image = My.Resources.bottleLonLonMilkHalf
                        allItems = allItems & "u"
                    Case 32
                        .Image = My.Resources.bottlePoe
                        allItems = allItems & "u"
                    Case 33
                        .Image = My.Resources.youngWeirdEgg
                        allItems = allItems & "y1"
                    Case 34
                        .Image = My.Resources.youngChicken
                        allItems = allItems & "y2"
                    Case 35
                        .Image = My.Resources.youngZeldasLetter
                        allItems = allItems & "y3"
                    Case 36
                        .Image = My.Resources.youngKeatonMask
                        allItems = allItems & "y4"
                    Case 37
                        .Image = My.Resources.youngSkullMask
                        allItems = allItems & "y5"
                    Case 38
                        .Image = My.Resources.youngSpookyMask
                        allItems = allItems & "y6"
                    Case 39
                        .Image = My.Resources.youngBunnyHood
                        allItems = allItems & "y7"
                    Case 40
                        .Image = My.Resources.youngGoronMask
                        allItems = allItems & "y8"
                    Case 41
                        .Image = My.Resources.youngZoraMask
                        allItems = allItems & "y9"
                    Case 42
                        .Image = My.Resources.youngGerudoMask
                        allItems = allItems & "yA"
                    Case 43
                        .Image = My.Resources.youngMaskOfTruth
                        allItems = allItems & "yB"
                    Case 44
                        .Image = My.Resources.youngSoldOut
                        allItems = allItems & "y0"
                    Case 45
                        .Image = My.Resources.adultPocketEgg
                        allItems = allItems & "z1"
                    Case 46
                        .Image = My.Resources.adultPocketCucco
                        allItems = allItems & "z2"
                    Case 47
                        .Image = My.Resources.adultCojiro
                        allItems = allItems & "z3"
                    Case 48
                        .Image = My.Resources.adultOddMushroom
                        allItems = allItems & "z4"
                    Case 49
                        .Image = My.Resources.adultOddPotion
                        allItems = allItems & "z5"
                    Case 50
                        .Image = My.Resources.adultPoachersSaw
                        allItems = allItems & "z6"
                    Case 51
                        .Image = My.Resources.adultGoronsSwordBroken
                        allItems = allItems & "z7"
                    Case 52
                        .Image = My.Resources.adultPrescription
                        allItems = allItems & "z8"
                    Case 53
                        .Image = My.Resources.adultEyeballFrog
                        allItems = allItems & "z9"
                    Case 54
                        .Image = My.Resources.adultEyeDrops
                        allItems = allItems & "zA"
                    Case 55
                        .Image = My.Resources.adultClaimCheck
                        allItems = allItems & "zB"
                    Case Else
                        .Image = My.Resources.emptySlot
                End Select
                If i <= 14 Then
                    Select Case cTemp
                        Case 0 To 3, 6, 9, 16
                            aGetQuantity(i) = True
                        Case Else
                            aGetQuantity(i) = False
                    End Select
                End If
            End With
        Next

        Dim iAmount As Byte = 0
        For i = 0 To 14
            iAmount = CByte("&H" & Mid(quantity, (i * 2) + 1, 2))
            If aGetQuantity(i) = True Then
                With aoInventory(i)
                    ' If single digits
                    Dim xPos As Byte = 34
                    ' If double digits
                    If iAmount > 9 Then xPos = 19
                    ' If triple digits
                    If iAmount > 99 Then xPos = 3
                    ' Font for items numbers
                    Dim fontItem = New Font("Lucida Console", 24, FontStyle.Bold, GraphicsUnit.Pixel)

                    ' Draw the value over the lower right of the item's picturebox, first in black to give it some definition, then in white
                    Graphics.FromImage(.Image).DrawString(iAmount.ToString, fontItem, New SolidBrush(Color.Black), xPos - 6, 28)
                    Graphics.FromImage(.Image).DrawString(iAmount.ToString, fontItem, New SolidBrush(Color.White), xPos - 5, 29)
                End With
            End If
        Next
        ' Since the last check is Magic Beans, store that value to our beans
        magicBeans = iAmount
    End Sub

    Private Sub endianFlip(ByRef input As String)
        Dim sDWORD As New List(Of String)
        For i = 1 To input.Length Step 8
            sDWORD.Add(Mid(input, i, 8))
        Next
        input = String.Empty
        Dim sCurrent As String = String.Empty
        For i = 0 To sDWORD.Count - 1
            input = input & Mid(sDWORD(i), 7, 2) & Mid(sDWORD(i), 5, 2) & Mid(sDWORD(i), 3, 2) & Mid(sDWORD(i), 1, 2)
        Next
    End Sub

    Private Sub updateDungeonItems()
        If isLoadedGame() = False Then Exit Sub
        getSmallKeys()
        getDungeonItems()
        getHearts()
        getMagic()
        getGoldSkulltulas()
        getTriforce()
        getPlayerName()
    End Sub

    Private Sub rtbOutputLeft_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles rtbOutputLeft.MouseDoubleClick
        Dim rtb As RichTextBox = CType(sender, RichTextBox)
        rtbDoubleClicks(rtb, e)
    End Sub
    Private Sub rtbOutputRight_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles rtbOutputRight.MouseDoubleClick
        Dim rtb As RichTextBox = CType(sender, RichTextBox)
        rtbDoubleClicks(rtb, e, True)
    End Sub
    Private Sub rtbDoubleClicks(ByVal rtb As RichTextBox, e As MouseEventArgs, Optional right As Boolean = False)
        If rtbOutputLeft.Text = Nothing Then Return
        With rtb
            ' Get the double-clicked location
            Dim clickPos As Integer = .GetCharIndexFromPosition(e.Location)
            ' Get the line that location was located on
            Dim linePos As Integer = .GetLineFromCharIndex(clickPos)
            ' Read the line into a string
            Dim readLine As String = .Lines(linePos).ToString
            ' Grab the first line from the left textbox, the location
            Dim readArea As String = rtbOutputLeft.Lines(0).ToString
            ' Clean up the location name
            readArea = Trim(readArea.Replace(":", "").Replace("MQ", "").Replace("(Found)", ""))
            ' Convert area name into the code used
            Dim areaCode As String = area2code(readArea)
            If right Then
                Select Case areaCode
                    Case "HF"
                        areaCode = "QBPH"
                    Case "ZR"
                        areaCode = "QF"
                End Select
            End If
            If flipKeyForced(Trim(readLine).Replace("(Y)", "").Replace("(A)", ""), areaCode) Then
                Dim newLine As String = readLine
                updateLabels()
                updateLabelsDungeons()
                btnFocus.Focus()
                If newLine.Contains("(Forced)") Then
                    newLine = newLine.Replace(" (Forced)", "")
                Else
                    newLine = newLine & " (Forced)"
                End If
                replaceRTB(rtb, linePos, " " & newLine.Replace("  ", " "))
                Application.DoEvents()
                ' If logic setting is set, and the embolden list is not empty
                If My.Settings.setLogic And emboldenList.Count > 0 Then
                    ' Run each line through the emboldening process
                    For Each line In emboldenList
                        embolden(line)
                    Next
                End If
            End If
        End With
    End Sub
    Private Sub replaceRTB(ByVal rtb As RichTextBox, ByVal line As Integer, ByVal newText As String)
        ' Replaces the targeted line with the newText
        Dim newRTB As String = String.Empty
        ' Step through the contents of the RichTextBox and remake its contents
        For i = 0 To rtb.Lines.Count - 1
            If i > 0 Then newRTB = newRTB & vbCrLf
            If Not i = line Then
                newRTB = newRTB & rtb.Lines(i)
            Else
                ' If we are at the selected line, instead replace it with newText
                newRTB = newRTB & newText
            End If
        Next
        ' Replace the text
        rtb.Text = newRTB
    End Sub
    Private Sub embolden(ByVal toBold As String)
        ' Trim up the text to bold
        toBold = Trim(toBold.Replace(vbCrLf, ""))
        Dim findStart As Integer = -1
        ' Start with left side
        Dim rtb As RichTextBox = rtbOutputLeft

        For ii = 1 To 2
            With rtb
                ' search for the line to bold
                For i = .Lines.Count - 1 To 0 Step -1
                    'If .Lines(i).Contains(toBold) Then
                    If Trim(.Lines(i)) = toBold Then
                        findStart = i
                        Exit For
                    End If
                Next
            End With
            If findStart = -1 Then
                ' If not found, check for the right side next
                rtb = rtbOutputRight
            Else
                ' If found, go ahead and exit the small loop
                Exit For
            End If
        Next
        ' If still not found, then exit sub
        If findStart = -1 Then Exit Sub

        ' Set up the font to be the same, select the line, and replace the font with just a bolded version
        Dim richFont As System.Drawing.Font = rtb.Font
        rtb.Select(rtb.GetFirstCharIndexFromLine(findStart), rtb.Lines(findStart).Length - 1)
        rtb.SelectionFont = New Font(richFont.FontFamily, richFont.Size, FontStyle.Bold)
    End Sub

    Private Sub outputSong(ByVal title As String, ByVal notes As String)
        rtbRefresh = 2
        lastOutput.Clear()
        rtbOutputLeft.Text = title & ":"
        rtbOutputRight.Clear()
        Dim outSong As String = "  "
        For i = 1 To notes.Length
            Select Case LCase(Mid(notes, i, 1))
                Case "a"
                    outSong = outSong & "A  "
                Case "u"
                    outSong = outSong & "▲ "
                Case "d"
                    outSong = outSong & "▼ "
                Case "l"
                    outSong = outSong & "◀   "
                Case "r"
                    outSong = outSong & "▶   "
            End Select
        Next
        rtbAddLine(outSong)
    End Sub
    Private Sub pbxZeldasLullaby_Click(sender As Object, e As EventArgs) Handles pbxZeldasLullaby.Click
        outputSong("Zelda's Lullaby", "LURLUR")
    End Sub
    Private Sub pbxEponasSong_Click(sender As Object, e As EventArgs) Handles pbxEponasSong.Click
        outputSong("Epona's Song", "ULRULR")
    End Sub
    Private Sub pbxSariasSong_Click(sender As Object, e As EventArgs) Handles pbxSariasSong.Click
        outputSong("Saria's Song", "DRLDRL")
    End Sub
    Private Sub pbxSunsSong_Click(sender As Object, e As EventArgs) Handles pbxSunsSong.Click
        outputSong("Sun's Song", "RDURDU")
    End Sub
    Private Sub pbxSongOfTime_Click(sender As Object, e As EventArgs) Handles pbxSongOfTime.Click
        outputSong("Song of Time", "RADRAD")
    End Sub
    Private Sub pbxSongOfStorms_Click(sender As Object, e As EventArgs) Handles pbxSongOfStorms.Click
        outputSong("Song of Storms", "ADUADU")
    End Sub
    Private Sub pbxMinuetOfForest_Click(sender As Object, e As EventArgs) Handles pbxMinuetOfForest.Click
        outputSong("Minuet of Forest", "AULRLR")
    End Sub
    Private Sub pbxBoleroOfFire_Click(sender As Object, e As EventArgs) Handles pbxBoleroOfFire.Click
        outputSong("Bolero of Fire", "DADARDRD")
    End Sub
    Private Sub pbxSerenadeOfWater_Click(sender As Object, e As EventArgs) Handles pbxSerenadeOfWater.Click
        outputSong("Serenade of Water", "ADRRL")
    End Sub
    Private Sub pbxRequiemOfSpirit_Click(sender As Object, e As EventArgs) Handles pbxRequiemOfSpirit.Click
        outputSong("Requiem of Spirit", "ADARDA")
    End Sub
    Private Sub pbxNocturneOfShadow_Click(sender As Object, e As EventArgs) Handles pbxNocturneOfShadow.Click
        outputSong("Nocturne of Shadow", "LRRALRD")
    End Sub
    Private Sub pbxPreludeOfLight_Click(sender As Object, e As EventArgs) Handles pbxPreludeOfLight.Click
        outputSong("Prelude of Light", "URURLU")
    End Sub
    Private Sub pbxStoneKokiri_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxStoneKokiri.MouseClick
        changeAddedText(18, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneKokiri_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxStoneKokiri.MouseDoubleClick
        changeAddedText(18, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneGoron_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxStoneGoron.MouseClick
        changeAddedText(19, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneGoron_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxStoneGoron.MouseDoubleClick
        changeAddedText(19, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneZora_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxStoneZora.MouseClick
        changeAddedText(20, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneZora_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxStoneZora.MouseDoubleClick
        changeAddedText(20, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalForest_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalForest.MouseClick
        changeAddedText(0, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalForest_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalForest.MouseDoubleClick
        changeAddedText(0, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalFire_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalFire.MouseClick
        changeAddedText(1, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalFire_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalFire.MouseDoubleClick
        changeAddedText(1, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalWater_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalWater.MouseClick
        changeAddedText(2, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalWater_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalWater.MouseDoubleClick
        changeAddedText(2, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalSpirit_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalSpirit.MouseClick
        changeAddedText(3, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalSpirit_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalSpirit.MouseDoubleClick
        changeAddedText(3, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalShadow_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalShadow.MouseClick
        changeAddedText(4, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalShadow_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalShadow.MouseDoubleClick
        changeAddedText(4, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalLight_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalLight.MouseDoubleClick
        changeAddedText(5, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalLight_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalLight.MouseClick
        changeAddedText(5, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub frmTrackerOfTime_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        btnFocus.Focus()
    End Sub
    Private Sub frmTrackerOfTime_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Dim pnFore As Pen = New Pen(Me.ForeColor, 1)
        If showSetting Then e.Graphics.DrawLine(pnFore, pnlMain.Width + 10, 0, pnlMain.Width + 10, Me.Height)
        updateSettingsPanel()
    End Sub
    Private Sub frmTrackerOfTime_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' This is the best solution I could come up with for it not drawing the checkboxes after a resize
        If firstRun Then Exit Sub
        For Each lbl In pnlSettings.Controls.OfType(Of Label)()
            lbl.Refresh()
        Next
    End Sub

    Private Sub handleLCXMouseClick(sender As Object, e As MouseEventArgs)
        Dim lcx As Label = CType(sender, Label)
        If Mid(lcx.Name, 1, 3) = "nlb" Then Exit Sub
        If e.Button = Windows.Forms.MouseButtons.Left Then
            clickLCX(lcx.Text)
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            infoLCX(lcx.Text)
        End If
    End Sub
    Private Sub clickLCX(ByVal text As String)
        ' Compare sent text with lcx text to figure out which one was clicked, flip its setting variable
        Select Case text
            Case lcxLogic.Text
                My.Settings.setLogic = Not My.Settings.setLogic
            Case lcxCowShuffle.Text
                My.Settings.setCow = Not My.Settings.setCow
            Case lcxScrubShuffle.Text
                ' Scrub shuffle requires a little more work, to disable the original 3 scrub keys and enable scrub shuffle keys. And vice versa
                My.Settings.setScrub = Not My.Settings.setScrub
                changeScrubs()
            Case lcxBombchuLogic.Text
                My.Settings.setBombchus = Not My.Settings.setBombchus
            Case lcxLoTLogic.Text
                My.Settings.setLensOfTruth = Not My.Settings.setLensOfTruth
            Case lcxSoALogic.Text
                My.Settings.setStoneOfAgony = Not My.Settings.setStoneOfAgony
            Case lcxHoverTricks.Text
                My.Settings.setHoverTricks = Not My.Settings.setHoverTricks
            Case lcxOpenForest.Text
                My.Settings.setOpenKF = Not My.Settings.setOpenKF
            Case lcxOpenFountain.Text
                My.Settings.setOpenZF = Not My.Settings.setOpenZF
            Case lcxFiTRustedSwitches.Text
                My.Settings.setFiTRustedSwitches = Not My.Settings.setFiTRustedSwitches
            Case lcxTunic.Text
                My.Settings.setFewerTunics = Not My.Settings.setFewerTunics
            Case lcxDekuB1Skip.Text
                My.Settings.setDekuB1Skip = Not My.Settings.setDekuB1Skip
            Case lcxDCStaircase.Text
                My.Settings.setDCStaircase = Not My.Settings.setDCStaircase
            Case lcxDCSpikeJump.Text
                My.Settings.setDCSpikeJump = Not My.Settings.setDCSpikeJump
            Case lcxDCSlingshotSkips.Text
                My.Settings.setDCSlingshotSkips = Not My.Settings.setDCSlingshotSkips
            Case lcxDCLightEyes.Text
                My.Settings.setDCLightEyes = Not My.Settings.setDCLightEyes
            Case lcxFoTVines.Text
                My.Settings.setFoTVines = Not My.Settings.setFoTVines
            Case lcxFoTFrame.Text
                My.Settings.setFoTFrame = Not My.Settings.setFoTFrame
            Case lcxFoTLedge.Text
                My.Settings.setFoTLedge = Not My.Settings.setFoTLedge
            Case lcxFoTBackdoor.Text
                My.Settings.setFoTBackdoor = Not My.Settings.setFoTBackdoor
            Case lcxFoTMQPuzzle.Text
                My.Settings.setFoTMQPuzzle = Not My.Settings.setFoTMQPuzzle
            Case lcxFoTMQTwisted.Text
                My.Settings.setFoTMQTwisted = Not My.Settings.setFoTMQTwisted
            Case lcxFoTMQWell.Text
                My.Settings.setFoTMQWell = Not My.Settings.setFoTMQWell
            Case lcxFiTClimb.Text
                My.Settings.setFiTClimb = Not My.Settings.setFiTClimb
            Case lcxFiTScarecrow.Text
                My.Settings.setFiTScarecrow = Not My.Settings.setFiTScarecrow
            Case lcxFiTMaze.Text
                My.Settings.setFiTMaze = Not My.Settings.setFiTMaze
            Case lcxFiTMQClimb.Text
                My.Settings.setFiTMQClimb = Not My.Settings.setFiTMQClimb
            Case lcxFiTMQMazeHovers.Text
                My.Settings.setFiTMQMazeHovers = Not My.Settings.setFiTMQMazeHovers
            Case lcxFiTMQMazeJump.Text
                My.Settings.setFiTMQMazeJump = Not My.Settings.setFiTMQMazeJump
            Case lcxWaTTorch.Text
                My.Settings.setWaTTorch = Not My.Settings.setWaTTorch
            Case lcxWaTDrgnSwitch.Text
                My.Settings.setWaTDrgnSwitch = Not My.Settings.setWaTDrgnSwitch
            Case lcxWaTDrgnDive.Text
                My.Settings.setWaTDragonDive = Not My.Settings.setWaTDragonDive
            Case lcxWaTBKR.Text
                My.Settings.setWaTBKR = Not My.Settings.setWaTBKR
            Case lcxWaTCrackNothing.Text
                My.Settings.setWaTCrackNothing = Not My.Settings.setWaTCrackNothing
            Case lcxWaTCrackHovers.Text
                My.Settings.setWaTCrackHovers = Not My.Settings.setWaTCrackHovers
            Case lcxSpTLensless.Text
                My.Settings.setSpTLensless = Not My.Settings.setSpTLensless
            Case lcxSpTWall.Text
                My.Settings.setSpTWall = Not My.Settings.setSpTWall
            Case lcxSpTBombchu.Text
                My.Settings.setSpTBombchu = Not My.Settings.setSpTBombchu
            Case lcxSpTMQLowAdult.Text
                My.Settings.setSpTBombchu = Not My.Settings.setSpTMQLowAdult
            Case lcxSpTMQSunRoom.Text
                My.Settings.setSpTMQSunRoom = Not My.Settings.setSpTMQSunRoom
            Case lcxShTLensless.Text
                My.Settings.setShTLensless = Not My.Settings.setShTLensless
            Case lcxShTPlatform.Text
                My.Settings.setShTPlatform = Not My.Settings.setShTPlatform
            Case lcxShTUmbrella.Text
                My.Settings.setShTUmbrella = Not My.Settings.setShTUmbrella
            Case lcxShTMQGap.Text
                My.Settings.setShTMQGap = Not My.Settings.setShTMQGap
            Case lcxShTMQPit.Text
                My.Settings.setShTMQPit = Not My.Settings.setShTMQPit
            Case lcxShTStatue.Text
                My.Settings.setShTStatue = Not My.Settings.setShTStatue
            Case lcxBotWLensless.Text
                My.Settings.setBotWLensless = Not My.Settings.setBotWLensless
            Case lcxBotWDeadHand.Text
                My.Settings.setBotWDeadHand = Not My.Settings.setBotWDeadHand
            Case lcxBotWMQPits.Text
                My.Settings.setBotWMQPits = Not My.Settings.setBotWMQPits
            Case lcxNavi.Text
                My.Settings.setNavi = Not My.Settings.setNavi
            Case lcxBeans.Text
                My.Settings.setBeans = Not My.Settings.setBeans
                changeSalesmans()
            Case lcxSaleman.Text
                My.Settings.setSalesman = Not My.Settings.setSalesman
                changeSalesmans()
            Case lcxWastelandBackstep.Text
                My.Settings.setWastelandBackstep = Not My.Settings.setWastelandBackstep
            Case lcxWastelandLensless.Text
                My.Settings.setWastelandLensless = Not My.Settings.setWastelandLensless
            Case lcxWastelandReverse.Text
                My.Settings.setWastelandReverse = Not My.Settings.setWastelandReverse
            Case lcxIGCMQRustedSwitches.Text
                My.Settings.setIGCMQRustedSwitches = Not My.Settings.setIGCMQRustedSwitches
            Case lcxIGCLensless.Text
                My.Settings.setIGCLensless = Not My.Settings.setIGCLensless
            Case lcxHideSpoiler.Text
                My.Settings.setHideSpoiler = Not My.Settings.setHideSpoiler
            Case lcxCucco.Text
                My.Settings.setCucco = Not My.Settings.setCucco
            Case lcxShowMap.Text
                My.Settings.setMap = Not My.Settings.setMap
                resizeForm()
                If My.Settings.setShortForm Then
                    showSetting = Not showSetting
                    updateShowSettings()
                    showSetting = Not showSetting
                    updateShowSettings()
                End If
            Case lcxExpand.Text
                My.Settings.setExpand = Not My.Settings.setExpand
                resizeForm()
                If My.Settings.setShortForm Then
                    showSetting = Not showSetting
                    updateShowSettings()
                    showSetting = Not showSetting
                    updateShowSettings()
                End If
                tmrFixIt.Enabled = True
            Case lcxShortForm.Text
                My.Settings.setShortForm = Not My.Settings.setShortForm
                resizeForm()
                tmrFixIt.Enabled = True
                showSetting = Not showSetting
                updateShowSettings()
                showSetting = Not showSetting
                updateShowSettings()
            Case lcxHideQuests.Text
                My.Settings.setHideQuests = Not My.Settings.setHideQuests
            Case lcxGTGLensless.Text
                My.Settings.setGTGLensless = Not My.Settings.setGTGLensless
                'Case lcxxx.Text
                'My.Settings.setxx = Not My.Settings.setxx
            Case lblGoldSkulltulas.Text, lblShopsanity.Text, lblSmallKeys.Text, lblInfo.Text

            Case Else
                rtbOutputLeft.Text = "Unhandled LCX: " & text
        End Select

        ' Save settings and update labels and graphics
        My.Settings.Save()
        updateSettingsPanel()
        updateLabels()
        updateLabelsDungeons()
    End Sub
    Private Sub infoLCX(ByVal text As String)
        ' Compare sent text with lcx text to figure out which one was right-clicked, output info for each
        Dim message As String = String.Empty
        Select Case text
            Case lcxLogic.Text
                message = "Attempts to bold any checks you can currently reach, and any area with checks. Minor glitches or tricks may get you more checks than detected." & vbCrLf & vbCrLf &
                            "Note: Enabling logic will slow down the scanning.  It may take 2-3 scans to fully update, and may act up when changing areas due to location shifting" &
                            "mid-scan.  It is also a lot of work that may not be perfect.  A few logic tricks are able to be set if you know how to do them to help display more."
            Case lblGoldSkulltulas.Text
                message = "Adds Gold Skulltulas to the checks for each area. Can be set to stop tracking at 50 Tokens."
            Case lcxCowShuffle.Text
                message = "Adds shuffled cows to the checks for each area."
            Case lcxScrubShuffle.Text
                message = "Adds shuffled Deku Scrubs to the checks for each area."
            Case lblShopsanity.Text
                message = "Adds shopsanity checks to each shoppe." & vbCrLf & "NOTE: OOTR coding for Shopsanity Random is not supported due to their particular design. AP randomizer works with '4 items per shop' setting."
            Case lblSmallKeys.Text
                message = "Sets where to find keys. Either in their own dungeon, outside their dungeon (keysanity), or remove them completely."
            Case lcxOpenForest.Text
                message = "Opens Kokiri Forest to Hyrule Fields."
            Case lcxOpenFountain.Text
                message = "Opens Zora's Domain to Zora's Fountain for adult Link."
            Case lcxFiTRustedSwitches.Text
                message = "Hit rusted switches through walls in Fire Temple and Fire Temple MQ."
            Case lcxBombchuLogic.Text
                message = "Considers Bombchus as viable as bombs, also requires Bombchus for Bombchu Bowling."
            Case lcxLoTLogic.Text
                message = "Require Lens of Truth for Treasure Chest Game and various dungeon checks."
            Case lcxSoALogic.Text
                message = "Require Stone of Agony for finding hidden Grottos."
            Case lcxHoverTricks.Text
                message = "Overworld Hover Boots tricks to reach the Gold Skulltula above Impa's House, the Volcano Piece of Heart, the Goron City Maze Left Chest, and the Gerudo Valley Crate Piece of Heart."
            Case lcxTunic.Text
                message = "Enables access to various checks without the Goron or Zora Tunic."
            Case lcxDekuB1Skip.Text
                message = "A jump trick to skip needing the slingshot."
            Case lcxDCStaircase.Text
                message = "Knock down the stairs with the bow and quick shots."
            Case lcxDCSpikeJump.Text
                message = "A corner jump to reach a platform in the spike trap room without Hover Boots."
            Case lcxDCSlingshotSkips.Text
                message = "A set of careful jumps to pass the flame circles as young Link without the Slingshot."
            Case lcxDCLightEyes.Text
                message = "Use bomb flowers to light eyes in the Master Quest version."
            Case lcxFoTVines.Text
                message = "Reach vines in eastern courtyard with just Hookshot."
            Case lcxFoTFrame.Text
                message = "Use Hover Boots from the upper balconies to land on a door frame, allowing you to reach the scarecrow in normal version of Forest Temple, or the gold skulltula in the MQ version."
            Case lcxFoTLedge.Text
                message = "Use Hover Boots to fall down to the North-East outdoor's ledge.."
            Case lcxFoTBackdoor.Text
                message = "Use a jump slash to reach the backdoor to the western courtyard."
            Case lcxFoTMQPuzzle.Text
                message = "In Master Quest version, use a Bombchu to trigger crystal, bypassing need for Bracelet or Gauntelets."
            Case lcxFoTMQTwisted.Text
                message = "In Master Quest version, use a jump slash to hit hallway switch through glass."
            Case lcxFoTMQWell.Text
                message = "In Master Quest version, use the Hookshot swim through the well before draining it."
            Case lcxFiTClimb.Text
                message = "Skip needing to push the block to reach grate."
            Case lcxFiTScarecrow.Text
                message = "Reach the elevator target with the Longshot, skipping the need for the scarecrow song."
            Case lcxFiTMaze.Text
                message = "Quickly move past the edge of the flame wall before it can block you. Both in normal and MQ versions."
            Case lcxFiTMQClimb.Text
                message = "Use Hover Boots to get reach climable wall without a fire source."
            Case lcxFiTMQMazeHovers.Text
                message = "Use Hover Boots to reach Upper Lizalfos Maze."
            Case lcxFiTMQMazeJump.Text
                message = "Use a precise jump off a crate to reach Upper Lizalfos Maze."
            Case lcxWaTTorch.Text
                message = "Use Longshot to reach bottom level torches and swim into corridor."
            Case lcxWaTDrgnSwitch.Text
                message = "While standing just before the water, use the Bow, Hookshot, or a Bombchu to hit the switch. Then carefully jump into the water to the right, allowing the current to take you into the tunnel."
            Case lcxWaTDrgnDive.Text
                message = "When entering Dragon Statue room from the river, a well aimed dive can get you into the tunnel without need for Iron Boots or Scale. Works in both normal and MQ versions, but need to hit the switch with the Bow in the normal version."
            Case lcxWaTBKR.Text
                message = "Use Hover Boots to reach Boss Key Region without the need for the Longshot."
            Case lcxWaTCrackNothing.Text
                message = "Reach the cracked wall without raising water level or Hover Boots."
            Case lcxWaTCrackHovers.Text
                message = "Reach the cracked wall with the Hover Boots, without raising the water level."
            Case lcxSpTLensless.Text
                message = "Navigate the Spirit Temple and Spirit Temple MQ without Lens of Truth."
            Case lcxSpTWall.Text
                message = "Climb the shifting wall without dealing with the Beamos or the Walltula first."
            Case lcxSpTBombchu.Text
                message = "Use a Bombchu to hit the bridge switch."
            Case lcxSpTMQLowAdult.Text
                message = "Use Din's Fire and Bow to quickly light all three torches to access the Lower Adult area."
            Case lcxSpTMQSunRoom.Text
                message = "You can toss and smash the crate on the switch to briefly unlock the door to the Sun Block Room."
            Case lcxShTLensless.Text
                message = "Navigate the Shadow Temple and Shadow Temple MQ without Lens of Truth, except for the invisible moving platform."
            Case lcxShTPlatform.Text
                message = "Get onto and off of the invisible moving platform in Shadow Temple and Shadow Temple MQ."
            Case lcxShTUmbrella.Text
                message = "Use Hover Boots and a precise jump off of the lower chest to get atop the crushing spikes, or a well timed backflip, in Shadow Temple and Shadow Temple MQ."
            Case lcxShTStatue.Text
                message = "Send a bombchu across the edge of the gorge to knock down the statue in Shadow Temple and Shadow Temple MQ."
            Case lcxShTMQGap.Text
                message = "Longshot to a torch and then a jump slash off bounce yourself onto the tongue."
            Case lcxShTMQPit.Text
                message = "Make it to the lower part of the huge pit in Spirit Temple MQ without triggering the frozen eye."
            Case lcxBotWLensless.Text
                message = "Navigate the Bottom of the Well without Lens of Truth."
            Case lcxBotWDeadHand.Text
                message = "Use Deku Sticks to defeat the Dead Hand instead of the Kokiri Sword. Applies to both versions of the Bottom of the Well."
            Case lcxBotWMQPits.Text
                message = "Use side-hops, backflips, or explosives to get over invisible holes in Bottom of the Well MQ."
            Case lcxNavi.Text
                message = "I know we all love hearing Navi, but this option will keep Navi from randomly wanting to yell for you. She will still shout at curious shiny rocks on the ground, but no more telling you to get a move on."
            Case lcxBeans.Text
                message = "Shuffles a 10 pack of Magic Beans into the item pool, and adds the Magic Bean Salesman as an extra check."
            Case lcxSaleman.Text
                message = "Shuffles the Broken Knife into the item pool, and adds the Medigoron and Carpet Salesmans as extra checks."
            Case lcxWastelandBackstep.Text
                message = "Use backsteps to cross the shifting sands, rather than Longshot or Hover Boots."
            Case lcxWastelandLensless.Text
                message = "Cross the Wasteland without need of the Lens of Truth."
            Case lcxWastelandReverse.Text
                message = "Cross the Wasteland backwards."
            Case lcxIGCMQRustedSwitches.Text
                message = "Hit the rusted switch through the wall in Ganon's Castle MQ - Spirit Trail."
            Case lcxIGCLensless.Text
                message = "Navigate Ganon's Castle and Ganon's Castle MQ without Lens of Truth."
            Case lcxHideSpoiler.Text
                message = "This will treat your Rainbow Bridge and Light Arrows Cutscene (LAC) as vanilla. Uncheck if you wish the tracker to detect and use your settings to determin logic. This will not tell you what you need to get the checks, just if you can."
            Case lcxCucco.Text
                message = "As young Link, use a cucco to fly behind the water to enter Zora's Domain."
            Case lcxShowMap.Text
                message = "Uses a visual map rather than text. You do not get to see the number of checks for each area, but it looks prettier for streaming."
            Case lcxExpand.Text
                message = "The application is trimmed down for 1920x1080 displays. This will expand the application to see the bottom part of the overwold map, and the output box will show all checks with no cut-off." & vbCrLf & vbCrLf &
                    "Note: Steps have been taken to reduce any cut-off of checks you can reach. It only happens in rare occasions, will be fixed once you collect a few checks from the area, and if they are reachable, they should show in place of ones you cannot reach."
            Case lcxShortForm.Text
                message = "Shorten the application height and adds scrollbars to allow it to fit into 1366x768 resolutions."
            Case lblInfo.Text
                message = "A list of things to know."
            Case lcxHideQuests.Text
                message = "Hide Big Poes from being displayed."
            Case lcxGTGLensless.Text
                message = "Navigate Gerudo Training Ground without Lens of Truth."
                'Case lcxxx.Text
                'message = "."
        End Select

        If Not message = String.Empty Then
            If text.Length < 100 Then
                rtbOutputLeft.Text = text.Replace(":", "") & ":"
            Else
                rtbOutputLeft.Text = "Instructions:"
            End If
            rtbAddLine("  " & message, True)
        End If
    End Sub
    Private Sub changeScrubs()
        ' Step through each entry in aKeys and search for the 3 Scrubs
        Dim foundScrubs As Integer = 0
        For Each key In aKeys
            Select Case key.loc
                ' The 3 .loc codes for the 3 Scrubs
                Case "6827", "7202", "7203"
                    ' Add up a found, and set them to a flipped Scrub settings
                    inc(foundScrubs)
                    key.scan = Not My.Settings.setScrub
            End Select
            ' On 3 founds, quit the loop
            If foundScrubs >= 3 Then Exit For
        Next
    End Sub
    Private Sub changeSalesmans()
        ' Change shuffled salesmans to be scannable or not scannable, depending on the setting
        For Each key In aKeys
            With key
                ' Magic Bean Salesman
                If .loc = "2301" Then .scan = My.Settings.setBeans
                ' Medigoron
                If .loc = "3001" Then .scan = My.Settings.setSalesman
                ' Carpet Salesman
                If .loc = "11401" Then .scan = My.Settings.setSalesman
            End With
        Next
    End Sub
    Private Sub handleLTBMouseClick(sender As Object, e As MouseEventArgs)
        Dim ltb As Label = CType(sender, Label)
        clickLTB(ltb.Name, CBool(IIf(e.X <= ltb.Width / 2, True, False)))
    End Sub
    Private Sub handleLTBMouseDoubleClick(sender As Object, e As MouseEventArgs)
        Dim ltb As Label = CType(sender, Label)
        clickLTB(ltb.Name, CBool(IIf(e.X <= ltb.Width / 2, True, False)))
    End Sub
    Private Sub clickLTB(ByVal sName As String, Optional lower As Boolean = False)
        Dim addVal As Double = 1
        If lower Then addVal = -1

        Select Case sName
            Case ltbShopsanity.Name
                If My.Settings.setShop = 0 And lower Then
                    My.Settings.setShop = 4
                Else
                    My.Settings.setShop = CByte(My.Settings.setShop + addVal)
                    If My.Settings.setShop > 4 Then My.Settings.setShop = 0
                End If
            Case ltbGoldSkulltulas.Name
                If My.Settings.setSkulltula = 0 And lower Then
                    My.Settings.setSkulltula = 2
                Else
                    My.Settings.setSkulltula = CByte(My.Settings.setSkulltula + addVal)
                    If My.Settings.setSkulltula > 2 Then My.Settings.setSkulltula = 0
                End If
            Case ltbGSLocation.Name
                If My.Settings.setGSLoc = 0 And lower Then
                    My.Settings.setGSLoc = 2
                Else
                    My.Settings.setGSLoc = CByte(My.Settings.setGSLoc + addVal)
                    If My.Settings.setGSLoc > 2 Then My.Settings.setGSLoc = 0
                End If
            Case ltbKeys.Name
                If My.Settings.setSmallKeys = 0 And lower Then
                    My.Settings.setSmallKeys = 2
                Else
                    My.Settings.setSmallKeys = CByte(My.Settings.setSmallKeys + addVal)
                    If My.Settings.setSmallKeys > 2 Then My.Settings.setSmallKeys = 0
                End If
                updateSmallKeys()
        End Select
        updateLTB(sName)
    End Sub
    Private Sub updateSmallKeys()
        Dim lbl As New Label
        For i = 3 To 10
            If i = 9 Then i = 10
            Select Case i
                Case 3
                    lbl = zFoT
                Case 4
                    lbl = zFiT
                Case 5
                    lbl = zWaT
                Case 6
                    lbl = zSpT
                Case 7
                    lbl = zShT
                Case 8
                    lbl = zBotW
                Case 10
                    lbl = zGTG
            End Select
            If canDungeon(i) Then lbl.Refresh()
        Next
        For i = 0 To canDungeon.Length - 1
            canDungeon(i) = False
        Next
    End Sub

    Public Sub updateLTB(ByVal ltbName As String)
        Select Case ltbName
            Case ltbShopsanity.Name
                Select Case My.Settings.setShop
                    Case 0
                        ltbShopsanity.Text = "Off"
                    Case Else
                        ltbShopsanity.Text = My.Settings.setShop.ToString & " Item" & IIf(My.Settings.setShop > 1, "s", "").ToString & " Per Shop"
                End Select
                updateShoppes()
            Case ltbGoldSkulltulas.Name
                Select Case My.Settings.setSkulltula
                    Case 0
                        ltbGoldSkulltulas.Text = "Off"
                    Case 1
                        ltbGoldSkulltulas.Text = "On"
                    Case Else
                        ltbGoldSkulltulas.Text = "Stop at 50"
                End Select
            Case ltbGSLocation.Name
                Select Case My.Settings.setGSLoc
                    Case 0
                        ltbGSLocation.Text = "Only Dungeons"
                    Case 1
                        ltbGSLocation.Text = "Everywhere"
                    Case Else
                        ltbGSLocation.Text = "Only Overworld"
                End Select
            Case ltbKeys.Name
                Select Case My.Settings.setSmallKeys
                    Case 0
                        ltbKeys.Text = "Dungeon"
                    Case 1
                        ltbKeys.Text = "Keysanity"
                    Case Else
                        ltbKeys.Text = "Remove"
                End Select
        End Select
    End Sub
    Private Sub drawArrows(sender As Object, e As PaintEventArgs)
        Dim ltb As Label = CType(sender, Label)
        With ltb
            Dim p As New Pen(Me.ForeColor, 7)
            p.EndCap = Drawing2D.LineCap.ArrowAnchor
            e.Graphics.DrawLine(p, 6, 6, 2, 6)
            e.Graphics.DrawLine(p, .Width - 6, 6, .Width - 2, 6)
        End With
    End Sub
    Private Sub pnlDrawBorder(sender As Object, e As PaintEventArgs)
        Dim pnl As Panel = CType(sender, Panel)
        e.Graphics.DrawRectangle(New Pen(Me.ForeColor, 1), 0, 0, pnl.Width - 1, pnl.Height - 1)
    End Sub
    Private Sub lblDrawBorder(sender As Object, e As PaintEventArgs)
        Dim lbl As Label = CType(sender, Label)
        e.Graphics.DrawRectangle(New Pen(Color.Black, 1), 0, 0, lbl.Width - 1, lbl.Height - 1)
        e.Graphics.DrawRectangle(New Pen(Color.White, 1), 1, 1, lbl.Width - 3, lbl.Height - 3)
    End Sub


    Public Sub updateSettingsPanel()
        'pnlSettings.Invalidate()
        ' LCX are the Label comboboxes I use for custom checkbox drawing. This updates how to draw them

        ' Set art to work with graphics, pen for single lines, and brushes for filled rectangles
        Dim art As Graphics
        Dim pn1 As Pen = New Pen(Me.ForeColor, 1)
        Dim brBack As Brush = New SolidBrush(Me.BackColor)
        Dim brFore As Brush = New SolidBrush(Me.ForeColor)
        Dim isTrue As Boolean = False

        For Each lbl As Label In pnlSettings.Controls.OfType(Of Label)()
            With lbl
                Select Case Mid(.Name, 1, 3)
                    Case "lcx"
                        ' lcx are Labels acting like a Checkbox for settings with just an on/off state
                        Select Case .Name
                            Case lcxLogic.Name
                                isTrue = My.Settings.setLogic
                            Case lcxCowShuffle.Name
                                isTrue = My.Settings.setCow
                            Case lcxScrubShuffle.Name
                                isTrue = My.Settings.setScrub
                            Case lcxBombchuLogic.Name
                                isTrue = My.Settings.setBombchus
                            Case lcxLoTLogic.Name
                                isTrue = My.Settings.setLensOfTruth
                            Case lcxSoALogic.Name
                                isTrue = My.Settings.setStoneOfAgony
                            Case lcxHoverTricks.Name
                                isTrue = My.Settings.setHoverTricks
                            Case lcxOpenForest.Name
                                isTrue = My.Settings.setOpenKF
                            Case lcxOpenFountain.Name
                                isTrue = My.Settings.setOpenZF
                            Case lcxFiTRustedSwitches.Name
                                isTrue = My.Settings.setFiTRustedSwitches
                            Case lcxTunic.Name
                                isTrue = My.Settings.setFewerTunics
                            Case lcxDekuB1Skip.Name
                                isTrue = My.Settings.setDekuB1Skip
                            Case lcxDCStaircase.Name
                                isTrue = My.Settings.setDCStaircase
                            Case lcxDCSpikeJump.Name
                                isTrue = My.Settings.setDCSpikeJump
                            Case lcxDCSlingshotSkips.Name
                                isTrue = My.Settings.setDCSlingshotSkips
                            Case lcxDCLightEyes.Name
                                isTrue = My.Settings.setDCLightEyes
                            Case lcxFoTVines.Name
                                isTrue = My.Settings.setFoTVines
                            Case lcxFoTLedge.Name
                                isTrue = My.Settings.setFoTLedge
                            Case lcxFoTFrame.Name
                                isTrue = My.Settings.setFoTFrame
                            Case lcxFoTBackdoor.Name
                                isTrue = My.Settings.setFoTBackdoor
                            Case lcxFoTMQPuzzle.Name
                                isTrue = My.Settings.setFoTMQPuzzle
                            Case lcxFoTMQTwisted.Name
                                isTrue = My.Settings.setFoTMQTwisted
                            Case lcxFoTMQWell.Name
                                isTrue = My.Settings.setFoTMQWell
                            Case lcxFiTClimb.Name
                                isTrue = My.Settings.setFiTClimb
                            Case lcxFiTScarecrow.Name
                                isTrue = My.Settings.setFiTScarecrow
                            Case lcxFiTMaze.Name
                                isTrue = My.Settings.setFiTMaze
                            Case lcxFiTMQClimb.Name
                                isTrue = My.Settings.setFiTMQClimb
                            Case lcxFiTMQMazeHovers.Name
                                isTrue = My.Settings.setFiTMQMazeHovers
                            Case lcxFiTMQMazeJump.Name
                                isTrue = My.Settings.setFiTMQMazeJump
                            Case lcxWaTTorch.Name
                                isTrue = My.Settings.setWaTTorch
                            Case lcxWaTDrgnSwitch.Name
                                isTrue = My.Settings.setWaTDrgnSwitch
                            Case lcxWaTDrgnDive.Name
                                isTrue = My.Settings.setWaTDragonDive
                            Case lcxWaTBKR.Name
                                isTrue = My.Settings.setWaTBKR
                            Case lcxWaTCrackNothing.Name
                                isTrue = My.Settings.setWaTCrackNothing
                            Case lcxWaTCrackHovers.Name
                                isTrue = My.Settings.setWaTCrackHovers
                            Case lcxSpTLensless.Name
                                isTrue = My.Settings.setSpTLensless
                            Case lcxSpTWall.Name
                                isTrue = My.Settings.setSpTWall
                            Case lcxSpTBombchu.Name
                                isTrue = My.Settings.setSpTBombchu
                            Case lcxSpTMQLowAdult.Name
                                isTrue = My.Settings.setSpTMQLowAdult
                            Case lcxSpTMQSunRoom.Name
                                isTrue = My.Settings.setSpTMQSunRoom
                            Case lcxShTLensless.Name
                                isTrue = My.Settings.setShTLensless
                            Case lcxShTPlatform.Name
                                isTrue = My.Settings.setShTPlatform
                            Case lcxShTUmbrella.Name
                                isTrue = My.Settings.setShTUmbrella
                            Case lcxShTStatue.Name
                                isTrue = My.Settings.setShTStatue
                            Case lcxShTMQGap.Name
                                isTrue = My.Settings.setShTMQGap
                            Case lcxShTMQPit.Name
                                isTrue = My.Settings.setShTMQPit
                            Case lcxBotWLensless.Name
                                isTrue = My.Settings.setBotWLensless
                            Case lcxBotWDeadHand.Name
                                isTrue = My.Settings.setBotWDeadHand
                            Case lcxBotWMQPits.Name
                                isTrue = My.Settings.setBotWMQPits
                            Case lcxNavi.Name
                                isTrue = My.Settings.setNavi
                            Case lcxBeans.Name
                                isTrue = My.Settings.setBeans
                            Case lcxSaleman.Name
                                isTrue = My.Settings.setSalesman
                            Case lcxWastelandBackstep.Name
                                isTrue = My.Settings.setWastelandBackstep
                            Case lcxWastelandLensless.Name
                                isTrue = My.Settings.setWastelandLensless
                            Case lcxWastelandReverse.Name
                                isTrue = My.Settings.setWastelandReverse
                            Case lcxIGCMQRustedSwitches.Name
                                isTrue = My.Settings.setIGCMQRustedSwitches
                            Case lcxIGCLensless.Name
                                isTrue = My.Settings.setIGCLensless
                            Case lcxHideSpoiler.Name
                                isTrue = My.Settings.setHideSpoiler
                            Case lcxCucco.Name
                                isTrue = My.Settings.setCucco
                            Case lcxShowMap.Name
                                isTrue = My.Settings.setMap
                            Case lcxExpand.Name
                                isTrue = My.Settings.setExpand
                            Case lcxShortForm.Name
                                isTrue = My.Settings.setShortForm
                            Case lcxHideQuests.Name
                                isTrue = My.Settings.setHideQuests
                            Case lcxGTGLensless.Name
                                isTrue = My.Settings.setGTGLensless
                                'Case lcxxx.Name
                                'isTrue = My.Settings.setxx
                            Case Else
                                isTrue = False
                        End Select
                        ' Set up its graphics and draw the basic outline square
                        art = .CreateGraphics
                        Dim startY As Integer = CInt(Math.Floor(.Height / 2) - 7)
                        art.DrawRectangle(pn1, 2, startY, 12, 12)

                        If isTrue Then
                            ' If true, draw a filled square
                            art.FillRectangle(brFore, 4, startY + 2, 9, 9)
                        Else
                            ' If not true, empty out the square
                            art.FillRectangle(brBack, 4, startY + 2, 9, 9)
                        End If

                        'Case "ltb"
                        ' ltb are Labels acting like a Track Bar for settings with a variety of options
                        'Dim p As New Pen(Brushes.Black, 4)
                        'p.StartCap = Drawing2D.LineCap.ArrowAnchor



                End Select
            End With
        Next

    End Sub
    Private Sub subMenu(ByVal strTheme As String)
        Dim valTheme As Byte = 0

        Select Case LCase(strTheme)
            Case "default"
                valTheme = 0
            Case "dark mode"
                valTheme = 1
            Case "lavender"
                valTheme = 2
            Case "midnight"
                valTheme = 3
            Case "hotdog stand"
                valTheme = 4
            Case "the hub"
                valTheme = 5
            Case Else
                rtbOutputLeft.Text = "-- Theme Error: " & strTheme
                Exit Sub
        End Select

        subMenuCheck(valTheme)
        changeTheme(valTheme)
    End Sub
    Private Sub subMenuCheck(ByVal valTheme As Byte)
        DefaultToolStripMenuItem.Checked = False
        DarkModeToolStripMenuItem.Checked = False
        LavenderToolStripMenuItem.Checked = False
        MidnightToolStripMenuItem.Checked = False
        HotdogStandToolStripMenuItem.Checked = False
        TheHubToolStripMenuItem.Checked = False
        Select Case valTheme
            Case 0
                DefaultToolStripMenuItem.Checked = True
            Case 1
                DarkModeToolStripMenuItem.Checked = True
            Case 2
                LavenderToolStripMenuItem.Checked = True
            Case 3
                MidnightToolStripMenuItem.Checked = True
            Case 4
                HotdogStandToolStripMenuItem.Checked = True
            Case 5
                TheHubToolStripMenuItem.Checked = True
            Case Else
                rtbOutputLeft.Text = "-- Check Theme Error: " & valTheme.ToString
        End Select
    End Sub
    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        subMenu(sender.ToString)
    End Sub
    Private Sub DarkModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DarkModeToolStripMenuItem.Click
        subMenu(sender.ToString)
    End Sub
    Private Sub LavenderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LavenderToolStripMenuItem.Click
        subMenu(sender.ToString)
    End Sub
    Private Sub MidnightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MidnightToolStripMenuItem.Click
        subMenu(sender.ToString)
    End Sub
    Private Sub HotdogStandToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HotdogStandToolStripMenuItem.Click
        subMenu(sender.ToString)
    End Sub
    Private Sub TheHubToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TheHubToolStripMenuItem.Click
        subMenu(sender.ToString)
    End Sub

    Private Function getKeyLogic(ByVal name As String, ByVal area As String) As String
        getKeyLogic = String.Empty
        ' Strip any extras we added to the display name to get the original key name
        name = Replace(name, "GS:", "")
        name = Replace(name, "Cow:", "")
        name = Replace(name, "Scrub:", "")
        name = Replace(name, "Shopsanity:", "")
        name = Replace(name, "(Forced)", "")
        name = Trim(name)
        ' Start with an empty string

        If Not Mid(area, 1, 3) = "DUN" Then
            ' Start with the non-dungeon checks
            For Each key In aKeys.Where(Function(k As keyCheck) k.name.Equals(name))
                With key
                    If .area = area Then
                        .forced = Not .forced
                        getKeyLogic = .logic
                        Exit Function
                    End If
                End With
            Next
        Else
            ' Convert the dungeon into a number for which array to look unto
            Dim dunNum As Byte = CByte(area.Replace("DUN", ""))
            For i = 0 To aKeysDungeons(dunNum).Length - 1
                With aKeysDungeons(dunNum)(i)
                    If .name = name Then
                        .forced = Not .forced
                        getKeyLogic = .logic
                        Exit Function
                    End If
                End With
            Next
        End If
    End Function
    Private Function flipKeyForced(ByVal name As String, ByVal area As String, Optional setAs As Byte = 0) As Boolean
        flipKeyForced = False
        ' Strip any extras we added to the display name to get the original key name
        name = Replace(name, "GS:", "")
        name = Replace(name, "Cow:", "")
        name = Replace(name, "Scrub:", "")
        name = Replace(name, "Shopsanity:", "")
        name = Replace(name, "(Forced)", "")
        name = Trim(name)

        If Not Mid(area, 1, 3) = "DUN" Then
            ' Start with the non-dungeon checks

            ' Need to work on the doubled areas for visual map
            Dim area2 As String = area
            Dim area3 As String = area
            If My.Settings.setMap Then
                Select Case area
                    Case "MK"
                        area2 = "QM"
                        area3 = "TT"
                    Case "HC"
                        area2 = "OGC"
                    Case "KV"
                        area2 = "QGS"
                End Select
            End If

            For Each key In aKeys.Where(Function(k As keyCheck) k.name.Equals(name))
                With key
                    If .area = area Or .area = area2 Or .area = area3 Then
                        Select Case setAs
                            Case 0  ' Default: Flips
                                .forced = Not .forced
                            Case 1  ' Unforces
                                .forced = False
                            Case 2  ' Forces
                                .forced = True
                        End Select

                        flipKeyForced = True
                        Exit Function
                    End If
                End With
            Next
        Else
            ' Convert the dungeon into a number for which array to look unto
            Dim dunNum As Byte = CByte(area.Replace("DUN", ""))
            For i = 0 To aKeysDungeons(dunNum).Length - 1
                With aKeysDungeons(dunNum)(i)
                    If .name = name Then
                        .forced = Not .forced
                        flipKeyForced = True
                        Exit Function
                    End If
                End With
            Next
        End If
    End Function

    Private Sub ScanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScanToolStripMenuItem.Click
        goScan(False)
    End Sub
    Private Sub AutoScanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoScanToolStripMenuItem.Click
        AutoScanToolStripMenuItem.Text = "Stop"
        goScan(True)
    End Sub
    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        pnlDungeonItems.BackgroundImage = My.Resources.backgroundDungeonItems
        stopScanning()
        pbxPoH.Image = My.Resources.poh0
        pbxMap.Image = My.Resources.mapBlank
        rtbOutputLeft.Clear()
        rtbOutputRight.Clear()
        lastRoomScan = 0
        populateLocations()
    End Sub
    Private Sub ExitScanToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Private Sub MiniMapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MiniMapToolStripMenuItem.Click
        pnlER.Visible = Not pnlER.Visible
    End Sub
    Private Sub ShowSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowSettingsToolStripMenuItem.Click
        showSetting = Not showSetting
        updateShowSettings()
        My.Settings.setFirstTime = False
    End Sub

    Private Sub changeAddedText(ByVal thisOne As Byte, Optional reverse As Boolean = False, Optional justRefresh As Boolean = False)
        refreshQuestItemImages(thisOne)
        With aoQuestItems(thisOne)
            ' We need to reset the image to clear off any drawn letters
            ' This checks if the reward was set to collected, if so we will use the coloured image, if not we use the greyed out image
            If aQuestRewardsCollected(thisOne) Then
                .Image = aoQuestItemImages(thisOne)
            Else
                .Image = aoQuestItemImagesEmpty(thisOne)
            End If

            ' Need to store value to an integer for a bit to work math on it that may make it go negative, which byte cannot handle
            Dim tempValue As Integer = aQuestRewardsText(thisOne)

            ' Sometimes we just need a refresh
            If Not justRefresh Then
                ' If reverse, then backstep one dungeon. If not, then forward step
                If reverse Then
                    dec(tempValue)
                Else
                    inc(tempValue)
                End If
            End If

            ' Make sure the value stays between 0 and 9
            If tempValue < 0 Then tempValue = 9
            If tempValue > 9 Then tempValue = 0

            ' And put it back when done
            aQuestRewardsText(thisOne) = CByte(tempValue)

            ' Get the new dungeon letters to draw
            Dim outputText As String = aDungeonLetters(tempValue)

            Dim xPos As Integer = 0
            Select Case outputText.Length
                Case 2
                    xPos = 11
                Case 3
                    xPos = 5
                Case 4
                    xPos = 0
            End Select

            ' Font for dungeon letters
            Dim fontDungeon = New Font("Lucida Console", 18, FontStyle.Regular, GraphicsUnit.Pixel)

            ' Draw the value over the lower right of the gold skulltula picturebox, first in black to give it some definition, then in white
            Graphics.FromImage(.Image).DrawString(outputText, fontDungeon, New SolidBrush(Color.Black), xPos - 1, 31)
            Graphics.FromImage(.Image).DrawString(outputText, fontDungeon, New SolidBrush(Color.White), xPos, 32)
        End With
    End Sub
    Private Sub scanDungeonRewards()
        ' If either are blank, abort
        If aAddresses(12) = 0 Or aAddresses(13) = 0 Then Return

        ' Start and finish range for reading the rewards
        Dim lowScan As Byte = 3
        Dim highScan As Byte = 2
        ' The dungeon rewards in string form
        Dim rewards As String = String.Empty
        Dim useCompass As Boolean = False

        ' Compass Info On
        If goRead(aAddresses(13), 1) = 1 Then
            useCompass = True
        Else
            ' Compass Info Off and Hint Info On
            If goRead(aAddresses(12), 1) = 1 Then
                Dim pedestal As Byte = CByte(goRead(&H11B4FC, 1))
                If pedestal = 1 Or pedestal = 3 Then highScan = 8
                If pedestal = 2 Or pedestal = 3 Then lowScan = 0
            End If
        End If

        ' Check for OOTRs odd byte check
        If Not aAddresses(17) = 0 Then
            If goRead(aAddresses(17), 1) = 0 Then
                ' If it is 0, then no info, and revert back to paradox range
                lowScan = 3
                highScan = 2
            End If
        End If

        Dim tempValue As Byte = 0
        Dim newValue As Byte = 0
        Dim findFree As String = "012345678"
        Dim iNew As Byte = 0

        rewards = Hex(goRead(aAddresses(9) + 4))
        fixHex(rewards)
        rewards = Hex(goRead(aAddresses(9))) & rewards
        fixHex(rewards, 16)

        For i = 0 To 7
            ' Store the dungeon reward value
            tempValue = CByte(Mid(rewards, (i * 2) + 1, 2))
            findFree = findFree.Replace(tempValue.ToString, "")
            iNew = CByte(i + 2)

            ' Convert the value to the autotrackers picturebox
            Select Case tempValue
                Case 0 To 2
                    ' Stones need to be converted to 18 to 20
                    newValue = CByte(tempValue + 18)
                Case 3 To 8
                    ' Medallions need to be converted to 0 to 4
                    newValue = CByte(tempValue - 3)
                    ' Flip the Spirit and Shadow Temple
            End Select

            If useCompass Then
                ' For compass info, if compass is not visible, set the reward to 0
                If Not aoCompasses(i).Visible Then iNew = 0
                aQuestRewardsText(newValue) = CByte(iNew)
                changeAddedText(newValue, , True)
            Else
                If tempValue < lowScan Or tempValue > highScan Then iNew = 0
                aQuestRewardsText(newValue) = CByte(iNew)
                changeAddedText(newValue, , True)
            End If
        Next

        ' Find the free dungeon reward
        If findFree.Length = 1 Then
            Dim freeValue = CByte(findFree)
            Select Case freeValue
                Case 0 To 2
                    incB(freeValue, 18)
                Case Else
                    decB(freeValue, 3)
            End Select
            aQuestRewardsText(freeValue) = 1
            changeAddedText(freeValue, , True)
        End If
    End Sub
    Private Function isWarpShuffle() As Boolean
        ' Checks to see if any warp songs have had optional settings. If so, then return true
        isWarpShuffle = False
        For i As Byte = 6 To 11
            If Not aQuestRewardsText(i) = 0 Then Return True
        Next
    End Function
    Private Sub getPlayerName()
        ' Grab the first 4 characters
        Dim sPlayerName1 As String = Hex(goRead(If(isSoH, soh.SAV(&H1E), &H11A5F4)))
        ' Grab the last 4 characters
        Dim sPlayerName2 As String = Hex(goRead(If(isSoH, soh.SAV(&H22), &H11A5F8)))

        fixHex(sPlayerName1) ' pad to 8 chars
        fixHex(sPlayerName2)

        ' Combine into one
        sPlayerName1 &= sPlayerName2

        If isSoH Then endianFlip(sPlayerName1)

        ' Check with the stored player name, if it is the same, exit sub and do not redraw it
        If playerName = sPlayerName1 Then Exit Sub
        playerName = sPlayerName1

        ' Will re-use second half of name varaible
        sPlayerName2 = String.Empty
        For i = 0 To 7
            ' Steps through the stored hex of the player name and convert it into readable characters
            sPlayerName2 = sPlayerName2 & decodeLetter(CByte("&H" & Mid(sPlayerName1, (i * 2) + 1, 2)))
        Next

        With pnlDungeonItems
            ' Restore the background image and prepare the font for the player name
            .BackgroundImage = My.Resources.backgroundDungeonItems
            Dim fontName As Font = New Font("Lucida Console", 40, FontStyle.Bold, GraphicsUnit.Pixel)

            ' Draw the value over the lower right of the gold skulltula picturebox, first in black to give it some definition, then in white
            Graphics.FromImage(.BackgroundImage).DrawString(sPlayerName2.TrimEnd, fontName, New SolidBrush(Color.White), -3, 141)
        End With
    End Sub
    Private Function decodeLetter(ByRef valLetter As Byte) As String
        decodeLetter = String.Empty
        Select Case valLetter
            Case 0 To 9
                ' 0 to 9, 48 to 57, so +48, but instead just turn it a string
                decodeLetter = valLetter.ToString
            Case 10 To 35
                ' A-Z for soh
                If isSoH Then
                    decodeLetter = Chr(valLetter + 55)
                End If
            Case 36 To 61
                ' a-z for soh
                If isSoH Then
                    decodeLetter = Chr(valLetter + 61)
                End If
            Case 62
                If isSoH Then
                    decodeLetter = " "
                End If
            Case 63
                If isSoH Then
                    decodeLetter = "-"
                End If
            Case 64
                If isSoH Then
                    decodeLetter = "."
                End If
            Case 171 To 196
                ' A-Z, normally 65 to 90, so -106
                decodeLetter = Chr(valLetter - 106)
            Case 197 To 222
                ' a-z, normally 97-122, so -100
                decodeLetter = Chr(valLetter - 100)
            Case 223
                decodeLetter = " "
            Case 228
                decodeLetter = "-"
            Case 234
                decodeLetter = "."
        End Select
    End Function
    Private Sub getWarps()
        If isSoH Then   ' SoH will load the default exit codes
            aWarps(0) = "5F4"
            aWarps(1) = "0BB"
            aWarps(2) = "600"
            aWarps(3) = "4F6"
            aWarps(4) = "604"
            aWarps(5) = "1F1"
            aWarps(6) = "568"
            aWarps(7) = "5F4"
            bSpawnWarps = False
            bSongWarps = False
        Else
            Dim arrOffsets() As Integer = {&H903D0, &H903E0, &H3AB22E, &H3AB22C, &H3AB232, &H3AB230, &H3AB236, &H3AB234}
            Dim sRead As String = String.Empty
            ' If the Minuet and Bolero warps are 0, then player is in a load screen or pause menu. Abort.
            If goRead(arrOffsets(3)) = 0 Then Exit Sub

            For i = 0 To aReachA.Length - 1
                aReachA(i) = False
                aReachY(i) = False
            Next

            For i = 0 To 7
                sRead = Hex(goRead(arrOffsets(i), 15))
                If Not sRead = "0" Then
                    fixHex(sRead, 3)
                    aWarps(i) = sRead
                End If
            Next
            Dim iStart As Byte = 0
            Dim iEnd As Byte = 7
            If aWarps(0) & aWarps(1) = "5F40BB" Then
                ' If both spawn warps are vanilla, then skip having to check them and set displaying them to false
                bSpawnWarps = False
            Else
                bSpawnWarps = True
            End If
            If aWarps(2) & aWarps(3) & aWarps(4) & aWarps(5) & aWarps(6) & aWarps(7) = "" Then
                ' If all warps are vanilla, then skip having to check them and set displaying them to false
                bSongWarps = False
            Else
                bSongWarps = True
            End If
        End If

        For i As Byte = 0 To 7
            Select Case aWarps(i)
                Case "09C", "0BB", "0C1", "0C9", "211", "266", "26A", "272", "286", "33C", "433", "437", "443", "447"
                    ' KF Main
                    aWarps(i) = "KF"
                    addReach(0, i)
                Case "20D"
                    ' KF Trapped
                    aWarps(i) = "KF"
                    addReach(8, i)
                Case "11E", "4D6", "4DA"
                    ' LW Front
                    aWarps(i) = "LW"
                    addReach(2, i)
                Case "1A9"
                    ' LW Behind Mido
                    aWarps(i) = "LW"
                    addReach(3, i)
                Case "4DE", "5E0"
                    ' LW Bridge
                    aWarps(i) = "LW"
                    addReach(4, i)
                Case "0FC", "600"
                    ' SFM Main
                    aWarps(i) = "SFM"
                    addReach(5, i)
                Case "17D", "181", "185", "189", "18D", "1F9", "1FD", "27E"
                    ' HF
                    aWarps(i) = "HF"
                    addReach(7, i)
                Case "04F", "157", "2F9", "378", "42F", "5D0", "5D4"
                    ' LLR
                    aWarps(i) = "LLR"
                    addReach(9, i)
                Case "033", "063", "067", "07E", "0B1", "16D", "1CD", "1D1", "1D5", "25A", "25E", "26E", "276", "2A2", "388",
                        "3B8", "3BC", "3C0", "43B", "507", "528", "52C", "530"
                    ' MK
                    aWarps(i) = "MK"
                    addReach(10, i)
                Case "053", "171", "472", "5F4"
                    ' ToT Front
                    aWarps(i) = "ToT"
                    addReach(11, i)
                Case "138"
                    ' HC
                    aWarps(i) = "HC"
                    addReach(13, i)
                Case "03B", "072", "0B7", "0DB", "195", "201", "2FD", "345", "349", "34D", "351", "384", "39C", "3EC", "44B",
                        "453", "463", "4EE", "4FF", "550", "5C8", "5DC"
                    ' KV Main
                    aWarps(i) = "KV"
                    addReach(15, i)
                Case "554"
                    ' KV Rooftops
                    aWarps(i) = "KV"
                    addReach(16, i)
                Case "191"
                    ' KV Behind Gate
                    aWarps(i) = "KV"
                    addReach(17, i)
                Case "0E4", "30D", "355"
                    ' GY Main
                    aWarps(i) = "GY"
                    addReach(18, i)
                Case "568"
                    ' GY Upper
                    aWarps(i) = "GY"
                    addReach(19, i)
                Case "13D", "1B9"
                    ' DMT Lower
                    aWarps(i) = "DMT"
                    addReach(20, i)
                Case "1BD", "315", "45B"
                    ' DMT Upper 
                    aWarps(i) = "DMT"
                    addReach(21, i)
                Case "147"
                    ' DMC Upper Local
                    aWarps(i) = "DMC"
                    addReach(22, i)
                Case "246", "482", "4BE"
                    ' DMC Lower Nearby
                    aWarps(i) = "DMC"
                    addReach(24, i)
                Case "4F6"
                    ' DMC Central Local
                    aWarps(i) = "DMC"
                    addReach(27, i)
                Case "14D", "3FC"
                    ' GC Main
                    aWarps(i) = "GC"
                    addReach(29, i)
                Case "4E2"
                    ' GC Shortcut
                    aWarps(i) = "GC"
                    addReach(30, i)
                Case "1C1"
                    ' GC Darunia
                    aWarps(i) = "GC"
                    addReach(31, i)
                Case "37C"
                    ' GC Shoppe
                    aWarps(i) = "GC"
                    addReach(32, i)
                Case "0EA"
                    ' ZR Front
                    aWarps(i) = "ZR"
                    addReach(34, i)
                Case "199", "1DD"
                    ' ZR Main
                    aWarps(i) = "ZR"
                    addReach(35, i)
                Case "19D"
                    ' ZR Behind Waterfall
                    aWarps(i) = "ZR"
                    addReach(36, i)
                Case "108", "153", "328", "3C4"
                    ' ZD Main
                    aWarps(i) = "ZD"
                    addReach(37, i)
                Case "1A1"
                    ' ZD Behind King
                    aWarps(i) = "ZD"
                    addReach(38, i)
                Case "380"
                    aWarps(i) = "ZD"
                    addReach(39, i)
                Case "225", "371", "394"
                    ' ZF Main
                    aWarps(i) = "ZF"
                    addReach(40, i)
                Case "043", "102", "219", "3CC", "560", "604"
                    ' LH Main
                    aWarps(i) = "LH"
                    addReach(42, i)
                Case "309", "45F"
                    ' LH Fishing Ledge
                    aWarps(i) = "LH"
                    addReach(43, i)
                Case "117"
                    ' GV Hyrule Side
                    aWarps(i) = "GV"
                    addReach(44, i)
                Case "22D", "3A0", "3D0"
                    ' GV Gerudo Side
                    aWarps(i) = "GV"
                    addReach(45, i)
                Case "129"
                    ' GF Main
                    aWarps(i) = "GF"
                    addReach(46, i)
                Case "3AC"
                    ' GF Behind Gate
                    aWarps(i) = "GF"
                    addReach(47, i)
                Case "130"
                    ' HW Gerudo Side
                    aWarps(i) = "HW"
                    addReach(48, i)
                Case "365"
                    ' HW Colossus Side
                    aWarps(i) = "HW"
                    addReach(49, i)
                Case "123", "1F1", "57C", "588"
                    ' DC Main
                    aWarps(i) = "DC"
                    addReach(50, i)
                Case "340"
                    ' Schrodinger's Castle
                    aWarps(i) = "CASL"
                    addReach(52, i)
                Case "4C2"
                    ' Schrodinger's Fairy Adult
                    aWarps(i) = "CASL"
                    addReach(53, i)
                Case "578"
                    ' Schrodinger's Fairy Young
                    aWarps(i) = "CASL"
                    addReach(54, i)
                Case Else
                    ' Unknown entry
                    aWarps(i) = aWarps(i) & "?"
            End Select
        Next
    End Sub
    Private Sub addReach(ByVal warp As Byte, ByVal type As Byte)
        Select Case type
            Case 0
                If canAdult Then
                    aReachA(warp) = True
                    reachAreas(warp, True)
                End If
            Case 1
                If canYoung Then
                    aReachY(warp) = True
                    reachAreas(warp, False)
                End If
            Case Else
                If aQuestItems(type + 4) And allItems.Contains("h") Then
                    If canAdult Then
                        aReachA(warp) = True
                        reachAreas(warp, True)
                    End If
                    If canYoung Then
                        aReachY(warp) = True
                        reachAreas(warp, False)
                    End If
                End If
        End Select
    End Sub
    Private Function canWarpReach(ByVal warp As Byte) As Boolean
        canWarpReach = False
        Select Case warp
            Case 0
                If canAdult Then Return True
            Case 1
                If canYoung Then Return True
            Case 2 To 7
                If aQuestItems(warp + 4) Then Return True
        End Select
    End Function
    Private Sub displaySpawns(ByVal spawn As Byte)
        ' Font for the letters
        Dim fontDungeon = New Font("Lucida Console", 24, FontStyle.Regular, GraphicsUnit.Pixel)
        Dim text As String = aWarps(spawn)
        Dim xPos As Integer = 0

        Select Case spawn
            Case 0
                With pbxSpawnAdult
                    .Image = My.Resources.spawnLocations
                    'If isAdult Or checkLoc("6405") Or Not My.Settings.setHideSpoiler Then Graphics.FromImage(.Image).DrawString("Adult: " & text, fontDungeon, New SolidBrush(Color.White), 0, 0)
                    If isAdult Or aReachA(12) Then Graphics.FromImage(.Image).DrawString("Adult: " & text, fontDungeon, New SolidBrush(Color.White), 0, 0)
                End With
            Case 1
                With pbxSpawnYoung
                    .Image = My.Resources.spawnLocations
                    'If Not isAdult Or checkLoc("6405") Or Not My.Settings.setHideSpoiler Then Graphics.FromImage(.Image).DrawString("Young: " & text, fontDungeon, New SolidBrush(Color.White), 0, 0)
                    If Not isAdult Or aReachY(12) Then Graphics.FromImage(.Image).DrawString("Young: " & text, fontDungeon, New SolidBrush(Color.White), 0, 0)
                End With
        End Select
    End Sub
    Private Sub displayWarps(ByVal warp As Byte)
        If Not item("ocarina") Then Exit Sub
        Dim text As String = aWarps(warp)
        incB(warp, 4)
        With aoQuestItems(warp)
            ' We need to reset the image to clear off any drawn letters
            ' This checks if the reward was set to collected, if so we will use the coloured image, if not we use the greyed out image
            If aQuestRewardsCollected(warp) Then
                refreshQuestItemImages(warp)
                .Image = aoQuestItemImages(warp)
            Else
                ' Exit sub if the song is not collected yet
                Exit Sub
            End If

            Dim xPos As Integer = 0
            Select Case text.Length
                Case 2
                    xPos = 11
                Case 3
                    xPos = 5
                Case 4
                    xPos = 0
                Case Else
                    ' Abort if not within the letter range
                    Exit Sub
            End Select

            ' Font for the letters
            Dim fontDungeon = New Font("Lucida Console", 18, FontStyle.Regular, GraphicsUnit.Pixel)

            ' Draw letters over with shadow first, then in white
            Graphics.FromImage(.Image).DrawString(text, fontDungeon, New SolidBrush(Color.Black), xPos - 1, 31)
            Graphics.FromImage(.Image).DrawString(text, fontDungeon, New SolidBrush(Color.White), xPos, 32)
        End With
    End Sub
    Private Sub getRandoVer()

        randoVer = String.Empty


        ' Only using C as there are no duplicates of it yet. Should it occur, can uncomment one and start using it

        'Dim readA As String = Hex(goRead(&H400000))
        'Dim readB As String = Hex(goRead(&H400004))
        Dim readC As String = Hex(goRead(&H400008))
        'Dim readD As String = Hex(goRead(&H40000C))
        'fixHex(readA)
        'fixHex(readB)
        fixHex(readC)
        'fixHex(readD)



        'Addr:              80400000 80400004 80400008 8040000C
        'AP1:               80400020 80400844 80409FD4 00000000
        'AP2:               80400020 80400844 8040A334 00000000
        'OOTR 6.2:          80400020 80400834 8040A474 00000000
        'OOTR 6.2.72:       80400020 80400834 8040AA7C 80400CD0
        'Roman 6.2.43:      80400020 80400834 8040ACC4 80400CD0
        'ROMAN 6.2.72-R2:   80400020 80400834 8040B11C 80400CD0

        Select Case readC
            Case "80409FD4"
                ' AP 0.3.1
                aAddresses(0) = &H40B220    ' MQs Addr
                aAddresses(1) = &H400CC4    ' LAC1 Addr
                aAddresses(2) = &H400CD0    ' LAC2 Addr
                aAddresses(3) = &H40B1B0    ' TH Addr
                aAddresses(4) = &H400CC0    ' RB1 Addr
                aAddresses(5) = &H400CD2    ' RB2 Addr
                aAddresses(6) = 1           ' Shopsanity Style
                aAddresses(7) = &H400CBC    ' Bombchu's in Logic
                aAddresses(8) = &H400CD4    ' Cow Shuffle
            Case "8040A334"
                ' AP 0.3.2
                aAddresses(0) = &H40B5C8    ' MQs Addr
                aAddresses(1) = &H400CC4    ' LAC1 Addr
                aAddresses(2) = &H400CD0    ' LAC2 Addr
                aAddresses(3) = &H40B550    ' TH Addr
                aAddresses(4) = &H400CC0    ' RB1 Addr
                aAddresses(5) = &H400CD2    ' RB2 Addr
                aAddresses(6) = 1           ' Shopsanity
                aAddresses(7) = &H400CBC    ' Bombchu's in Logic
                aAddresses(8) = &H400CD4    ' Cow Shuffle
                aAddresses(9) = &H40A274    ' Dungeon Rewards
                aAddresses(10) = &H400CE8   ' Scrub Shuffle
                aAddresses(11) = &H400CE9   ' Key Mode
                aAddresses(12) = &H400CEA   ' Pedestal DRs
                aAddresses(13) = &H400CEB   ' Compass DRs
                aAddresses(14) = &H400CED   ' Big Poe Goal
                aAddresses(15) = &H400CEF   ' KF | 0: Open | 1: Closed Deku | 2: Closed
                aAddresses(16) = &H400CEE   ' ZF | 0: Open | 1: Adult | 2: Closed
                aAddresses(18) = &H400CDD   ' Overworld ER
                aAddresses(19) = &H400CDE   ' Dungeon ER
            Case "8040A474"
                ' OOTR 6.2
                aAddresses(0) = &H40B6E0    ' MQs Addr
                aAddresses(1) = &H400CB4    ' LAC1 Addr
                aAddresses(2) = &H400CC0    ' LAC2 Addr
                aAddresses(3) = &H40B668    ' TH Addr
                aAddresses(4) = &H400CB0    ' RB1 Addr
                aAddresses(5) = &H400CC2    ' RB2 Addr
                aAddresses(7) = &H400CAC    ' Bombchu's in Logic
                aAddresses(8) = &H400CC4    ' Cow Shuffle
                'aAddresses(9) = &H40A3B4    ' Dungeon Rewards
                aAddresses(18) = &H400CCD   ' Overworld ER
                aAddresses(19) = &H400CCE   ' Dungeon ER
            Case "8040AA7C"
                ' OOTR 6.2.72
                aAddresses(0) = &H400CF8    ' MQs Addr
                aAddresses(1) = &H400CB4    ' LAC1 Addr
                aAddresses(2) = &H400CC0    ' LAC2 Addr
                aAddresses(3) = &H40BD38    ' TH Addr
                aAddresses(4) = &H400CB0    ' RB1 Addr
                aAddresses(5) = &H400CC2    ' RB2 Addr
                aAddresses(7) = &H400CAC    ' Bombchu's in Logic
                aAddresses(8) = &H400CC4    ' Cow Shuffle
                aAddresses(9) = &H400CEC    ' Dungeon Rewards
                'aAddresses(10) = &H   ' Scrub Shuffle
                'aAddresses(11) = &H   ' Key Mode
                aAddresses(12) = &H400CE8   ' Pedestal DRs
                aAddresses(13) = &H400CE4   ' Compass DRs
                'aAddresses(14) = &H   ' Big Poe Goal
                'aAddresses(15) = &H   ' KF | 0: Open | 1: Closed Deku | 2: Closed
                'aAddresses(16) = &H   ' ZF | 0: Closed | 1: Adult | 2: Open
                aAddresses(17) = &H400CE0   ' OOTR Info On/Off
                aAddresses(18) = &H400CCD   ' Overworld ER
                aAddresses(19) = &H400CCE   ' Dungeon ER
            Case "8040ACC4"
                ' ROMAN 6.2.43
                aAddresses(0) = &H400CF8    ' MQs Addr
                aAddresses(1) = &H400CB4    ' LAC1 Addr
                aAddresses(2) = &H400CC0    ' LAC2 Addr
                aAddresses(3) = &H40BF80    ' TH Addr
                aAddresses(4) = &H400CB0    ' RB1 Addr
                aAddresses(5) = &H400CC2    ' RB2 Addr
                aAddresses(7) = &H400CAC    ' Bombchu's in Logic
                aAddresses(8) = &H400CC4    ' Cow Shuffle
                aAddresses(18) = &H400CCD   ' Overworld ER
                aAddresses(19) = &H400CCE   ' Dungeon ER
            Case "8040B11C"
                ' ROMAN 6.2.72-R2
                aAddresses(0) = &H400CF8    ' MQs Addr
                aAddresses(1) = &H400CB4    ' LAC1 Addr
                aAddresses(2) = &H400CC0    ' LAC2 Addr
                aAddresses(3) = &H40C3D8    ' TH Addr
                aAddresses(4) = &H400CB0    ' RB1 Addr
                aAddresses(5) = &H400CC2    ' RB2 Addr
                aAddresses(7) = &H400CAC    ' Bombchu's in Logic
                aAddresses(8) = &H400CC4    ' Cow Shuffle
                aAddresses(18) = &H400CCD   ' Overworld ER
                aAddresses(19) = &H400CCE   ' Dungeon ER
        End Select
    End Sub
    Private Sub getRainbowBridge()
        ' If Not detected, then set to default
        If aAddresses(4) = 0 Then
            rainbowBridge(0) = 4
            Exit Sub
        End If
        ' Rainbow Bridge Condition
        rainbowBridge(0) = CByte(goRead(aAddresses(4), 1))
        ' Rainbow Bridge Condition Count
        rainbowBridge(1) = CByte(goRead(aAddresses(5), 1))
    End Sub
    Private Sub getER()
        'If isSoH Then Exit Sub ' soh 3.0.0 has no entrance rando

        iER = 0
        ' Overworld ER check
        If Not aAddresses(18) = 0 Then
            If goRead(aAddresses(18), 1) = 1 Then incB(iER)
        End If
        ' Dungeon ER check
        If Not aAddresses(19) = 0 Then
            If goRead(aAddresses(19), 1) = 1 Then incB(iER, 2)
        End If
        If Not iER = iOldER Then
            ' If a change in ER is detected (basically first scan), clear the appropriate exits
            If iER Mod 2 = 1 Then clearArrayExitsOverworld()
            If iER > 1 Then clearArrayExitsDungeons()
            iOldER = iER
        End If
        'iER = 3 ' Debug REMOVE
        scanER()
    End Sub
    Private Sub pnlSettings_Paint(sender As Object, e As PaintEventArgs) Handles pnlSettings.Paint
        updateSettingsPanel()
    End Sub
    Private Sub rtbAddLine(ByVal line As String, Optional hideRight As Boolean = False)
        rtbOutputRight.Visible = Not hideRight
        If rtbOutputLeft.Lines.Count <= rtbLines Then
            rtbOutputLeft.AppendText(vbCrLf & line)
        ElseIf rtbOutputRight.Lines.Count <= rtbLines Then
            rtbOutputRight.AppendText(vbCrLf & line)
        End If
    End Sub
    Private Sub changeSong(ByVal area As String)
        Dim song() As String = {"", "", "", ""}
        Select Case area
            Case "KF"
                song = {"00", "Kokiri Forest"}
            Case "LW"
                song = {"01", "Lost Woods"}
            Case "SFM"
                song = {"02", "Sacred Forest Meadow"}
            Case "HF"
                song = {"03", "Hyrule Field"}
            Case "LLR"
                song = {"04", "Lon Lon Ranch"}
            Case "MK"
                song = {"05", "Market"}
            Case "TT"
                song = {"06", "Temple of Time"}
            Case "HC"
                song = {"07", "Hyrule Castle"}
            Case "KV"
                song = {"08", "Kakariko Village"}
            Case "GY"
                song = {"09", "Graveyard"}
            Case "DMT"
                song = {"10", "Death Mountain Trail"}
            Case "DMC"
                song = {"11", "Death Mountain Crater"}
            Case "GC"
                song = {"12", "Goron City"}
            Case "ZR"
                song = {"13", "Zora's River"}
            Case "ZD"
                song = {"14", "Zora's Domain"}
            Case "ZF"
                song = {"15", "Zora's Fountain"}
            Case "LH"
                song = {"16", "Lake Hylia"}
            Case "GV"
                song = {"17", "Gerudo Valley"}
            Case "GF"
                song = {"18", "Gerudo Fotress"}
            Case "HW"
                song = {"19", "Haunted Wasteland"}
            Case "DC"
                song = {"20", "Desert Colossus"}
            Case "OGC"
                song = {"21", "Outside Ganon's Castle"}
            Case "DUN0"
                song = {"22", "Deku Tree"}
            Case "BOSS0"
                song = {"23", "Queen Gohma"}
            Case "DUN1"
                song = {"24", "Dodongo's Cavern"}
            Case "BOSS1"
                song = {"25", "King Dodongo"}
            Case "DUN2"
                song = {"26", "Jabu-Jabu's Belly"}
            Case "BOSS2"
                song = {"27", "Barinade"}
            Case "DUN3"
                song = {"28", "Forest Temple"}
            Case "BOSS3"
                song = {"29", "Phantom Ganon"}
            Case "DUN4"
                song = {"30", "Fire Temple"}
            Case "BOSS4"
                song = {"31", "Volvagia"}
            Case "DUN5"
                song = {"32", "Water Temple"}
            Case "BOSS5"
                song = {"33", "Morpha"}
            Case "DUN6"
                song = {"34", "Spirit Temple"}
            Case "BOSS6"
                song = {"35", "Twinrova"}
            Case "DUN7"
                song = {"36", "Shadow Temple"}
            Case "BOSS7"
                song = {"37", "Bongo Bongo"}
            Case "DUN8"
                song = {"38", "Bottom of the Well"}
            Case "DUN9"
                song = {"39", "Ice Cavern"}
            Case "DUN10"
                song = {"40", "Gerudo Training Ground"}
            Case "DUN11"
                song = {"41", "Ganon's Castle"}
            Case "BOSS8"
                song = {"42", "Ganondorf"}
            Case "ESCAPE"
                song = {"43", "Falling Tower"}
            Case "BOSS8"
                song = {"44", "Ganon"}
        End Select

        Dim file As String = String.Empty
        Dim passed As Boolean = False
        Dim iType As Byte = 0
        For i = 0 To 3
            passed = True
            Select Case i
                Case 0
                    file = song(0)
                Case 1
                    file = area
                    Select Case Mid(area, 1, 3)
                        Case "DUN", "ESC"
                            iType = 2
                        Case "BOS"
                            iType = 1
                    End Select

                Case 2
                    file = song(1)
                Case 3
                    file = song(1).Replace(" ", "").Replace("'", "").Replace("-", "")
            End Select
            Try
                file = Application.StartupPath & "\" & file & ".mp3"
                Dim playSong As New ProcessStartInfo(file)
                playSong.UseShellExecute = True
                Process.Start(playSong)
            Catch ex As Exception
                passed = False
            End Try
            If passed Then Exit For
            If i = 3 Then
                Select Case iType
                    Case 1
                        Try
                            file = Application.StartupPath & "\BOSS.mp3"
                            Dim playSong As New ProcessStartInfo(file)
                            playSong.UseShellExecute = True
                            Process.Start(playSong)
                        Catch ex As Exception
                        End Try
                    Case 2
                        Try
                            file = Application.StartupPath & "\DUN.mp3"
                            Dim playSong As New ProcessStartInfo(file)
                            playSong.UseShellExecute = True
                            Process.Start(playSong)
                        Catch ex As Exception
                        End Try
                End Select
            End If
        Next
    End Sub

    Private Sub zKF_MouseClick(sender As Object, e As MouseEventArgs) Handles zKF.MouseClick
        displayChecks("KF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zLW_MouseClick(sender As Object, e As MouseEventArgs) Handles zLW.MouseClick
        displayChecks("LW", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zSFM_MouseClick(sender As Object, e As MouseEventArgs) Handles zSFM.MouseClick
        displayChecks("SFM", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zHF_MouseClick(sender As Object, e As MouseEventArgs) Handles zHF.MouseClick
        displayChecks("HF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zLLR_MouseClick(sender As Object, e As MouseEventArgs) Handles zLLR.MouseClick
        displayChecks("LLR", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zMK_MouseClick(sender As Object, e As MouseEventArgs) Handles zMK.MouseClick
        displayChecks("MK", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zCastle_MouseClick(sender As Object, e As MouseEventArgs) Handles zCastle.MouseClick
        displayChecks("HC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zKV_MouseClick(sender As Object, e As MouseEventArgs) Handles zKV.MouseClick
        displayChecks("KV", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zGY_MouseClick(sender As Object, e As MouseEventArgs) Handles zGY.MouseClick
        displayChecks("GY", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zDMT_MouseClick(sender As Object, e As MouseEventArgs) Handles zDMT.MouseClick
        displayChecks("DMT", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zDMC_MouseClick(sender As Object, e As MouseEventArgs) Handles zDMC.MouseClick
        displayChecks("DMC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zGC_MouseClick(sender As Object, e As MouseEventArgs) Handles zGC.MouseClick
        displayChecks("GC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zZR_MouseClick(sender As Object, e As MouseEventArgs) Handles zZR.MouseClick
        displayChecks("ZR", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zZD_MouseClick(sender As Object, e As MouseEventArgs) Handles zZD.MouseClick
        displayChecks("ZD", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zZF_MouseClick(sender As Object, e As MouseEventArgs) Handles zZF.MouseClick
        displayChecks("ZF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zLH_MouseClick(sender As Object, e As MouseEventArgs) Handles zLH.MouseClick
        displayChecks("LH", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zGV_MouseClick(sender As Object, e As MouseEventArgs) Handles zGV.MouseClick
        displayChecks("GV", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zGF_MouseClick(sender As Object, e As MouseEventArgs) Handles zGF.MouseClick
        displayChecks("GF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zHW_MouseClick(sender As Object, e As MouseEventArgs) Handles zHW.MouseClick
        displayChecks("HW", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zDC_MouseClick(sender As Object, e As MouseEventArgs) Handles zDC.MouseClick
        displayChecks("DC", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zDT_MouseClick(sender As Object, e As MouseEventArgs) Handles zDT.MouseClick
        displayChecksDungeons(0, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zDDC_MouseClick(sender As Object, e As MouseEventArgs) Handles zDDC.MouseClick
        displayChecksDungeons(1, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zJB_MouseClick(sender As Object, e As MouseEventArgs) Handles zJB.MouseClick
        displayChecksDungeons(2, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zFoT_MouseClick(sender As Object, e As MouseEventArgs) Handles zFoT.MouseClick
        displayChecksDungeons(3, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zFiT_MouseClick(sender As Object, e As MouseEventArgs) Handles zFiT.MouseClick
        displayChecksDungeons(4, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zWaT_MouseClick(sender As Object, e As MouseEventArgs) Handles zWaT.MouseClick
        displayChecksDungeons(5, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zSpT_MouseClick(sender As Object, e As MouseEventArgs) Handles zSpT.MouseClick
        displayChecksDungeons(6, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zShT_MouseClick(sender As Object, e As MouseEventArgs) Handles zShT.MouseClick
        displayChecksDungeons(7, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zBotW_MouseClick(sender As Object, e As MouseEventArgs) Handles zBotW.MouseClick
        displayChecksDungeons(8, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zIC_MouseClick(sender As Object, e As MouseEventArgs) Handles zIC.MouseClick
        displayChecksDungeons(9, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zGTG_MouseClick(sender As Object, e As MouseEventArgs) Handles zGTG.MouseClick
        displayChecksDungeons(10, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub zOGC_MouseClick(sender As Object, e As MouseEventArgs) Handles zIGC.MouseClick
        displayChecksDungeons(11, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub mnuOptions_MouseEnter(sender As Object, e As EventArgs) Handles mnuOptions.MouseEnter
        mnuOptions.Height = 34
    End Sub

    Private Sub mnuOptions_MouseLeave(sender As Object, e As EventArgs) Handles mnuOptions.MouseLeave
        mnuOptions.Height = 8
    End Sub

    Private Sub tmrFixIt_Tick(sender As Object, e As EventArgs) Handles tmrFixIt.Tick
        updateSettingsPanel()
        resizeForm()
        redrawOutputBoarder()
        tmrFixIt.Enabled = False
    End Sub

    Private Sub pnlMain_Paint(sender As Object, e As PaintEventArgs) Handles pnlMain.Paint
        redrawOutputBoarder()
    End Sub

    Private Sub tmrTT_Tick(sender As Object, e As EventArgs) Handles tmrTT.Tick
        justTheTip.RemoveAll()
        tmrTT.Enabled = False
    End Sub

    Private Sub pbxMap_MouseMove(sender As Object, e As MouseEventArgs) Handles pbxMap.MouseMove
        Dim reg As Integer = lRegions.FindIndex(Function(rng) rng.Contains(e.Location))
        If reg < 0 Or reg = lastTip Then Exit Sub
        justTheTip.SetToolTip(pbxMap, aIconName(reg))
        tmrTT.Enabled = False
        tmrTT.Enabled = True
    End Sub

    Private Sub overrideExits()
        Dim writeExits(7) As Integer
        For i = 0 To 6
            writeExits(i) = 0
        Next
        Select Case iLastMinimap
            Case 0
                writeExits(0) = &H377116     ' KF from Deku Tree
                If aMQ(0) Then writeExits(0) = &H3770B6 ' MQ version
            Case 1
                writeExits(0) = &H36FABE     ' DMT from Dodongo's Cavern
                If aMQ(1) Then writeExits(0) = &H36FA8E ' MQ version
            Case 2
                writeExits(0) = &H36F43E     ' ZD from Jabu-Jabu's Belly
                If aMQ(2) Then writeExits(0) = &H36F40E ' MQ version
            Case 3
                writeExits(0) = &H36ED12     ' SFM from Forest Temple
                If aMQ(3) Then writeExits(0) = &H36ED12 ' MQ version
            Case 4
                writeExits(0) = &H36A3C6     ' DMC from Fire Temple
                If aMQ(4) Then writeExits(0) = &H36A376 ' MQ version
            Case 5
                writeExits(0) = &H36EF76     ' LH from Water Temple
                If aMQ(5) Then writeExits(0) = &H36EF36 ' MQ version
            Case 6
                writeExits(0) = &H36B1EA     ' DC from Spirit Temple
                If aMQ(6) Then writeExits(0) = &H36B12A ' MQ version
            Case 7
                writeExits(0) = &H36C86E     ' GY from Shadow Temple
                If aMQ(7) Then writeExits(0) = &H36C86E ' MQ version
            Case 8
                writeExits(0) = &H3785C6     ' KV from BotW
                If aMQ(8) Then writeExits(0) = &H378566 ' MQ version
            Case 9
                writeExits(0) = &H37352E     ' ZF from Ice Cavern
                If aMQ(9) Then writeExits(0) = &H37344E ' MQ version
            Case 10
                writeExits(0) = &H374348     ' Ganon's Castle from Ganon's Tower
            Case 11
                writeExits(0) = &H373686     ' GF from GTG
                If aMQ(10) Then writeExits(0) = &H373686 ' MQ version
            Case 13
                writeExits(0) = &H3634BC     ' Ganon's Tower from Ganon's Castle
                writeExits(1) = &H3634BE     ' OGC from Ganon's Castle
            Case 27 To 29
                writeExits(0) = &H384650     ' HF from MK Entrance Day
                writeExits(1) = &H384652     ' MK from MK Entrance Day
                If iLastMinimap = 28 Then
                    writeExits(0) = writeExits(0) - &H48
                    writeExits(1) = writeExits(1) - &H48
                End If
            Case 30, 31
                writeExits(0) = &H383824     ' MK from Back Alley Left Day
                writeExits(1) = &H383826     ' MK from Back Alley Right Day
                If iLastMinimap = 31 Then
                    writeExits(0) = writeExits(0) - &H98
                    writeExits(1) = writeExits(1) - &H98
                End If
            Case 32, 33
                writeExits(0) = &H3824A8     ' HC from MK Day
                writeExits(1) = &H3824AA     ' MK Entrance from MK Day
                writeExits(2) = &H3824AE     ' Outside ToT from MK Day
                writeExits(3) = &H3824AC     ' Back Alley Right from MK Day
                writeExits(4) = &H3824B2     ' Back Alley Left from MK Day
                If iLastMinimap = 33 Then
                    writeExits(0) = writeExits(0) + &H40
                    writeExits(1) = writeExits(1) + &H40
                    writeExits(2) = writeExits(2) + &H40
                    writeExits(3) = writeExits(3) + &H40
                    writeExits(4) = writeExits(4) + &H40
                End If
            Case 34
                writeExits(0) = &H383478     ' OGC from MK Day
                writeExits(1) = &H38347A     ' MK Entrance from MK
                writeExits(2) = &H38347E     ' Outside ToT from MK
            Case 35, 36
                writeExits(0) = &H383524     ' ToT from Outside ToT Day
                writeExits(1) = &H383526     ' MK from Outside ToT Day
                If iLastMinimap = 36 Then
                    writeExits(0) = writeExits(0) - &H18
                    writeExits(1) = writeExits(1) - &H18
                End If
            Case 37
                writeExits(0) = &H383574     ' ToT from Outside ToT
                writeExits(1) = &H383576     ' MK from Outside ToT 
            Case 67
                writeExits(0) = &H372332     ' Outside ToT from ToT
            Case 81
                writeExits(0) = &H36BF9C     ' KV from HF
                writeExits(1) = &H36BFA0     ' LW Bridge from HF
                writeExits(2) = &H36BFA2     ' ZR from HF
                writeExits(3) = &H36BFA4     ' GV from HF
                writeExits(4) = &H36BFA8     ' MK from HF
                writeExits(5) = &H36BFAA     ' LLR from HF
                writeExits(6) = &H36BFAE     ' LH from HF
            Case 82
                writeExits(0) = &H368A82     ' BotW from KV
                writeExits(1) = &H368A78     ' DMT from KV
                writeExits(2) = &H368A7A     ' HF from KV
                writeExits(3) = &H368A7E     ' GY from KV
            Case 83
                writeExits(0) = &H378E44     ' Shadow Temple from GY
                writeExits(1) = &H378E46     ' KV from GY
            Case 84
                writeExits(0) = &H37951C     ' HF from ZR
                writeExits(1) = &H37951E     ' ZD from ZR
                writeExits(2) = &H379522     ' LW from ZR
            Case 85
                writeExits(0) = &H3738F4     ' Deku Tree from KF
                writeExits(1) = &H3738FA     ' LW Bridge from KF
                writeExits(2) = &H373902     ' LW from KF
            Case 86
                writeExits(0) = &H36FCD4     ' Forest Temple from SFM
                writeExits(1) = &H36FCD6     ' LW from SFM
            Case 87
                writeExits(0) = &H3696A6     ' Water Temple from LH
                writeExits(1) = &H3696A2     ' HF from LH
                writeExits(2) = &H3696AE     ' ZD from LH
            Case 88
                writeExits(0) = &H37B26C     ' ZF from ZD
                writeExits(1) = &H37B26E     ' ZR from ZD
                writeExits(2) = &H37B270     ' LH from ZD
            Case 89
                writeExits(0) = &H3733CC     ' Jabu-Jabu's Belly from ZF
                writeExits(1) = &H3733D0     ' Ice Cavern from ZF
                writeExits(2) = &H3733D2     ' ZD from ZF
            Case 90
                writeExits(0) = &H373904     ' LH from GV
                writeExits(1) = &H373906     ' HF from GV
                writeExits(2) = &H373908     ' GF from GV
            Case 91
                writeExits(0) = &H37475C     ' SFM from LW
                writeExits(1) = &H37475E     ' KF from LW
                writeExits(2) = &H374768     ' ZR from LW
                writeExits(3) = &H37476A     ' GC from LW
                writeExits(4) = &H37476C     ' KF from LW Bridge
                writeExits(5) = &H37476E     ' HF from LW Bridge
            Case 92
                writeExits(0) = &H36B5AC     ' Spirit Temple from DC
                writeExits(1) = &H36B5AE     ' HW from DC
            Case 93
                writeExits(0) = &H374D02     ' GTG from GF
                writeExits(1) = &H374CE6     ' GV from GF
                writeExits(2) = &H374D00     ' HW from GF
            Case 94
                writeExits(0) = &H37EBFC     ' DC from HW
                writeExits(1) = &H37EBFE     ' GF from HW
            Case 95
                writeExits(0) = &H36C55E     ' MK from HC
                writeExits(1) = &H36C562     ' GFF from HC
                writeExits(2) = &H36C55C     ' Castle Courtyard from HC
            Case 96
                writeExits(0) = &H365FF0     ' Dodongo's Cavern from DMT
                writeExits(1) = &H365FEC     ' GC from DMT
                writeExits(2) = &H365FEE     ' KV from DMT
                writeExits(3) = &H365FF2     ' DMC from DMT
            Case 97
                writeExits(0) = &H374B9E     ' Fire Temple from DMC
                writeExits(1) = &H374B98     ' GC from DMC
                writeExits(2) = &H374B9A     ' DMT from DMC
            Case 98
                writeExits(0) = &H37A64C     ' DMC from GC
                writeExits(1) = &H37A64E     ' DMT from GC
                writeExits(2) = &H37A650     ' LW from GC
            Case 99
                writeExits(0) = &H377C12     ' HF from LLR
            Case 100
                writeExits(0) = &H37FEC0     ' Ganon's Castle from OGC
                writeExits(1) = &H37FEC2     ' MK from OGC
        End Select

        For i = 0 To 6
            If writeExits(i) = 0 Then Exit For
            quickWrite16(writeExits(i), 0, emulator)
        Next
    End Sub

    Private Function getGanonMap() As Byte
        getGanonMap = 0

        ' 0: Main Region Upper
        ' 1: Main Region Lower
        ' 2: Forest Trial
        ' 3: Water Trial
        ' 4: Shadow Trial
        ' 5: Fire Trial
        ' 6: Light Trial
        ' 7: Spirit Trial

        ' Get the XYZ position
        Dim linkPOS As Double() = getPosition()

        ' First determine if player is on the upper or lower region of the castle
        ' Divide line being -37, the height of the landing halfway down the stairs
        If linkPOS(1) > -37 Then
            ' Top half of the castle
            Dim testDistance As Double = 0
            Dim ptLink As New Point(CInt(linkPOS(0)), CInt(linkPOS(2)))
            ' Split into 4 regions
            If linkPOS(2) > -840 Then
                ' This is the southern region of the castle

                ' Create a gap to ignore the area you enter in at
                If linkPOS(0) < -370 Then
                    ' This is the south-western region: Spirit
                    If lineSide(New Point(-147, 454), New Point(-1194, -151), ptLink) < -50000 Then getGanonMap = 7
                ElseIf linkPOS(0) > 370 Then
                    ' This is the south-eastern region: Forest
                    If lineSide(New Point(1194, -151), New Point(146, 455), ptLink) < -50000 Then getGanonMap = 2
                End If
            Else
                ' This is the northern region of the castle
                If linkPOS(0) < 0 Then
                    ' This is the north-western region: Fire
                    If lineSide(New Point(-1197, -1529), New Point(0, -2219), ptLink) < -50000 Then getGanonMap = 5
                Else
                    ' This is the south-western region: Shadow
                    If lineSide(New Point(0, -2219), New Point(1194, -1529), ptLink) < -50000 Then getGanonMap = 4
                End If
            End If
        Else
            ' Lower half of the castle
            If linkPOS(0) < -1319 Then
                ' The divide between the Castle into the Light Trial door is from -1291 to -1347
                ' Using -1319 as the halfway point: Light
                getGanonMap = 6
            ElseIf linkPOS(0) > 1232 Then
                ' The divide between the Castle into the Water Trial door is from 1204 to 1260
                ' Using 1232 as the halfway point: Water
                getGanonMap = 3
            Else
                ' Else, on the lower level, just use the main region
                getGanonMap = 1
            End If
        End If
    End Function

    Private Function lineSide(ByVal pt1 As Point, ByVal pt2 As Point, ByVal ptTest As Point) As Double
        lineSide = (pt2.X - pt1.X) * (ptTest.Y - pt1.Y) - (pt2.Y - pt1.Y) * (ptTest.X - pt1.X)
    End Function

    Private Function getPosition() As Double()
        ' Grab values for XYZ
        Dim valX As String = String.Empty
        Dim valY As String = String.Empty
        Dim valZ As String = String.Empty

        If isSoH Then
            valX = Convert.ToString(GDATA(&H17264), 2)
            valY = Convert.ToString(GDATA(&H17268), 2)
            valZ = Convert.ToString(GDATA(&H1726C), 2)
        Else
            valX = Convert.ToString(goRead(&H1DAA54), 2)
            valY = Convert.ToString(goRead(&H1DAA54), 2)
            valZ = Convert.ToString(goRead(&H1DAA54), 2)
        End If

        fixBinaryLength(valX, valY, valZ)

        ' Convert values into IEEE-754 floating points
        Dim coordX As Double = bin2float(valX)
        Dim coordY As Double = bin2float(valY)
        Dim coordZ As Double = bin2float(valZ)

        ' Return the whole array
        Return New Double() {coordX, coordY, coordZ}
    End Function

    Private Sub updateTrials()
        pbxTrialForest.Visible = checkLoc("6611")
        pbxTrialFire.Visible = checkLoc("6614")
        pbxTrialWater.Visible = checkLoc("6612")
        pbxTrialSpirit.Visible = checkLoc("6629")
        pbxTrialShadow.Visible = checkLoc("6613")
        pbxTrialLight.Visible = checkLoc("6615")
    End Sub

    Private Sub wastelandPOS()
        Dim linkPOS As Double() = getPosition()

        Dim coordX As Double = ((linkPOS(0) + 4550) / 8200) * 400 + 53
        Dim coordZ As Double = ((linkPOS(2) + 3750) / 8200) * 400 + 17


        Dim linkRot As Integer = 0

        If isSoH Then
            linkRot = CInt(GDATA(&H1730A, 2))
        Else
            linkRot = goRead(&H1DAA74, 15)
        End If
        Dim headA As Double = (((linkRot / 65535 * 360) - 90) * -1) * Math.PI / 180
        Dim tailA As Double = (((linkRot / 65535 * 360) + 90) * -1) * Math.PI / 180

        Dim headX As Double = (8 * Math.Cos(headA)) + coordX
        Dim headZ As Double = (8 * Math.Sin(headA)) + coordZ
        Dim tailX As Double = (8 * Math.Cos(tailA)) + coordX
        Dim tailZ As Double = (8 * Math.Sin(tailA)) + coordZ

        Dim p As New Pen(Color.Yellow, 5)
        p.EndCap = Drawing2D.LineCap.ArrowAnchor

        Graphics.FromImage(pbxMap.Image).DrawLine(p, CInt(tailX), CInt(tailZ), CInt(headX), CInt(headZ))
    End Sub

    Private Sub tmrFastScan_Tick(sender As Object, e As EventArgs) Handles tmrFastScan.Tick
        'tmrFastScan.Interval = 1000
        ' Timer is ran every 3 seconds. Often enough for accuracy but not so much that it bogs things down. May Lower to 5 seconds.

        ' Checks that things should still be running and ready
        If Not keepRunning Or emulator = String.Empty Then
            stopScanning()
            Exit Sub
        End If

        ' Runs a check for "ZELDAZ" in the memory where it is expected
        If Not checkZeldaz() = 0 Then
            updateFast()
            If Not keepRunning Then
                stopScanning()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub dump()
        ' Dump values to help debug issues for users, start with the current items
        Dim sText As String = allItems & vbCrLf
        Dim sHex As String = String.Empty
        ' Next read all the values of arrLocation
        For i = 0 To arrLocation.Length - 1
            sHex = Hex(goRead(arrLocation(i)))
            fixHex(sHex)
            sText = sText & i.ToString & ": " & sHex & vbCrLf
        Next
        ' Add in the single checks that do not make it to arrLocations, mainly Bean Plants and a few events
        sText = sText & checkLoc("B0") & checkLoc("B1") & checkLoc("B2") & checkLoc("B3") & checkLoc("B4") & vbCrLf
        sText = sText & checkLoc("C00") & checkLoc("C01") & checkLoc("C02") & checkLoc("C03") & checkLoc("C04") & checkLoc("C05")
        Clipboard.SetText(sText)
    End Sub
End Class

Public Class custMenu
    Inherits ToolStripProfessionalRenderer
    Shared Property highlight As Color = Color.Blue
    Shared Property backColour As Color = Color.WhiteSmoke
    Shared Property foreColour As Color = Color.Black
    Shared Property font As New Font("Segeo UI", 9)
    Protected Overloads Overrides Sub OnRenderItemCheck(e As ToolStripItemImageRenderEventArgs)
        e.Graphics.DrawRectangle(New Pen(foreColour, 1), 10, 3, e.Item.Size.Height - 7, e.Item.Size.Height - 7)
        e.Graphics.FillRectangle(New SolidBrush(foreColour), 13, 6, e.Item.Size.Height - 12, e.Item.Size.Height - 12)
    End Sub
    Protected Overloads Overrides Sub OnRenderMenuItemBackground(ByVal e As ToolStripItemRenderEventArgs)
        With e.Item
            Dim cBack As Color = CType(IIf(.Selected, highlight, backColour), Color)
            e.Graphics.FillRectangle(New SolidBrush(cBack), New Rectangle(Point.Empty, .Size))
            Select Case LCase(.Text)
                Case "scan", "auto scan", "stop", "reset", "exit", "themes", "settings >", "settings <", "mini-map"
                    e.Graphics.DrawRectangle(New Pen(foreColour, 1), New Rectangle(0, 0, .Width - 1, .Height - 1))
            End Select
            .ForeColor = foreColour
            .Font = font
        End With
    End Sub
End Class
