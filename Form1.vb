Option Explicit On
Option Strict On

Public Class frmTrackerOfTime
    Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Public Declare Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Integer, ByVal lpBaseAddress As Integer, ByRef lpBuffer As Integer, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer

    ' Constant variables used throughout the app. The most important here is the 'IS_64BIT' as this needs to be set if compiling in x64
    Private Const PROCESS_ALL_ACCESS = &H1F0FFF
    Private Const CHECK_COUNT = 113
    Private Const IS_64BIT = True

    ' Variables used to determine what emulator is connected, its state, and its starting memory address
    Private Const romAddrStart As Integer = &HDFE40000
    Private romAddrStart64 As Int64 = 0
    Private emulator As String = String.Empty
    Private keepRunning As Boolean = False
    Private zeldazFails As Integer = 0

    ' Variables for a variety of rom info used in the scan
    Private pedestalRead As Byte = 0
    Private playerName As String = String.Empty
    Private randoVer As String = String.Empty
    Private rainbowBridge(1) As Byte
    Private aGetQuantity(14) As Boolean

    ' Arrays for tracking checks
    Private keyCount As Integer = 325
    Private aKeys(keyCount) As keyCheck
    Private aKeysDungeons(11)() As keyCheck
    Private aQuestRewardsCollected(22) As Boolean
    Private aQIChecks(22) As Boolean

    ' RTB variables
    Private emboldenList As New List(Of String)
    Private rtbRefresh As Byte = 0
    Private lastArea As String = String.Empty
    Private lastOutput As New List(Of String)

    ' Arrays for location scanning and their settings
    Private arrLocation(CHECK_COUNT) As Integer
    Private arrChests(CHECK_COUNT) As Integer
    Private arrHigh(CHECK_COUNT) As Byte
    Private arrLow(CHECK_COUNT) As Byte

    ' Variables for from settings
    Private firstRun As Boolean = True
    Private showSetting As Boolean = False

    ' Variable for the colour of the highlight, blended from both front and background colour
    Private cBlend As New Color

    ' Arrays for the displaying of the dungeon name over a quest reward
    Private aQuestRewardsText(22) As Byte
    Private aDungeonLetters() As String = {"", "Free", "Deku", "DC", "Jabu", "Frst", "Fire", "Wtr", "Sprt", "Shdw"}

    ' Variables for tracking logic checks
    Private allItems As String = String.Empty
    Private goldSkulltulas As Byte = 0
    Private canMagic As Boolean = False
    Private canAdult As Boolean = False
    Private canYoung As Boolean = False
    Private youngZD As Boolean = False
    Private adultZD As Boolean = False
    Private youngDC As Boolean = False
    Private adultDC As Boolean = False
    Private youngSpTCentre As Boolean = False
    Private adultSpTCentre As Boolean = False
    Private magicBeans As Byte = 0
    Private aDungeonKeys(7) As Byte
    Private aBossKeys(7) As Boolean

    Private arrDAC(11)() As Boolean

    ' Variables for detecting the royal tomb's Sun's Song check 
    Private Const CUR_ROOM_ADDR As Integer = &H1C8544
    Private lastRoomScan As Integer = 0
    Private royalSongScan As Integer = 0

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
    Private aoInventory(23) As PictureBox

    Private Sub frmTrackerOfTime_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        stopScanning()
    End Sub

    ' On load, populate the locations array
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i = 0 To aKeys.Length - 1
            aKeys(i) = New keyCheck
        Next

        makeArrayObjects()
        custMenu.highlight = Me.ForeColor
        custMenu.backColour = Me.BackColor
        custMenu.foreColour = Me.ForeColor
        mnuOptions.Renderer = New custMenu

        'custComboBox.highlight = Me.ForeColor
        'custComboBox.backColour = Me.BackColor
        'custComboBox.foreColour = Me.ForeColor
        'ComboBox1 = New custCombobox
        'custComboBox.
        loadSettings()

        For i As Integer = 0 To arrLocation.Length - 1
            arrChests(i) = 0
            arrHigh(i) = 0
            arrLow(i) = 31
        Next

        For Each checkBox In pnlHidden.Controls.OfType(Of CheckBox)().Where(Function(cb As CheckBox) cb.Name.Contains("cb65"))
            AddHandler checkBox.CheckedChanged, AddressOf cccCarpenters
        Next
        For Each checkBox In pnlHidden.Controls.OfType(Of CheckBox)().Where(Function(cb As CheckBox) cb.Name.Contains("cb74"))
            AddHandler checkBox.CheckedChanged, AddressOf cccEquipment
        Next
        For Each checkBox In pnlHidden.Controls.OfType(Of CheckBox)().Where(Function(cb As CheckBox) cb.Name.Contains("cb75"))
            AddHandler checkBox.CheckedChanged, AddressOf cccEquipment
        Next
        For Each checkBox In pnlHidden.Controls.OfType(Of CheckBox)().Where(Function(cb As CheckBox) cb.Name.Contains("cb76"))
            AddHandler checkBox.CheckedChanged, AddressOf cccUpgrades
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
        ' Used to handle the mouse enter/leave and click event for all of the label dropdowns
        'For Each Label In pnlSettings.Controls.OfType(Of Label)().Where(Function(lbl As Label) Mid(lbl.Name, 1, 3) = "ldd")
        'Label.BorderStyle = BorderStyle.None
        'Label.Padding = New Padding(1)
        'AddHandler Label.MouseEnter, AddressOf handleLDDMouseEnter
        'AddHandler Label.MouseLeave, AddressOf handleLDDMouseLeave
        'AddHandler Label.MouseClick, AddressOf handleLDDMouseClick
        'AddHandler Label.Paint, AddressOf handleLDDPaint
        'Next
        ' Used to give a border to each panel below the item display
        For Each Panel In Me.Controls.OfType(Of Panel)().Where(Function(pnl As Panel) pnl.Location.Y > 300)
            AddHandler Panel.Paint, AddressOf pnlDrawBorder
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

        ' Link  the pictureboxes of each inventory item into an array
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
            Case 12
                aoQuestItemImages(img) = My.Resources.songNormal
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyNormal
            Case 13
                aoQuestItemImages(img) = My.Resources.songNormal
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyNormal
            Case 14
                aoQuestItemImages(img) = My.Resources.songNormal
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyNormal
            Case 15
                aoQuestItemImages(img) = My.Resources.songNormal
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyNormal
            Case 16
                aoQuestItemImages(img) = My.Resources.songNormal
                aoQuestItemImagesEmpty(img) = My.Resources.songEmptyNormal
            Case 17
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
        youngZD = False
        adultZD = False
        youngDC = False
        adultDC = False
        youngSpTCentre = False
        adultSpTCentre = False
        rainbowBridge(0) = 4
        rainbowBridge(1) = 64
        magicBeans = 0
        goldSkulltulas = 0
        lastArea = String.Empty
        lastOutput.Clear()

        resetDAC()
        For i = 0 To 7
            aDungeonKeys(i) = 0
            aBossKeys(i) = False
        Next

        ' Clean up the UI
        clearItems()
        setupKeys()
        updateLabels()
        updateLabelsDungeons()

        ' Process the high/lows, but only the first time
        If firstRun Then
            getHighLows()
            firstRun = False
        End If

        arrLocation(0) = &H11AD18 + 4       ' DMC/DMT/OGC Great Fairy Fountain
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


        For i As Integer = 0 To arrLocation.Length - 1
            arrChests(i) = 0
        Next
        updateShoppes()

    End Sub

    Private Function checkSunSong(Optional flipOn As Boolean = False) As Boolean
        checkSunSong = False
        Dim sunCheck As Double = goRead(&H11B4AC) 'romAddrStart + arrLocation(64))
        If sunCheck = 0 Then Return False
        Dim sunSongHex As String = Hex(sunCheck)
        While sunSongHex.Length < 8
            sunSongHex = "0" & sunSongHex
        End While
        Dim sunSongBit = Mid(sunSongHex, 6, 1)
        Select Case sunSongBit
            Case "4", "5", "6", "7", "C", "D", "E", "F"
                Return True
        End Select
        If Not flipOn Then Return False
        sunCheck = CInt("&H" & sunSongBit)
        sunCheck = sunCheck + 4
        sunSongBit = Hex(sunCheck)
        sunSongHex = Mid(sunSongHex, 1, 5) & sunSongBit & Mid(sunSongHex, 7)
        sunCheck = CDbl("&H" & sunSongHex)
        If sunCheck > 2147483648 Then
            sunCheck = sunCheck - 2147483648 - 2147483648
        End If
        If emulator = "variousX64" Then
            Try
                WriteMemory(Of Integer)(romAddrStart64 + &H11B4AC, CInt(sunCheck))
            Catch ex As Exception
                stopScanning()
                MessageBox.Show("checkSunSong Problem: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        ElseIf emulator = "project64" Then
            If IS_64BIT = False Then quickWrite(romAddrStart + &H11B4AC, CInt(sunCheck))
        End If
    End Function

    Private Sub getAge()
        ' Checks for the current age variable, starts off with setting ages to not accessable
        canAdult = False
        canYoung = False

        ' If 0 or 1 , set the age as accessable
        Select Case CByte(goRead(&H11A5D4, 1))
            Case 0
                canAdult = True
                If entranceCheck("TT") Then
                    ' If you have an ocarina and the song of time, then you can become young Link
                    If allItems.Contains("h") And aQIChecks(16) Then canYoung = True
                    ' If Temple of Time is open, then you can become young Link
                    If checkLoc("6427") Then canYoung = True
                End If
            Case 1
                canYoung = True
                If entranceCheck("TT") Then
                    ' If you have an ocarina and the song of time, then you can become adult Link
                    If allItems.Contains("h") And aQIChecks(16) Then canAdult = True
                    ' If Temple of Time is open, then you can become adult Link
                    If checkLoc("6427") Then canAdult = True
                End If
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
            tempItems = Hex(goRead(&H11A678 + (i * 4)))

            ' Make sure all leading 0's are put back
            While tempItems.Length < 8
                tempItems = "0" & tempItems
            End While

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
    End Sub
    Private Sub getGoldSkulltulas()
        ' Get gold skulltula count
        Dim bGS As Byte = CByte(goRead(&H11A6A0 + 2, 1))

        ' Checks to see if the number of gold skulltula's have changed, if not then do not bother to do all that
        If bGS = goldSkulltulas Then Exit Sub

        ' Update your gold skulltulas
        goldSkulltulas = bGS

        With pbxGoldSkulltula
            ' If not visible, make visible, and reset image to default to remove drawn numbers
            If Not .Visible Then .Visible = True
            .Image = My.Resources.goldSkulltula

            ' 23 is the starting position for single digit numbers, but will be reducing it by 5 since byte varaibles cannot be negative and triple digits needs to start at -5
            Dim xPos As Byte = 23
            ' If double digits, set starting position to 12(-5 = 7)
            If bGS > 9 Then xPos = 12
            ' If triple digits, set starting position to 0(-5 = -5)
            If bGS > 99 Then xPos = 0
            ' Font for gold skulltula numbers
            Dim fontGS = New Font("Lucida Console", 18, FontStyle.Bold, GraphicsUnit.Pixel)

            ' Draw the value over the lower right of the gold skulltula picturebox, first in black to give it some definition, then in white
            Graphics.FromImage(.Image).DrawString(bGS.ToString, fontGS, New SolidBrush(Color.Black), xPos - 6, 17)
            Graphics.FromImage(.Image).DrawString(bGS.ToString, fontGS, New SolidBrush(Color.White), xPos - 5, 18)
        End With
    End Sub
    Private Sub getHearts()
        ' Check if player has enhanced defence
        Dim isEnhanced As Boolean = CBool(IIf(goRead(&H11A60C + 2, 1) > 0, True, False))

        ' Get max heart container value, divide by 16 to undo their 16x multiplyer
        Dim bHearts As Byte = CByte(goRead(&H11A5FC, 15) / 16)

        ' Just a percaution, limit hearts to 99, even though 20 is the max. Never know what people may do to their save file 
        If bHearts > 99 Then bHearts = 99

        ' 18 is the starting position for single digit numbers
        Dim xPos As Byte = 18
        ' If double digits, set starting position to 7
        If bHearts > 9 Then xPos = 7

        ' Set up the font for the number of hearts
        Dim fontHearts = New Font("Lucida Console", 18, FontStyle.Bold, GraphicsUnit.Pixel)

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
            Graphics.FromImage(.Image).DrawString(bHearts.ToString, fontHearts, New SolidBrush(Color.Black), xPos - 1, 17)
            Graphics.FromImage(.Image).DrawString(bHearts.ToString, fontHearts, New SolidBrush(Color.White), xPos, 18)
            '.Invalidate()
        End With

        ' Now for the heart pieces
        bHearts = CByte(goRead(&H11A674 + 3, 1))

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
    Private Sub getItemAmounts()
        ' Get the amount of each item
        Dim items1 As String = Hex(goRead(&H11A65C))
        Dim items2 As String = Hex(goRead(&H11A660))
        Dim items3 As String = Hex(goRead(&H11A664))
        Dim items4 As String = Hex(goRead(&H11A668))

        fixHex(items1)
        fixHex(items2)
        fixHex(items3)
        fixHex(items4)

        Dim itemsAll As String = items1 & items2 & items3 & Mid(items4, 1, 6)
        Dim itemAmount As Byte = 0

        For i = 0 To 14
            If aGetQuantity(i) = True Then
                itemAmount = CByte("&H" & Mid(itemsAll, (i * 2) + 1, 2))
                If itemAmount > 0 Then
                    With aoInventory(i)
                        ' 23 is the starting position for single digit numbers, but will be reducing it by 5 since byte varaibles cannot be negative and triple digits needs to start at -5
                        Dim xPos As Byte = 23
                        ' If double digits, set starting position to 12(-5 = 7)
                        If itemAmount > 9 Then xPos = 12
                        ' If triple digits, set starting position to 0(-5 = -5)
                        If itemAmount > 99 Then xPos = 0
                        ' Font for items numbers
                        Dim fontGS = New Font("Lucida Console", 18, FontStyle.Bold, GraphicsUnit.Pixel)

                        ' Draw the value over the lower right of the item's picturebox, first in black to give it some definition, then in white
                        Graphics.FromImage(.Image).DrawString(itemAmount.ToString, fontGS, New SolidBrush(Color.Black), xPos - 6, 17)
                        Graphics.FromImage(.Image).DrawString(itemAmount.ToString, fontGS, New SolidBrush(Color.White), xPos - 5, 18)
                    End With
                End If
            End If
        Next
        ' Since the last check is Magic Beans, store that value to our beans
        magicBeans = itemAmount
    End Sub
    Private Sub getMagic()
        ' Get magic check. This one is a curious one as there are many areas to check, and none perfect. Should someone hack their file to have magic, this should hopefully find it though
        Dim bMagic As Byte = CByte(goRead(&H11A600 + 1, 1))

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
        ' Variables used to grab values and store the needed parts into an easy to use string
        Dim tempKeys As String = String.Empty
        Dim stringKeys As String = String.Empty

        '  Grab keys for just Forest Temple
        tempKeys = Hex(goRead(&H11A68C, 1))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys, 2)

        ' Set string to the Forest Temple keys
        stringKeys = tempKeys

        ' Grab keys for Fire, Water, Spirit, and Shadow Temple
        tempKeys = Hex(goRead(&H11A690))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys)

        ' Add all four of the grabbed Temple keys
        stringKeys = stringKeys & tempKeys

        ' Grab keys for Bottom of the Well and Gerudo Training Ground
        tempKeys = Hex(goRead(&H11A694))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys)

        ' Add Bottom of the Well and Gerudo Training Ground keys
        stringKeys = stringKeys & Mid(tempKeys, 1, 2) & Mid(tempKeys, 7, 2)

        ' Grab keys for Ganon's Castle
        tempKeys = Hex(goRead(&H11A698))

        ' Make sure all leading 0's are put back
        fixHex(tempKeys)

        ' Add Ganon's Castle keys
        stringKeys = stringKeys & Mid(tempKeys, 3, 2)

        ' Set up the font for the number of keys
        Dim fontSmallKeys = New Font("Lucida Console", 14, FontStyle.Bold, GraphicsUnit.Pixel)
        ' Variable for temp storage for keys
        Dim valKeys As Byte = 0
        ' Position for x: 2 for double digits, 11 for single
        Dim xPos As Byte = 0

        ' Step through each stored small keys
        For i = 0 To 7
            With aoSmallKeys(i)
                ' Default position 11 since technically no key should reach double digits
                xPos = 11

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
                    If valKeys > 9 Then xPos = 2

                    ' Draw the value over the lower right of the key picturebox, first in black to give it some definition, then in white
                    Graphics.FromImage(.Image).DrawString(valKeys.ToString, fontSmallKeys, New SolidBrush(Color.Black), xPos - 1, 10)
                    Graphics.FromImage(.Image).DrawString(valKeys.ToString, fontSmallKeys, New SolidBrush(Color.White), xPos, 11)
                    '.Invalidate()
                End If
                aDungeonKeys(i) = valKeys
            End With
        Next
    End Sub
    Private Function isTriforceHunt() As Boolean
        ' Default to false, and grab at where the varaible for needed triforce pieces is at
        isTriforceHunt = False

        ' Checks for different rom generators
        Dim isOOTR As Boolean = False
        Dim isAP As Boolean = False

        ' Both use the pattern of "0032####" for thieir requirement check location
        If goRead(&H40B668 + 2, 15) = 50 Then isOOTR = True

        ' If not OOTR, check for AP
        If Not isOOTR Then
            If goRead(&H40B1B0 + 2, 15) = 50 Then isAP = True
        End If

        ' Variable to store required triforce pieces
        Dim neededTriforce As Int16 = 0

        Select Case True
            Case isAP
                ' If rom is generated by archipelago.gg
                neededTriforce = CShort(goRead(&H40B1B0, 15))
            Case isOOTR
                ' If rom is generated by ootrandomizer.com
                neededTriforce = CShort(goRead(&H40B668, 15))
            Case Else
                ' Anything else, return false
                Return False
        End Select

        ' If value is 'FFFF' then it is just a placeholder, not a triforce hunt, so return false
        If neededTriforce = -1 Then Return False

        ' If it makes it this far, return true
        Return True
    End Function
    Private Sub getTriforce()
        ' Get gold skulltula count
        Dim bTriforce As Byte = CByte(goRead(&H11AE94, 1))

        With pbxTriforce
            If bTriforce > 0 Then
                If Not .Visible Then .Visible = True
                ' Reset the image to the default triforce piece to remove drawn numbers
                .Image = My.Resources.triforce

                ' 23 is the starting position for single digit numbers, but will be reducing it by 5 since byte varaibles cannot be negative and triple digits needs to start at -5
                Dim xPos As Byte = 23
                ' If double digits, set starting position to 12(-5 = 7)
                If bTriforce > 9 Then xPos = 12
                ' If triple digits, set starting position to 0(-5 = -5)
                If bTriforce > 99 Then xPos = 0
                ' Font for triforce numbers
                Dim fontTriforce = New Font("Lucida Console", 18, FontStyle.Bold, GraphicsUnit.Pixel)

                ' Draw the value over the lower right of the triforce picturebox, first in black to give it some definition, then in white
                Graphics.FromImage(.Image).DrawString(bTriforce.ToString, fontTriforce, New SolidBrush(Color.Black), xPos - 6, 17)
                Graphics.FromImage(.Image).DrawString(bTriforce.ToString, fontTriforce, New SolidBrush(Color.White), xPos - 5, 18)
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

    ' Scan each of the chests data
    Private Sub readChestData()
        If emulator = String.Empty Then
            If IS_64BIT = False Then
                emulator = "project64"
            Else
                attachToBizHawk()
                If emulator = String.Empty Then attachToRMG()
                If emulator = String.Empty Then attachToM64P()
                If emulator = String.Empty Then attachToRetroArch()
            End If
            If Not emulator = String.Empty Then
                Me.Text = "Tracker of Time (" & emulator & ")"
                Select Case LCase(emulator)
                    Case "emuhawk", "rmg", "mupen64plus-gui", "retroarch - mupen64plus", "retroarch - parallel"
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
        Dim locationCode As Integer = goRead(CUR_ROOM_ADDR + 2, 15)
        Dim doMath As Integer = 0
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If

        Dim chestCheck As Integer
        Dim foundChests As Double = 0
        Dim compareTo As Double = 0
        Dim strI As String = String.Empty
        Dim strII As String = String.Empty
        Dim doCheck As Boolean = False
        Dim checkAgain = False

        For i = 0 To arrLocation.Length - 1
            If i Mod 5 = 0 Then Application.DoEvents()
            checkAgain = True
            Select Case i
                Case 0 To 59, Is >= 100
                    ' These are the area checks, either chest, standing items, area events, as they will need to be checked as they happen

                    doMath = (locationCode * 28) + 212 + &H11A5D0
                    tempVar = &H1CA1D8

                    ' Royal Tomb Only
                    If locationCode = 65 And randoVer = "OOTR" Then
                        If Not cb6410.Checked Then
                            If royalSongScan = 0 Then royalSongScan = goRead(&H11A674)
                            If Not royalSongScan = goRead(&H11A674) Then
                                royalSongScan = 0
                                cb6410.Checked = True
                                checkSunSong(True)
                            End If
                        End If
                    ElseIf royalSongScan > 0 Then
                        royalSongScan = 0
                    End If

                    Select Case i
                        Case 0 To 2, Is >= 103
                            ' Scene Checks
                            inc(doMath, 4)
                            tempVar = &H1CA1C8
                        Case 3 To 30, 100 To 102
                            ' Standing Checks
                            inc(doMath, 12)
                            tempVar = &H1CA1E4
                    End Select

                    If doMath = arrLocation(i) Then
                        checkAgain = False
                        chestCheck = goRead(tempVar, arrHigh(i))
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
                chestCheck = goRead(arrLocation(i), arrHigh(i))
                If Not keepRunning Then
                    stopScanning()
                    Exit Sub
                End If
                If Not chestCheck = arrChests(i) Then
                    arrChests(i) = chestCheck
                    parseChestData(i)
                End If
            End If
        Next
        scanSingleChecks()

        If rtbRefresh = 0 Then
            Select Case locationCode
                Case 0 To 9
                    ' Dungeons from DT to IC
                    lastArea = "DN" & locationCode.ToString
                Case 11
                    ' Gerudo Training Grounds
                    lastArea = "DN10"
                Case 13, 10
                    ' Ganon's Castle 
                    lastArea = "DN11"
                Case 85, 52
                    ' Kakori Forest
                    lastArea = "KF"
                Case 91
                    ' Lost Woods
                    lastArea = "LW"
                Case 86
                    ' Sacred Forest Meadow
                    lastArea = "SFM"
                Case 81
                    ' Hyrule Field
                    lastArea = "HF"
                Case 99
                    ' Lon Lon Ranch
                    lastArea = "LLR"
                Case 32, 27
                    ' The Market
                    lastArea = "MK"
                Case 67, 35
                    ' Temple of Time
                    lastArea = "TT"
                Case 95, 74
                    ' Hyrule Castle
                    lastArea = "HC"
                Case 82
                    ' Kakakori Village
                    lastArea = "KV"
                Case 83
                    ' Graveyard
                    lastArea = "GY"
                Case 96
                    ' Death Mountain Trail
                    lastArea = "DMT"
                Case 97
                    ' Death Mountain Crater
                    lastArea = "DMC"
                Case 98
                    ' Goron City
                    lastArea = "GC"
                Case 84
                    ' Zora's River
                    lastArea = "ZR"
                Case 88
                    ' Zora's Domain
                    lastArea = "ZD"
                Case 89
                    ' Zora's Fountain
                    lastArea = "ZF"
                Case 87
                    ' Lake Hylia
                    lastArea = "LH"
                Case 90
                    ' Gerudo Valley
                    lastArea = "GV"
                Case 93, 12
                    ' Gerudo Fortress
                    lastArea = "GF"
                Case 94
                    ' Haunted Wasteland
                    lastArea = "HW"
                Case 92
                    ' Desert Colossus
                    lastArea = "DC"
            End Select

            If Not lastArea = String.Empty Then
                If Mid(lastArea, 1, 2) = "DN" Then
                    displayChecksDungeons(CByte(Mid(lastArea, 3)), False, False)
                Else
                    displayChecks(lastArea, False, False)
                End If
            End If
        Else
            decB(rtbRefresh)
        End If
    End Sub
    Private Sub scanSingleChecks()
        ' LW Bean Planted
        If My.Settings.setSkulltula > 0 Then setLoc("B0", checkBit(&H11B09C, 22))
        ' DC Bean Planted
        setLoc("B1", checkBit(&H11B0B8, 24))
        ' DMT Bean Planted
        setLoc("B2", checkBit(&H11B128, 6))
        ' DMC Bean Planted
        setLoc("B3", checkBit(&H11B144, 3))

        ' GV Opened Gate to Haunted Wasteland
        setLoc("C00", checkBit(&H11B0D4, 3))
        ' DMC Deku Near Ladder
        If My.Settings.setScrub Then setLoc("C01", checkBit(&H11B150, 6))
        ' EV: KV Well Drained
        setLoc("C02", checkBit(&H11B4B0, 23))
        ' EV: LH Restored
        setLoc("C04", checkBit(&H11B4B0, 25))
        ' Deliver Zelda's Letter | Unlock Mask Shoppe
        setLoc("C05", checkBit(&H11B4D4, 6))


    End Sub

    Private Sub parseChestData(ByVal loc As Integer)
        Dim foundChests As Double = arrChests(loc)
        Dim compareTo As Double = 2147483648
        Dim strI As String = loc.ToString
        Dim strII As String = String.Empty
        Dim doCheck As Boolean = False
        Dim gotHit As Boolean = False
        Dim startI As Byte = 31

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

        If IS_64BIT Then
            'startI = 31
            'compareTo = 2147483648
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
                        Case "6410"
                            If doCheck = True Then cb6410.Checked = True
                    End Select
                    If .area = "INV" Then data2Checkbox(.loc, .checked)
                    gotHit = True
                End With
            Next

            If gotHit = False Then
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

    Private Sub data2Checkbox(ByVal box As String, ByVal val As Boolean)
        ' This sub converts the key's check data into a checkbox's checked
        ' This is literally just more lazy programming... I should use variables instead of invisible checkboxes
        For Each chk In pnlHidden.Controls.OfType(Of CheckBox)()
            ' Search all the checkboxes in the hidden panel, set its value to the key's check, and exit the sub
            If chk.Name = "cb" & box Then
                chk.Checked = val
                Exit Sub
            End If
        Next
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
                        Select Case My.Settings.setSkulltula
                            Case 0
                                countCheck = False
                            Case 1
                                countCheck = True
                            Case Else
                                countCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                        End Select
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
                            If Not .checked Then
                                If checkLogic(.logic, .area) Then aBoldLabels(area2num(.area)) = True
                            End If
                        End If
                    End If
                End If
            End With
        Next

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
                            Select Case My.Settings.setSkulltula
                                Case 0
                                    countCheck = False
                                Case 1
                                    countCheck = True
                                Case Else
                                    countCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                            End Select
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
                                    If checkLogic(.logic, .area) Then aBoldLabels(i) = True
                                End If
                            End If
                        End If
                    End If
                End With
            Next
        Next

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
            Case "Market"
                Return "MK"
            Case "Temple of Time"
                Return "TT"
            Case "Hyrule Castle"
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
        AutoScanToolStripMenuItem.Text = "Auto Scan"
        'Me.Controls.Find("xButtonAutoScan", True)(0).Text = "Auto Scan"
        Me.Text = "Tracker of Time"
        emulator = String.Empty
    End Sub

    Private Function goRead(ByVal offsetAddress As Integer, Optional bitType As Byte = 31) As Integer
        goRead = 0
        If emulator = "variousX64" Then
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
        ElseIf emulator = "project64" Then
            If IS_64BIT = False Then
                Select Case bitType
                    Case 0 To 7
                        goRead = quickRead8(romAddrStart + offsetAddress, emulator)
                    Case 8 To 15
                        goRead = quickRead16(romAddrStart + offsetAddress, emulator)
                    Case Else
                        goRead = quickRead32(romAddrStart + offsetAddress, emulator)
                End Select
            End If
        End If
    End Function
    Private Function quickRead8(ByVal readAddress As Integer, Optional ByVal sTarget As String = "project64") As Integer
        quickRead8 = 0

        Dim p As Process = Nothing
        If Process.GetProcessesByName(sTarget).Count > 0 Then
            p = Process.GetProcessesByName(sTarget)(0)
        Else
            stopScanning()
            MessageBox.Show(sTarget & " is not open!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Function
        End If
        Try
            quickRead8 = Memory.ReadInt8(p, readAddress)
        Catch ex As Exception
            stopScanning()
            MessageBox.Show("quickRead Problem: " & vbCrLf & ex.Message & vbCrLf & readAddress.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function
    Private Function quickRead16(ByVal readAddress As Integer, Optional ByVal sTarget As String = "project64") As Integer
        quickRead16 = 0

        Dim p As Process = Nothing
        If Process.GetProcessesByName(sTarget).Count > 0 Then
            p = Process.GetProcessesByName(sTarget)(0)
        Else
            stopScanning()
            MessageBox.Show(sTarget & " is not open!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Function
        End If
        Try
            quickRead16 = Memory.ReadInt16(p, readAddress)
        Catch ex As Exception
            stopScanning()
            MessageBox.Show("quickRead Problem: " & vbCrLf & ex.Message & vbCrLf & readAddress.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function
    Private Function quickRead32(ByVal readAddress As Integer, Optional ByVal sTarget As String = "project64") As Integer
        quickRead32 = 0

        Dim p As Process = Nothing
        If Process.GetProcessesByName(sTarget).Count > 0 Then
            p = Process.GetProcessesByName(sTarget)(0)
        Else
            stopScanning()
            MessageBox.Show(sTarget & " is not open!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Function
        End If
        Try
            quickRead32 = Memory.ReadInt32(p, readAddress)
        Catch ex As Exception
            stopScanning()
            MessageBox.Show("quickRead Problem: " & vbCrLf & ex.Message & vbCrLf & readAddress.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function
    Private Sub quickWrite16(ByVal writeAddress As Integer, ByVal writeValue As Int16, Optional ByVal sTarget As String = "project64")
        writeAddress = romAddrStart + writeAddress
        Dim p As Process = Nothing
        If Process.GetProcessesByName(sTarget).Count > 0 Then
            p = Process.GetProcessesByName(sTarget)(0)
        Else
            MessageBox.Show(sTarget & " is not open!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            WriteInt16(p, writeAddress, writeValue)
        Catch ex As Exception
            MessageBox.Show("quickWrite Problem: " & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    Private Sub quickWrite(ByVal writeAddress As Integer, ByVal writeValue As Integer, Optional ByVal sTarget As String = "project64")
        Dim p As Process = Nothing
        If Process.GetProcessesByName(sTarget).Count > 0 Then
            p = Process.GetProcessesByName(sTarget)(0)
        Else
            MessageBox.Show(sTarget & " is not open!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
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
            readChestData()
            If Not keepRunning Then
                stopScanning()
                Exit Sub
            End If
            updateEverything()
        End If
    End Sub
    Private Sub goScan(ByVal auto As Boolean)
        zeldazFails = 0
        keepRunning = True
        readChestData()
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If
        getAge()
        checkMQs()
        readChestData()
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If
        updateEverything()
        If Not auto Then Exit Sub
        If tmrAutoScan.Enabled = False Then
            tmrAutoScan.Enabled = True
            ' checkMQs()
        Else
            stopScanning()
        End If
    End Sub

    Private Function checkZeldaz() As Byte
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
        If zeldaz1 = 1514490948 Then
            If zeldaz2 = 16730 Then checkZeldaz = 2
        End If

        ' Check for the rando version only if a full-true
        If checkZeldaz = 2 Then
            getRandoVer()
            getRainbowBridge()
        End If
    End Function
    Private Function isLoadedGame() As Boolean
        ' Checks the game state (2=game menu, 1=title screen, 0=gameplay), if 0 and a successful ZELDAZ check, then true
        isLoadedGame = False
        If goRead(&H11B92C, 1) = 0 And checkZeldaz() = 2 Then isLoadedGame = True
    End Function
    Private Sub debugInfo()
        emulator = String.Empty
        If IS_64BIT = False Then
            emulator = "project64"
        Else
            attachToBizHawk()
            If emulator = String.Empty Then attachToRMG()
            If emulator = String.Empty Then attachToM64P()
            If emulator = String.Empty Then attachToRetroArch()
        End If
        If Not emulator = String.Empty Then
            Me.Text = "Tracker of Time (" & emulator & ")"
            Select Case LCase(emulator)
                Case "emuhawk", "rmg", "mupen64plus-gui", "retroarch - mupen64plus", "retroarch - parallel"
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

        ' Update the MQ Dungeons
        Dim MQs As String = String.Empty
        Dim testSpot As Integer = &H40B6E0
        Dim hits As Integer = 0
        Dim allZeros As Integer = 0

        For i = 0 To aMQ.Length - 1
            aMQ(i) = False
        Next

        For i = 0 To 1
            If randoVer = "AP" Then i = 1
            allZeros = 0
            If i = 1 Then testSpot = &H40B220
            MQs = Hex(goRead(testSpot))
            If Not keepRunning Then
                stopScanning()
                Exit Sub
            End If
            If MQs = "0" Then inc(allZeros)
            While MQs.Length < 8
                MQs = "0" & MQs
            End While
            If Mid(MQs, 1, 2) = "01" Then
                inc(hits)
                aMQ(0) = True
            End If
            If Mid(MQs, 3, 2) = "01" Then
                inc(hits)
                aMQ(1) = True
            End If
            If Mid(MQs, 5, 2) = "01" Then
                inc(hits)
                aMQ(2) = True
            End If
            If Mid(MQs, 7, 2) = "01" Then
                inc(hits)
                aMQ(3) = True
            End If

            MQs = Hex(goRead(testSpot))
            If Not keepRunning Then
                stopScanning()
                Exit Sub
            End If
            If MQs = "0" Then inc(allZeros)
            While MQs.Length < 8
                MQs = "0" & MQs
            End While
            If Mid(MQs, 1, 2) = "01" Then
                inc(hits)
                aMQ(4) = True
            End If
            If Mid(MQs, 3, 2) = "01" Then
                inc(hits)
                aMQ(5) = True
            End If
            If Mid(MQs, 5, 2) = "01" Then
                inc(hits)
                aMQ(6) = True
            End If
            If Mid(MQs, 7, 2) = "01" Then
                inc(hits)
                aMQ(7) = True
            End If

            MQs = Hex(goRead(testSpot))
            If Not keepRunning Then
                stopScanning()
                Exit Sub
            End If
            If MQs = "0" Then inc(allZeros)
            While MQs.Length < 8
                MQs = "0" & MQs
            End While
            If Mid(MQs, 1, 2) = "01" Then
                inc(hits)
                aMQ(8) = True
            End If
            If Mid(MQs, 3, 2) = "01" Then
                inc(hits)
                aMQ(9) = True
            End If
            If Mid(MQs, 7, 2) = "01" Then
                inc(hits)
                aMQ(10) = True
            End If

            MQs = Hex(goRead(testSpot))
            If Not keepRunning Then
                stopScanning()
                Exit Sub
            End If
            If MQs = "0" Then inc(allZeros)
            While MQs.Length < 8
                MQs = "0" & MQs
            End While
            If Mid(MQs, 3, 2) = "01" Then
                inc(hits)
                aMQ(11) = True
            End If

            If hits > 0 Or allZeros = 4 Then Exit For
            If randoVer = "OOTR" Then Exit For
        Next

        updateMQs()
    End Sub
    Private Sub updateEverything()
        If checkZeldaz() = 2 And isLoadedGame() Then
            updateItems()
            updateQuestItems()
            updateDungeonItems()
            updateDACs()
            updateLabels()
            updateLabelsDungeons()
        End If
    End Sub

    Private Function checkLoc(ByVal loc As String) As Boolean
        ' Checks for a specific key by location to see if it is checked
        checkLoc = False

        ' Checks normal keys
        For Each key In aKeys.Where(Function(k As keyCheck) k.loc.Equals(loc))
            Return key.checked
        Next

        ' Checks dungeon keys
        For i = 0 To 11
            For Each key In aKeysDungeons(i).Where(Function(k As keyCheck) k.loc.Equals(loc))
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

    Private Sub shutupNavi()
        ' With great power, comes little care for what others have to day. Shut up Navi's timed complaints.
        If emulator = "project64" Then
            quickWrite16(&H11A608 + 2, 0)
        ElseIf emulator = "variousX64" Then
            WriteMemory(Of Int16)(romAddrStart64 + &H11A608 + 2, 0)
        End If
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        'Dim roomCode As String = Hex(goRead(CUR_ROOM_ADDR))
        'fixHex(roomCode)
        'Dim locationCode As Integer = Convert.ToInt16(Mid(roomCode, 1, 4), 16)
        'locationCode = (locationCode * 28) + 212 + &H11A5D0
        'MsgBox(roomCode & vbCrLf & Hex(locationCode))
        Dim roomCode As String = Hex(goRead(CUR_ROOM_ADDR))
        fixHex(roomCode)
        Dim locationCode As Integer = goRead(CUR_ROOM_ADDR + 2, 15)
        locationCode = (locationCode * 28) + 212 + &H11A5D0
        MsgBox(roomCode & vbCrLf & Hex(locationCode))
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
        'Exit Sub
        pnlHidden.Visible = False
        Button2.Visible = False

        If showSetting Then
            Me.Width = pnlDekuTree.Location.X + pnlDekuTree.Width + 423
        Else
            Me.Width = pnlDekuTree.Location.X + pnlDekuTree.Width + 22
        End If
        Me.Height = rtbOutputRight.Location.Y + rtbOutputRight.Height + 46
    End Sub

    Private Sub displayChecks(ByVal area As String, ByVal showChecked As Boolean, Optional setCounter As Boolean = True)
        Dim displayName As String = String.Empty
        Dim sOut As String = String.Empty

        ' Counter set to delay refreshing the output
        If setCounter Then rtbRefresh = 3

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
                displayName = "Market"
            Case "TT"
                displayName = "Temple of Time"
            Case "HC"
                displayName = "Hyrule Castle"
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
                displayName = "Outside Ganon's Castle"
            Case "QGS"
                displayName = "Quest: Gold Skulltulas"
            Case "QM"
                displayName = "Quest: Masks"
            Case Else
                displayName = "Bad Area Code"
        End Select

        ' If right-clicked to show checked, not that it is found checks
        If showChecked Then displayName = displayName & " (Found)"

        ' Create a list and fill it with checks
        Dim outputLines As New List(Of String)
        scanArea(outputLines, area, showChecked)

        Dim theyLookSoGodDamnLikeTheSameList As Boolean = True
        For i = 0 To outputLines.Count - 1
            If i >= lastOutput.Count Then
                theyLookSoGodDamnLikeTheSameList = False
                Exit For
            End If
            If Not outputLines(i) = lastOutput(i) Then
                theyLookSoGodDamnLikeTheSameList = False
                Exit For
            End If
        Next

        If Not theyLookSoGodDamnLikeTheSameList Then
            ' Clear out the output boxes and set the display name
            emboldenList.Clear()
            rtbOutputLeft.Text = displayName & ":"
            Select Case area
                Case "HF"
                    rtbOutputRight.Text = "Big Poes:" & IIf(showChecked, " (Found)", "").ToString
                    Do While outputLines.Count < 13
                        outputLines.Add(String.Empty)
                    Loop
                    scanArea(outputLines, "QBPH", showChecked)
                Case "ZR"
                    rtbOutputRight.Text = "Fabulous Five Froggish Tenors:" & IIf(showChecked, " (Found)", "").ToString
                    Do While outputLines.Count < 13
                        outputLines.Add(String.Empty)
                    Loop
                    scanArea(outputLines, "QF", showChecked)
                Case Else
                    rtbOutputRight.Clear()
            End Select

            ' Output each line
            For Each line In outputLines
                rtbAddLine(line)
            Next
        End If
        lastOutput = outputLines

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
                                Select Case My.Settings.setSkulltula
                                    Case 0
                                        addCheck = False
                                    Case 1
                                        addCheck = True
                                    Case Else
                                        addCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                                End Select
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
                                If My.Settings.setLogic And checkLogic(.logic, .area) Then
                                    emboldenList.Add(prefix & .name & suffix)
                                End If
                            End If

                            If addCheck Then lines.Add(prefix & .name & suffix & Chr(32))
                        End If
                    End If
                End If
            End With
        Next
    End Sub
    Private Sub displayChecksDungeons(ByVal dungeon As Byte, ByVal showChecked As Boolean, Optional setCounter As Boolean = True)
        Dim displayName As String = String.Empty
        Dim sOut As String = String.Empty


        ' Counter set to delay refreshing the output
        If setCounter Then rtbRefresh = 3

        displayName = dungeonNumber2name(dungeon)

        ' Add 'MQ' to the display if it is a Master Quest dungeon
        If aMQ(dungeon) Then displayName = displayName & " MQ"

        ' If right-clicked to show checked, not that it is found checks
        If showChecked Then displayName = displayName & " (Found)"

        ' Create a list and fill it with checks
        Dim outputLines As New List(Of String)
        scanDungeon(outputLines, dungeon, showChecked)

        Dim theyLookSoGodDamnLikeTheSameList As Boolean = True
        For i = 0 To outputLines.Count - 1
            If i >= lastOutput.Count Then
                theyLookSoGodDamnLikeTheSameList = False
                Exit For
            End If
            If Not outputLines(i) = lastOutput(i) Then
                theyLookSoGodDamnLikeTheSameList = False
                Exit For
            End If
        Next

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
                            Select Case My.Settings.setSkulltula
                                Case 0
                                    addCheck = False
                                Case 1
                                    addCheck = True
                                Case Else
                                    addCheck = CBool(IIf(goldSkulltulas < 50, True, False))
                            End Select
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
                            If My.Settings.setLogic And checkLogic(.logic, .area) Then
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

        For i = 0 To aKeys.Length - 1
            aKeys(i) = New keyCheck
            aKeys(i).scan = True
        Next

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

    Private Function entranceCheck(ByVal area As String) As Boolean
        If randoVer = String.Empty Then Return False
        entranceCheck = False
        Application.DoEvents()

        ' Ways inside each area
        Select Case area
            Case "KF"
                ' Should always be true, at least for now. Young Link starts here, adult Link can get here with nothing
                'If canYoung Or canAdult Then Return True
                Return True
            Case "LW"
                ' If we can get to KF, return true. It should always be true, until entrance randomizer
                'If entranceCheck("KF") Then Return True
                Return True
            Case "SFM"
                ' LW to SFM will always be open, but check anyhow
                'If entranceCheck("LW") Then Return True
                ' Also can play Minuet of Forest
                'If allItems.Contains("h") And checkLoc("7706") Then Return True

                If entranceCheck("LW") Or item("minuet of forest") Then Return True
            Case "HF"
                ' This is a wide one...
                ' If can be an adult, you have access to Hyrule Field
                If canAdult Then Return True
                ' If settings is set to open, or the kid has moved from the entrance
                If My.Settings.setOpenKF Or checkLoc("6223") Then Return True
                ' Ocarina and Serenade of Water can get you to Lake Hylia and then out
                If item("serenade of water") Then Return True
                ' Ocarina and Nocturne of Shadow can get you to Graveyard, Kakariko Village, then Hyrule Field
                If item("nocturne of shadow") Then Return True
                ' Ocarina and Prelude of Light can get you to Temple of Time, Market, then Hyrule Field
                If item("prelude of light") Then Return True

                ' This one here is dumb, but to prevent an endless loop of checks, giving Hyrule Field all the checks

                ' Check for scales to swim from LW to Lake Hylia
                'If entranceCheck("LW") Then
                ''If checkLoc("7609") Or checkLoc("7610") Then Return True
                'End If
            Case "LLR", "MK", "TT", "KV", "GY", "ZR", "LH", "GV"
                ' Simply tie all of these to Hyrule Field
                'If entranceCheck("HF") Then Return True
                Return entranceCheck("HF")
            Case "HC"
                ' Need to be young Link and reach Market
                If entranceCheck("HF") Then Return canYoung

            Case "DMT"
                ' Adult Link can always reach Death Mountain Trail
                If canAdult Then Return True
                ' As young Link, can enter from Kakariko Village, but you will need Zelda's Letter or have had delivered it
                If entranceCheck("KV") Then
                    If allItems.Contains("y3") Or checkLoc("C05") Then Return True
                End If
                ' You can also reach it from Lost Woods to Goron City to Death Mountain Trail
                If entranceCheck("LW") Then
                    ' Check with being able to explode those rocks blocking the LW/GC shortcut
                    If canExplode() Then Return True
                    ' That will be all for now since you need to go one direction
                End If
            Case "DMC"
                ' Reach the top entrance through climbing Death Mountain
                If canTopDeathMountain() Then Return True
                ' Reach the warp pad with an ocarina and the Bolero of Fire
                If item("bolero of fire") Then Return True
                ' Enter from Goron City after stopping the rolling Goron
                If canAdult And checkLoc("7025") Then Return True
            Case "GC"
                ' If you can get to Death Mountain Trail, you can get into Goron City
                If entranceCheck("DMT") Then Return True
                ' If you can get to LW and break the rocks between the two areas
                If entranceCheck("LW") And canBreakRocks() Then Return True
                ' Checks if player has an ocarina and the Bolero of Fire and can get halfway up from Warp
                If item("bolero of fire") And canCraterWarpToTop(True) Then Return True
            Case "ZD"
                ' If you can reach Lake Hylia as young Link
                youngZD = False
                adultZD = False

                ' Lake Hylia and diving with young Link gets you in
                If entranceCheck("LH") And canYoung And item("dive") Then youngZD = True

                ' The other way in, you will need to reach ZR
                If Not entranceCheck("ZR") Then Return False
                ' An ocarina with Zelda's Lullaby
                If item("zelda's lullaby") Then
                    youngZD = True
                    adultZD = True
                End If
                ' If you can be young Link and have the cucco trick enabled

                ' If you can be adult Link and can do Hover Boot tricks
                If canAdult And canHoverTricks() Then adultZD = True

                If youngZD Or adultZD Then Return True
            Case "ZF"
                ' If you can enter as young or adult Link
                If canEnterZF(False) Or canEnterZF(True) Then Return True
            Case "GF"
                ' Auto False if you cannot adult
                If Not canAdult Then Return False
                ' Check Gerudo Valley
                If entranceCheck("GV") Then
                    ' Check for Longshot or if bridge is built
                    If item("longshot") Or checkLoc("6208") Or checkLoc("CARD") Then Return True
                    ' Check for an ocarina, Epona's Song, and Epona
                    If item("epona") Then Return True
                End If
                ' For the other direction, you will need an ocarina, Requiem of Spirit, and Lens of Truth access
                If item("requiem of spirit") And canLens() Then
                    ' Check for either Hover Boots or Longshot
                    If item("hover boots") Or item("longshot") Then Return True
                End If
            Case "HW"
                ' Checks the Gerudo Fortress side of things
                If entranceCheck("GF") Then
                    ' As an adult, you can open the gate. Or check that the gate is already open
                    If canAdult Or checkLoc("C00") Then Return True
                End If
                ' Check for crossing the desert to the front of the Haunted Wasteland
                If item("requiem of spirit") And canLens() Then
                    ' Check for either Hover Boots or Longshot
                    If item("hover boots") Or item("longshot") Then Return True
                End If
            Case "DC"
                ' Check for ocarina and Requiem of Spirit
                If item("requiem of spirit") Then
                    youngDC = True
                    adultDC = True
                End If
                ' Check HW side of things, will need a Lens of Truth check
                If entranceCheck("HW") And canLens() Then
                    ' Check for either Hover Boots or Longshot
                    If item("hover boots") Or item("longshot") Then adultDC = True
                End If
                If youngDC Or adultDC Then Return True
            Case "OGC"
                ' Just need to be an adult
                If canAdult Then Return True
            Case "DT0", "DT1", "DT2", "DT3", "DT4", "DT5"
                Dim DAC As Byte = CByte("&H" & Mid(area, 3))
                If arrDAC(0)(DAC) Then Return True
            Case "DDC0", "DDC1", "DDC2", "DDC3", "DDC4", "DDC5", "DDC6"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(1)(DAC) Then Return True
            Case "JB0", "JB1", "JB2", "JB3", "JB4", "JB5"
                Dim DAC As Byte = CByte("&H" & Mid(area, 3))
                If arrDAC(2)(DAC) Then Return True
            Case "FOT0", "FOT1", "FOT2", "FOT3", "FOT4", "FOT5", "FOT6", "FOT7", "FOT8", "FOT9", "FOTA", "FOTB"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(3)(DAC) Then Return True
            Case "FIT0", "FIT1", "FIT2", "FIT3", "FIT4", "FIT5", "FIT6"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(4)(DAC) Then Return True
            Case "WAT0", "WAT1", "WAT2", "WAT3", "WAT4", "WAT5", "WAT6", "WAT7", "WAT8"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(5)(DAC) Then Return True
            Case "SPT0", "SPT1", "SPT2", "SPT3", "SPT4", "SPT5", "SPT7", "SPT6", "SPT8", "SPT9"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(6)(DAC) Then Return True
            Case "SHT0", "SHT1", "SHT2", "SHT3", "SHT4", "SHT5", "SHT6", "SHT7", "SHT8", "SHT9", "SHTA", "SHTB", "SHTC", "SHTD"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(7)(DAC) Then Return True
            Case "BW0", "BW1", "BW2"
                Dim DAC As Byte = CByte("&H" & Mid(area, 3))
                If arrDAC(8)(DAC) Then Return True
            Case "IC"
                ' Need to reach Zora's Fountain as an adult
                Return canEnterZF(True)
            Case "GTG0", "GTG1", "GTG2", "GTG3", "GTG4", "GTG5", "GTG6", "GTG7"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(10)(DAC) Then Return True
            Case "IGC0", "IGC1", "IGC2", "IGC3"
                Dim DAC As Byte = CByte("&H" & Mid(area, 4))
                If arrDAC(11)(DAC) Then Return True
            Case "QBPH"
                ' All Quest: Big Poe Hunt checks need adult Link in the Hyrule Field with Epona access, bow, and a bottle
                If canAdult And entranceCheck("HF") And item("epona") And item("bow") And item("bottle") Then Return True
            Case "QF"
                ' All Quest: Frogs checks require you reach the area (same as buying beans) and have an ocarina
                If canBuyBeans() And allItems.Contains("h") Then Return True
            Case "QGS"
                ' All Quest: Gold Skulltulas need to have access to Kakariko Village
                If entranceCheck("KV") Then Return True
            Case "QM"
                ' All Quest: Masks need young Link, can access the Market
                If canYoung And entranceCheck("MK") Then Return True
            Case Else
                ' All entrances are entered in, now return false for any unexpected tests
                Return False
        End Select
    End Function
    Private Sub resetDAC(Optional dungeon As Byte = 99)
        ' Resets all DAC to false. The 0 is for Ice Cavern, as the current level of logic does not consider entrance rando, so IC becomes just 1 area
        Dim arrMaxes() As Byte = {5, 6, 5, 11, 6, 8, 9, 13, 2, 0, 7, 3}

        ' Loop's range variables, default all dungeons
        Dim iStart As Byte = 0
        Dim iEnd As Byte = CByte(arrDAC.Length - 1)
        ' If a dungeon number was passed into function, set it as the only range
        If dungeon <= 11 Then
            iStart = dungeon
            iEnd = dungeon
        End If

        ' Step through range
        For i = iStart To iEnd
            ReDim arrDAC(i)(arrMaxes(i))
            ' Set through dungeon's arrays
            For ii = 0 To arrDAC(i).Length - 1
                ' Reset to False
                arrDAC(i)(ii) = False
            Next
        Next
    End Sub
    Private Sub fullDAC(ByVal dungeon As Byte)
        For i = 0 To arrDAC(dungeon).Length - 1
            arrDAC(dungeon)(i) = True
        Next
    End Sub
    Private Sub updateDACs()
        dacDekuTree()
        dacDodongosCavern()
        dacJabuJabusBelly()
        dacForestTemple()
        dacFireTemple()
        dacWaterTemple()
        dacSpiritTemple()
        dacShadowTemple()
        dacBottomOfTheWell()
        dacGerudoTrainingGround()
        dacGanonsCastle()
    End Sub
    Private Sub dacDekuTree()
        ' Set up DAC for the Deku Tree
        resetDAC(0)
        ' Check if Mido has moved, either by the player or by Forest Open settings
        ' Check for both Kokiri Sword and the Deku Shield, if so you can move Mido
        If canYoung And (checkLoc("6220") Or (checkLoc("7416") And checkLoc("7420"))) Then
            ' 0: Deku Tree Lobby
            arrDAC(0)(0) = True
            If Not aMQ(0) Then
                ' Non-Master Quest checks
                ' 1: Deku Tree Slingshot Room
                ' 2: Deku Tree Basement Backroom
                ' 3: Deku Tree Boss Room

                ' 0 to 1
                If item("deku shield") Then arrDAC(0)(1) = True
                ' 0 to 2 | 0 to 3
                If (item("slingshot") And canBurn(True)) Or My.Settings.setDekuB1Skip Or checkLoc("10316") Then
                    arrDAC(0)(2) = True
                    arrDAC(0)(3) = True
                End If
            Else
                ' Master Quest checks
                ' 1: MQ Deku Tree Compass Room
                ' 2: MQ Deku Tree Basement Water Room Front
                ' 3: MQ Deku Tree Basement Water Room Back
                ' 4: MQ Deku Tree Basement Back Room
                ' 5: MQ Deku Tree Basement Ledge

                ' 0 to 1 | 0 to 2
                If item("slingshot") And canBurn(True) Then
                    arrDAC(0)(1) = True
                    arrDAC(0)(2) = True
                    ' 2 to 3
                    If arrDAC(0)(2) Then
                        If item("deku shield") Or item("hylian shield") Then
                            arrDAC(0)(3) = True
                            ' 3 to 4 | 3 to 5
                            If item("deku stick") Or item("din's fire") Then
                                arrDAC(0)(4) = True
                                arrDAC(0)(5) = True
                            End If
                        End If
                    End If
                End If

                ' 0 to 5
                If My.Settings.setDekuB1Skip Or checkLoc("10316") Then arrDAC(0)(5) = True

                ' 5 to 4
                If arrDAC(0)(5) Then
                    arrDAC(0)(4) = True
                    ' 4 to 3
                    If item("kokiri sword") Or canProjectile(0) Or (item("deku nuts") And item("deku stick")) Then arrDAC(0)(3) = True
                End If
            End If
        End If
    End Sub
    Private Sub dacDodongosCavern()
        ' Set up DAC for Dodongo's Cavern
        resetDAC(1)
        ' If you cannot get to Death Mountain Trail
        If entranceCheck("DMT") And (canAdult Or canExplode() Or item("lift")) Then
            ' 0: Dodongo's Cavern Beginning
            arrDAC(1)(0) = True
            If Not aMQ(1) Then
                ' Non-Master Quest checks
                ' 1: Dodongo's Cavern Lobby
                ' 2: Dodongo's Cavern Staircase Room
                ' 3: Dodongo's Cavern Climb
                ' 4: Dodongo's Cavern Far Bridge
                ' 5: Dodongo's Cavern Boss Area

                ' 0 to 1
                If canBreakRocks() Or item("lift") Then
                    arrDAC(1)(1) = True
                    ' 1 to 2
                    If canAdult Or (canExplode() Or item("lift") And (item("deku stick") Or (item("din's fire") And (item("slingshot") Or item("bomb") Or item("kokiri sword"))))) Then
                        arrDAC(1)(2) = True
                        ' 2 to 3
                        If canExplode() Or item("lift") Or item("din's fire") Or (My.Settings.setDCStaircase And item("bow")) Then
                            arrDAC(1)(3) = True
                            ' 3 to 4
                            If (item("hover boots") Or item("longshot") Or My.Settings.setDCSpikeJump) Or _
                                ((canBreakRocks() Or item("lift")) And (item("slingshot") Or (My.Settings.setDCSlingshotSkips And (item("deku stick") Or item("bomb") Or item("kokiri sword"))) Or item("bow"))) Then
                                arrDAC(1)(4) = True
                                ' 4 to 5
                                If canExplode() Then arrDAC(1)(5) = True
                            End If
                        End If
                    End If
                    If Not arrDAC(1)(5) Then
                        ' 1 to 5
                        If (checkLoc("10410") And canExplode()) Or checkLoc("10426") Then arrDAC(1)(5) = True
                    End If
                End If
            End If
        Else
            ' Master Quest checks
            ' 1: MQ Dodongo's Cavern Lobby
            ' 2: MQ Dodongo's Cavern Elevator
            ' 3: MQ Dodongo's Cavern Lower Right Side
            ' 4: MQ Dodongo's Cavern Bomb Bag Area
            ' 5: MQ Dodongo's Cavern Boss Area
            ' 6: MQ Dodongo's Cavern Boss Room

            ' 0 to 1
            If canBreakRocks() Or item("lift") Then
                arrDAC(1)(1) = True
                ' 1 to 2
                If canBreakRocks() Or item("lift") Then
                    arrDAC(1)(2) = True
                    ' 2 to 3
                    If (item("deku stick") Or item("din's fire")) And item("lift") Then arrDAC(1)(3) = True
                    ' 2 to 5
                    If (checkLoc("10410") And canExplode()) Or checkLoc("10426") Then
                        arrDAC(1)(5) = True
                    ElseIf canExplode() Or (item("lift") And My.Settings.setDCLightEyes) And (item("deku stick") Or item("din's fire") Or (canAdult And (My.Settings.setDCSpikeJump Or item("hammer") Or item("hookshot")))) Then
                        arrDAC(1)(5) = True
                    End If
                End If
                ' 1 to 3
                If canBreakRocks() Then
                    arrDAC(1)(3) = True
                    ' 3 to 4
                    If (item("lift") Or item("din's fire") Or canExplode()) And item("slingshot") Then arrDAC(1)(4) = True
                End If
                ' 1 to 4
                If canAdult Then
                    arrDAC(1)(4) = True
                    ' 4 to 3
                    arrDAC(1)(3) = True
                End If
                ' 5 to 6
                If arrDAC(1)(5) Then
                    If canAdult Or canExplode() Or item("din's fire") Or (item("deku stick") Or ((item("deku nuts") Or item("boomerang")) And (item("kokiri sword") Or item("slingshot")))) Then arrDAC(1)(6) = True
                End If
            End If
        End If
    End Sub
    Private Sub dacJabuJabusBelly()
        ' Set up DAC for Jabu-Jabu's Belly
        resetDAC(2)
        ' If you cannot get to Zora's Domain as a child
        If youngZD And allItems.Contains("u") Then

            ' 0: Jabu-Jabu's Belly Beginning
            arrDAC(2)(0) = True
            If Not aMQ(2) Then
                ' 1: Jabu-Jabu's Belly Main
                ' 2: Jabu-Jabu's Belly Depths
                ' 3: Jabu-Jabu's Belly Boss Area

                ' 0 to 1
                If canExplode() Or item("slingshot") Or item("boomerang") Then
                    arrDAC(2)(1) = True
                    ' 1 to 2
                    If item("boomerang") Then
                        arrDAC(2)(2) = True
                        ' 2 to 3
                        If item("kokiri sword") Or item("deku stick") Then arrDAC(2)(3) = True
                    End If
                End If
                ' 0 to 3
                If Not arrDAC(2)(3) Then
                    ' Save for the end to prevent scans and hopefully save cycles
                    arrDAC(2)(3) = checkLoc("10529")
                End If
            Else
                ' 1: MQ Jabu-Jabu's Belly Elevator Room
                ' 2: MQ Jabu-Jabu's Belly Main
                ' 3: MQ Jabu-Jabu's Belly Depth
                ' 4: MQ Jabu-Jabu's Belly Past Big Octo
                ' 5: MQ Jabu-Jabu's Belly Boss Area

                ' 0 to 1
                If item("slingshot") Then
                    arrDAC(2)(1) = True
                    ' 1 to 2
                    arrDAC(2)(2) = True
                    ' 2 to 3
                    If canExplode() And item("boomerang") And item("slingshot") Then
                        arrDAC(2)(3) = True
                        ' 3 to 4
                        If item("deku stick") Or (item("kokiri sword") And item("din's fire")) Then
                            arrDAC(2)(4) = True
                            ' 4 to 5
                            arrDAC(2)(5) = True
                        End If
                    End If
                    ' 1 to 5
                    If Not arrDAC(2)(5) Then
                        arrDAC(2)(5) = checkLoc("10529")
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub dacForestTemple()
        ' Set up DAC for Forest Temple
        resetDAC(3)
        Dim reach As Boolean = False
        ' Need to get to SFM as an adult with the hookshot
        If entranceCheck("SFM") And canAdult And allItems.Contains("k") Then
            ' 0: Forest Temple Lobby
            arrDAC(3)(0) = True
            If Not aMQ(3) Then
                ' 1: Forest Temple NW Outdoors
                ' 2: Forest Temple NE Outdoors
                ' 3: Forest Temple Outdoors High Balconies
                ' 4: Forest Temple Falling Room
                ' 5: Forest Temple Block Push Room
                ' 6: Forest Temple Straightened Hall
                ' 7: Forest Temple Outside Upper Ledge
                ' 8: Forest Temple Bow Region
                ' 9: Forest Temple Boss Region

                'If Not My.Settings.setSmallKeys = 1 Then
                ' If not keysanity, do a check for reaching all areas
                'If item("bow") And item("lift", 1) And (item("song of time") Or item("hover boots") Or ((item("longshot") Or (item("hookshot") And My.Settings.setFoTVines)) And (item("iron boots") Or item("dive", 2)) Or My.Settings.setFoTBackdoor)) Then
                'fullDAC(3)
                'Exit Sub
                'End If
                'End If

                ' 0 to 1
                If item("song of time") Then
                    arrDAC(3)(1) = True
                    ' 1 to 3
                    arrDAC(3)(3) = True
                    ' 3 to 2
                    arrDAC(3)(2) = True
                ElseIf item("bow") Then
                    ' 0 to 2
                    arrDAC(3)(2) = True
                End If
                If arrDAC(3)(2) Then
                    ' 2 to 3
                    If item("longshot") Or (item("hookshot") And My.Settings.setFoTVines) Then arrDAC(3)(3) = True
                    ' 2 to 1
                    If item("iron boots") Or item("dive", 2) Then arrDAC(3)(1) = True
                End If
                ' 0 to 5
                If dungeonKeyCounter(3, "01") Then
                    arrDAC(3)(5) = True
                    ' This reach is for checking access to area 7 (Outside Upper Ledge)
                    ' 5 to 7
                    If item("hover boots") Then reach = True
                    If item("lift") Then
                        ' 5 to 6
                        If dungeonKeyCounter(3, "0102") And item("bow") Then
                            arrDAC(3)(6) = True
                            reach = True
                        ElseIf My.Settings.setFoTBackdoor Then
                            ' 5 to 7: No need to check if 5 to 6 covers it as well
                            reach = True
                        End If
                        ' Check all those area 7 reaches because it gives access to many areas checked previously
                        If reach Then
                            arrDAC(3)(7) = True
                            arrDAC(3)(1) = True
                            arrDAC(3)(2) = True
                            arrDAC(3)(3) = True
                        End If
                        ' 5 to 8
                        If dungeonKeyCounter(3, "010203") Then
                            arrDAC(3)(8) = True
                            ' 8 to 4
                            If (item("bow") Or item("din's fire")) And dungeonKeyCounter(3, "0001020304") Then arrDAC(3)(4) = True
                        End If
                    End If
                End If
                ' 3 to 4
                If arrDAC(3)(3) And item("hover boots") And item("scarecrow") And My.Settings.setFoTFrame Then arrDAC(3)(4) = True
                ' 4 to 2
                If arrDAC(3)(4) Then arrDAC(3)(2) = True
                ' 0 to 9
                If checkLoc("10628") Or (checkLoc("10629") And checkLoc("10630") And checkLoc("10631") And item("bow")) Then arrDAC(3)(9) = True
            Else
                ' 1: MQ Forest Temple Central Ara
                ' 2: MQ Forest Temple After Block Puzzle
                ' 3: MQ Forest Temple Outdoor Ledge
                ' 4: MQ Forest Temple NW Outdoors
                ' 5: MQ Forest Temple NE Outdoors
                ' 6: MQ Forest Temple Outdoors Top Ledges
                ' 7: MQ Forest Temple NE Outdoor Ledge
                ' 8: MQ Forest Temple Bow Region
                ' 9: MQ Forest Temple Falling Room
                ' A: MQ Forest Temple Boss Region
                ' B: MQ Forest Temple Big Key Chest Area        I added in this area to give an extra check without having to code more checkLogic for keys

                ' 0 to 1
                If dungeonKeyCounter(3, "06") Then
                    arrDAC(3)(1) = True
                    If item("bow") Then
                        ' 1 to 4 | 1 to 5
                        arrDAC(3)(4) = True
                        arrDAC(3)(5) = True
                    End If

                    ' 1 to 2
                    If item("lift") Or (My.Settings.setFoTMQPuzzle And item("bombchu") And item("hookshot")) Then
                        arrDAC(3)(2) = True
                        ' 2 to B | 2 to 3
                        If dungeonKeyCounter(3, "0602") Then ', "03") Then
                            arrDAC(3)(11) = True
                            arrDAC(3)(3) = True
                        ElseIf My.Settings.setFoTMQTwisted Then
                            If My.Settings.setFoTBackdoor Or item("hookshot") Then arrDAC(3)(3) = True
                        End If
                        ' 2 to 4
                        If dungeonKeyCounter(3, "0602") Then arrDAC(3)(4) = True
                        ' 2 to 8
                        If dungeonKeyCounter(3, "060201") Then ', "03") Then
                            arrDAC(3)(8) = True
                            ' 8 to 9
                            If dungeonKeyCounter(3, "06020100") And (item("bow") Or item("din's fire")) Then arrDAC(3)(9) = True ', "04") Then arrDAC(3)(9) = True
                        End If
                    End If
                    ' 0 to 3
                    If My.Settings.setFoTMQTwisted And item("hover boots") Then arrDAC(3)(3) = True
                    ' 3 to 4
                    If arrDAC(3)(3) Then arrDAC(3)(4) = True
                    If arrDAC(3)(4) Then
                        ' 4 to 5
                        If item("iron boots") Or item("longshot") Or item("dive", 2) Or (My.Settings.setFoTMQWell And item("hookshot")) Then arrDAC(3)(5) = True
                        ' 4 to 6
                        If item("fire arrows") Then
                            arrDAC(3)(6) = True
                            ' 6 to 5
                            arrDAC(3)(5) = True
                        End If
                    End If
                    If arrDAC(3)(5) Then
                        If item("longshot") Then
                            ' 5 to 6 | 5 to 7
                            arrDAC(3)(6) = True
                            arrDAC(3)(7) = True
                        End If
                        ' 5 to 6
                        If Not arrDAC(3)(6) Then
                            If item("hookshot") And (item("hover boots") Or item("song of time") Or My.Settings.setFoTVines) Then arrDAC(3)(6) = True
                        End If
                    End If
                    If arrDAC(3)(6) Then
                        ' 6 to 5
                        arrDAC(3)(5) = True
                        ' 6 to 7
                        If My.Settings.setFoTLedge Then arrDAC(3)(7) = True
                    End If
                    If arrDAC(3)(7) Then
                        ' 7 to 5
                        arrDAC(3)(5) = True
                        ' 7 to 9
                        If item("song of time") Then
                            arrDAC(3)(9) = True
                        End If
                    End If
                    If arrDAC(3)(9) Then
                        ' 9 to 7
                        arrDAC(3)(7) = True
                        ' 7 to 5
                        arrDAC(3)(5) = True
                    End If
                    ' 0 to A
                    If checkLoc("10628") Or (checkLoc("10629") And checkLoc("10630") And checkLoc("10631") And item("bow")) Then arrDAC(3)(9) = True
                End If
            End If
        End If
    End Sub
    Private Sub dacFireTemple()
        ' Set up DAC for Fire Temple
        resetDAC(4)
        Dim reach As Boolean = False
        ' Need to get to Fire Temple
        If canFewerGoron() And (item("bolero of fire") Or (entranceCheck("DMC") And item("goron tunic") And (item("hover boots") Or item("hookshot")))) Then
            ' 0: Fire Temple Lower
            arrDAC(4)(0) = True
            If Not aMQ(4) Then
                ' 1: Fire Temple Big Lava Room
                ' 2: Fire Temple Middle
                ' 3: Fire Temple Upper
                ' 4: Fire Temple Lobby Back Room

                ' 0 to 1
                If dungeonKeyCounter(4, "29") Then ' , "23") Then
                    arrDAC(4)(1) = True
                    ' 1 to 2
                    If item("goron tunic") And dungeonKeyCounter(4, "293024") And (item("lift") Or My.Settings.setFiTClimb) And (canExplode() Or item("bow") Or item("hookshot")) Then
                        arrDAC(4)(2) = True
                        ' 2 to 3
                        If dungeonKeyCounter(4, "29302427263125") Then ' , "23") Then
                            arrDAC(4)(3) = True
                        ElseIf dungeonKeyCounter(4, "293024272631") And (My.Settings.setFiTMaze Or (item("hover boots") And item("hammer"))) Then
                            arrDAC(4)(3) = True
                        End If
                    End If
                End If
                ' 0 to 4
                If item("hammer") Then
                    If My.Settings.setSmallKeys = 1 Then
                        arrDAC(4)(4) = dungeonKeyCounter(4, "2930242726312523")
                    Else
                        arrDAC(4)(4) = True
                    End If
                End If
            Else
                ' 1: MQ Fire Temple Locked Door
                ' 2: MQ Fire Temple Big Lava Room
                ' 3: MQ Fire Temple Lower Maze
                ' 4: MQ Fire Temple Upper Maze
                ' 5: MQ Fire Temple Upper
                ' 6: MQ Fire Temple Boss Room

                ' 0 to 1
                If dungeonKeyCounter(4, "23") Then arrDAC(4)(1) = True
                ' 0 to 2
                If item("hammer") Then
                    arrDAC(4)(2) = True
                    ' 2 to 3
                    reach = False
                    If item("goron tunic") And dungeonKeyCounter(4, "30") And (canBurn() Or (My.Settings.setFiTMQClimb And item("hover boots"))) Then
                        arrDAC(4)(3) = True
                        ' 3 to 4
                        If ((canExplode() Or My.Settings.setRustedSwitches) And item("hookshot")) Or (My.Settings.setFiTMQMazeHovers And item("hover boots")) Or My.Settings.setFiTMQMazeJump Then
                            arrDAC(4)(4) = True
                            ' 4 to 5
                            If dungeonKeyCounter(4, "3026") And ((item("bow") And item("hookshot")) Or item("fire arrows")) Then arrDAC(4)(5) = True
                        End If
                    End If
                    ' 0 to 6
                    If item("goron tunic") And item("hammer") And ((canBurn() And (item("hover boots") Or checkLoc("10710"))) Or arrDAC(4)(5)) Then arrDAC(4)(6) = True
                End If
            End If
        End If
    End Sub
    Private Sub dacWaterTemple()
        ' Set up DAC for Water Temple
        resetDAC(5)
        Dim reach As Boolean = False
        ' Need to be adult Link, who can always reach Lake Hylia unless ER
        If item("iron boots") And (checkLoc("231") Or allItems.Contains("k")) Then
            ' Iron Boots with either already opened or Hookshot
            reach = True
        ElseIf item("dive", 2) And (checkLoc("231") Or allItems.Contains("l")) Then
            ' Golden Scale with already open or Longshot
            reach = True
        End If
        If reach Then
            ' 0: Water Temple Lobby
            arrDAC(5)(0) = True
            If Not aMQ(5) Then
                ' 1: Water Temple Highest Water Level
                ' 2: Water Temple Dive
                ' 3: Water Temple North Basement
                ' 4: Water Temple Cracked Wall
                ' 5: Water Temple Dragon Statue
                ' 6: Water Temple Middle Water Level
                ' 7: Water Temple Falling Platform Room
                ' 8: Water Temple Dark Link Region

                ' 0 to 1
                If item("bow") Or item("hookshot") Or item("hover boots") Or (canBurn() And canExplode()) Then
                    arrDAC(5)(1) = True
                    ' 1 to 7
                    If dungeonKeyCounter(5, "05") Then
                        arrDAC(5)(7) = True
                        ' 7 to 8
                        If item("hookshot") And dungeonKeyCounter(5, "0502") Then
                            arrDAC(5)(8) = True
                            ' 8 to 5
                            If canFewerZora() And item("bow") And item("song of time") And (My.Settings.setWaTDragonDive Or item("iron boots")) Then arrDAC(5)(5) = True
                        End If
                    End If
                End If
                ' 0 to 2
                If canFewerZora() And ((My.Settings.setWaTTorch And item("longshot")) Or item("iron boots")) Then
                    arrDAC(5)(2) = True
                    ' 2 to 3
                    If dungeonKeyCounter(5, "01") And (item("longshot") Or (My.Settings.setWaTBKR And item("hover boots"))) And (item("iron boots") Or item("zelda's lullaby")) Then arrDAC(5)(3) = True
                    ' 2 to 4
                    If item("zelda's lullaby") And (item("hookshot") Or item("hover boots")) And (My.Settings.setWaTCrackNothing Or (My.Settings.setWaTCrackHovers And item("hover boots"))) Then arrDAC(5)(4) = True
                    ' 2 to 5
                    If item("zelda's lullaby") And item("lift") And item("iron boots") And item("hookshot") Then arrDAC(5)(5) = True
                    ' 2 to 6
                    If (item("bow") Or item("din's fire") Or (dungeonKeyCounter(5, "06") And item("hookshot"))) And item("zelda's lullaby") Then
                        arrDAC(5)(6) = True
                        ' 6 to 4
                        arrDAC(5)(4) = True
                    End If
                End If
            Else
                ' 1: MQ Water Temple Dive
                ' 2: MQ Water Temple Lowered Water Levels
                ' 3: MQ Water Temple Dark Link Region
                ' 4: MQ Water Temple Basement Gated Areas

                ' 0 to 1
                If canFewerZora() And item("iron boots") Then
                    arrDAC(5)(1) = True
                    ' 1 to 2
                    If item("zelda's lullaby") Then arrDAC(5)(2) = True
                End If
                ' 0 to 3
                If dungeonKeyCounter(5, "21") And item("longshot") Then
                    arrDAC(5)(3) = True
                    ' 3 to 4
                    If canFewerZora() And item("din's fire") And item("iron boots") Then arrDAC(5)(4) = True
                End If
            End If
        End If

    End Sub
    Private Sub dacSpiritTemple()
        ' Set up DAC for Spirit Temple
        resetDAC(6)

        If youngDC Or adultDC Then
            ' 0: Spirit Temple Lobby
            arrDAC(6)(0) = True

            youngSpTCentre = False
            adultSpTCentre = False

            If Not aMQ(6) Then
                ' 1: Child Spirit Temple
                ' 2: Child Spirit Before Locked Door
                ' 3: Child Spirit Temple Climb
                ' 4: Early Adult Spirit Temple
                ' 5: Spirit Temple Central Chamber
                ' 6: Spirit Temple Outdoor Hands
                ' 7: Spirit Temple Beyond Central Locked Door
                ' 8: Spirit Temple Beyond Final Locked Door
                ' 9: Spirit Temple Boss Platform

                ' For various ways to reach area 8
                Dim reach As Boolean = False
                ' 0 to 1
                If youngDC Then
                    arrDAC(6)(1) = True
                    ' 1 to 2
                    arrDAC(6)(2) = True
                    ' 2 to 3
                    If dungeonKeyCounter(6, "30") Then
                        arrDAC(6)(3) = True
                        ' If young Link can make it here
                        youngSpTCentre = True
                        ' 3 to 5
                        If canExplode() Then arrDAC(6)(5) = True
                    End If
                End If
                ' 0 to 5 Shortcut
                If checkLoc("10909") And checkLoc("10911") And checkLoc("10912") Then
                    arrDAC(6)(5) = True
                    ' 5 to 3 Shortcut
                    arrDAC(6)(3) = True
                    ' If adult Link can make it here
                    adultSpTCentre = True
                    ' 5 to 7 via Centre Path
                    If dungeonKeyCounter(6, "27") And item("hookshot") Then
                        arrDAC(6)(7) = True
                        ' 7 to 8's 2 doors
                        If dungeonKeyCounter(6, "2724") Then reach = True
                    End If
                    ' 5 to 4 without Gauntlets
                    If item("hookshot") Then arrDAC(6)(4) = True
                End If

                ' 0 to 4
                If item("lift", 2) Then
                    arrDAC(6)(4) = True
                    ' 4 to 5
                    If dungeonKeyCounter(6, "13") Then
                        arrDAC(6)(5) = True
                        ' 5 to 3 Adult
                        arrDAC(6)(3) = True
                        ' If adult Link can make it here
                        adultSpTCentre = True
                        ' 5 to 7 via Right Path
                        If dungeonKeyCounter(6, "1327") And item("lift", 2) Then
                            arrDAC(6)(7) = True
                            ' 7 to 8' 3 doors
                            If dungeonKeyCounter(6, "132724") Then reach = True
                        End If
                    End If
                End If
                If arrDAC(6)(5) Then
                    ' 5 to 6
                    arrDAC(6)(6) = True
                    ' 5 to 9
                    If item("longshot") And checkLoc("10923") And checkLoc("10904") Then arrDAC(6)(9) = True
                End If
                If reach Then
                    ' 7 to 8
                    If My.Settings.setSpTWall Or item("longshot") Or item("bombchu") Or ((item("bomb") Or item("deku nuts") Or item("din's fire")) And (item("bow") Or item("hookshot") Or item("hammer"))) Then
                        arrDAC(6)(8) = True
                        ' 8 to 9
                        If item("mirror shield") And canExplode() Then arrDAC(6)(9) = True
                    End If
                End If
            Else
                ' 1: MQ Child Spirit Temple
                ' 2: MQ Adult Spirit Temple
                ' 3: MQ Spirit Temple Shared
                ' 4: MQ Lower Adult Spirit Temple
                ' 5: MQ Spirit Temple Boss Area
                ' 6: MQ Spirit Temple Boss Platform
                ' 7: MQ Mirror Shield Hand
                ' 8: MQ Silver Gauntlets Hand

                ' 0 to 1
                If youngDC Then
                    arrDAC(6)(1) = True
                    ' 1 to 3
                    If item("bombchu") And dungeonKeyCounter(6, "3018") Then
                        arrDAC(6)(3) = True
                        ' If young Link can make it here
                        youngSpTCentre = True
                        ' 3 to 8 as young Link
                        If dungeonKeyCounter(6, "301821") And (item("song of time") Or My.Settings.setSpTMQSunRoom) Then arrDAC(6)(8) = True
                    End If
                End If
                If item("longshot") And (item("lift", 2) Or checkLoc("10909")) And (item("bombchu") Or checkLoc("10912")) Then
                    arrDAC(6)(2) = True
                    ' 2 to 3
                    arrDAC(6)(3) = True
                    ' If adult Link can make it here
                    adultSpTCentre = True
                    ' 3 to 1: Added this one in to get adult path into the area for checks
                    If dungeonKeyCounter(6, "18") Then arrDAC(6)(1) = True
                    ' 3 to 8 as adult Link
                    If dungeonKeyCounter(6, "21") Then arrDAC(6)(8) = True
                    ' 2 to 4
                    If item("mirror shield") And (item("fire arrows") Or (My.Settings.setSpTMQLowAdult And item("din's fire") And item("bow"))) Then arrDAC(6)(4) = True
                    ' 2 to 5
                    If dungeonKeyCounter(6, "2728") And item("zelda's lullaby") And item("hammer") Then
                        arrDAC(6)(5) = True
                        ' 5 to 6
                        If item("mirror shield") Then arrDAC(6)(6) = True
                    End If
                    ' 2 to 6
                    If checkLoc("10923") And checkLoc("10904") Then arrDAC(6)(6) = True
                    ' 2 to 7
                    If dungeonKeyCounter(6, "27") And item("song of time") And canLens() And (canExplode() Or item("deku nuts")) Then arrDAC(6)(7) = True
                End If
            End If
        End If
    End Sub
    Private Sub dacShadowTemple()
        ' Set up DAC for Shadow Temple
        resetDAC(7)
        If item("nocturne of shadow") And item("din's fire") Then
            ' 0: Shadow Temple Entryway
            arrDAC(7)(0) = True
            If Not aMQ(7) Then
                ' 1: Shadow Temple Beginning
                ' 2: Shadow Temple First Beamos
                ' 3: Shadow Temple Huge Pit
                ' 4: Shadow Temple Invisible Spikes
                ' 5: Shadow Temple Wind Tunnel
                ' 6: Shadow Temple After Wind
                ' 7: Shadow Temple Boat
                ' 8: Shadow Temple Beyond Boat
                ' 9: Shadow Temple Boss Room

                ' 0 to 1
                If (item("lens of truth") Or My.Settings.setShTLensless) And (item("hover boots") Or item("hookshot")) Then
                    arrDAC(7)(1) = True
                    ' 1 to 2
                    If item("hover boots") Then
                        arrDAC(7)(2) = True
                        ' 2 to 3
                        If canExplode() And dungeonKeyCounter(7, "22") Then
                            arrDAC(7)(3) = True
                            ' 3 to 4
                            If dungeonKeyCounter(7, "2223") Then
                                arrDAC(7)(4) = True
                                ' 4 to 5
                                If item("hookshot") And dungeonKeyCounter(7, "222324") Then
                                    arrDAC(7)(5) = True
                                    ' 5 to 6
                                    arrDAC(7)(6) = True
                                    ' 6 to 7
                                    If dungeonKeyCounter(7, "22232421") Then
                                        arrDAC(7)(7) = True
                                        ' 7 to 8
                                        If item("zelda's lullaby") Then
                                            arrDAC(7)(8) = True
                                            ' 8 to 9
                                            If dungeonKeyCounter(7, "2223242125") And (item("bow") Or item("scarecrow", 2) Or checkLoc("11016") Or (My.Settings.setShTStatue And item("bombchu"))) Then arrDAC(7)(9) = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' 1: MQ Shadow Temple Beginning
                ' 2: MQ Shadow Temple Dead Hand Area
                ' 3: MQ Shadow Temple First Beamos
                ' 4: MQ Shadow Temple Upper Huge Pit
                ' 5: MQ Shadow Temple Invisible Blades
                ' 6: MQ Shadow Temple Lower Huge Pit
                ' 7: MQ Shadow Temple Falling Spikes
                ' 8: MQ Shadow Temple Invisible Spikes
                ' 9: MQ Shadow Temple Wind Tunnel and After Wind
                ' A: MQ Shadow Temple Boat
                ' B: MQ Shadow Temple Beyond Boat
                ' C: MQ Shadow Temple Invisible Maze
                ' D: MQ Shadow Temple Near Boss

                ' 0 to 1
                If (item("lens of truth") Or My.Settings.setShTLensless) And (item("hover boots") Or item("hookshot")) Then
                    arrDAC(7)(1) = True
                    ' 1 to 2
                    If canExplode() And dungeonKeyCounter(7, "25") Then arrDAC(7)(2) = True
                    ' 1 to 3
                    If item("fire arrows") Or item("hover boots") Or (My.Settings.setShTMQGap And item("longshot")) Then
                        arrDAC(7)(3) = True
                        ' 3 to 4
                        If canExplode() And dungeonKeyCounter(7, "22") Then
                            arrDAC(7)(4) = True
                            ' 4 to 5
                            If item("song of time") Then arrDAC(7)(5) = True
                            ' 4 to 6
                            If canBurn() Then
                                arrDAC(7)(6) = True
                                ' 6 to 7
                                arrDAC(7)(7) = True
                                ' 6 to 8
                                If item("hover boots") And dungeonKeyCounter(7, "2223") Then
                                    arrDAC(7)(8) = True
                                    ' 8 to 9
                                    If item("hookshot") And dungeonKeyCounter(7, "222324") Then
                                        arrDAC(7)(9) = True
                                        ' 9 to A
                                        If dungeonKeyCounter(7, "22232421") Then
                                            arrDAC(7)(10) = True
                                            ' A to B
                                            If item("zelda's lullaby") Then
                                                arrDAC(7)(11) = True
                                                ' B to C
                                                If item("bow") And item("song of time") And item("longshot") Then arrDAC(7)(12) = True
                                                ' B to D
                                                If item("bow") Or (My.Settings.setShTStatue And item("bombchu")) Or checkLoc("11016") Then arrDAC(7)(13) = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub dacBottomOfTheWell()
        ' Set up DAC for the Bottom of the Well
        resetDAC(8)
        ' Need to be young Link, be able to access Kakariko Village, and either have the well drained or the Song of Storms to drain it
        If canYoung And entranceCheck("KV") And (checkLoc("C02") Or item("song of storms")) Then
            ' 0: Bottom of the Well
            arrDAC(8)(0) = True
            If Not aMQ(8) Then
                ' 1: Bottom of the Well Main Area
                ' 2: Bottom of the Well Main Area with Lens of Truth checks

                ' 0 to 1
                If canAttackYoung(True) Or item("deku nuts") Then
                    arrDAC(8)(1) = True
                    ' 1 to 2
                    If My.Settings.setBotWLensless Or item("lens of truth") Then arrDAC(8)(2) = True
                End If
            Else
                ' 0: MQ Bottom of the Well and Perimeter
                ' 1: MQ Bottom of the Well Middle

                ' 0 to 1
                If item("zelda's lullaby") Or (True And canExplode()) Then arrDAC(8)(1) = True
            End If
        End If
    End Sub
    Private Sub dacGerudoTrainingGround()
        ' Set up DAC for the Gerudo Training Ground
        resetDAC(10)
        ' Need to be adult Link, able to access Gerudo Fortress, and have the Gerudo's Membership Card
        If canAdult And entranceCheck("GF") And checkLoc("7722") Then
            ' 0: Gerudo Training Ground Lobby
            arrDAC(10)(0) = True
            If Not aMQ(10) Then
                ' 1: Gerudo Training Ground Central Maze Right
                ' 2: Gerudo Training Ground Lava Room
                ' 3: Gerudo Training Ground Hammer Room
                ' 4: Gerudo Training Ground Eye Statue Lower
                ' 5: Gerudo Training Ground Eye Statue Upper
                ' 6: Gerudo Training Ground Heavy Block Room
                ' 7: Gerudo Training Ground Like Like Room

                ' 0 to 2
                If canExplode() Then arrDAC(10)(2) = True
                ' 0 to 6
                If item("hookshot") Then
                    arrDAC(10)(6) = True
                    ' 6 to 5
                    If canLens() And (item("hookshot") Or item("hover boots")) Then
                        arrDAC(10)(5) = True
                        ' 5 to 4
                        arrDAC(10)(4) = True
                        ' 4 to 3
                        arrDAC(10)(3) = True
                    End If
                    ' 6 to 7
                    If item("lift", 2) And canLens() And (item("hookshot") Or item("hover boots")) Then arrDAC(10)(7) = True
                End If
                ' 0 to 1
                If dungeonKeyCounter(10, "1003") Then
                    arrDAC(10)(1) = True
                    ' 1 to 2
                    arrDAC(10)(2) = True
                End If
                If arrDAC(10)(2) Then
                    ' 2 to 1
                    If item("song of time") Then arrDAC(10)(1) = True
                    ' 2 to 3
                    If item("longshot") Or (item("hookshot") And item("hover boots")) Then arrDAC(10)(3) = True
                End If
                ' 1 to 3
                If arrDAC(10)(1) And item("hookshot") Then arrDAC(10)(3) = True
                If arrDAC(10)(3) Then
                    ' 3 to 2
                    arrDAC(10)(2) = True
                    ' 3 to 2 to 1 recheck
                    If item("song of time") Then arrDAC(10)(1) = True
                    ' 3 to 4
                    If item("hammer") And item("bow") Then arrDAC(10)(4) = True
                End If
            Else
                ' 1: MQ Gerudo Training Ground Right Side
                ' 2: MQ Gerudo Training Ground Underwater
                ' 3: MQ Gerudo Training Ground Left Side
                ' 4: MQ Gerudo Training Ground Stalfos Room
                ' 5: MQ Gerudo Training Ground Back Areas
                ' 6: MQ Gerudo Training Ground Central Maze Right

                ' 0 to 1
                If item("bow") Then arrDAC(10)(1) = True

                ' 0 to 3
                If canBurn() Then
                    arrDAC(10)(3) = True
                    ' 3 to 4
                    If item("longshot") Then
                        arrDAC(10)(4) = True
                        ' 4 to 5
                        If canLens() And item("blue fire") And item("song of time") Then
                            arrDAC(10)(5) = True
                            ' 5 to 6
                            If item("hammer") Then arrDAC(10)(6) = True
                            ' 5 to 1 Since we already have the Longshot check
                            arrDAC(10)(1) = True
                        End If
                    End If
                End If
                ' 1 to 2 modified to complete the check
                If arrDAC(10)(1) Then
                    If (item("fire arrows") Or (item("longshot") And item("din's fire"))) And item("hover boots") And item("iron boots") And canFewerZora() Then arrDAC(10)(2) = True
                End If
            End If
        End If
    End Sub
    Private Sub dacGanonsCastle()
        ' Set up DAC for the Ganon's Castle
        resetDAC(11)

        If canEnterGanonsCastle() Then
            arrDAC(11)(0) = True
            ' Normal and MQ share same setup
            ' 0: Ganon's Castle
            ' 1: Ganon's Castle Deku Room
            ' 2: Ganon's Castle Light Trial
            ' 3: Ganon's Tower

            ' 0 to 1
            If canLens() Then arrDAC(11)(1) = True
            ' 0 to 2
            If item("lift", 3) Then arrDAC(11)(2) = True
            ' 0 to 3
            If checkLoc("6719") Then
                arrDAC(11)(3) = True
            Else
                If checkLoc("6611") And checkLoc("6612") And checkLoc("6613") And checkLoc("6614") And checkLoc("6615") And checkLoc("6629") Then arrDAC(11)(3) = True
            End If
        End If
    End Sub
    Private Function dungeonKeyCounter(ByVal dungeon As Byte, ByVal theseDoors As String) As Boolean        ', Optional safeDoors As String = "x") As Boolean
        ' Works to count dungeon keys and unlocked doors
        dungeonKeyCounter = False
        Dim countDoors As Byte = 0

        ' If set to remove small keys, always true. This setting is only for OOTR, as AP handles bit flags better
        If My.Settings.setSmallKeys = 2 Then Return True

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
                dunArea = "111"
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
                If allItems.Contains("u") And (allItems.Contains("w") Or canEnterZF(True)) Then Return True
            Case "bomb"
                If allItems.Contains("c") Then Return True
            Case "bombchu"
                If allItems.Contains("j") Then Return True
            Case "boomerang"
                If canYoung And allItems.Contains("o") Then Return True
            Case "bottle"
                If allItems.Contains("u") Then Return True
            Case "bow"
                If canAdult And allItems.Contains("d") Then Return True
            Case "deku nuts"
                If allItems.Contains("b") Then Return True
            Case "deku stick"
                If canYoung And allItems.Contains("a") Then Return True
            Case "din's fire"
                If canMagic And allItems.Contains("f") Then Return True
            Case "fire arrows"
                If canAdult And allItems.Contains("d") And allItems.Contains("e") And canMagic Then Return True
            Case "hammer", "megaton hammer"
                If canAdult And allItems.Contains("r") Then Return True
            Case "hookshot"
                If canAdult And allItems.Contains("k") Then Return True
            Case "lens of truth"
                If canMagic And allItems.Contains("p") Then Return True
            Case "longshot"
                If canAdult And allItems.Contains("l") Then Return True
            Case "slingshot"
                If canYoung And allItems.Contains("g") Then Return True
            Case "zelda's lullaby"
                If allItems.Contains("h") And aQIChecks(12) Then Return True
            Case "epona's song"
                If allItems.Contains("h") And aQIChecks(13) Then Return True
            Case "saria's song"
                If allItems.Contains("h") And aQIChecks(14) Then Return True
            Case "sun's song"
                If allItems.Contains("h") And aQIChecks(15) Then Return True
            Case "song of time"
                If allItems.Contains("h") And aQIChecks(16) Then Return True
            Case "song of storms"
                If allItems.Contains("h") And aQIChecks(17) Then Return True
            Case "minuet of forest"
                If allItems.Contains("h") And aQIChecks(6) Then Return True
            Case "bolero of fire"
                If allItems.Contains("h") And aQIChecks(7) Then Return True
            Case "serenade of water"
                If allItems.Contains("h") And aQIChecks(8) Then Return True
            Case "requiem of spirit"
                If allItems.Contains("h") And aQIChecks(9) Then Return True
            Case "nocturne of shadow"
                If allItems.Contains("h") And aQIChecks(10) Then Return True
            Case "prelude of light"
                If allItems.Contains("h") And aQIChecks(11) Then Return True
            Case "scarecrow"
                If canAdult And checkLoc("6512") Then
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
        ' Check for Kokiri Sword, Deku Sticks, Fairy Slingshot, or can explode
        If checkLoc("7416") Or allItems.Contains("a") Or allItems.Contains("g") Or canExplode() Then Return True
        ' If counting Boomerang and you have it
        If countBoomerang And allItems.Contains("o") Then Return True
        ' Finally, Din's Fire and magic
        If allItems.Contains("f") And canMagic Then Return True
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
        ' Check for Megaton Hammer and adult Link
        If allItems.Contains("r") And canAdult Then Return True
    End Function
    Private Function canBurn(Optional allowStick As Boolean = False) As Boolean
        canBurn = False
        ' Check for Din's Fire access
        If item("din's fire") Then Return True
        ' Check for Fire Arrow access
        If item("fire arrows") Then Return True
        ' If sticks are allowed, check they can be used
        If allowStick And item("deku stick") Then Return True
    End Function
    Private Function canBuyBeans() As Boolean
        ' Check if you can buy beans
        canBuyBeans = False
        ' Auto False if you cannot be young Link
        If Not canYoung Then Return False
        ' Can enter Zora's River and use bombs
        If entranceCheck("ZR") And canExplode() Then Return True
        ' Can come from Lost Woods side
        If canLostWoodsToZorasRiver() Then Return True
    End Function
    Private Function canCraterWarpToTop(Optional halfWay As Boolean = False) As Boolean
        ' Checks that you can from Crater Warp to top of Crater
        canCraterWarpToTop = False
        ' Check if you can even reach Death Mountain Crater
        If Not entranceCheck("DMC") Then Return False
        ' If you cannot be an adult, auto False
        If Not canAdult Then Return False
        ' Need to have the tunic or only be going halfway
        If checkLoc("7425") Or halfWay Then
            ' If those, then also a hookshot or hover boots
            If allItems.Contains("k") Or checkLoc("7430") Then Return True
            ' Checks if beans have been planted
            If checkLoc("B3") Then Return True
            ' If you can be young link and get or have seeds
            If canYoung Then
                If canBuyBeans() Or magicBeans > 0 Then Return True
            End If
        End If
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
        ' First check that you can be adult Link
        If Not canAdult Then Return False
        ' This is redundant, but check that you can reach Outside Ganon's Castle (if I ever do entrance rando, it will be more important)
        If Not entranceCheck("OGC") Then Return False
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
    Private Function canEnterZF(Optional asAdult As Boolean = False) As Boolean
        ' Check if you can enter Zora's Fountain as young or adult Link, the separation is important
        canEnterZF = False

        ' Both sides can check if the King has already moved, by letter or by force
        If checkLoc("6303") Then Return True

        If asAdult Then
            If Not adultZD Then Return False
            ' Adult side needs Open Zora's Fountain setting
            If canAdult And My.Settings.setOpenZF Then Return True
        Else
            If Not youngDC Then Return False
            ' Young side needs Ruto's Letter
            If canYoung And allItems.Contains("v") Then Return True
        End If
    End Function
    Private Function canExplode() As Boolean
        ' Check to see if you can explode things
        canExplode = False
        ' Check for bombs
        If allItems.Contains("c") Then Return True
        ' Check for bombchus and that they are in logic
        If allItems.Contains("j") And My.Settings.setBombchus Then Return True
    End Function
    Private Function canGCGrotto() As Boolean
        ' Check if you can reach the Goron City Grotto
        canGCGrotto = False
        ' Adult Link with an ocarina and the Song of Time
        If canAdult And allItems.Contains("h") And checkLoc("7716") Then Return True
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
        ' Check that you can be an adult and reach Temple of Time
        If Not canAdult Then Return False
        If Not entranceCheck("TT") Then Return False

        ' Determine what is needed for LACS
        Dim startAddress As Integer = &H400CB4
        If randoVer = "AP" Then
            ' AP is offset by 16 bytes
            inc(startAddress, 16)
        ElseIf Not randoVer = "OOTR" Then
            ' If not AP or OOTR, abort this check
            Return False
        End If

        Dim conditionLACS As Byte = CByte(goRead(startAddress, 1))
        Dim countLACS As Byte = CByte(goRead(startAddress + 12, 1))
        Dim countChecks As Byte = 0
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
        If My.Settings.setFewerTunics Or checkLoc("7425") Then Return canAdult
    End Function
    Private Function canFewerZora() As Boolean
        canFewerZora = False
        If My.Settings.setFewerTunics Or checkLoc("7426") Then Return canAdult
    End Function
    Private Function canGetBugs() As Boolean
        ' Checks if you can reach bugs
        canGetBugs = False
        ' Auto False if you do not have a bottle (does not count Ruto's Letter)
        If Not allItems.Contains("u") Then Return False

        ' Areas both young and adult Link can get bugs with just a bottle
        If entranceCheck("DC") Or entranceCheck("KV") Or entranceCheck("GY") Then Return True

        ' Goron City with breaking rocks or Silver Gauntlets
        If entranceCheck("GC") Then
            If canBreakRocks() Or checkLoc("7607") Then Return True
        End If

        ' Storms Grottos with bugs
        If canStormsGrotto() Then
            If entranceCheck("KF") Or entranceCheck("DMT") Then
                If canAdult Or allItems.Contains("a") Or allItems.Contains("o") Or canExplode() Or checkLoc("7416") Then Return True
            End If
        End If

        ' All checks after this require young Link
        If Not canYoung Then Return False

        ' Area young Link can get with just a bottle
        If entranceCheck("HC") Or entranceCheck("GV") Then Return True

        'Reaching Death Mountain Summit
        If canTopDeathMountain() Or canCraterWarpToTop() Then Return True

        ' Areas where young Link can find bugs in bushes
        If entranceCheck("LW") Or entranceCheck("HF") Or entranceCheck("LH") Or canBuyBeans() Then
            ' Need sticks, boomerang, explosions, or the Kokiri Sword
            If allItems.Contains("a") Or allItems.Contains("o") Or canExplode() Or checkLoc("7416") Then Return True
        End If
    End Function
    Private Function canHoverTricks() As Boolean
        canHoverTricks = False
        If checkLoc("7430") And My.Settings.setHoverTricks Then Return True
    End Function
    Private Function canLostWoodsToZorasRiver() As Boolean
        ' Checks shortcut in Lost Woods to Zora's River
        canLostWoodsToZorasRiver = False
        ' Lost Woods access needed
        If entranceCheck("LW") Then
            ' Need any dive rank
            Return item("dive")
        End If
    End Function
    Private Function canProjectile(Optional onlyAge As Byte = 1) As Boolean
        canProjectile = False
        If canExplode() Then Return True
        If onlyAge >= 1 And item("bow") Or item("hookshot") Then Return True
        If onlyAge <= 1 And item("boomerang") Or item("slingshot") Then Return True
    End Function
    Private Function canReachStructure() As Boolean
        ' Checks if you can reach the Haunted Wasteland Structure
        canReachStructure = False
        ' Auto False if you cannot reach the Haunted Wasteland from either side
        If Not entranceCheck("HW") Then Return False
        ' Desert Colossus warp and Lens of Truth access
        If allItems.Contains("h") And checkLoc("7709") And canLens() Then Return True
        ' Gerudo Fortress side needs adult access
        If Not canAdult Then Return False
        ' Next check for Longshot of Hover Boots
        If allItems.Contains("l") Or checkLoc("7430") Then Return True
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
        ' Check for ocarina and Sun's Song
        If allItems.Contains("h") And checkLoc("7715") Then Return True

        ' Going to only check 3 of the areas because the others all depend on Hyrule Field
        If entranceCheck("HF") Then Return True
        If entranceCheck("DMT") Then Return True
        If entranceCheck("DC") Then Return True

        'If entranceCheck("HC") Then Return True
        'If entranceCheck("ZR") Then Return True
        'If entranceCheck("LH") Then Return True
        'If entranceCheck("GV") Then Return True
    End Function
    Private Function canTopDeathMountain() As Boolean
        ' Checks that you can get atop Death Mountain Trail
        canTopDeathMountain = False
        ' Check if you can even reach Death Mountain Trail
        If Not entranceCheck("DMT") Then Return False
        ' If you can break rocks with bombs or hammer
        If canBreakRocks() Then Return True
        ' From here, everything checks on being able to adult
        If canAdult Then
            ' If you can do Hover Boot tricks, you can get up
            If canHoverTricks() Then Return True
            ' Checks if the beans have already been planted
            If checkLoc("B2") Then Return True
            ' Check if you can plant the beans
            If canYoung Then
                ' Disable adult Link for these tests, since you will need to be young Link to plant the beans
                canAdult = False
                ' Check if young Link can reach DMT
                If entranceCheck("DMT") Then
                    ' Check that you can reach ZR to get the beans, or have beans, and set to true
                    ' We will not Return True because we need to switch adult status back
                    If canBuyBeans() Or magicBeans > 0 Then canTopDeathMountain = True
                End If
                ' Restore adult status
                canAdult = True
            End If
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

    Private Function checkLogic(ByVal logicKey As String, Optional area As String = "X") As Boolean
        ' Checks the logic key to see if it is available. Start with false
        checkLogic = False
        Dim canDoThis As Boolean = False

        ' First check if you can reach that area
        If Not area = "X" Then
            If Not entranceCheck(area) Then Return False
        End If

        ' If there is no logic key, then item is always accessable, return true
        If logicKey = String.Empty Then Return True

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
        ' O is for checking for fire with deku sticks
        If logicKey.Contains("O") And canBurn(True) Then logicKey = logicKey.Replace("O", "")
        ' P is for checking for fire 
        If logicKey.Contains("P") And canBurn() Then logicKey = logicKey.Replace("P", "")
        ' W is for checking Light Arrow Cutscene requirements and access
        If logicKey.Contains("W") And canLACS() Then logicKey = logicKey.Replace("W", "")
        ' S is for checking if you can stun a deku scrub
        If logicKey.Contains("S") And canDeku() Then logicKey = logicKey.Replace("S", "")
        ' T is for checking if you need the lens of truth
        If logicKey.Contains("T") And canLens() Then logicKey = logicKey.Replace("T", "")
        ' U is for Saria's Gift, either opened Forest setting or cleared Deku Tree
        If logicKey.Contains("U") And My.Settings.setOpenKF Or checkLoc("6223") Then logicKey = logicKey.Replace("U", "")
        ' x is for checking if you can explode things
        If logicKey.Contains("x") And canExplode() Then logicKey = logicKey.Replace("x", "")
        ' X is for checking if Bombchus are in logic. If they are, test them, if not, test bomb bag
        If logicKey.Contains("X") And inlogicBombchus() Then logicKey = logicKey.Replace("X", "")
        ' Y is for young Link access
        If logicKey.Contains("Y") And canYoung Then logicKey = logicKey.Replace("Y", "")
        ' Z is for adult Link access
        If logicKey.Contains("Z") And canAdult Then logicKey = logicKey.Replace("Z", "")

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
        If isLogicGood(logicKey) Then Return True

        ' Check for various keys
        For i = 1 To logicKey.Length
            Select Case Mid(logicKey, i, 1)
                Case "G"
                    ' G# = Test Types:  00 = Break Rock 01 = Break Rock with SoA    02 = Storms Grotto      03 = Ascending Death Mountain   04 = Crater Entrance Check  05 = Half Crater Climb  06 = Full Crater Climb
                    '                   07 = LW to ZR   08 = Can Buy Beans          09 = Fewer Goron Tunic  0A = Fewer Zora Tunic           0B = Entrance Check GV      0C = Reach Structure    0D = Can Get Bugs
                    '                   0E = Young ZD   0F = Adult ZD               10 = Young DC           11 = Adult DC                   12 = Young DC Centre        13 = Adult DC Centre    14 = DC Shortcut
                    '                   15 = Young ZF   16 = Adult ZF               17 = Blue Fire Access   18 = KV Entrance
                    '                   
                    '                   20 = Can Projectile Young   21 = Can Projectile Adult
                    '                   
                    '                   A# = Logics
                    '                   A0 = DC Jump        A1 = FoT Ledge      A2 = FoT Doorframe      A3 = Rusted Switches        A4 = FiT Flame Maze Skip        A5 = SpT Young Bombchu          A6 = SpT MQ Sun Block without Song
                    '                   A7 = ShT Umbrella   A8 = BotW Dead Hand
                    '                   B# = Boss Key: # Dungeon
                    '                   C0 = Fire Temple Keys: 3        C1 = Fire Temple Keys: 4    C2 = Fire Temple Keys: 5    C3 = Fire Temple MQ Keys: 3     C4 = Fire Temple MQ Keys: 4     C5 = Water Temple Pillar
                    '                   C6 = Water Temple BK Chest      C7 = Water Temple MQ All    C8 = Sp. Temple Young All   C9 = Sp. Temple Young Last Key  CA = Sp. Temple MQ Young Door 1 CB = Sp Temple MQ Statue Left Door 1
                    '                   CC = Sp. Temple MQ 9 Thrones    CD = Sp. Temple MQ Symphony CE = Sh. Temple MQ Spikes   CF = BotW Right                 D0 = BotW Centre Left           D1 = BotW Centre Right
                    '                   D2 = BotW MQ Left Side          D3 = BotW MQ Right Side     D4 = GTG Hidden Ceiling     D5 = GTG 1st Chest              D6 = GTG 2nd Chest              D7 = GTG 3rd Chest
                    '                   D8 = GTG Final Chest            D9 = GTG MQ First Door      DA = GTG MQ All Doors       DB = GC Light Trail First Door

                    tempString = Mid(logicKey, i + 1, 2)
                    candothis = False
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
                            canDoThis = canTopDeathMountain()
                        Case "04"
                            canDoThis = entranceCheck("DMC")
                        Case "05"
                            canDoThis = canCraterWarpToTop(False)
                        Case "06"
                            canDoThis = canCraterWarpToTop(True)
                        Case "07"
                            canDoThis = canLostWoodsToZorasRiver()
                        Case "08"
                            canDoThis = canBuyBeans()
                        Case "09"
                            canDoThis = canFewerGoron()
                        Case "0A"
                            canDoThis = canFewerZora()
                        Case "0B"
                            ' This is to check if you can reach GV, thus GF from the GV side
                            canDoThis = entranceCheck("GF")
                        Case "0C"
                            canDoThis = canReachStructure()
                        Case "0D"
                            canDoThis = canGetBugs()
                        Case "0E"
                            canDoThis = youngZD
                        Case "0F"
                            canDoThis = adultZD
                        Case "10"
                            canDoThis = youngDC
                        Case "11"
                            canDoThis = adultDC
                        Case "12"
                            canDoThis = youngSpTCentre
                        Case "13"
                            canDoThis = adultSpTCentre
                        Case "14"
                            If checkLoc("10909") And checkLoc("10911") And checkLoc("10912") Then canDoThis = True
                        Case "15"
                            canDoThis = canEnterZF(False)
                        Case "16"
                            canDoThis = canEnterZF(True)
                        Case "17"
                            canDoThis = item("blue fire")
                        Case "18"
                            canDoThis = entranceCheck("KV")
                        Case "20"
                            canDoThis = canProjectile(0)
                        Case "21"
                            canDoThis = canProjectile(2)
                        Case "A0"
                            canDoThis = My.Settings.setDCSpikeJump
                        Case "A1"
                            canDoThis = My.Settings.setFoTLedge
                        Case "A2"
                            canDoThis = My.Settings.setFoTFrame
                        Case "A3"
                            canDoThis = My.Settings.setRustedSwitches
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
                        Case "B0", "B1", "B2", "B3", "B4"
                            canDoThis = aBossKeys(CByte(Mid(tempString, 2)))
                            If Not canDoThis Then
                                ' If the key was not found, check each boss door
                                Select Case Mid(tempString, 2)
                                    Case "0"    ' Forest
                                        canDoThis = checkLoc("10620")
                                    Case "1"    ' Fire
                                        canDoThis = checkLoc("10720")
                                    Case "2"    ' Water
                                        canDoThis = checkLoc("10820")
                                    Case "3"    ' Spirit
                                        canDoThis = checkLoc("10920")
                                    Case "4"    ' Shadow
                                        canDoThis = checkLoc("11020")
                                    Case Else
                                        MsgBox("G" & tempString)
                                End Select
                            End If
                        Case "C0"
                            canDoThis = dungeonKeyCounter(4, "293024") ', "23")
                        Case "C1"
                            canDoThis = dungeonKeyCounter(4, "29302427") ', "23")
                        Case "C2"
                            canDoThis = dungeonKeyCounter(4, "2930242726") ', "23")
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
                        Case Else
                            MsgBox(logicKey & ": G" & tempString)
                    End Select
                    If canDoThis Then
                        logicKey = logicKey.Replace("G" & tempString, "")
                        ' Step i back a step to not miss from the removed text
                        i = i - 1
                    Else
                        i = i + 2
                    End If
                Case "K"
                    ' K# = Gold Skulltulas x10
                    tempString = Mid(logicKey, i + 1, 1)
                    ' Compare the number of gold skulltulas to the need times 10
                    If goldSkulltulas >= CByte("&H" & tempString) * 10 Then
                        logicKey = logicKey.Replace("K" & tempString, "")
                        ' Step i back a step to not miss from the removed text
                        i = i - 1
                    Else
                        i = i + 1
                    End If
                Case "L"
                    ' L###### is for checking a key by loc value
                    tempString = Mid(logicKey, i + 1, 5)
                    ' Since a key's loc can be 3-5 characters long, Ls are used to fill the space, remove them from the check request
                    If checkLoc(tempString.Replace("L", "")) Then
                        logicKey = logicKey.Replace("L" & tempString, "")
                        ' Step i back a step to not miss from the removed text
                        i = i - 1
                    Else
                        i = i + 5
                    End If
                Case "Q"
                    ' Q### = Check DAC #|## = Dungeon|Area
                    tempString = Mid(logicKey, i + 1, 3)
                    If arrDAC(CByte("&H" & (Mid(tempString, 1, 1))))(CByte("&H" & (Mid(tempString, 2)))) Then
                        logicKey = logicKey.Replace("Q" & tempString, "")
                        i = 1 - 1
                    Else
                        i = i + 3
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
                    End Select

                    ' Compare item check and back up i
                    If item(tempItem, CByte(Mid(tempString, 2))) Then
                        logicKey = logicKey.Replace("V" & tempString, "")
                        i = i - 1
                    Else
                        i = i + 2
                    End If
                Case "y", "z"
                    ' y# = Young Link Quest Item
                    ' z# = Adult Link Quest Item
                    tempString = Mid(logicKey, i, 2)
                    If allItems.Contains(tempString) Then
                        logicKey = logicKey.Replace(tempString, "")
                        i = i - 1
                    Else
                        i = i + 1
                    End If
            End Select
        Next
        ' After all the removals, check the key
        Return isLogicGood(logicKey)
    End Function
    Private Function isLogicGood(ByVal logicString As String) As Boolean
        ' If logic key is empty, contains a '..', or starts/end with '.' then all or some logic must have been cleared, return true
        isLogicGood = False
        If logicString = String.Empty Then Return True
        If logicString.Contains("..") Or Mid(logicString, 1, 1) = "." Or Mid(logicString, logicString.Length, 1) = "." Then Return True
    End Function

    ' Make the keys for each area
    Private Sub makeKeysKF(ByRef tK As Integer)
        ' Set up keys for Kokiri 
        With aKeys(tK)
            .loc = "4500"
            .area = "KF"
            .name = "Mido's House Upper Left Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4501"
            .area = "KF"
            .name = "Mido's House Upper Right Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4502"
            .area = "KF"
            .name = "Mido's House Lower Left Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4503"
            .area = "KF"
            .name = "Mido's House Lower Right Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "5100"
            .area = "KF"
            .name = "Kokiri Sword Chest (Young)"
            .logic = "Y"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4612"
            .area = "KF"
            .name = "Storms Grotto Chest"
            .logic = "G02"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8100"
            .area = "KF"
            .name = "Soft Soil Near Shoppe (Young)"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8101"
            .area = "KF"
            .name = "On Know-it-All Brother's House (N) (Young)"
            .gs = True
            .logic = "YJN"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8102"
            .area = "KF"
            .name = "Above the House of Twins (N) (Adult)"
            .gs = True
            .logic = "ZNk"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "10024"
            .area = "KF"
            .name = "Link's House (Adult)"
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
        With aKeys(tK)
            .loc = "7202"
            .area = "LW"
            .name = "Deku Scrub Near Bridge (Young)"
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4620"
            .area = "LW"
            .name = "Near Shortcuts Grotto Chest"
            .logic = "G00"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "7203"
            .area = "LW"
            .name = "Deku Scrub in Grotto"
            .logic = "G00S.G00ZhLL7714.G00ZhLL7706"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6717"
            .area = "LW"
            .name = "Gift from Saria"
            .logic = "U"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6813"
            .area = "LW"
            .name = "Target in Woods (Young)"
            .logic = "Yg"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6807"
            .area = "LW"
            .name = "Ocarina Memory Game (Young)"
            .logic = "Yh"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6806"
            .area = "LW"
            .name = "Skull Kid (Young)"
            .logic = "YhLL7714"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6814"
            .area = "LW"
            .name = "Deku Theatre Skull Mask (Young)"
            .logic = "YLL6920.YLL6926"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6815"
            .area = "LW"
            .name = "Deku Theatre Mask of Truth (Young)"
            .logic = "YLL6926"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8108"
            .area = "LW"
            .name = "Soft Soil Near Bridge (Young)"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8109"
            .area = "LW"
            .name = "Soft Soil Near Theatre (Young)"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8110"
            .area = "LW"
            .name = "Theatre Bean Plant Ride (N) (Adult)"
            .gs = True
            .logic = "ZNG08hLL7714.YZNqhLL7714.YZNLLLLB0hLL7714.ZNG08hLL7706.YZNqhLL7706.YZNLLLLB0hLL7706"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9810"
            .area = "LW"
            .name = "Near Bridge (Young)"
            .scrub = True
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9801"
            .area = "LW"
            .name = "Outside Theatre Right (Young)"
            .scrub = True
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9802"
            .area = "LW"
            .name = "Outside Theatre Left (Young)"
            .scrub = True
            .logic = "YS"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9311"
            .area = "LW"
            .name = "Grotto Left"
            .scrub = True
            .logic = "G00S.G00ZhLL7714.G00ZhLL7706"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9304"
            .area = "LW"
            .name = "Grotto Right"
            .scrub = True
            .logic = "G00S.G00ZhLL7714.G00ZhLL7706"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "B0"
            .area = "EVENT"
            .name = "Beans Planted Near Theatre"
        End With
        inc(tK)
    End Sub
    Private Sub makeKeysSFM(ByRef tk As Integer)
        ' Set up keys for the Sacred Forest Meadow
        With aKeys(tk)
            .loc = "4617"
            .area = "SFM"
            .name = "Wolfos Grotto Chest"
            .logic = "G01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6407"
            .area = "SFM"
            .name = "Song from Saria (Young)"
            .logic = "YLL6416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6400"
            .area = "SFM"
            .name = "Song from Sheik (Adult)"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8111"
            .area = "SFM"
            .name = "On East Wall (N) (Adult)"
            .gs = True
            .logic = "ZNk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9009"
            .area = "SFM"
            .name = "Storms Grotto Front"
            .scrub = True
            .logic = "G02S.G02Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9008"
            .area = "SFM"
            .name = "Storms Grotto Back"
            .scrub = True
            .logic = "G02S.G02Z"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysHF(ByRef tk As Integer)
        ' Set up keys for Hyrule Field
        With aKeys(tk)
            .loc = "4600"
            .area = "HF"
            .name = "Near Market Grotto Chest"
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4602"
            .area = "HF"
            .name = "Southeast Grotto Chest"
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4603"
            .area = "HF"
            .name = "Open Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6827"
            .area = "HF"
            .name = "Deku Scrub Inside Fence Grotto"
            .logic = "G01S.G01Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1901"
            .area = "HF"
            .name = "Tektite Grotto Piece of Heart (Adult)"
            .logic = "G01LL7429.GLL17610"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "103"
            .area = "HF"
            .name = "Get Ocarina of Time (Young)"
            .logic = "YLL7718LL7719LL7720"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6625"
            .area = "HF"
            .name = "Song from Ocarina of Time (Young)"
            .logic = "YLL7718LL7719LL7720"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8016"
            .area = "HF"
            .name = "Grotto Near Gerudo Valley"
            .gs = True
            .logic = "YG01Po.ZG00Pk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8017"
            .area = "HF"
            .name = "Grotto Near Kakariko Village"
            .gs = True
            .logic = "G01o.G01k"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1925"
            .area = "HF"
            .name = "Grotto Near Gerudo Valley"
            .cow = True
            .logic = "YG01PhLL7713.ZG00PhLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8803"
            .area = "HF"
            .name = "Inside Fence Grotto"
            .scrub = True
            .logic = "G01S.G01Z"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysLLR(ByRef tk As Integer)
        ' Set up keys for Lon Lon Ranch
        With aKeys(tk)
            .loc = "2101"
            .area = "LLR"
            .name = "Tower Piece of Heart (Young)"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6818"
            .area = "LLR"
            .name = "Talon's Chickens (Young)"
            .logic = "YLL6204h"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6408"
            .area = "LLR"
            .name = "Song from Malon (Young)"
            .logic = "YLL6204h"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6208"
            .area = "LLR"
            .name = "Get Epona (Adult)"
            .logic = "ZhLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8024"
            .area = "LLR"
            .name = "Wall Near Tower (N) (Young)"
            .gs = True
            .logic = "YNo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8025"
            .area = "LLR"
            .name = "Behind Coral (N) (Young)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8026"
            .area = "LLR"
            .name = "On House Window (N) (Young)"
            .gs = True
            .logic = "YNo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8027"
            .area = "LLR"
            .name = "Tree Near House (Young)"
            .gs = True
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "10124"
            .area = "LLR"
            .name = "Stables Left"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "10125"
            .area = "LLR"
            .name = "Stables Right"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2125"
            .area = "LLR"
            .name = "Tower Left"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2124"
            .area = "LLR"
            .name = "Tower Right"
            .cow = True
            .logic = "hLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6214"
            .area = "LLR"
            .name = "Win Cow from Malon (Adult)"
            .cow = True
            .logic = "ZhLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9601"
            .area = "LLR"
            .name = "Open Grotto Left (Young)"
            .scrub = True
            .logic = "YS"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9604"
            .area = "LLR"
            .name = "Open Grotto Centre (Young)"
            .scrub = True
            .logic = "YS"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9606"
            .area = "LLR"
            .name = "Open Grotto Right (Young)"
            .scrub = True
            .logic = "YS"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysMK(ByRef tk As Integer)
        ' Set up keys for the Market
        With aKeys(tk)
            .loc = "6829"
            .area = "MK"
            .name = "Shooting Gallery Reward (Young)"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6801"
            .area = "MK"
            .name = "Bombchu Bowling First Prize (Young)"
            .logic = "YX"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6802"
            .area = "MK"
            .name = "Bombchu Bowling Second Prize (Young)"
            .logic = "YX"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7201"
            .area = "MK"
            .name = "Lost Dog (Young)"
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6811"
            .area = "MK"
            .name = "Treasure Chest Game (Young)"
            .logic = "YNT"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8119"
            .area = "MK"
            .name = "Crate in Guard House (Young)"
            .gs = True
            .logic = "Y"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysTT(ByRef tk As Integer)
        ' Set up keys for the Temple of Time
        With aKeys(tk)
            .loc = "6405"
            .area = "TT"
            .name = "Song from Sheik (Adult)"
            .logic = "ZLL7700"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6720"
            .area = "TT"
            .name = "Light Arrows Cutscene (Adult)"
            .logic = "ZW"
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
        With aKeys(tk)
            .loc = "6202"
            .area = "HC"
            .name = "Malon's Egg (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6416"
            .area = "HC"
            .name = "Zelda's Letter (Young)"
            .logic = "y1.y2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6409"
            .area = "HC"
            .name = "Song from Impa (Young)"
            .logic = "y1.y2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6809"
            .area = "HC"
            .name = "Great Fairy Fountain (Young)"
            .logic = "hxLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8117"
            .area = "HC"
            .name = "Inside Storms Grotto (Young)"
            .gs = True
            .logic = "G02o"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8118"
            .area = "HC"
            .name = "Tree Near Entrance (Young)"
            .gs = True
            .logic = "J"
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
        With aKeys(tk)
            .loc = "4610"
            .area = "KV"
            .name = "Redead Grotto Chest"
            .logic = "G01a.G01f.G01LL7416.G01Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4608"
            .area = "KV"
            .name = "Open Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6828"
            .area = "KV"
            .name = "Anju's Chickens (Young)"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6928"
            .area = "KV"
            .name = "Talk to Anju (Adult)"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6805"
            .area = "KV"
            .name = "Talk to Man on Roof"
            .logic = "Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1801"
            .area = "KV"
            .name = "Impa's House Piece of Heart"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6830"
            .area = "KV"
            .name = "Shooting Gallery (Adult)"
            .logic = "Zd"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2001"
            .area = "KV"
            .name = "Windmill Piece of Heart"
            .logic = "Yo.ZhLL7716"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6411"
            .area = "KV"
            .name = "Song from Windmill (Adult)"
            .logic = "Zh"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6404"
            .area = "KV"
            .name = "Song from Shiek (Adult)"
            .logic = "ZLL7700LL7701LL7702"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8201"
            .area = "KV"
            .name = "On House Near Death Mountain (N) (Young)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8202"
            .area = "KV"
            .name = "Watchtower Ladder (N) (Young)"
            .gs = True
            .logic = "YNg.YNj"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8203"
            .area = "KV"
            .name = "Construction Site (N) (Young)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8204"
            .area = "KV"
            .name = "On Skulltula House (N) (Young)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8205"
            .area = "KV"
            .name = "Tree Near Entrance (N) (Young)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8206"
            .area = "KV"
            .name = "Above Impa's House (N) (Adult)"
            .gs = True
            .logic = "ZNk.ZNH"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1824"
            .area = "KV"
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
        With aKeys(tk)
            .loc = "2208"
            .area = "GY"
            .name = "Dampe's Gravedigging Tour (Young)"
            .logic = "YN"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4800"
            .area = "GY"
            .name = "Shield Grave Chest"
            .logic = "N.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2204"
            .area = "GY"
            .name = "Ledge Piece of Heart"
            .logic = "Zl.ZG08.ZYq"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4700"
            .area = "GY"
            .name = "Heart Piece Grave Chest"
            .logic = "NhLL7715.ZhLL7715"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4900"
            .area = "GY"
            .name = "Royal Family's Tomb Chest"
            .logic = "hLL7712P"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6410"
            .area = "GY"
            .name = "Song from Royal Family's Tomb"
            .logic = "ZhLL7712.YJhLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5000"
            .area = "GY"
            .name = "Hookshot Chest (Adult)"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2007"
            .area = "GY"
            .name = "Dampe's Race Piece of Heart (Adult)"
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8200"
            .area = "GY"
            .name = "Soft Soil (Young)"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8207"
            .area = "GY"
            .name = "On South Wall (N) (Young)"
            .gs = True
            .logic = "YN"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDMT(ByRef tk As Integer)
        ' Set up keys for Death Mountain Trail
        With aKeys(tk)
            .loc = "5801"
            .area = "DMT"
            .name = "Blast Wall Chest"
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2830"
            .area = "DMT"
            .name = "Piece of Heart Above Cavern"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "024"
            .area = "DMT"
            .name = "Great Fairy Fountain"
            .logic = "G00G03hLL7712.G00G04hLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4623"
            .area = "DMT"
            .name = "Storms Grotto Chest"
            .logic = "G02G03.G02G04"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6008"
            .area = "DMT"
            .name = "Help Biggoron (Adult)"
            .logic = "ZG03z9.ZG03zA.ZG03zB.ZG04z9.ZG04zA.ZG04zB"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8125"
            .area = "DMT"
            .name = "Soft Soil Near Dodongo's Cavern (Young)"
            .gs = True
            .logic = "YxG0D.YJG0DV01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8126"
            .area = "DMT"
            .name = "Blast Wall Near Village"
            .gs = True
            .logic = "G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8127"
            .area = "DMT"
            .name = "Rock Near Bomb Flower (N) (Adult)"
            .gs = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8128"
            .area = "DMT"
            .name = "Rock Near Climb Wall (N) (Adult)"
            .gs = True
            .logic = "ZG03r.ZG04r"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2824"
            .area = "DMT"
            .name = "In Grotto"
            .cow = True
            .logic = "G00G03hLL7713.G00G04hLL7713"
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
        With aKeys(tk)
            .loc = "4626"
            .area = "DMC"
            .name = "Upper Grotto Chest"
            .logic = "G03G00.G05G00"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2902"
            .area = "DMC"
            .name = "Wall Piece of Heart"
            .logic = "G03.G05"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2908"
            .area = "DMC"
            .name = "Volcano Piece of Heart"
            .logic = "ZH.ZG08hLL7707.ZYqhLL7707.ZLLLLB3hLL7707.ZG06LLLLB3"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "016"
            .area = "DMC"
            .name = "Great Fairy Fountain"
            .logic = "ZkhLL7707rLL7712.ZLL7430hLL7707rLL7712.ZLL7025rLL7712.ZLL7425LL7430hrLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6401"
            .area = "DMC"
            .name = "Song from Sheik (Adult)"
            .logic = "ZkLL7425.ZLL7430LL7425.ZLL7025LL7430.ZkLL7025.ZhLL7707"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8124"
            .area = "DMC"
            .name = "Soft Soil Near Warp (Young)"
            .gs = True
            .logic = "YJG0DhLL7707"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8131"
            .area = "DMC"
            .name = "Crate at Entrance (Young)"
            .gs = True
            .logic = "YJG03"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "C01"
            .area = "DMC"
            .name = "Near Ladder (Young)"
            .scrub = True
            .logic = "YSG03"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9401"
            .area = "DMC"
            .name = "Grotto Left (Adult)"
            .scrub = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9404"
            .area = "DMC"
            .name = "Grotto Centre (Adult)"
            .scrub = True
            .logic = "Zr"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9406"
            .area = "DMC"
            .name = "Grotto Right (Adult)"
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
        With aKeys(tk)
            .loc = "5900"
            .area = "GC"
            .name = "Maze Left Chest (Adult)"
            .logic = "Zr.ZV02.ZHx"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5902"
            .area = "GC"
            .name = "Maze Centre Chest"
            .logic = "G00.V02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5901"
            .area = "GC"
            .name = "Maze Right Chest"
            .logic = "G00.V02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7014"
            .area = "GC"
            .name = "Rolling Goron (Young)"
            .logic = "Yx"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7025"
            .area = "GC"
            .name = "Rolling Goron (Adult)"
            .logic = "Zx.Zd.ZV01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "3031"
            .area = "GC"
            .name = "Pot Piece of Heart (Young)"
            .logic = "Yfx.YfV01.YhLL7712x.YhLL7712V01"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6306"
            .area = "GC"
            .name = "Darunia's Joy (Young)"
            .logic = "YhLL7712LL7714"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8129"
            .area = "GC"
            .name = "Rope Platform (Adult)"
            .gs = True
            .logic = "Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8130"
            .area = "GC"
            .name = "Crate in Maze Near Chests (Young)"
            .gs = True
            .logic = "Yx"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9501"
            .area = "GC"
            .name = "Open Grotto Left (Adult)"
            .scrub = True
            .logic = "ZShLL7716"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9504"
            .area = "GC"
            .name = "Open Grotto Centre (Adult)"
            .scrub = True
            .logic = "ZShLL7716"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9506"
            .area = "GC"
            .name = "Open Grotto Right (Adult)"
            .scrub = True
            .logic = "ZShLL7716"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZR(ByRef tk As Integer)
        ' Set up keys for Zora's River
        With aKeys(tk)
            .loc = "2304"
            .area = "ZR"
            .name = "Near Open Grotto Piece of Heart"
            .logic = "G08.ZLL7430"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4609"
            .area = "ZR"
            .name = "Open Grotto Chest"
            .logic = "G08.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2311"
            .area = "ZR"
            .name = "Near Domain Piece of Heart"
            .logic = "G08.ZLL7430"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8208"
            .area = "ZR"
            .name = "On Ladder Near Zora's Domain (N) (Young)"
            .gs = True
            .logic = "G08J"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8209"
            .area = "ZR"
            .name = "Tree Near Entrance (Young)"
            .gs = True
            .logic = "YJ"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8211"
            .area = "ZR"
            .name = "Wall Above Bridge (N) (Adult)"
            .gs = True
            .logic = "Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8212"
            .area = "ZR"
            .name = "Wall Near Open Grotto (N) (Adult)"
            .gs = True
            .logic = "Zk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8909"
            .area = "ZR"
            .name = "Storms Grotto Front"
            .scrub = True
            .logic = "G08G02S.G02Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8908"
            .area = "ZR"
            .name = "Storms Grotto Back"
            .scrub = True
            .logic = "G08G02S.G02Z"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZD(ByRef tk As Integer)
        ' Set up keys for Zora's Domain
        With aKeys(tk)
            .loc = "5300"
            .area = "ZD"
            .name = "Torches Chest (Young)"
            .logic = "aG0E"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6308"
            .area = "ZD"
            .name = "Diving Minigame (Young)"
            .logic = "G0E"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6303"
            .area = "EVENT"
            .name = "Delivered Ruto's Letter"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7109"
            .area = "ZD"
            .name = "Zora King (Adult)"
            .logic = "G0FG17"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8214"
            .area = "ZD"
            .name = "Frozen Waterfall Top (N) (Adult)"
            .gs = True
            .logic = "G0FNk.G0FNd.G0FNM"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZF(ByRef tk As Integer)
        ' Set up keys for Zora's Fountain
        With aKeys(tk)
            .loc = "2501"
            .area = "ZF"
            .name = "Iceberg Piece of Heart (Adult)"
            .logic = "G16"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2520"
            .area = "ZF"
            .name = "Bottom Piece of Heart (Adult)"
            .logic = "G16LL7426LL7429.G16G0ALL7429"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6808"
            .area = "ZF"
            .name = "Great Fairy Fountain"
            .logic = "xhLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8210"
            .area = "ZF"
            .name = "Lower West Wall (N) (Young)"
            .gs = True
            .logic = "G15No"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8215"
            .area = "ZF"
            .name = "Southeast Corner Log (Young)"
            .gs = True
            .logic = "G15"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8213"
            .area = "ZF"
            .name = "Hidden Cave (N) (Adult)"
            .gs = True
            .logic = "G16V02G00Nk"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysLH(ByRef tk As Integer)
        ' Set up keys for Lake Hylia
        With aKeys(tk)
            .loc = "7316"
            .area = "LH"
            .name = "Scarecrow (Young)"
            .logic = "Yh"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6512"
            .area = "LH"
            .name = "Scarecrow (Adult)"
            .logic = "YLL7316"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "211"
            .area = "LH"
            .name = "Underwater Item (Young)"
            .logic = "YV11"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6110"
            .area = "LH"
            .name = "Big Fish (Young)"
            .logic = "Y"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6111"
            .area = "LH"
            .name = "Big Fish (Adult)"
            .logic = "ZkhLL6512.ZLLLC04.ZG08.ZYq.ZLLL202"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2430"
            .area = "LH"
            .name = "Lab Tower Piece of Heart (Adult)"
            .logic = "ZkhLL6512.ZG08.ZYq.ZLLL202"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6800"
            .area = "LH"
            .name = "Lab Dive (Adult)"
            .logic = "ZLL7610"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5200"
            .area = "LH"
            .name = "Shoot the Sun (Adult)"
            .logic = "ZdkhLL6512.ZdLLLC04"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8216"
            .area = "LH"
            .name = "Soft Soil Near Lab (Young)"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8217"
            .area = "LH"
            .name = "On Small Island Pillar (N) (Young)"
            .gs = True
            .logic = "YJ"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8218"
            .area = "LH"
            .name = "Behind Lab (N) (Young)"
            .gs = True
            .logic = "Yo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8219"
            .area = "LH"
            .name = "Crate in Lab Pool (Adult)"
            .gs = True
            .logic = "ZLL7429k"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8220"
            .area = "LH"
            .name = "On Tree Near Warp (N) (Adult)"
            .gs = True
            .logic = "Zl"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9101"
            .area = "LH"
            .name = "Grave Grotto Left"
            .scrub = True
            .logic = "S.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9104"
            .area = "LH"
            .name = "Grave Grotto Centre"
            .scrub = True
            .logic = "S.Z"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9106"
            .area = "LH"
            .name = "Grave Grotto Right"
            .scrub = True
            .logic = "S.Z"
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
        With aKeys(tk)
            .loc = "2602"
            .area = "GV"
            .name = "Crate Piece of Heart"
            .logic = "Y.Zl.ZH"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2601"
            .area = "GV"
            .name = "Waterfall Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5400"
            .area = "GV"
            .name = "Chest Behind Rocks(Adult)"
            .logic = "Zlr.ZhLL7713LL6208r.ZLLCARDr.ZG0Br"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8224"
            .area = "GV"
            .name = "Soft Soil Near Cow (Young)"
            .gs = True
            .logic = "YJG0D"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8225"
            .area = "GV"
            .name = "Waterfall Near Entrance (N) (Young)"
            .gs = True
            .logic = "Yo"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8226"
            .area = "GV"
            .name = "Pillar Near Tents (N) (Adult)"
            .gs = True
            .logic = "Zl.ZhLL7713LL6208k.ZLLCARDk.ZG0Bk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8227"
            .area = "GV"
            .name = "Behind Tents (N) (Adult)"
            .gs = True
            .logic = "Zl.ZhLL7713LL6208k.ZLLCARDk.ZG0Bk"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2624"
            .area = "GV"
            .name = "Near Soft Soil"
            .cow = True
            .logic = "YhLL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9209"
            .area = "GV"
            .name = "Storms Grotto Front (Adult)"
            .scrub = True
            .logic = "ZlG02.ZhLL7713LL6208G02.ZLLCARDG02.ZG0BG02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9208"
            .area = "GV"
            .name = "Storms Grotto Back (Adult)"
            .scrub = True
            .logic = "ZlG02.ZhLL7713LL6208G02.ZLLCARDG02.ZG0BG02"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGF(ByRef tk As Integer)
        ' Set up keys for Gerudo's Fortress
        With aKeys(tk)
            .loc = "5600"
            .area = "GF"
            .name = "Chest on Top(Adult)"
            .logic = "Zl.ZLL7430.ZkhLL6512"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "CARD"
            .area = "GF"
            .name = "Membership Card (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7200"
            .area = "GF"
            .name = "Archery 1000 points (Adult)"
            .logic = "ZhLL7713LL6208LL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6831"
            .area = "GF"
            .name = "Archery 1500 points (Adult)"
            .logic = "ZhLL7713LL6208LL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6500"
            .area = "GF"
            .name = "Carpenter with 1 Torch"
            .scan = False
            .logic = "Z.LL7416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6502"
            .area = "GF"
            .name = "Carpenter with 2 Torches"
            .scan = False
            .logic = "Z.LL7416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6503"
            .area = "GF"
            .name = "Carpenter with 3 Torches"
            .scan = False
            .logic = "ZLL7722.Zd.Zk.ZLL7430.LL7416LL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6501"
            .area = "GF"
            .name = "Carpenter with 4 Torches"
            .scan = False
            .logic = "Z.LL7416"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8300"
            .area = "GF"
            .name = "Top Wall Near Scarecrow (N) (Adult)"
            .gs = True
            .logic = "ZNd.ZNk.ZNLL7430.ZNLL7722"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8301"
            .area = "GF"
            .name = "Far Archery Target (N) (Adult)"
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
        With aKeys(tk)
            .loc = "5700"
            .area = "HW"
            .name = "Structure Chest"
            .logic = "G0CP"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8309"
            .area = "HW"
            .name = "Structure Basement"
            .gs = True
            .logic = "YG0Co.ZG0Ck"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDC(ByRef tk As Integer)
        ' Set up keys for the Desert Colossus
        With aKeys(tk)
            .loc = "6810"
            .area = "DC"
            .name = "Great Fairy Fountain"
            .logic = "xhLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6628"
            .area = "DC"
            .name = "Song from Shiek"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2713"
            .area = "DC"
            .name = "Piece of Heart (Adult)"
            .logic = "G10G11q.G08G1011.G11LLLLB1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8308"
            .area = "DC"
            .name = "Soft Soil By Temple Entrance (Young)"
            .gs = True
            .logic = "G10JG0DhLL7709"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8310"
            .area = "DC"
            .name = "Bean Plant Ride Hill (N) (Adult)"
            .gs = True
            .logic = "G10G11ql.G08G10G11l.G11LLLLB1l"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8311"
            .area = "DC"
            .name = "Southern Edge On Tree (N) (Adult)"
            .gs = True
            .logic = "G11k"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9709"
            .area = "DC"
            .name = "Grotto Front (Adult)"
            .scrub = True
            .logic = "G11V02"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9708"
            .area = "DC"
            .name = "Grotto Back (Adult)"
            .scrub = True
            .logic = "G11V02"
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
        With aKeys(tk)
            .loc = "008"
            .area = "OGC"
            .name = "Great Fairy Fountain"
            .logic = "ZV03hLL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8116"
            .area = "OGC"
            .name = "On Pillar (Adult)"
            .gs = True
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
        With aKeys(tk)
            .loc = "124"
            .area = "QBPH"
            .name = "#1: Near Castle Gate"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "123"
            .area = "QBPH"
            .name = "#2: Near Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "122"
            .area = "QBPH"
            .name = "#3: West of Castle"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "130"
            .area = "QBPH"
            .name = "#4: Between Gerudo Valley and Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "131"
            .area = "QBPH"
            .name = "#5: Near Gerudo Valley"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "128"
            .area = "QBPH"
            .name = "#6: Southeast Field Near Path"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "129"
            .area = "QBPH"
            .name = "#7: Southeast Field Near Grotto"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "127"
            .area = "QBPH"
            .name = "#8: Between Kokiri Forest and Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "126"
            .area = "QBPH"
            .name = "#9: Wall East of Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "125"
            .area = "QBPH"
            .name = "#10: Near Kakariko Village"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQF(ByRef tk As Integer)
        ' Set up keys for Quest: Frogs
        With aKeys(tk)
            .loc = "6706"
            .area = "QF"
            .name = "Play Song of Storms (Young)"
            .logic = "LL7717"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6705"
            .area = "QF"
            .name = "Play Song of Time (Young)"
            .logic = "LL7716"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6704"
            .area = "QF"
            .name = "Play Saria's Song (Young)"
            .logic = "LL7714"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6703"
            .area = "QF"
            .name = "Play Sun's Song (Young)"
            .logic = "LL7715"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6702"
            .area = "QF"
            .name = "Play Epona's Song (Young)"
            .logic = "LL7713"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6701"
            .area = "QF"
            .name = "Play Zelda's Lullaby (Young)"
            .logic = "LL7712"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6700"
            .area = "QF"
            .name = "Ocarina Game (Young)"
            .logic = "LL7712LL7713LL7714LL7715LL7716LL7717"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQGS(ByRef tk As Integer)
        ' Set up keys for Quest: Gold Skulltula Rewards
        With aKeys(tk)
            .loc = "6710"
            .area = "QGS"
            .name = "10 Gold Skulltulas"
            .logic = "K1"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6711"
            .area = "QGS"
            .name = "20 Gold Skulltulas"
            .logic = "K2"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6712"
            .area = "QGS"
            .name = "30 Gold Skulltulas"
            .logic = "K3"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6713"
            .area = "QGS"
            .name = "40 Gold Skulltulas"
            .logic = "K4"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6714"
            .area = "QGS"
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
            .name = "Deliver Zelda's Letter to Unlock Mask Shoppe"
            .logic = "G18y3"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6908"
            .area = "QM"
            .name = "Sell Keaton Mask"
            .logic = "LLLC05"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6920"
            .area = "QM"
            .name = "Get Skull Mask"
            .logic = "LL6908"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6909"
            .area = "QM"
            .name = "Sell Skull Mask"
            .logic = "LL6920"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6910"
            .area = "QM"
            .name = "Sell Spooky Mask"
            .logic = "LL6909"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6911"
            .area = "QM"
            .name = "Sell Bunny Hood"
            .logic = "LL6910"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6926"
            .area = "QM"
            .name = "Get Mask of Truth"
            .logic = "LL6911"
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
        Select Case randoVer
            Case "AP"
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
        If items > 4 Then items = My.Settings.setShop

        ' Set starting key for shopsanity. Being the last 32 keys, -31 (need to count the one it is on)
        Dim startKey = keyCount - 31
        Dim totalItems As Byte = My.Settings.setShop
        Dim curShop As Integer = 0
        Dim shopArea() As String = {"MK", "MK", "ZD", "KV", "MK", "GC", "KF", "KV"}
        Dim shopName() As String = {"Potion Shop", "Bombchu Shop", "Shop", "Potion Shop", "Bazaar", "Shop", "Shop", "Bazaar"}
        Dim shopLogic() As String = {"Y", "Y", "Y.w", "Z", "Y", "Yx.YV01.YhLL7712.ZLL7025", "", "Z"}
        Dim curShelf As Integer = 0
        Dim shelfLoc() As String = {"Lower-Left", "Lower-Right", "Upper-Left", "Upper-Right"}

        ' Step through the last 32 keys, which will have to be kept the Shopsanity keys
        For i = 0 To 31

            ' Start with resetting the key
            aKeys(startKey + i) = New keyCheck

            If i < (items * 8) Then
                With aKeys(startKey + i)
                    .loc = "99" & IIf(i < 10, "0", "").ToString & i.ToString
                    .area = shopArea(curShop)
                    .name = shopName(curShop) & ": " & shelfLoc(curShelf)
                    .logic = shopLogic(curShop)
                    .shop = True
                    .scan = True
                    inc(curShelf)
                    If curShelf >= totalItems Then
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
                .name = "Map Chest"
            End With
            With aKeysDungeons(0)(1)
                .loc = "3101"
                .area = "DT1"
                .name = "Slingshot Chest"
            End With
            With aKeysDungeons(0)(2)
                .loc = "3105"
                .area = "DT1"
                .name = "Slingshot Room Side Chest"
            End With
            With aKeysDungeons(0)(3)
                .loc = "3102"
                .area = "DT0"
                .name = "Compass Chest"
            End With
            With aKeysDungeons(0)(4)
                .loc = "3106"
                .area = "DT0"
                .name = "Compass Room Side Chest"
            End With
            With aKeysDungeons(0)(5)
                .loc = "3104"
                .area = "DT0"
                .name = "Basement Chest"
                .logic = "J.b"
            End With
            With aKeysDungeons(0)(6)
                .loc = "1031"
                .area = "DT3"
                .name = "Queen Gohma"
                .logic = "LL7420LL7416.LL7420a"
            End With
            With aKeysDungeons(0)(7)
                .loc = "7803"
                .area = "DT0"
                .name = "Compass Room"
                .gs = True
                .logic = "J"
            End With
            With aKeysDungeons(0)(8)
                .loc = "7802"
                .area = "DT0"
                .name = "Basement Vines"
                .gs = True
                .logic = "f.x.Yg.Yo"
            End With
            With aKeysDungeons(0)(9)
                .loc = "7801"
                .area = "DT0"
                .name = "Basement Gate"
                .gs = True
                .logic = "J"
            End With
            With aKeysDungeons(0)(10)
                .loc = "7800"
                .area = "DT2"
                .name = "Basement Back Room"
                .gs = True
                .logic = "OG00o"
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
                .name = "Map Chest"
            End With
            With aKeysDungeons(0)(1)
                .loc = "3102"
                .area = "DT0"
                .name = "Slingshot Chest"
                .logic = "J"
            End With
            With aKeysDungeons(0)(2)
                .loc = "3106"
                .area = "DT0"
                .name = "Slingshot Back Room Chest"
                .logic = "O"
            End With
            With aKeysDungeons(0)(3)
                .loc = "3101"
                .area = "DT1"
                .name = "Compass Chest"
            End With
            With aKeysDungeons(0)(4)
                .loc = "3104"
                .area = "DT0"
                .name = "Basement Chest"
                .logic = "O"
            End With
            With aKeysDungeons(0)(5)
                .loc = "3105"
                .area = "DT2"
                .name = "Before Spinning Log Chest"
            End With
            With aKeysDungeons(0)(6)
                .loc = "3100"
                .area = "DT3"
                .name = "After Spinning Log Chest"
                .logic = "hLL7716"
            End With
            With aKeysDungeons(0)(7)
                .loc = "1031"
                .area = "DT5"
                .name = "Queen Gohma"
                .logic = "OLL7420"
            End With
            With aKeysDungeons(0)(8)
                .loc = "7801"
                .area = "DT0"
                .name = "Lobby"
                .gs = True
                .logic = "J"
            End With
            With aKeysDungeons(0)(9)
                .loc = "7803"
                .area = "DT1"
                .name = "Compass Room"
                .gs = True
                .logic = "oj.ochLL7716"
            End With
            With aKeysDungeons(0)(10)
                .loc = "7802"
                .area = "DT4"
                .name = "Basement Graves Room"
                .gs = True
                .logic = "hLL7716o"
            End With
            With aKeysDungeons(0)(11)
                .loc = "7800"
                .area = "DT4"
                .name = "Basement Back Room"
                .gs = True
                .logic = "Oo"
            End With
            With aKeysDungeons(0)(12)
                .loc = "8405"
                .area = "DT5"
                .name = "Basement"
                .scrub = True
                .logic = "S"
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
                .name = "Map Chest"
                .logic = "x.V01"
            End With
            With aKeysDungeons(1)(1)
                .loc = "3205"
                .area = "DDC2"
                .name = "Compass Chest"
                .logic = "x.V01"
            End With
            With aKeysDungeons(1)(2)
                .loc = "3206"
                .area = "DDC3"
                .name = "Bomb Flower Platform Chest"
            End With
            With aKeysDungeons(1)(3)
                .loc = "3204"
                .area = "DDC4"
                .name = "Bomb Bag Chest"
            End With
            With aKeysDungeons(1)(4)
                .loc = "3210"
                .area = "DDC4"
                .name = "End of Bridge Chest"
                .logic = "G00"
            End With
            With aKeysDungeons(1)(5)
                .loc = "4400"
                .area = "DDC5"
                .name = "Boss Room Chest"
            End With
            With aKeysDungeons(1)(6)
                .loc = "1131"
                .area = "DDC5"
                .name = "King Dodongo"
                .logic = "Zc.ZV01.cLL7416.V01LL7416.ca.V01a"
            End With
            With aKeysDungeons(1)(7)
                .loc = "7812"
                .area = "DDC1"
                .name = "Side Room Near Lower Lizalfos"
                .gs = True
                .logic = "Z.x.V01LL7416.V01a.V01g.V01o"
            End With
            With aKeysDungeons(1)(8)
                .loc = "7809"
                .area = "DDC1"
                .name = "Scarecrow"
                .gs = True
                .logic = "ZkhLL6512.Zl"
            End With
            With aKeysDungeons(1)(9)
                .loc = "7810"
                .area = "DDC4"
                .name = "Alcove Above Stairs"
                .gs = True
                .logic = "Zk.Yo"
            End With
            With aKeysDungeons(1)(10)
                .loc = "7808"
                .area = "DDC3"
                .name = "Vines Above Stairs"
                .gs = True
            End With
            With aKeysDungeons(1)(11)
                .loc = "7811"
                .area = "DDC5"
                .name = "Back Room"
                .gs = True
                .logic = "G00"
            End With
            With aKeysDungeons(1)(12)
                .loc = "8505"
                .area = "DDC1"
                .name = "Lobby"
                .scrub = True
                .logic = "Z.S"
            End With
            With aKeysDungeons(1)(13)
                .loc = "8502"
                .area = "DDC1"
                .name = "Side Room Near Dodongos"
                .scrub = True
                .logic = "ZG00.ZV01.G00LL7416.V01LL7416.G00a.V01a.G00c.V01c.G00g.V01g"
            End With
            With aKeysDungeons(1)(14)
                .loc = "8501"
                .area = "DDC3"
                .name = "Near Bomb Bag Left"
                .scrub = True
                .logic = "G00"
            End With
            With aKeysDungeons(1)(15)
                .loc = "8504"
                .area = "DDC3"
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
                .name = "Map Chest"
                .logic = "G00.V01"
            End With
            With aKeysDungeons(1)(1)
                .loc = "3204"
                .area = "DDC4"
                .name = "Bomb Bag Chest"
            End With
            With aKeysDungeons(1)(2)
                .loc = "3203"
                .area = "DDC2"
                .name = "Torch Puzzle Room Chest"
                .logic = "G00.a.f.ZA0.ZLL7430.k"
            End With
            With aKeysDungeons(1)(3)
                .loc = "3202"
                .area = "DDC2"
                .name = "Larvae Room Chest"
                .logic = "P.a"
            End With
            With aKeysDungeons(1)(4)
                .loc = "3205"
                .area = "DDC2"
                .name = "Compass Chest"
                .logic = "Z.J.b"
            End With
            With aKeysDungeons(1)(5)
                .loc = "3201"
                .area = "DDC5"
                .name = "Under Grave Chest"
            End With
            With aKeysDungeons(1)(6)
                .loc = "4400"
                .area = "DDC6"
                .name = "Boss Room Chest"
            End With
            With aKeysDungeons(1)(7)
                .loc = "1131"
                .area = "DDC6"
                .name = "King Dodongo"
                .logic = "rc.rV01.xa.xLL7416.xZ.rV01"
            End With
            With aKeysDungeons(1)(8)
                .loc = "7810"
                .area = "DDC4"
                .name = "Scrub Room"
                .gs = True
                .logic = "dl.V01l.fl.xl.do.V01o.fo.xo"
            End With
            With aKeysDungeons(1)(9)
                .loc = "7812"
                .area = "DDC2"
                .name = "Larvae Room"
                .gs = True
                .logic = "P.a"
            End With
            With aKeysDungeons(1)(10)
                .loc = "7809"
                .area = "DDC2"
                .name = "Lizalfos Room"
                .gs = True
                .logic = "G00"
            End With
            With aKeysDungeons(1)(11)
                .loc = "7811"
                .area = "DDC2"
                .name = "Song of Time Block Room"
                .gs = True
                .logic = "JhLL7716.ZhLL7716"
            End With
            With aKeysDungeons(1)(12)
                .loc = "7808"
                .area = "DDC5"
                .name = "Back Area"
                .gs = True
                .logic = "Z.x.o.f.abLL7416.aoLL7416.abg.aog"
            End With
            With aKeysDungeons(1)(13)
                .loc = "8504"
                .area = "DDC1"
                .name = "Lobby Front"
                .scrub = True
                .logic = "Z.S"
            End With
            With aKeysDungeons(1)(14)
                .loc = "8502"
                .area = "DDC1"
                .name = "Lobby Rear"
                .scrub = True
                .logic = "Z.S"
            End With
            With aKeysDungeons(1)(15)
                .loc = "8508"
                .area = "DDC3"
                .name = "Side Room Near Lower Lizalfos"
                .scrub = True
                .logic = "G00Z.G00J.V01Z.V01J"
            End With
            With aKeysDungeons(1)(16)
                .loc = "8505"
                .area = "DDC2"
                .name = "Staircase Room"
                .scrub = True
                .logic = "Z.S"
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
                .name = "Boomerang Chest"
            End With
            With aKeysDungeons(2)(1)
                .loc = "3302"
                .area = "JB2"
                .name = "Map Chest"
            End With
            With aKeysDungeons(2)(2)
                .loc = "3304"
                .area = "JB2"
                .name = "Compass Chest"
            End With
            With aKeysDungeons(2)(3)
                .loc = "1231"
                .area = "JB3"
                .name = "Barinade"
                .logic = "LL7416o.ao"
            End With
            With aKeysDungeons(2)(4)
                .loc = "7819"
                .area = "JB1"
                .name = "Water Switch Room"
                .gs = True
            End With
            With aKeysDungeons(2)(5)
                .loc = "7817"
                .area = "JB1"
                .name = "Lobby Basement Upper"
                .gs = True
                .logic = "o"
            End With
            With aKeysDungeons(2)(6)
                .loc = "7816"
                .area = "JB1"
                .name = "Lobby Basement Lower"
                .gs = True
                .logic = "o"
            End With
            With aKeysDungeons(2)(7)
                .loc = "7818"
                .area = "JB3"
                .name = "Near Boss"
                .gs = True
            End With
            With aKeysDungeons(2)(8)
                .loc = "8601"
                .area = "JB1"
                .name = "Through Water Passage"
                .scrub = True
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
                .name = "Map Chest"
                .logic = "G00"
            End With
            With aKeysDungeons(2)(1)
                .loc = "3305"
                .area = "JB0"
                .name = "First Room Side Chest"
                .logic = "g"
            End With
            With aKeysDungeons(2)(2)
                .loc = "3302"
                .area = "JB1"
                .name = "Second Room Lower Chest"
            End With
            With aKeysDungeons(2)(3)
                .loc = "3300"
                .area = "JB1"
                .name = "Compass Chest"
                .logic = "g.j"
            End With
            With aKeysDungeons(2)(4)
                .loc = "3308"
                .area = "JB2"
                .name = "Basement Near Vines Chest"
                .logic = "g"
            End With
            With aKeysDungeons(2)(5)
                .loc = "3304"
                .area = "JB2"
                .name = "Basement Near Switches Chest"
                .logic = "g"
            End With
            With aKeysDungeons(2)(6)
                .loc = "3306"
                .area = "JB2"
                .name = "Boomerang Chest"
                .logic = "LL7416.a.c.g"
            End With
            With aKeysDungeons(2)(7)
                .loc = "3301"
                .area = "JB2"
                .name = "Boomerang Room Small Chest"
            End With
            With aKeysDungeons(2)(8)
                .loc = "3309"
                .area = "JB3"
                .name = "Falling Like Like Room"
            End With
            With aKeysDungeons(2)(9)
                .loc = "3307"
                .area = "JB5"
                .name = "Second Room Upper Chest"
                .logic = "g"
            End With
            With aKeysDungeons(2)(10)
                .loc = "3310"
                .area = "JB5"
                .name = "Near Boss Chest"
                .logic = "g"
            End With
            With aKeysDungeons(2)(11)
                .loc = "1231"
                .area = "JB5"
                .name = "Barinade"
                .logic = "oLL7416.oa"
            End With
            With aKeysDungeons(2)(12)
                .loc = "7816"
                .area = "JB2"
                .name = "Boomerang Chest Room"
                .gs = True
                .logic = "hLL7716J"
            End With
            With aKeysDungeons(2)(13)
                .loc = "7818"
                .area = "JB3"
                .name = "Tailpasaran Room"
                .gs = True
                .logic = "a.f"
            End With
            With aKeysDungeons(2)(14)
                .loc = "7819"
                .area = "JB3"
                .name = "Invisible Enemies Room"
                .gs = True
                .logic = "Tgo"
            End With
            With aKeysDungeons(2)(15)
                .loc = "7817"
                .area = "JB5"
                .name = "Near Boss"
                .gs = True
                .logic = "o"
            End With
            With aKeysDungeons(2)(16)
                .loc = "10224"
                .area = "JB4"
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
                .name = "First Room Chest"
            End With
            With aKeysDungeons(3)(1)
                .loc = "3400"
                .area = "FOT0"
                .name = "First Stalfos Chest"
            End With
            With aKeysDungeons(3)(2)
                .loc = "3405"
                .area = "FOT2"
                .name = "Raised Island Courtyard Chest"
                .logic = "k.Q304.A1LL7430Q303"
            End With
            With aKeysDungeons(3)(3)
                .loc = "3401"
                .area = "FOT3"
                .name = "Map Chest"
            End With
            With aKeysDungeons(3)(4)
                .loc = "3409"
                .area = "FOT3"
                .name = "Well Chest"
            End With
            With aKeysDungeons(3)(5)
                .loc = "3404"
                .area = "FOT5"
                .name = "Eye Switch Chest"
                .logic = "V01d"
            End With
            With aKeysDungeons(3)(6)
                .loc = "3414"
                .area = "FOT6"
                .name = "Boss Key Chest"
            End With
            With aKeysDungeons(3)(7)
                .loc = "3402"
                .area = "FOT7"
                .name = "Floormaster Chest"
            End With
            With aKeysDungeons(3)(8)
                .loc = "3413"
                .area = "FOT8"
                .name = "Red Poe Chest"
                .logic = "d"
            End With
            With aKeysDungeons(3)(9)
                .loc = "3412"
                .area = "FOT8"
                .name = "Bow Chest"
            End With
            With aKeysDungeons(3)(10)
                .loc = "3415"
                .area = "FOT8"
                .name = "Blue Poe Chest"
                .logic = "d"
            End With
            With aKeysDungeons(3)(11)
                .loc = "3407"
                .area = "FOT4"
                .name = "Falling Ceiling Room Chest"
            End With
            With aKeysDungeons(3)(12)
                .loc = "3411"
                .area = "FOT9"
                .name = "Basement Chest"
            End With
            With aKeysDungeons(3)(13)
                .loc = "1331"
                .area = "FOT9"
                .name = "Phantom Ganon"
                .logic = "GB0d.GB0k"
            End With
            With aKeysDungeons(3)(14)
                .loc = "7825"
                .area = "FOT0"
                .name = "First Room"
                .gs = True
                .logic = "c.d.f.j.k"
            End With
            With aKeysDungeons(3)(15)
                .loc = "7827"
                .area = "FOT0"
                .name = "Lobby"
                .gs = True
                .logic = "k"
            End With
            With aKeysDungeons(3)(16)
                .loc = "7826"
                .area = "FOT1"
                .name = "Level Island Courtyard"
                .gs = True
                .logic = "l.Q307k"
            End With
            With aKeysDungeons(3)(17)
                .loc = "7824"
                .area = "FOT2"
                .name = "Raised Island Courtyard"
                .gs = True
                .logic = "k.Q304d.Q304f.Q304x"
            End With
            With aKeysDungeons(3)(18)
                .loc = "7828"
                .area = "FOT9"
                .name = "Basement"
                .gs = True
                .logic = "k"
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
            ' 1: MQ Forest Temple Central Ara
            ' 2: MQ Forest Temple After Block Puzzle
            ' 3: MQ Forest Temple Outdoor Ledge
            ' 4: MQ Forest Temple NW Outdoors
            ' 5: MQ Forest Temple NE Outdoors
            ' 6: MQ Forest Temple Outdoors Top Ledges
            ' 7: MQ Forest Temple NW Outdoor Ledge
            ' 8: MQ Forest Temple Bow Region
            ' 9: MQ Forest Temple Falling Room
            ' A: MQ Forest Temple Boss Region
            ' B: MQ Forest Temple Big Key Chest Area

            With aKeysDungeons(3)(0)
                .loc = "3403"
                .area = "FOT0"
                .name = "First Room Chest"
            End With
            With aKeysDungeons(3)(1)
                .loc = "3400"
                .area = "FOT1"
                .name = "Wolfos Chest"
                .logic = "hLL7716"
            End With
            With aKeysDungeons(3)(2)
                .loc = "3409"
                .area = "FOT5"
                .name = "Well Chest"
                .logic = "d"
            End With
            With aKeysDungeons(3)(3)
                .loc = "3401"
                .area = "FOT7"
                .name = "Raised Island Courtyard Lower Chest"
            End With
            With aKeysDungeons(3)(4)
                .loc = "3405"
                .area = "FOT6"
                .name = "Raised Island Courtyard Upper Chest"
            End With
            With aKeysDungeons(3)(5)
                .loc = "3414"
                .area = "FOTB"
                .name = "Boss Key Chest"
            End With
            With aKeysDungeons(3)(6)
                .loc = "3402"
                .area = "FOT3"
                .name = "Redead Chest"
            End With
            With aKeysDungeons(3)(7)
                .loc = "3413"
                .area = "FOT8"
                .name = "Map Chest"
                .logic = "d"
            End With
            With aKeysDungeons(3)(8)
                .loc = "3412"
                .area = "FOT8"
                .name = "Bow Chest"
            End With
            With aKeysDungeons(3)(9)
                .loc = "3415"
                .area = "FOT8"
                .name = "Compass Chest"
                .logic = "d"
            End With
            With aKeysDungeons(3)(10)
                .loc = "3406"
                .area = "FOT9"
                .name = "Falling Ceiling Room Chest"
            End With
            With aKeysDungeons(3)(11)
                .loc = "3411"
                .area = "FOTA"
                .name = "Basement Chest"
            End With
            With aKeysDungeons(3)(12)
                .loc = "1331"
                .area = "FOTA"
                .name = "Phantom Ganon"
                .logic = "GB0d.GB0k"
            End With
            With aKeysDungeons(3)(13)
                .loc = "7825"
                .area = "FOT0"
                .name = "First Hallway"
                .gs = True
                .logic = "k"
            End With
            With aKeysDungeons(3)(14)
                .loc = "7824"
                .area = "FOT5"
                .name = "Raised Island Courtyard"
                .gs = True
                .logic = "k.deLL7716.deLL7430A2"
            End With
            With aKeysDungeons(3)(15)
                .loc = "7826"
                .area = "FOT4"
                .name = "Level Island Courtyard"
                .gs = True
            End With
            With aKeysDungeons(3)(16)
                .loc = "7827"
                .area = "FOT5"
                .name = "Well"
                .gs = True
                .logic = "LL7729k.d"
            End With
            With aKeysDungeons(3)(17)
                .loc = "7828"
                .area = "FOT1"
                .name = "Block Push Room"
                .gs = True
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
                .name = "Near Boss Chest"
            End With
            With aKeysDungeons(4)(1)
                .loc = "3500"
                .area = "FIT4"
                .name = "Flare Dancer Chest"
            End With
            With aKeysDungeons(4)(2)
                .loc = "3512"
                .area = "FIT4"
                .name = "Boss Key Chest"
            End With
            With aKeysDungeons(4)(3)
                .loc = "3504"
                .area = "FIT1"
                .name = "Big Lava Room Lower Open Door Chest"
            End With
            With aKeysDungeons(4)(4)
                .loc = "3502"
                .area = "FIT1"
                .name = "Big Lava Room Blocked Door Chest"
                .logic = "x"
            End With
            With aKeysDungeons(4)(5)
                .loc = "3503"
                .area = "FIT2"
                .name = "Boulder Maze Lower Chest"
            End With
            With aKeysDungeons(4)(6)
                .loc = "3508"
                .area = "FIT2"
                .name = "Boulder Maze Side Room Chest"
            End With
            With aKeysDungeons(4)(7)
                .loc = "3510"
                .area = "FIT2"
                .name = "Map Chest"
                .logic = "GC2.GC1d"
            End With
            With aKeysDungeons(4)(8)
                .loc = "3506"
                .area = "FIT2"
                .name = "Boulder Maze Upper Chest"
                .logic = "GC2"
            End With
            With aKeysDungeons(4)(9)
                .loc = "3511"
                .area = "FIT2"
                .name = "Boulder Maze Shortcut Chest"
                .logic = "GC2x"
            End With
            With aKeysDungeons(4)(10)
                .loc = "3513"
                .area = "FIT2"
                .name = "Scarecrow Chest"
                .logic = "GC2hkLL6512"
            End With
            With aKeysDungeons(4)(11)
                .loc = "3507"
                .area = "FIT2"
                .name = "Compass Chest"
                .logic = "GC3"
            End With
            With aKeysDungeons(4)(12)
                .loc = "3509"
                .area = "FIT3"
                .name = "Highest Goron Chest"
                .logic = "rhLL7716.rA3LL7430.rA3x"
            End With
            With aKeysDungeons(4)(13)
                .loc = "3505"
                .area = "FIT3"
                .name = "Megaton Hammer Chest"
                .logic = "x"
            End With
            With aKeysDungeons(4)(14)
                .loc = "1431"
                .area = "FIT0"
                .name = "Volvagia"
                .logic = "GB1LL7425rLL7430.GB1LL7425rL10718.GB1LL7425rQ403hLL7716.GB1LL7425rQ403x"
            End With
            With aKeysDungeons(4)(15)
                .loc = "7901"
                .area = "FIT4"
                .name = "Boss Key Loop"
                .gs = True
            End With
            With aKeysDungeons(4)(16)
                .loc = "7900"
                .area = "FIT1"
                .name = "Song of Time Room"
                .gs = True
                .logic = "hLL7716"
            End With
            With aKeysDungeons(4)(17)
                .loc = "7902"
                .area = "FIT2"
                .name = "Boulder Maze"
                .gs = True
                .logic = "GC0x"
            End With
            With aKeysDungeons(4)(18)
                .loc = "7904"
                .area = "FIT2"
                .name = "Scarecrow Climb"
                .gs = True
                .logic = "GC2hkLL6512"
            End With
            With aKeysDungeons(4)(19)
                .loc = "7903"
                .area = "FIT2"
                .name = "Scar ecrow Top"
                .gs = True
                .logic = "GC2hkLL6512"
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
                .name = "Near Boss Chest"
                .logic = "LL7430P.kP"
            End With
            With aKeysDungeons(4)(1)
                .loc = "3502"
                .area = "FIT0"
                .name = "Map Room Side Chest"
            End With
            With aKeysDungeons(4)(2)
                .loc = "3500"
                .area = "FIT1"
                .name = "Megaton Hammer Chest"
                .logic = "k.r.x"
            End With
            With aKeysDungeons(4)(3)
                .loc = "3512"
                .area = "FIT1"
                .name = "Map Chest"
                .logic = "r"
            End With
            With aKeysDungeons(4)(4)
                .loc = "3501"
                .area = "FIT2"
                .name = "Big Lava Room Blocked Door Chest"
                .logic = "kxP"
            End With
            With aKeysDungeons(4)(5)
                .loc = "3504"
                .area = "FIT2"
                .name = "Boss Key Chest"
                .logic = "dkP"
            End With
            With aKeysDungeons(4)(6)
                .loc = "3503"
                .area = "FIT3"
                .name = "Lizalfos Maze Lower Chest"
            End With
            With aKeysDungeons(4)(7)
                .loc = "3508"
                .area = "FIT3"
                .name = "Lizalfos Maze Side Room Chest"
                .logic = "Q404x"
            End With
            With aKeysDungeons(4)(8)
                .loc = "3506"
                .area = "FIT4"
                .name = "Lizalfos Maze Upper Chest"
            End With
            With aKeysDungeons(4)(9)
                .loc = "3511"
                .area = "FIT4"
                .name = "Compass Chest"
                .logic = "x"
            End With
            With aKeysDungeons(4)(10)
                .loc = "328"
                .area = "FIT5"
                .name = "Under Platform Freestanding Key"
                .logic = "k.A4"
            End With
            With aKeysDungeons(4)(11)
                .loc = "3505"
                .area = "FIT5"
                .name = "Chest on Fire"
                .logic = "C3k.C3A4"
            End With
            With aKeysDungeons(4)(12)
                .loc = "1431"
                .area = "FIT6"
                .name = "Volvagia"
                .logic = "GB1"
            End With
            With aKeysDungeons(4)(13)
                .loc = "7900"
                .area = "FIT2"
                .name = "Big Lava Room Open Door"
                .gs = True
            End With
            With aKeysDungeons(4)(14)
                .loc = "7902"
                .area = "FIT4"
                .name = "Skull On Fire"
                .gs = True
                .logic = "l.hxkLL7716.A3khLL7716"
            End With
            With aKeysDungeons(4)(15)
                .loc = "7903"
                .area = "FIT5"
                .name = "Fire Wall Maze Centre"
                .gs = True
                .logic = "x"
            End With
            With aKeysDungeons(4)(16)
                .loc = "7904"
                .area = "FIT5"
                .name = "Fire Wall Maze Side Room"
                .gs = True
                .logic = "A4.LL7430.hLL7716"
            End With
            With aKeysDungeons(4)(17)
                .loc = "7901"
                .area = "FIT5"
                .name = "Above Fire Wall Maze"
                .gs = True
                .logic = "C4k"
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
                .name = "Map Chest"
                .logic = "Q501"
            End With
            With aKeysDungeons(5)(1)
                .loc = "3600"
                .area = "WAT4"
                .name = "Cracked Wall Chest"
                .logic = "x"
            End With
            With aKeysDungeons(5)(2)
                .loc = "3601"
                .area = "WAT2"
                .name = "Torches Chest"
                .logic = "ehLL7712.fhLL7712"
            End With
            With aKeysDungeons(5)(3)
                .loc = "3609"
                .area = "WAT2"
                .name = "Compass Chest"
                .logic = "hLL7716k.LL7429k"
            End With
            With aKeysDungeons(5)(4)
                .loc = "3608"
                .area = "WAT2"
                .name = "Central Bow Target Chest"
                .logic = "V01hLL7712dLL7730.V01hLL7712dl"
            End With
            With aKeysDungeons(5)(5)
                .loc = "3606"
                .area = "WAT6"
                .name = "Central Pillar Chest"
                .logic = "LL7429LL7426kGC5.LL7429LL7426kd.LL7429LL7426kf"
            End With
            With aKeysDungeons(5)(6)
                .loc = "3607"
                .area = "WAT8"
                .name = "Longshot Chest"
            End With
            With aKeysDungeons(5)(7)
                .loc = "3603"
                .area = "WAT8"
                .name = "River Chest"
                .logic = "hLL7716d"
            End With
            With aKeysDungeons(5)(8)
                .loc = "3610"
                .area = "WAT5"
                .name = "Dragon Chest"
            End With
            With aKeysDungeons(5)(9)
                .loc = "3605"
                .area = "WAT3"
                .name = "Boss Key Chest"
                .logic = "GC6LL7429LL7430.GC6LL7429xV01"
            End With
            With aKeysDungeons(5)(10)
                .loc = "1531"
                .area = "WAT1"
                .name = "Morpha"
                .logic = "GB2l"
            End With
            With aKeysDungeons(5)(11)
                .loc = "7908"
                .area = "WAT2"
                .name = "Behind Gate"
                .gs = True
                .logic = "kxhLL7712LL7429.LL7430xhLL7712LL7429.kxhLL7712V11.LL7430xhLL7712V11"
            End With
            With aKeysDungeons(5)(12)
                .loc = "7911"
                .area = "WAT3"
                .name = "Near Boss Key Chest"
                .gs = True
            End With
            With aKeysDungeons(5)(13)
                .loc = "7910"
                .area = "WAT2"
                .name = "Central Pillar"
                .gs = True
                .logic = "hLL7712lGC5.hLL7712ld.hLL7712lf"
            End With
            With aKeysDungeons(5)(14)
                .loc = "7909"
                .area = "WAT7"
                .name = "Falling Platform Room"
                .gs = True
                .logic = "l"
            End With
            With aKeysDungeons(5)(15)
                .loc = "7912"
                .area = "WAT8"
                .name = "River"
                .gs = True
                .logic = "hLL7716LL7429G0A"
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
                .name = "Map Chest"
                .logic = "Pk"
            End With
            With aKeysDungeons(5)(1)
                .loc = "3600"
                .area = "WAT2"
                .name = "Longshot Chest"
                .logic = "k"
            End With
            With aKeysDungeons(5)(2)
                .loc = "3601"
                .area = "WAT2"
                .name = "Compass Chest"
                .logic = "d.f"
            End With
            With aKeysDungeons(5)(3)
                .loc = "3606"
                .area = "WAT1"
                .name = "Central Pillar Chest"
                .logic = "LL7426kfhLL7716"
            End With
            With aKeysDungeons(5)(4)
                .loc = "3605"
                .area = "WAT3"
                .name = "Boss Key Chest"
                .logic = "G0AfV11.G0AfLL7429"
            End With
            With aKeysDungeons(5)(5)
                .loc = "401"
                .area = "WAT4"
                .name = "Stalfos Room Freestanding Key"
                .logic = "LL7430.hLL6512k"
            End With
            With aKeysDungeons(5)(6)
                .loc = "1531"
                .area = "WAT0"
                .name = "Morpha"
                .logic = "GB2l"
            End With
            With aKeysDungeons(5)(7)
                .loc = "7908"
                .area = "WAT2"
                .name = "Lizalfos Hallway"
                .gs = True
                .logic = "f"
            End With
            With aKeysDungeons(5)(8)
                .loc = "7909"
                .area = "WAT3"
                .name = "River"
                .gs = True
            End With
            With aKeysDungeons(5)(9)
                .loc = "7910"
                .area = "WAT2"
                .name = "Before Upper Water Switch"
                .gs = True
                .logic = "l"
            End With
            With aKeysDungeons(5)(10)
                .loc = "7911"
                .area = "WAT4"
                .name = "Freestanding Key Area"
                .gs = True
                .logic = "GC7LL7430.GC7hLL6512k"
            End With
            With aKeysDungeons(5)(11)
                .loc = "7912"
                .area = "WAT4"
                .name = "Triple Wall Torch"
                .gs = True
                .logic = "deLL7430.dehLL6512k"
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
                .name = "Child Bridge Chest"
                .logic = "ga.gb.go.gx.oa.ox.oLL7416.jA5a.jA5x.jA5bLL7416"
            End With
            With aKeysDungeons(6)(1)
                .loc = "3700"
                .area = "SPT1"
                .name = "Child Early Torches Chest"
                .logic = "ga.gbf.gof.gxf.oa.oxf.oLL7416f.jA5a.jA5xf.jA5bLL7416f"
            End With
            With aKeysDungeons(6)(2)
                .loc = "3706"
                .area = "SPT3"
                .name = "Child Climb North Chest"
                .logic = "G12G20.G13G21"
            End With
            With aKeysDungeons(6)(3)
                .loc = "3712"
                .area = "SPT3"
                .name = "Child Climb East Chest"
                .logic = "G12G20.G13G21"
            End With
            With aKeysDungeons(6)(4)
                .loc = "3701"
                .area = "SPT5"
                .name = "Sun Block Room Chest"
                .logic = "G12a.G13de.f"
            End With
            With aKeysDungeons(6)(5)
                .loc = "3703"
                .area = "SPT5"
                .name = "Map Chest"
                .logic = "xf.G13xde"
            End With
            With aKeysDungeons(6)(6)
                .loc = "5511"
                .area = "SPT6"
                .name = "Silver Gauntlets Chest"
                .logic = "G12GC8.G13GC9"
            End With
            With aKeysDungeons(6)(7)
                .loc = "3704"
                .area = "SPT4"
                .name = "Compass Chest"
                .logic = "khLL7712"
            End With
            With aKeysDungeons(6)(8)
                .loc = "3707"
                .area = "SPT4"
                .name = "Early Adult Right Chest"
                .logic = "d.j.k"
            End With
            With aKeysDungeons(6)(9)
                .loc = "3713"
                .area = "SPT5"
                .name = "First Mirror Left Chest"
                .logic = "V02G13.G14k"
            End With
            With aKeysDungeons(6)(10)
                .loc = "3714"
                .area = "SPT5"
                .name = "First Mirror Right Chest"
                .logic = "V02G13.G14k"
            End With
            With aKeysDungeons(6)(11)
                .loc = "3702"
                .area = "SPT5"
                .name = "Statue Room Hand Chest"
                .logic = "G13Q609hLL7712.G13LL7430hLL7712"
            End With
            With aKeysDungeons(6)(12)
                .loc = "3715"
                .area = "SPT5"
                .name = "Statue Room Northeast Chest"
                .logic = "G13Q609hLL7712k.G13hLL7712LL7430"
            End With
            With aKeysDungeons(6)(13)
                .loc = "3705"
                .area = "SPT7"
                .name = "Near Four Armos Chest"
                .logic = "LL7422x"
            End With
            With aKeysDungeons(6)(14)
                .loc = "3721"
                .area = "SPT7"
                .name = "Hallway Left Invisible Chest"
                .logic = "Lx"
            End With
            With aKeysDungeons(6)(15)
                .loc = "3720"
                .area = "SPT7"
                .name = "Hallway Right Invisible Chest"
                .logic = "Lx"
            End With
            With aKeysDungeons(6)(16)
                .loc = "5509"
                .area = "SPT7"
                .name = "Mirror Shield Chest"
                .logic = "x"
            End With
            With aKeysDungeons(6)(17)
                .loc = "3710"
                .area = "SPT8"
                .name = "Boss Key Chest"
                .logic = "hLL7712dk"
            End With
            With aKeysDungeons(6)(18)
                .loc = "3718"
                .area = "SPT8"
                .name = "Topmost Chest"
                .logic = "LL7422"
            End With
            With aKeysDungeons(6)(19)
                .loc = "1631"
                .area = "SPT9"
                .name = "Twinrova"
                .logic = "GB3LL7422k"
            End With
            With aKeysDungeons(6)(20)
                .loc = "7920"
                .area = "SPT1"
                .name = "Metal Fence"
                .gs = True
                .logic = "ga.gb.go.gx.oa.ox.oLL7416.jA5a.jA5x.jA5bLL7416"
            End With
            With aKeysDungeons(6)(21)
                .loc = "7919"
                .area = "SPT3"
                .name = "Sun on Floor Room"
                .gs = True
                .logic = "f.G12G20.G13"
            End With
            With aKeysDungeons(6)(22)
                .loc = "7916"
                .area = "SPT5"
                .name = "Hall After Sun Block Room"
                .gs = True
                .logic = "G12o.G13k"
            End With
            With aKeysDungeons(6)(23)
                .loc = "7918"
                .area = "SPT5"
                .name = "Lobby"
                .gs = True
                .logic = "G13LL7430.G13khLL6512"
            End With
            With aKeysDungeons(6)(24)
                .loc = "7917"
                .area = "SPT4"
                .name = "Boulder Room"
                .gs = True
                .logic = "hLL7716d.hLL7716j.hLL7716k"
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
                .loc = "10924"
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
                .name = "Entrance Front Left Chest"
            End With
            With aKeysDungeons(6)(1)
                .loc = "3730"
                .area = "SPT0"
                .name = "Entrance Back Left Chest"
                .logic = "G00g.G00d"
            End With
            With aKeysDungeons(6)(2)
                .loc = "3731"
                .area = "SPT0"
                .name = "Entrance Back Right Chest"
                .logic = "j.d.k.g.o"
            End With
            With aKeysDungeons(6)(3)
                .loc = "3708"
                .area = "SPT1"
                .name = "Map Room Enemy Chest"
                .logic = "G12ajgf.G12LL7416jgf"
            End With
            With aKeysDungeons(6)(4)
                .loc = "3700"
                .area = "SPT1"
                .name = "Map Chest"
                .logic = "G12a.G12LL7416.G12c"
            End With
            With aKeysDungeons(6)(5)
                .loc = "3729"
                .area = "SPT1"
                .name = "Child Hammer Switch Chest"
                .logic = "G12G13GCAr"
            End With
            With aKeysDungeons(6)(6)
                .loc = "3728"
                .area = "SPT3"
                .name = "Silver Block Hallway Chest"
                .logic = "G12g"
            End With
            With aKeysDungeons(6)(7)
                .loc = "3703"
                .area = "SPT3"
                .name = "Compass Chest"
                .logic = "G12g.G13b"
            End With
            With aKeysDungeons(6)(8)
                .loc = "3701"
                .area = "SPT3"
                .name = "Sun Block Room Chest"
                .logic = "G13.G12hLL7716.G12GA6"
            End With
            With aKeysDungeons(6)(9)
                .loc = "5511"
                .area = "SPT8"
                .name = "Silver Gauntlets Chest"
            End With
            With aKeysDungeons(6)(10)
                .loc = "3706"
                .area = "SPT1"
                .name = "Child Climb North Chest"
                .logic = "x"
            End With
            With aKeysDungeons(6)(11)
                .loc = "3712"
                .area = "SPT1"
                .name = "Child Climb South Chest"
                .logic = "G13x"
            End With
            With aKeysDungeons(6)(12)
                .loc = "3715"
                .area = "SPT2"
                .name = "Statue Room Lullaby Chest"
                .logic = "hLL7712"
            End With
            With aKeysDungeons(6)(13)
                .loc = "3702"
                .area = "SPT2"
                .name = "Statue Room Invisible Chest"
                .logic = "T"
            End With
            With aKeysDungeons(6)(14)
                .loc = "3704"
                .area = "SPT4"
                .name = "Leever Room Chest"
            End With
            With aKeysDungeons(6)(15)
                .loc = "3707"
                .area = "SPT4"
                .name = "Symphony Room Chest"
                .logic = "GCDrhLL7712LL7713LL7715LL7716LL7717"
            End With
            With aKeysDungeons(6)(16)
                .loc = "3727"
                .area = "SPT4"
                .name = "Entrance Front Right Chest"
                .logic = "r"
            End With
            With aKeysDungeons(6)(17)
                .loc = "3725"
                .area = "SPT2"
                .name = "Beamos Room Chest"
                .logic = "GCBx"
            End With
            With aKeysDungeons(6)(18)
                .loc = "3724"
                .area = "SPT2"
                .name = "Chest Switch Chest"
                .logic = "GCBhLL7716x.GCBhLL7716b"
            End With
            With aKeysDungeons(6)(19)
                .loc = "3705"
                .area = "SPT2"
                .name = "Boss Key Chest"
                .logic = "GCBhLL7716LL7422x.GCBhLL7716LL7422b"
            End With
            With aKeysDungeons(6)(20)
                .loc = "5509"
                .area = "SPT7"
                .name = "Mirror Shield Chest"
            End With
            With aKeysDungeons(6)(21)
                .loc = "3718"
                .area = "SPT5"
                .name = "Mirror Puzzle Invisible Chest"
                .logic = "T"
            End With
            With aKeysDungeons(6)(22)
                .loc = "1631"
                .area = "SPT6"
                .name = "Twinrova"
                .logic = "GB3LL7422"
            End With
            With aKeysDungeons(6)(23)
                .loc = "7916"
                .area = "SPT3"
                .name = "Sun Block Room"
                .gs = True
                .logic = "G13"
            End With
            With aKeysDungeons(6)(24)
                .loc = "7917"
                .area = "SPT4"
                .name = "Leever Room"
                .gs = True
            End With
            With aKeysDungeons(6)(25)
                .loc = "7919"
                .area = "SPT4"
                .name = "Symphony Room"
                .gs = True
                .logic = "GCDrhLL7712LL7713LL7715LL7716LL7717"
            End With
            With aKeysDungeons(6)(26)
                .loc = "7918"
                .area = "SPT2"
                .name = "Nine Thrones Room West"
                .gs = True
                .logic = "GCC"
            End With
            With aKeysDungeons(6)(27)
                .loc = "7920"
                .area = "SPT2"
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
        ReDim aKeysDungeons(7)(CInt(IIf(isMQ, 33, 29)))

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
                .name = "Map Chest"
            End With
            With aKeysDungeons(7)(1)
                .loc = "3807"
                .area = "SHT1"
                .name = "Hover Boots Chest"
            End With
            With aKeysDungeons(7)(2)
                .loc = "3803"
                .area = "SHT2"
                .name = "Compass Chest"
            End With
            With aKeysDungeons(7)(3)
                .loc = "3802"
                .area = "SHT2"
                .name = "Early Silver Rupee Chest"
            End With
            With aKeysDungeons(7)(4)
                .loc = "3812"
                .area = "SHT3"
                .name = "Invisible Blades Visible Chest"
            End With
            With aKeysDungeons(7)(5)
                .loc = "3822"
                .area = "SHT3"
                .name = "Invisible Blades Invisible Chest"
            End With
            With aKeysDungeons(7)(6)
                .loc = "3805"
                .area = "SHT3"
                .name = "Falling Spikes Lower Chest"
            End With
            With aKeysDungeons(7)(7)
                .loc = "3804"
                .area = "SHT3"
                .name = "Falling Spikes Switch Chest"
                .logic = "GA3.V01"
            End With
            With aKeysDungeons(7)(8)
                .loc = "3806"
                .area = "SHT3"
                .name = "Falling Spikes Upper Chest"
                .logic = "GA3.V01"
            End With
            With aKeysDungeons(7)(9)
                .loc = "3809"
                .area = "SHT4"
                .name = "Invisible Spikes Chest"
            End With
            With aKeysDungeons(7)(10)
                .loc = "501"
                .area = "SHT4"
                .name = "Skull Pot Room Freestanding Key"
                .logic = "kx.kV01"
            End With
            With aKeysDungeons(7)(11)
                .loc = "3821"
                .area = "SHT5"
                .name = "Wind Hint Chest"
            End With
            With aKeysDungeons(7)(12)
                .loc = "3808"
                .area = "SHT6"
                .name = "After Wind Enemy Chest"
            End With
            With aKeysDungeons(7)(13)
                .loc = "3820"
                .area = "SHT6"
                .name = "After Wind Hidden Chest"
                .logic = "x"
            End With
            With aKeysDungeons(7)(14)
                .loc = "3810"
                .area = "SHT8"
                .name = "Spike Walls Left Chest"
                .logic = "f"
            End With
            With aKeysDungeons(7)(15)
                .loc = "3811"
                .area = "SHT8"
                .name = "Boss Key Chest"
                .logic = "f"
            End With
            With aKeysDungeons(7)(16)
                .loc = "3813"
                .area = "SHT8"
                .name = "Invisible Floormaster Chest"
            End With
            With aKeysDungeons(7)(17)
                .loc = "1731"
                .area = "SHT9"
                .name = "Bongo Bongo"
                .logic = "GB4d.GB4h"
            End With
            With aKeysDungeons(7)(18)
                .loc = "7927"
                .area = "SHT3"
                .name = "Like Like Room"
                .gs = True
            End With
            With aKeysDungeons(7)(19)
                .loc = "7925"
                .area = "SHT3"
                .name = "Falling Spikes Room"
                .gs = True
                .logic = "k"
            End With
            With aKeysDungeons(7)(20)
                .loc = "7924"
                .area = "SHT4"
                .name = "Single Giant Pot"
                .gs = True
                .logic = "k"
            End With
            With aKeysDungeons(7)(21)
                .loc = "7928"
                .area = "SHT7"
                .name = "Near Ship"
                .gs = True
                .logic = "l"
            End With
            With aKeysDungeons(7)(22)
                .loc = "7926"
                .area = "SHT8"
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
            ' 9: MQ Shadow Temple Wind Tunnel and After Wind
            ' A: MQ Shadow Temple Boat
            ' B: MQ Shadow Temple Beyond Boat
            ' C: MQ Shadow Temple Invisible Maze
            ' D: MQ Shadow Temple Near Boss

            With aKeysDungeons(7)(0)
                .loc = "3801"
                .area = "SHT2"
                .name = "Compass Chest"
            End With
            With aKeysDungeons(7)(1)
                .loc = "3807"
                .area = "SHT2"
                .name = "Hover Boots Chest"
                .logic = "hLL7716d"
            End With
            With aKeysDungeons(7)(2)
                .loc = "3802"
                .area = "SHT3"
                .name = "Map Chest"
            End With
            With aKeysDungeons(7)(3)
                .loc = "3803"
                .area = "SHT3"
                .name = "Early Gibdos Chest"
            End With
            With aKeysDungeons(7)(4)
                .loc = "3814"
                .area = "SHT3"
                .name = "Near Ship Invisible Chest"
            End With
            With aKeysDungeons(7)(5)
                .loc = "3812"
                .area = "SHT5"
                .name = "Invisible Blades Visible Chest"
            End With
            With aKeysDungeons(7)(6)
                .loc = "3822"
                .area = "SHT5"
                .name = "Invisible Blades Invisible Chest"
            End With
            With aKeysDungeons(7)(7)
                .loc = "3815"
                .area = "SHT7"
                .name = "Beamos Silver Rupee Chest"
                .logic = "l"
            End With
            With aKeysDungeons(7)(8)
                .loc = "3805"
                .area = "SHT7"
                .name = "Falling Spikes Lower Chest"
            End With
            With aKeysDungeons(7)(9)
                .loc = "3806"
                .area = "SHT7"
                .name = "Falling Spikes Upper Chest"
                .logic = "GA3.V01"
            End With
            With aKeysDungeons(7)(10)
                .loc = "3804"
                .area = "SHT7"
                .name = "Falling Spikes Switch Chest"
                .logic = "GA3.V01"
            End With
            With aKeysDungeons(7)(11)
                .loc = "3809"
                .area = "SHT8"
                .name = "Invisible Spikes Chest"
            End With
            With aKeysDungeons(7)(12)
                .loc = "3816"
                .area = "SHT8"
                .name = "Stalfos Room Chest"
                .logic = "h"
            End With
            With aKeysDungeons(7)(13)
                .loc = "3821"
                .area = "SHT9"
                .name = "Wind Hint Chest"
            End With
            With aKeysDungeons(7)(14)
                .loc = "3808"
                .area = "SHT9"
                .name = "After Wind Enemy Chest"
            End With
            With aKeysDungeons(7)(15)
                .loc = "3820"
                .area = "SHT9"
                .name = "After Wind Hidden Chest"
                .logic = "x"
            End With
            With aKeysDungeons(7)(16)
                .loc = "3810"
                .area = "SHTC"
                .name = "Spike Walls Left Chest"
                .logic = "fGCE"
            End With
            With aKeysDungeons(7)(17)
                .loc = "3811"
                .area = "SHTC"
                .name = "Boss Key Chest"
                .logic = "fGCE"
            End With
            With aKeysDungeons(7)(18)
                .loc = "3813"
                .area = "SHTC"
                .name = "Bomb Flower Chest"
            End With
            With aKeysDungeons(7)(19)
                .loc = "506"
                .area = "SHTC"
                .name = "Skull Pots Room Freestanding Key"
            End With
            With aKeysDungeons(7)(20)
                .loc = "1731"
                .area = "SHTD"
                .name = "Bongo Bongo"
                .logic = "GB4"
            End With
            With aKeysDungeons(7)(21)
                .loc = "7925"
                .area = "SHT7"
                .name = "Falling Spikes Room"
                .gs = True
                .logic = "h"
            End With
            With aKeysDungeons(7)(22)
                .loc = "7924"
                .area = "SHT9"
                .name = "Wind Hint Room"
                .gs = True
            End With
            With aKeysDungeons(7)(23)
                .loc = "7927"
                .area = "SHT9"
                .name = "After Wind"
                .gs = True
                .logic = "x"
            End With
            With aKeysDungeons(7)(24)
                .loc = "7928"
                .area = "SHTB"
                .name = "After Ship"
                .gs = True
            End With
            With aKeysDungeons(7)(25)
                .loc = "7926"
                .area = "SHTD"
                .name = "Near Boss"
                .gs = True
                .logic = "G21.f"
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
                .name = "Front Left Fake Wall Chest"
            End With
            With aKeysDungeons(8)(1)
                .loc = "3902"
                .area = "BW1"
                .name = "Front Centre Bombable Chest"
                .logic = "x"
            End With
            With aKeysDungeons(8)(2)
                .loc = "3905"
                .area = "BW2"
                .name = "Right Bottom Fake Wall Chest"
            End With
            With aKeysDungeons(8)(3)
                .loc = "3901"
                .area = "BW2"
                .name = "Compass Chest"
            End With
            With aKeysDungeons(8)(4)
                .loc = "3914"
                .area = "BW2"
                .name = "Centre Skulltula Chest"
            End With
            With aKeysDungeons(8)(5)
                .loc = "3904"
                .area = "BW2"
                .name = "Back Left Bombable Chest"
                .logic = "x"
            End With
            With aKeysDungeons(8)(6)
                .loc = "601"
                .area = "BW1"
                .name = "Coffin Freestanding Key"
                .logic = "a.f"
            End With
            With aKeysDungeons(8)(7)
                .loc = "3903"
                .area = "BW1"
                .name = "Lens of Truth Chest"
                .logic = "hLL7712LL7416.hLL7712GA8a"
            End With
            With aKeysDungeons(8)(8)
                .loc = "3920"
                .area = "BW2"
                .name = "Invisible Chest"
                .logic = "hLL7712"
            End With
            With aKeysDungeons(8)(9)
                .loc = "3916"
                .area = "BW1"
                .name = "Underwater Front Chest"
                .logic = "hLL7712"
            End With
            With aKeysDungeons(8)(10)
                .loc = "3909"
                .area = "BW1"
                .name = "Underwater Left Chest"
                .logic = "hLL7712"
            End With
            With aKeysDungeons(8)(11)
                .loc = "3907"
                .area = "BW1"
                .name = "Map Chest"
                .logic = "x.Q802V01.fV01"
            End With
            With aKeysDungeons(8)(12)
                .loc = "3910"
                .area = "BW2"
                .name = "Fire Keese Chest"
                .logic = "GCF"
            End With
            With aKeysDungeons(8)(13)
                .loc = "3912"
                .area = "BW2"
                .name = "Like Like Chest"
                .logic = "GCF"
            End With
            With aKeysDungeons(8)(14)
                .loc = "8000"
                .area = "BW2"
                .name = "Like Like Cage"
                .gs = True
                .logic = "GCFo"
            End With
            With aKeysDungeons(8)(15)
                .loc = "8002"
                .area = "BW2"
                .name = "West Inner Room"
                .gs = True
                .logic = "GD0o"
            End With
            With aKeysDungeons(8)(16)
                .loc = "8001"
                .area = "BW2"
                .name = "East Inner Room"
                .gs = True
                .logic = "GD1o"
            End With
            With aKeysDungeons(8)(17)
                .loc = "11127"
                .area = "EVENT"
                .name = "Door Fire Keese and Like Like Cage"
            End With
            With aKeysDungeons(8)(18)
                .loc = "11128"
                .area = "EVENT"
                .name = "Door West GS"
            End With
            With aKeysDungeons(8)(19)
                .loc = "11129"
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
                .name = "Map Chest"
            End With
            With aKeysDungeons(8)(1)
                .loc = "601"
                .area = "BW1"
                .name = "East Inner Room Freestanding Key"
            End With
            With aKeysDungeons(8)(2)
                .loc = "3902"
                .area = "BW0"
                .name = "Compass Chest"
                .logic = "LL7416.GA8a"
            End With
            With aKeysDungeons(8)(3)
                .loc = "602"
                .area = "BW0"
                .name = "Dead Hand Freestanding Key"
                .logic = "x"
            End With
            With aKeysDungeons(8)(4)
                .loc = "3901"
                .area = "BW1"
                .name = "Lens of Truth Chest"
                .logic = "GD3x"
            End With
            With aKeysDungeons(8)(5)
                .loc = "8000"
                .area = "BW0"
                .name = "Basement"
                .gs = True
                .logic = "J"
            End With
            With aKeysDungeons(8)(6)
                .loc = "8001"
                .area = "BW1"
                .name = "West Inner Room"
                .gs = True
                .logic = "x"
            End With
            With aKeysDungeons(8)(7)
                .loc = "8002"
                .area = "BW0"
                .name = "Coffin Room"
                .gs = True
                .logic = "JGD1"
            End With
            With aKeysDungeons(8)(8)
                .loc = "11120"
                .area = "EVENT"
                .name = "Door Coffin Room"
            End With
            With aKeysDungeons(8)(9)
                .loc = "11129"
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
                .name = "Map Chest"
                .logic = "u"
            End With
            With aKeysDungeons(9)(1)
                .loc = "4001"
                .area = "IC"
                .name = "Compass Chest"
                .logic = "u"
            End With
            With aKeysDungeons(9)(2)
                .loc = "701"
                .area = "IC"
                .name = "Piece of Heart"
                .logic = "u"
            End With
            With aKeysDungeons(9)(3)
                .loc = "4002"
                .area = "IC"
                .name = "Iron Boots Chest"
                .logic = "u"
            End With
            With aKeysDungeons(9)(4)
                .loc = "6402"
                .area = "IC"
                .name = "Song from Sheik"
                .logic = "u"
            End With
            With aKeysDungeons(9)(5)
                .loc = "8009"
                .area = "IC"
                .name = "Spinning Scythe Room"
                .gs = True
                .logic = "k"
            End With
            With aKeysDungeons(9)(6)
                .loc = "8010"
                .area = "IC"
                .name = "Heart Piece Room"
                .gs = True
                .logic = "uk"
            End With
            With aKeysDungeons(9)(7)
                .loc = "8008"
                .area = "IC"
                .name = "Push Block Room"
                .gs = True
                .logic = "uk"
            End With
        Else
            ' The Master Quest keys for the Ice Carvern
            With aKeysDungeons(9)(0)
                .loc = "4001"
                .area = "IC"
                .name = "Map Chest"
                .logic = "u"
            End With
            With aKeysDungeons(9)(1)
                .loc = "4000"
                .area = "IC"
                .name = "Compass Chest"
                .logic = "u"
            End With
            With aKeysDungeons(9)(2)
                .loc = "701"
                .area = "IC"
                .name = "Piece of Heart"
                .logic = "ux"
            End With
            With aKeysDungeons(9)(3)
                .loc = "4002"
                .area = "IC"
                .name = "Iron Boots Chest"
                .logic = "u"
            End With
            With aKeysDungeons(9)(4)
                .loc = "6402"
                .area = "IC"
                .name = "Song from Sheik"
                .logic = "u"
            End With
            With aKeysDungeons(9)(5)
                .loc = "8009"
                .area = "IC"
                .name = "Red Ice"
                .gs = True
                .logic = "uhLL7716"
            End With
            With aKeysDungeons(9)(6)
                .loc = "8010"
                .area = "IC"
                .name = "Ice Block"
                .gs = True
                .logic = "u"
            End With
            With aKeysDungeons(9)(7)
                .loc = "8008"
                .area = "IC"
                .name = "Scarecrow"
                .gs = True
                .logic = "uhLL6512k.uLL7730l"
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
                .name = "Lobby Left Chest"
                .logic = "d"
            End With
            With aKeysDungeons(10)(1)
                .loc = "4207"
                .area = "GTG0"
                .name = "Lobby Right Chest"
                .logic = "d"
            End With
            With aKeysDungeons(10)(2)
                .loc = "4200"
                .area = "GTG0"
                .name = "Stalfos Chest"
            End With
            With aKeysDungeons(10)(3)
                .loc = "4201"
                .area = "GTG0"
                .name = "Beamos Chest"
                .logic = "x"
            End With
            With aKeysDungeons(10)(4)
                .loc = "4217"
                .area = "GTG6"
                .name = "Before Heavy Block Chest"
            End With
            With aKeysDungeons(10)(5)
                .loc = "4215"
                .area = "GTG7"
                .name = "Heavy Block First Chest"
            End With
            With aKeysDungeons(10)(6)
                .loc = "4214"
                .area = "GTG7"
                .name = "Heavy Block Second Chest"
            End With
            With aKeysDungeons(10)(7)
                .loc = "4220"
                .area = "GTG7"
                .name = "Heaby Block Third Chest"
            End With
            With aKeysDungeons(10)(8)
                .loc = "4202"
                .area = "GTG7"
                .name = "Heavy Block Fourth Chest"
            End With
            With aKeysDungeons(10)(9)
                .loc = "4203"
                .area = "GTG4"
                .name = "Eye Statue Chest"
                .logic = "d"
            End With
            With aKeysDungeons(10)(10)
                .loc = "4204"
                .area = "GTG5"
                .name = "Near Scarecrow Chest"
                .logic = "d"
            End With
            With aKeysDungeons(10)(11)
                .loc = "4218"
                .area = "GTG3"
                .name = "Hammer Room Clear Chest"
            End With
            With aKeysDungeons(10)(12)
                .loc = "4216"
                .area = "GTG3"
                .name = "Hammer Room Switch Chest"
                .logic = "r"
            End With
            With aKeysDungeons(10)(13)
                .loc = "4205"
                .area = "GTG1"
                .name = "Maze Right Central Chest"
            End With
            With aKeysDungeons(10)(14)
                .loc = "4208"
                .area = "GTG1"
                .name = "Maze Right Side Chest"
            End With
            With aKeysDungeons(10)(15)
                .loc = "801"
                .area = "GTG1"
                .name = "Lava Room Freestanding Key"
            End With
            With aKeysDungeons(10)(16)
                .loc = "4213"
                .area = "GTG2"
                .name = "Underwater Silver Rupee Chest"
                .logic = "khLL7716LL7429G0A"
            End With
            With aKeysDungeons(10)(17)
                .loc = "4211"
                .area = "GTG0"
                .name = "Hidden Ceiling Chest"
                .logic = "GD4T"
            End With
            With aKeysDungeons(10)(18)
                .loc = "4206"
                .area = "GTG0"
                .name = "Maze Path First Chest"
                .logic = "GD5"
            End With
            With aKeysDungeons(10)(19)
                .loc = "4210"
                .area = "GTG0"
                .name = "Maze Path Second Chest"
                .logic = "GD6"
            End With
            With aKeysDungeons(10)(20)
                .loc = "4209"
                .area = "GTG0"
                .name = "Maze Path Third Chest"
                .logic = "GD7"
            End With
            With aKeysDungeons(10)(21)
                .loc = "4212"
                .area = "GTG0"
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
                .name = "Lobby Left Chest"
            End With
            With aKeysDungeons(10)(1)
                .loc = "4207"
                .area = "GTG0"
                .name = "Lobby Right Chest"
            End With
            With aKeysDungeons(10)(2)
                .loc = "4211"
                .area = "GTG0"
                .name = "Hidden Ceiling Chest"
                .logic = "T"
            End With
            With aKeysDungeons(10)(3)
                .loc = "4206"
                .area = "GTG0"
                .name = "Maze Path First Chest"
            End With
            With aKeysDungeons(10)(4)
                .loc = "4210"
                .area = "GTG0"
                .name = "Maze Oath Second Chest"
            End With
            With aKeysDungeons(10)(5)
                .loc = "4209"
                .area = "GTG0"
                .name = "Maze Path Third Chest"
                .logic = "GD9"
            End With
            With aKeysDungeons(10)(6)
                .loc = "4201"
                .area = "GTG1"
                .name = "Dinolfos Chest"
            End With
            With aKeysDungeons(10)(7)
                .loc = "4213"
                .area = "GTG2"
                .name = "Underwater Silver Rupee Chest"
            End With
            With aKeysDungeons(10)(8)
                .loc = "4200"
                .area = "GTG3"
                .name = "First Iron Knuckle Chest"
            End With
            With aKeysDungeons(10)(9)
                .loc = "4217"
                .area = "GTG4"
                .name = "Before Heavy Block Chest"
            End With
            With aKeysDungeons(10)(10)
                .loc = "4202"
                .area = "GTG4"
                .name = "Heavy Block Chest"
                .logic = "V02"
            End With
            With aKeysDungeons(10)(11)
                .loc = "4203"
                .area = "GTG5"
                .name = "Eye Statue Chest"
                .logic = "d"
            End With
            With aKeysDungeons(10)(12)
                .loc = "4218"
                .area = "GTG5"
                .name = "Second Iron Knuckle Chest"
            End With
            With aKeysDungeons(10)(13)
                .loc = "4214"
                .area = "GTG5"
                .name = "Flame Circle Chest"
                .logic = "k.d.x"
            End With
            With aKeysDungeons(10)(14)
                .loc = "4205"
                .area = "GTG6"
                .name = "Maze Right Central Chest"
            End With
            With aKeysDungeons(10)(15)
                .loc = "4208"
                .area = "GTG6"
                .name = "Maze Right Side Chest"
            End With
            With aKeysDungeons(10)(16)
                .loc = "4204"
                .area = "GTG6"
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
                .name = "Forest Trial Chest"
            End With
            With aKeysDungeons(11)(1)
                .loc = "4307"
                .area = "IGC0"
                .name = "Water Trial Left Chest"
            End With
            With aKeysDungeons(11)(2)
                .loc = "4306"
                .area = "IGC0"
                .name = "Water Trial Right Chest"
            End With
            With aKeysDungeons(11)(3)
                .loc = "4308"
                .area = "IGC0"
                .name = "Shadow Trial Front Chest"
                .logic = "de.h.LL7430.hLL7716"
            End With
            With aKeysDungeons(11)(4)
                .loc = "4305"
                .area = "IGC0"
                .name = "Shadow Trial Golden Gauntlets Chest"
                .logic = "de.lLL7439.lf"
            End With
            With aKeysDungeons(11)(5)
                .loc = "4318"
                .area = "IGC0"
                .name = "Spirit Trial Crystal Switch Chest"
                .logic = "k"
            End With
            With aKeysDungeons(11)(6)
                .loc = "4320"
                .area = "IGC0"
                .name = "Spirit Trial Invisible Chest"
                .logic = "jkT"
            End With
            With aKeysDungeons(11)(7)
                .loc = "4312"
                .area = "IGC1"
                .name = "Light Trial First Left Chest"
            End With
            With aKeysDungeons(11)(8)
                .loc = "4311"
                .area = "IGC1"
                .name = "Light Trial Second Left Chest"
            End With
            With aKeysDungeons(11)(9)
                .loc = "4313"
                .area = "IGC1"
                .name = "Light Trial Third Left Chest"
            End With
            With aKeysDungeons(11)(10)
                .loc = "4314"
                .area = "IGC1"
                .name = "Light Trial First Right Chest"
            End With
            With aKeysDungeons(11)(11)
                .loc = "4310"
                .area = "IGC1"
                .name = "Light Trial Second Right Chest"
            End With
            With aKeysDungeons(11)(12)
                .loc = "4315"
                .area = "IGC1"
                .name = "Light Trial Third Right Chest"
            End With
            With aKeysDungeons(11)(13)
                .loc = "4316"
                .area = "IGC1"
                .name = "Light Trial Invisible Enemies Chest"
                .logic = "T"
            End With
            With aKeysDungeons(11)(14)
                .loc = "4317"
                .area = "IGC1"
                .name = "Light Trail Lullaby Chest"
                .logic = "GDBhLL7712"
            End With
            With aKeysDungeons(11)(15)
                .loc = "4111"
                .area = "IGC3"
                .name = "Ganon’s Tower Boss Key Chest"
            End With
            With aKeysDungeons(11)(16)
                .loc = "8709"
                .area = "IGC1"
                .name = "Deku Room Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(17)
                .loc = "8706"
                .area = "IGC1"
                .name = "Deku Room Centre Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(18)
                .loc = "8704"
                .area = "IGC1"
                .name = "Deku Room Centre Right"
                .scrub = True
            End With
            With aKeysDungeons(11)(19)
                .loc = "8708"
                .area = "IGC1"
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
                .name = "Forest Trial Eye Switch Chest"
                .logic = "d"
            End With
            With aKeysDungeons(11)(1)
                .loc = "4303"
                .area = "IGC0"
                .name = "Forest Trial Frozen Eye Switch Chest"
                .logic = "P"
            End With
            With aKeysDungeons(11)(2)
                .loc = "901"
                .area = "IGC0"
                .name = "Forest Trial Stalfos Room Freestanding Key"
                .logic = "k"
            End With
            With aKeysDungeons(11)(3)
                .loc = "4301"
                .area = "IGC0"
                .name = "Water Trial Chest"
                .logic = "u"
            End With
            With aKeysDungeons(11)(4)
                .loc = "4300"
                .area = "IGC0"
                .name = "Shadow Trial Bomb Flower Chest"
                .logic = "dk.dLL7430.LL7430Tx.LL7430TV01.LL7430Tf"
            End With
            With aKeysDungeons(11)(5)
                .loc = "4305"
                .area = "IGC0"
                .name = "Shadow Trial Eye Switch Chest"
                .logic = "dTLL7430.dTkP"
            End With
            With aKeysDungeons(11)(6)
                .loc = "4310"
                .area = "IGC0"
                .name = "Spirit Trial First Chest"
                .logic = "dr.GA3r"
            End With
            With aKeysDungeons(11)(7)
                .loc = "4320"
                .area = "IGC0"
                .name = "Spirit Trial Invisible Chest"
                .logic = "drjT.GA3rjT"
            End With
            With aKeysDungeons(11)(8)
                .loc = "4309"
                .area = "IGC0"
                .name = "Spirit Trial Sun Front Left Chest"
                .logic = "dejrLL7422"
            End With
            With aKeysDungeons(11)(9)
                .loc = "4308"
                .area = "IGC0"
                .name = "Spirit Trial Sun Back Left Chest"
                .logic = "dejrLL7422"
            End With
            With aKeysDungeons(11)(10)
                .loc = "4307"
                .area = "IGC0"
                .name = "Spirit Trial Sun Back Right Chest"
                .logic = "dejrLL7422"
            End With
            With aKeysDungeons(11)(11)
                .loc = "4306"
                .area = "IGC0"
                .name = "Spirit Trial Golden Gauntlets Chest"
                .logic = "dejrLL7422"
            End With
            With aKeysDungeons(11)(12)
                .loc = "4304"
                .area = "IGC2"
                .name = "Light Trial Lullaby Chest"
                .logic = "hLL7712"
            End With
            With aKeysDungeons(11)(13)
                .loc = "4111"
                .area = "IGC3"
                .name = "Ganon’s Tower Boss Key Chest"
            End With
            With aKeysDungeons(11)(14)
                .loc = "8709"
                .area = "IGC1"
                .name = "Deku Room Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(15)
                .loc = "8706"
                .area = "IGC1"
                .name = "Deku Room Centre Left"
                .scrub = True
            End With
            With aKeysDungeons(11)(16)
                .loc = "8704"
                .area = "IGC1"
                .name = "Deku Room Centre Right"
                .scrub = True
            End With
            With aKeysDungeons(11)(17)
                .loc = "8708"
                .area = "IGC1"
                .name = "Deku Room Right"
                .scrub = True
            End With
            With aKeysDungeons(11)(18)
                .loc = "8701"
                .area = "IGC1"
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
    Private Sub inc(ByRef value As Integer, Optional ByVal amount As Byte = 1)
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
        pbxKokiriSword.Visible = cb7416.Checked
        pbxMasterSword.Visible = cb7417.Checked
        pbxDekuShield.Visible = cb7420.Checked
        pbxHylianShield.Visible = cb7421.Checked
        pbxMirrorShield.Visible = cb7422.Checked
        pbxKokiriTunic.Visible = cb7424.Checked
        pbxGoronTunic.Visible = cb7425.Checked
        pbxZoraTunic.Visible = cb7426.Checked
        pbxKokiriBoots.Visible = cb7428.Checked
        pbxIronBoots.Visible = cb7429.Checked
        pbxHoverBoots.Visible = cb7430.Checked

        ' Run the check to see if you have the Biggoron's Sword or the Knife
        biggoronSwordCheck()
    End Sub
    Private Sub biggoronSwordCheck()
        '0 = none, 1 = Biggoron's Sword, 2 = Broken Knife
        Dim who As Byte = 0
        If cb7418.Checked Then
            who = 1
            If cb7508.Checked Then
                who = 1
            ElseIf cb7419.Checked Then
                who = 2
            End If
        Else
            If cb7419.Checked Then
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
        If cb7600.Checked Then upgrades(0) = upgrades(0) + 1
        If cb7601.Checked Then upgrades(0) = upgrades(0) + 2

        ' Bomb Bags
        If cb7603.Checked Then upgrades(1) = upgrades(1) + 1
        If cb7604.Checked Then upgrades(1) = upgrades(1) + 2

        ' Gauntlets
        If cb7606.Checked Then upgrades(2) = upgrades(2) + 1
        If cb7607.Checked Then upgrades(2) = upgrades(2) + 2

        ' Scales
        If cb7609.Checked Then upgrades(3) = upgrades(3) + 1
        If cb7610.Checked Then upgrades(3) = upgrades(3) + 2

        ' Wallets
        If cb7612.Checked Then upgrades(4) = upgrades(4) + 1
        If cb7613.Checked Then upgrades(4) = upgrades(4) + 2

        ' Bullet Bag
        If cb7614.Checked Then upgrades(5) = upgrades(5) + 1
        If cb7615.Checked Then upgrades(5) = upgrades(5) + 2

        ' Sticks
        If cb7617.Checked Then upgrades(6) = upgrades(6) + 1
        If cb7618.Checked Then upgrades(6) = upgrades(6) + 2

        ' Nuts
        If cb7620.Checked Then upgrades(7) = upgrades(7) + 1
        If cb7621.Checked Then upgrades(7) = upgrades(7) + 2

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
        For i = 0 To 22
            For Each key In aKeys
                With key
                    If .loc = "77" & IIf(i > 9, "", "0").ToString & i.ToString Then aQIChecks(i) = .checked
                End With
            Next
        Next

        For i As Byte = 0 To 22
            With aoQuestItems(i)
                If aQIChecks(i) Then
                    .Image = aoQuestItemImages(i)
                Else
                    .Image = aoQuestItemImagesEmpty(i)
                End If
            End With

            ' Store for use when drawing text over each quest reward
            aQuestRewardsCollected(i) = aQIChecks(i)

            ' Update the added dungeon text on each medallion and stone
            Select Case i
                Case 0 To 5, 18 To 20
                    changeDungeonRewardText(i, , True)
            End Select
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


        Select Case attempt
            Case 0
                ' Builds October 27, 2021 to March 31, 2022
                attemptOffset = &H29C95D8
                attemptAdded = 2147483648
            Case 1
                ' Builds July 13, 2021 to October 11, 2021
                attemptOffset = &HCA6B8
            Case Else
                Return
        End Select

        ' Step through all modules to find mupen64plus.dll's base address
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next
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
                If ReadMemory(Of Integer)(romAddrStart64 + &H11A5D0) = 529 Then nextAttempt = False
            End If
        End If
        If nextAttempt Then attachToM64P(CByte(attempt + 1))
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
            romAddrStart64 = addressDLL + &H845000
            SetProcessName("retroarch")
            emulator = "retroarch - parallel"
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'updateItems()
        'Exit Sub

        'attachToBizHawk()
        If emulator = "variousX64" Then
            'WriteMemory(Of Integer)(romAddrStart64 + &H11B4AC, &H10203)
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
        ElseIf emulator = "project64" And IS_64BIT = False Then
            quickWrite(romAddrStart + &H11A644, &H10203)
            quickWrite(romAddrStart + &H11A648, &H4050608)
            quickWrite(romAddrStart + &H11A64C, &H90B0C0D)
            quickWrite(romAddrStart + &H11A650, &HE0F1011)
            quickWrite(romAddrStart + &H11A654, &H12131818)
            quickWrite(romAddrStart + &H11A658, &H1818372B)
            quickWrite(romAddrStart + &H11A65C, &H1E28281B)
            quickWrite(romAddrStart + &H11A660, &H2B00)
            quickWrite(romAddrStart + &H11A664, &H2B000000)
            quickWrite(romAddrStart + &H11A668, &HA0A)
            quickWrite(romAddrStart + &H11A66C, &H77770000)
            quickWrite(romAddrStart + &H11A670, &H36E4DB)
            quickWrite(romAddrStart + &H11A674, &H7FFFFF)
        End If
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
                Case 0 To 5, 18 To 20
                    changeDungeonRewardText(i, , True)
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
        Dim temp As String = String.Empty
        Dim cTemp As Integer = 0

        For i = &H11A644 To &H11A658 Step 4
            temp = Hex(goRead(i))
            While temp.Length < 8
                temp = "0" & temp
            End While
            items = items & temp
        Next

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
                        allItems = allItems & "a"
                    Case 1
                        .Image = My.Resources.dekuNut
                        allItems = allItems & "b"
                    Case 2
                        .Image = My.Resources.bombs
                        allItems = allItems & "c"
                    Case 3
                        .Image = My.Resources.bow
                        allItems = allItems & "d"
                    Case 4
                        .Image = My.Resources.fireArrow
                        allItems = allItems & "e"
                    Case 5
                        .Image = My.Resources.dinsFire
                        allItems = allItems & "f"
                    Case 6
                        .Image = My.Resources.fairySlingshot
                        allItems = allItems & "g"
                    Case 7
                        .Image = My.Resources.fairyOcarina
                        allItems = allItems & "h"
                    Case 8
                        .Image = My.Resources.ocarinaOfTime
                        allItems = allItems & "hi"
                    Case 9
                        .Image = My.Resources.bombchus
                        allItems = allItems & "j"
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
                        .Image = My.Resources.naryusLove
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
        getItemAmounts()
    End Sub
    Private Sub updateDungeonItems()
        If isLoadedGame() = False Then Exit Sub
        getSmallKeys()
        getDungeonItems()
        getHearts()
        getMagic()
        getGoldSkulltulas()

        ' This is for detecting if it is a triforce hunt, which people did not want. Oddly enough it is also the only solid check for OOTR vs AP, might use it later.
        'With pbxTriforce
        'If isTriforceHunt() Then
        ' Make sure the triforce is visible
        'If Not .Visible Then .Visible = True
        getTriforce()
        'Else
        ' If not a triforce hunt, hide the triforce picturebox
        'If .Visible Then .Visible = False
        'End If
        'End With
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
            If flipKeyForced(Trim(readLine), areaCode) Then
                Dim newLine As String = readLine
                updateLabels()
                updateLabelsDungeons()
                btnFocus.Focus()
                If newLine.Contains("(Forced)") Then
                    newLine = newLine.Replace(" (Forced)", "")
                Else
                    newLine = newLine & " (Forced)"
                End If
                replaceRTB(rtb, linePos, newLine)
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
                    If .Lines(i).Contains(toBold) Then
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
    Private Sub emboldenOld(ByVal rtb As RichTextBox, ByVal toBold As String)
        ' Trim up the text to bold
        toBold = Trim(toBold.Replace(vbCrLf, ""))
        Dim findStart As Integer = -1

        ' Start at the end of the lines and check for all that should be bolded
        For i = rtb.Lines.Count - 1 To 0 Step -1
            If rtb.Lines(i).Contains(toBold) Then
                findStart = i
                Exit For
            End If
        Next

        ' Exit sub if no index was found
        If findStart = -1 Then Exit Sub

        ' Set up the font to be the same, select the line, and replace the font with just a bolded version
        Dim richFont As System.Drawing.Font = rtb.Font
        rtb.Select(rtb.GetFirstCharIndexFromLine(findStart), rtb.Lines(findStart).Length - 1)
        rtb.SelectionFont = New Font(richFont.FontFamily, richFont.Size, FontStyle.Bold)
    End Sub

    Private Sub outputSong(ByVal title As String, ByVal notes As String)
        rtbOutputLeft.Text = title & ":"

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
        changeDungeonRewardText(18, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneKokiri_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxStoneKokiri.MouseDoubleClick
        changeDungeonRewardText(18, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneGoron_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxStoneGoron.MouseClick
        changeDungeonRewardText(19, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneGoron_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxStoneGoron.MouseDoubleClick
        changeDungeonRewardText(19, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneZora_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxStoneZora.MouseClick
        changeDungeonRewardText(20, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxStoneZora_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxStoneZora.MouseDoubleClick
        changeDungeonRewardText(20, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalForest_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalForest.MouseClick
        changeDungeonRewardText(0, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalForest_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalForest.MouseDoubleClick
        changeDungeonRewardText(0, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalFire_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalFire.MouseClick
        changeDungeonRewardText(1, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalFire_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalFire.MouseDoubleClick
        changeDungeonRewardText(1, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalWater_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalWater.MouseClick
        changeDungeonRewardText(2, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalWater_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalWater.MouseDoubleClick
        changeDungeonRewardText(2, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalSpirit_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalSpirit.MouseClick
        changeDungeonRewardText(3, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalSpirit_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalSpirit.MouseDoubleClick
        changeDungeonRewardText(3, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalShadow_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalShadow.MouseClick
        changeDungeonRewardText(4, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalShadow_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalShadow.MouseDoubleClick
        changeDungeonRewardText(4, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalLight_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles pbxMedalLight.MouseDoubleClick
        changeDungeonRewardText(5, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pbxMedalLight_MouseClick(sender As Object, e As MouseEventArgs) Handles pbxMedalLight.MouseClick
        changeDungeonRewardText(5, CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub

    Private Sub frmTrackerOfTime_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        btnFocus.Focus()
    End Sub
    Private Sub frmTrackerOfTime_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Dim pnFore As Pen = New Pen(Me.ForeColor, 1)
        If showSetting Then e.Graphics.DrawLine(pnFore, pnlDekuTree.Location.X + pnlDekuTree.Width + 6, 0, pnlDekuTree.Location.X + pnlDekuTree.Width + 6, Me.Height)
        With rtbOutputLeft
            e.Graphics.DrawRectangle(pnFore, .Location.X - 1, .Location.Y - 1, .Width + 1, .Height + 1)
        End With
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
        If lcx.Text = lblLogicWarning.Text Then Exit Sub
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
            Case lcxRustedSwitches.Text
                My.Settings.setRustedSwitches = Not My.Settings.setRustedSwitches
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
            Case lcxWaTDrgnDive.Text
                My.Settings.setWaTDragonDive = Not My.Settings.setWaTDragonDive
            Case lcxWaTBKR.Text
                My.Settings.setWaTBKR = Not My.Settings.setWaTBKR
            Case lcxWaTCrackNothing.Text
                My.Settings.setWaTCrackNothing = Not My.Settings.setWaTCrackNothing
            Case lcxWaTCrackHovers.Text
                My.Settings.setWaTCrackHovers = Not My.Settings.setWaTCrackHovers
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
            Case lcxShTUmbrella.Text
                My.Settings.setShTUmbrella = Not My.Settings.setShTUmbrella
            Case lcxShTMQGap.Text
                My.Settings.setShTMQGap = Not My.Settings.setShTMQGap
            Case lcxShTStatue.Text
                My.Settings.setShTStatue = Not My.Settings.setShTStatue
            Case lcxBotWLensless.Text
                My.Settings.setBotWLensless = Not My.Settings.setBotWLensless
            Case lcxBotWDeadHand.Text
                My.Settings.setBotWDeadHand = Not My.Settings.setBotWDeadHand
            Case lcxNavi.Text
                My.Settings.setNavi = Not My.Settings.setNavi
                'Case lcxxx.Text
                'My.Settings.setxx = Not My.Settings.setxx
            Case Else
                rtbOutputLeft.Text = "Unhandled LCX: " & text
        End Select

        ' Save settings and update labels and graphics
        My.Settings.Save()
        updateLabels()
        updateLabelsDungeons()
        updateSettingsPanel()
    End Sub
    Private Sub infoLCX(ByVal text As String)
        ' Compare sent text with lcx text to figure out which one was right-clicked, output info for each
        Dim message As String = String.Empty
        Select Case text
            Case lcxLogic.Text
                message = "Attempts to bold any checks you can currently reach, and any area with checks. Minor glitches or tricks may get you more checks than detected."
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
            Case lcxRustedSwitches.Text
                message = "Hit rusted switched through walls in Fire Temple, Fire Temple MQ, and Spirit Trial MQ."
            Case lcxBombchuLogic.Text
                message = "Considers Bombchus as viable as bombs, also requires Bombchus for Bombchu Bowling."
            Case lcxLoTLogic.Text
                message = "Require Lens of Truth for Treasure Chest Game, Wasteland, and various dungeons."
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
                message = "Use bomb flowers to light eyes."
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
            Case lcxWaTDrgnDive.Text
                message = "When entering Dragon Statue room from the river, a well aimed dive can get you into the tunnel without need for Iron Boots or Scale. Works in both normal and MQ versions."
            Case lcxWaTBKR.Text
                message = "Use Hover Boots to reach Boss Key Region without the need for the Longshot."
            Case lcxWaTCrackNothing.Text
                message = "Reach the cracked wall without raising water level or Hover Boots."
            Case lcxWaTCrackHovers.Text
                message = "Reach the cracked wall with the Hover Boots, without raising the water level."
            Case lcxSpTWall.Text
                message = "Climb the shifting wall without dealing with the Beamos or the Walltula first."
            Case lcxSpTBombchu.Text
                message = "Use a Bombchu to hit the bridge switch."
            Case lcxSpTMQLowAdult.Text
                message = "Use Din's Fire and Bow to quickly light all three torches to access the Lower Adult area."
            Case lcxSpTMQSunRoom.Text
                message = "You can toss and smash the crate on the switch to briefly unlock the door to the Sun Block Room."
            Case lcxShTLensless.Text
                message = "Navigate the Shadow Temple without Lens of Truth."
            Case lcxShTUmbrella.Text
                message = "Use Hover Boots and a precise jump off of the lower chest to get atop the crushing spikes. Works in normal and MQ dungeon."
            Case lcxShTStatue.Text
                message = "Send a bombchu across the edge of the gorge to knock down the statue. Works in normal and MQ dungeon."
            Case lcxShTMQGap.Text
                message = "Longshot to a torch and then a jump slash off bounce yourself onto the tongue."
            Case lcxBotWLensless.Text
                message = "Navigate the Bottom of the Well without Lens of Truth."
            Case lcxBotWDeadHand.Text
                message = "Use Deku Sticks to defeat the Dead Hand instead of the Kokiri Sword. Applies to both normal and MQ dungeon."
            Case lcxNavi.Text
                message = "I know we all love hearing Navi, but this option will keep Navi from randomly wanting to yell for you. She will still shout at curious shiny rocks on the ground, but no more telling you to get a move on."
                'Case lcxxx.Text
                'message = "."
        End Select

        If Not message = String.Empty Then
            rtbOutputLeft.Text = text.Replace(":", "") & ":"
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
            Case ltbKeys.Name
                If My.Settings.setSmallKeys = 0 And lower Then
                    My.Settings.setSmallKeys = 2
                Else
                    My.Settings.setSmallKeys = CByte(My.Settings.setSmallKeys + addVal)
                    If My.Settings.setSmallKeys > 2 Then My.Settings.setSmallKeys = 0
                End If
        End Select
        updateLTB(sName)
    End Sub
    Private Sub updateLTB(ByVal ltbName As String)
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


    Private Sub updateSettingsPanel()
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
                            Case lcxRustedSwitches.Name
                                isTrue = My.Settings.setRustedSwitches
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
                            Case lcxWaTDrgnDive.Name
                                isTrue = My.Settings.setWaTDragonDive
                            Case lcxWaTBKR.Name
                                isTrue = My.Settings.setWaTBKR
                            Case lcxWaTCrackNothing.Name
                                isTrue = My.Settings.setWaTCrackNothing
                            Case lcxWaTCrackHovers.Name
                                isTrue = My.Settings.setWaTCrackHovers
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
                            Case lcxShTUmbrella.Name
                                isTrue = My.Settings.setShTUmbrella
                            Case lcxShTStatue.Name
                                isTrue = My.Settings.setShTStatue
                            Case lcxShTMQGap.Name
                                isTrue = My.Settings.setShTMQGap
                            Case lcxBotWLensless.Name
                                isTrue = My.Settings.setBotWLensless
                            Case lcxBotWDeadHand.Name
                                isTrue = My.Settings.setBotWDeadHand
                            Case lcxNavi.Name
                                isTrue = My.Settings.setNavi
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
    Private Function flipKeyForced(ByVal name As String, ByVal area As String) As Boolean
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
            For Each key In aKeys.Where(Function(k As keyCheck) k.name.Equals(name))
                With key
                    If .area = area Then
                        .forced = Not .forced
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
        stopScanning()
        pbxPoH.Image = My.Resources.poh0
        For Each chk In pnlHidden.Controls.OfType(Of CheckBox)()
            chk.Checked = False
        Next
        rtbOutputLeft.Clear()
        rtbOutputRight.Clear()
        lastRoomScan = 0
        populateLocations()
    End Sub
    Private Sub ShowSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowSettingsToolStripMenuItem.Click
        showSetting = Not showSetting
        updateShowSettings()
    End Sub

    Private Sub changeDungeonRewardText(ByVal thisOne As Byte, Optional reverse As Boolean = False, Optional justRefresh As Boolean = False)
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
                    tempValue = tempValue - 1
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
                    xPos = 6
                Case 3
                    xPos = 3
                Case 4
                    xPos = -1
            End Select

            ' Font for dungeon letters
            Dim fontDungeon = New Font("Lucida Console", 12, FontStyle.Regular, GraphicsUnit.Pixel)

            ' Draw the value over the lower right of the gold skulltula picturebox, first in black to give it some definition, then in white
            Graphics.FromImage(.Image).DrawString(outputText, fontDungeon, New SolidBrush(Color.Black), xPos - 1, 20)
            Graphics.FromImage(.Image).DrawString(outputText, fontDungeon, New SolidBrush(Color.White), xPos, 21)
        End With
    End Sub
    Private Sub getPlayerName()
        ' Grab the first 4 characters
        Dim sPlayerName1 As String = Hex(goRead(&H11A5F4))
        ' Grab the last 4 characters
        Dim sPlayerName2 As String = Hex(goRead(&H11A5F8))

        ' Make sure to fill the 0's
        While sPlayerName1.Length < 8
            sPlayerName1 = "0" & sPlayerName1
        End While

        ' Make sure to fill the 0's
        While sPlayerName2.Length < 8
            sPlayerName2 = "0" & sPlayerName2
        End While

        ' Combine into one
        sPlayerName1 = sPlayerName1 & sPlayerName2

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
            Dim fontName As Font = New Font("Lucida Console", 26, FontStyle.Bold, GraphicsUnit.Pixel)

            ' Draw the value over the lower right of the gold skulltula picturebox, first in black to give it some definition, then in white
            Graphics.FromImage(.BackgroundImage).DrawString(sPlayerName2.TrimEnd, fontName, New SolidBrush(Color.White), -3, 90)
        End With


    End Sub
    Private Function decodeLetter(ByRef valLetter As Byte) As String
        decodeLetter = String.Empty
        Select Case valLetter
            Case 0 To 9
                ' 0 to 9, 48 to 57, so +48, but instead just turn it a string
                decodeLetter = valLetter.ToString
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
    Private Sub getRandoVer()
        ' Checks for different rom generators
        Dim isOOTR As Boolean = False
        Dim isAP As Boolean = False

        ' Both use the pattern of "0032####" for thieir requirement check location
        If goRead(&H40B668 + 2, 15) = 50 Then isOOTR = True

        ' If not OOTR, check for AP
        If Not isOOTR Then
            If goRead(&H40B1B0 + 2, 15) = 50 Then isAP = True
        End If

        randoVer = String.Empty
        If isOOTR And Not isAP Then randoVer = "OOTR"
        If isAP And Not isOOTR Then randoVer = "AP"
        If Not isOOTR And Not isAP Then randoVer = "NONE"
    End Sub
    Private Sub getRainbowBridge()
        ' Determine what is needed for Rainbow Bridge
        Dim startAddress As Integer = &H400CB0
        If randoVer = "AP" Then
            ' AP is offset by 16 bytes
            inc(startAddress, 16)
        ElseIf Not randoVer = "OOTR" Then
            ' If not AP or OOTR, abort this check
            Exit Sub
        End If

        ' Rainbow Bridge Condition
        rainbowBridge(0) = CByte(goRead(startAddress, 1))
        ' Rainbow Bridge Condition Count
        rainbowBridge(1) = CByte(goRead(startAddress + 18, 1))
    End Sub

    Private Sub pnlSettings_Paint(sender As Object, e As PaintEventArgs) Handles pnlSettings.Paint
        updateSettingsPanel()
    End Sub
    Private Sub rtbAddLine(ByVal line As String, Optional hideRight As Boolean = False)
        rtbOutputRight.Visible = Not hideRight
        If rtbOutputLeft.Lines.Count < 14 Then
            'rtbOutputLeft.AppendText(IIf(rtbOutputLeft.TextLength > 0, vbCrLf, "").ToString & line)
            rtbOutputLeft.AppendText(vbCrLf & line)
        Else
            rtbOutputRight.AppendText(vbCrLf & line)
        End If
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
                Case "scan", "auto scan", "stop", "reset", "themes", "settings >", "settings <"
                    e.Graphics.DrawRectangle(New Pen(foreColour, 1), New Rectangle(0, 0, .Width - 1, .Height - 1))
            End Select
            .ForeColor = foreColour
            .Font = font
        End With
    End Sub
End Class

