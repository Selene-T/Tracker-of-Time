Option Explicit On
Option Strict On

Public Class frmTrackerOfTime
    Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Public Declare Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Integer, ByVal lpBaseAddress As Integer, ByRef lpBuffer As Integer, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer

    Const PROCESS_ALL_ACCESS = &H1F0FFF
    Const CHECK_COUNT = 103
    Const IS_64BIT = False
    Const DO_COUNT = False
    'Const ONLY_GS = False

    Private doNotScroll As Boolean = False
    Private keepRunning As Boolean = False
    Private arrLocation(CHECK_COUNT) As Integer
    Private arrChests(CHECK_COUNT) As Integer
    Private arrHigh(CHECK_COUNT) As Byte
    Private arrLow(CHECK_COUNT) As Byte
    Private firstRun As Boolean = True
    Private firstDraw As Boolean = True
    Private cBlend As New Color
    Private aKeys(311) As keyCheck
    Private aKeysDungeons(11)() As keyCheck

    Private Const romAddrStart As Integer = &HDFE40000
    Private romAddrStart64 As Int64 = 0
    Private emulator As String = String.Empty

    Private Const currentRoomAddr As Integer = &H1C8544
    Private lastRoomScan As Integer = 0
    Private royalSongScan As Integer = 0

    Private aMQ(11) As Boolean
    Private aMQOld(11) As Boolean

    ' Attempts to make sure graphics are drawn upon keypresses
    Private Sub frmTrackerOfTime_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        'drawGraphics()
    End Sub
    Private Sub frmTrackerOfTime_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        'drawGraphics()
    End Sub
    Private Sub frmTrackerOfTime_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        'drawGraphics()
    End Sub

    ' On load, populate the locations array
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        custMenu.highlight = Me.ForeColor
        custMenu.backColour = Me.BackColor
        custMenu.foreColour = Me.ForeColor
        mnuOptions.Renderer = New custMenu

        For i As Integer = 0 To arrLocation.Length - 1
            arrChests(i) = 0
            arrHigh(i) = 0
            arrLow(i) = 31
        Next

        'emulator = "project64"
        compressGUI()

        If DO_COUNT Then ReDim aKeys(500)

        For i = 0 To aKeys.Length - 1
            aKeys(i) = New keyCheck
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
        For Each checkBox In pnlHidden.Controls.OfType(Of CheckBox)().Where(Function(cb As CheckBox) cb.Name.Contains("cb77"))
            AddHandler checkBox.CheckedChanged, AddressOf cccQuestItems
        Next

        ' Unset all MQs
        For i = 0 To aMQ.Length - 1
            aMQ(i) = False
            aMQOld(i) = True
        Next
        updateMQs()
        populateLocations()
        loadSettings()
    End Sub

    Private Sub loadSettings()
        'ddThemes.SelectedIndex = My.Settings.setTheme
        changeTheme(My.Settings.setTheme)
        subMenuCheck(My.Settings.setTheme)
    End Sub

    ' Populate location's addresses and clear up the chest data
    Private Sub populateLocations()
        clearItems()
        setupKeys()
        updateLabels()
        updateLabelsDungeons()
        If firstRun Then
            getHighLows()
            firstRun = False
        End If
        arrLocation(0) = &H11AD18 + 4       ' DMC/DMT/OGC Great Fairy Fountain
        arrLocation(1) = &H11AF80 + 4       ' Hyrule Field (Events) Big Poes Captured and Ocarina of Time
        arrLocation(2) = &H11B028 + 4       ' Lake Hylia (Underwater) 
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
        arrLocation(62) = &H11B4A4          ' *Items: Egg from Malon, Obtained Epona
        arrLocation(63) = &H11B4A8          ' *Item: Zora Diving Game, Darunia’s Joy
        arrLocation(64) = &H11B4AC          ' *Events 3: Zelda’s Letter, Song from Impa, Sun Song??
        arrLocation(65) = &H11B4B4          ' *Items: Scarecrow as Adult
        arrLocation(66) = &H11B4B8          ' 32b-25 *Events 6: Song at Colossus
        arrLocation(67) = &H11B4BC          ' *Events 7: Saria Gift, Skulltula trades
        arrLocation(68) = &H11B4C0          ' *Item Collect #1
        arrLocation(69) = &H11B4C4          ' *Item Collection #2
        arrLocation(70) = &H11B4E8          ' *Item: Rolling Goron as Young + Adult Link
        arrLocation(71) = &H11B4EC          ' *Thaw Zora King
        arrLocation(72) = &H11B4F8          ' *Itmes: 1st and 2nd Scrubs, Lost Dog
        arrLocation(73) = &H11B894          ' *Scarecrow Song
        arrLocation(74) = &H11A66C          ' *Equipment checks, figured this would be easier
        arrLocation(75) = &H11A60C          ' *Check for Biggoron's Sword
        arrLocation(76) = &H11A670          ' *Upgrades
        arrLocation(77) = &H11A674          ' *Quest Items
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
        arrLocation(99) = &H11B150          ' ***Scrub Shuffle (Death Mountain Crater)
        arrLocation(100) = &H11AC54 + 12    ' Link's House (Standing)
        arrLocation(101) = &H11AC8C + 12    ' Lon Lon Ranch Stables (Standing)
        arrLocation(102) = &H11A6DC + 12    ' Jabu-Jabu's Belly (Standing)
        arrLocation(103) = &H11AB84         ' *Shop Checks

        For i As Integer = 0 To arrLocation.Length - 1
            arrChests(i) = 0
        Next
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
        If emulator = "emuhawk" Or emulator = "rmg" Or emulator = "mupen64plus-gui" Then
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

    Private Sub cccCarpenters(chx As Object, e As EventArgs)
        checkCarpenters()
    End Sub
    Private Sub cccEquipment(chx As Object, e As EventArgs)
        checkEquipment()
    End Sub
    Private Sub cccUpgrades(chx As Object, e As EventArgs)
        checkUpgrades()
    End Sub
    Private Sub cccQuestItems(chx As Object, e As EventArgs)
        checkQuestItems()
    End Sub

    ' Scan each of the chests data
    Private Sub readChestData()
        If emulator = String.Empty Then
            If IS_64BIT = False Then
                emulator = "project64"
            Else
                attachToBizHawk()
                If emulator = String.Empty Then attachToRMG()
                If emulator = String.Empty Then attachToM64P()
            End If
        End If
        If emulator = String.Empty Then Exit Sub
        Me.Text = "Tracker of Time (" & emulator & ")"


        If zeldaCheck() = False Then Exit Sub
        ' Get current room code
        Dim roomCode As String = String.Empty
        Dim tempVar As Integer = 0
        roomCode = Hex(goRead(currentRoomAddr))
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If
        While roomCode.Length < 8
            roomCode = "0" & roomCode
        End While
        roomCode = Mid(roomCode, 1, 4)
        Dim dontMath As Integer = Convert.ToInt16(roomCode, 16)
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

        For i As Integer = 0 To arrLocation.Length - 1
            If i Mod 5 = 0 Then Application.DoEvents()
            checkAgain = True
            If i <= 59 Or i >= 100 Then
                doMath = (dontMath * 28) + 212 + &H11A5D0
                tempVar = &H1CA1D8

                ' Royal Tomb Only
                If Hex(doMath) = "11ADC0" Then
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

                If i <= 2 Then
                    ' Scene Checks
                    inc(doMath, 4)
                    tempVar = &H1CA1C8
                ElseIf i <= 30 Or i >= 100 Then
                    ' Standing Checks
                    inc(doMath, 12)
                    tempVar = &H1CA1E4
                End If

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
            ElseIf i >= 78 And i <= 83 And My.Settings.setSkulltula = False Then
                checkAgain = False
            ElseIf i >= 84 And i <= 99 And My.Settings.setScrub = False Then
                checkAgain = False
            End If

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
    End Sub

    Private Function zeldaCheck() As Boolean
        ' Checks for the 'ZELDAZ' within the memory to make sure you are playing Ocarina of Time, and that it is still reading the correct memory region
        zeldaCheck = False
        Dim zeld As Integer = goRead(CInt("&H11A5EC"))
        If zeld = 1514490948 Then
            Dim az As Integer = goRead(CInt("&H11A5F0"))
            If az >= 1096417280 And az <= 1096482815 Then zeldaCheck = True
        End If
    End Function

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

        For ii = startI To 0 Step -1
            strII = ii.ToString
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
                For i = 0 To 11
                    For j = 0 To aKeysDungeons(i).Length - 1
                        With aKeysDungeons(i)(j)
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

            If ii <= arrLow(loc) Then Exit Sub
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
        If DO_COUNT Then Exit Sub
        Dim aTotal(26) As Integer
        Dim aCheck(26) As Integer
        Dim countCheck As Boolean = False

        For i = 0 To aTotal.Length - 1
            aTotal(i) = 0
            aCheck(i) = 0
        Next

        For i = 0 To aKeys.Length - 1
            With aKeys(i)
                If .scan = True Then
                    countCheck = False
                    If .gs Then
                        If My.Settings.setSkulltula Then countCheck = True
                        'inc(aTotal(area2num(.area)))
                        'If .checked Then inc(aCheck(area2num(.area)))
                        'End If
                    ElseIf .cow Then
                        If My.Settings.setCow Then countCheck = True
                        'inc(aTotal(area2num(.area)))
                        'If .checked Then inc(aCheck(area2num(.area)))
                        'End If
                    ElseIf .scrub Then
                        If My.Settings.setScrub Then countCheck = True
                        'inc(aTotal(area2num(.area)))
                        'If .checked Then inc(aCheck(area2num(.area)))
                        'End If
                    ElseIf .shop Then
                        If My.Settings.setShop Then countCheck = True
                        'inc(aTotal(area2num(.area)))
                        'If .checked Then inc(aCheck(area2num(.area)))
                        'End If
                    Else
                        countCheck = True
                        'inc(aTotal(area2num(.area)))
                        'If .checked Then inc(aCheck(area2num(.area)))
                    End If

                    If countCheck Then
                        inc(aTotal(area2num(.area)))
                        If .checked Or .forced Then inc(aCheck(area2num(.area)))
                    End If
                End If
            End With
        Next

        Dim output As String = String.Empty

        output = "Kokiri Forest: " & aCheck(0).ToString & "/" & aTotal(0).ToString
        If Not lblKokiriForest.Text = output Then lblKokiriForest.Text = output
        output = "Lost Woods: " & aCheck(1).ToString & "/" & aTotal(1).ToString
        If Not lblLostWoods.Text = output Then lblLostWoods.Text = output
        output = "Sacred Forest Meadow: " & aCheck(2).ToString & "/" & aTotal(2).ToString
        If Not lblSacredForestMeadow.Text = output Then lblSacredForestMeadow.Text = output
        output = "Hyrule Field: " & aCheck(3).ToString & "/" & aTotal(3).ToString
        If Not lblHyruleField.Text = output Then lblHyruleField.Text = output
        output = "Lon Lon Ranch: " & aCheck(4).ToString & "/" & aTotal(4).ToString
        If Not lblLonLonRanch.Text = output Then lblLonLonRanch.Text = output
        output = "Market: " & aCheck(5).ToString & "/" & aTotal(5).ToString
        If Not lblMarket.Text = output Then lblMarket.Text = output
        output = "Temple of Time: " & aCheck(6).ToString & "/" & aTotal(6).ToString
        If Not lblTempleOfTime.Text = output Then lblTempleOfTime.Text = output
        output = "Hyrule Castle: " & aCheck(7).ToString & "/" & aTotal(7).ToString
        If Not lblHyruleCastle.Text = output Then lblHyruleCastle.Text = output
        output = "Kakariko Village: " & aCheck(8).ToString & "/" & aTotal(8).ToString
        If Not lblKakarikoVillage.Text = output Then lblKakarikoVillage.Text = output
        output = "Graveyard: " & aCheck(9).ToString & "/" & aTotal(9).ToString
        If Not lblGraveyard.Text = output Then lblGraveyard.Text = output
        output = "Death Mountain Trail: " & aCheck(10).ToString & "/" & aTotal(10).ToString
        If Not lblDMTrail.Text = output Then lblDMTrail.Text = output
        output = "Death Mountain Crater: " & aCheck(11).ToString & "/" & aTotal(11).ToString
        If Not lblDMCrater.Text = output Then lblDMCrater.Text = output
        output = "Goron City: " & aCheck(12).ToString & "/" & aTotal(12).ToString
        If Not lblGoronCity.Text = output Then lblGoronCity.Text = output
        output = "Zora's River: " & aCheck(13).ToString & "/" & aTotal(13).ToString
        If Not lblZorasRiver.Text = output Then lblZorasRiver.Text = output
        output = "Zora's Domain: " & aCheck(14).ToString & "/" & aTotal(14).ToString
        If Not lblZorasDomain.Text = output Then lblZorasDomain.Text = output
        output = "Zora's Fountain: " & aCheck(15).ToString & "/" & aTotal(15).ToString
        If Not lblZorasFountain.Text = output Then lblZorasFountain.Text = output
        output = "Lake Hylia: " & aCheck(16).ToString & "/" & aTotal(16).ToString
        If Not lblLakeHylia.Text = output Then lblLakeHylia.Text = output
        output = "Gerudo Valley: " & aCheck(17).ToString & "/" & aTotal(17).ToString
        If Not lblGerudoValley.Text = output Then lblGerudoValley.Text = output
        output = "Gerudo Fortress: " & aCheck(18).ToString & "/" & aTotal(18).ToString
        If Not lblGerudoFortress.Text = output Then lblGerudoFortress.Text = output
        output = "Haunted Wasteland: " & aCheck(19).ToString & "/" & aTotal(19).ToString
        If Not lblHauntedWasteland.Text = output Then lblHauntedWasteland.Text = output
        output = "Desert Colossus: " & aCheck(20).ToString & "/" & aTotal(20).ToString
        If Not lblDesertColossus.Text = output Then lblDesertColossus.Text = output
        output = "Outside Ganon's Castle: " & aCheck(21).ToString & "/" & aTotal(21).ToString
        If Not lblOutsideGanonsCastle.Text = output Then lblOutsideGanonsCastle.Text = output
        output = "Quest: Big Poe Hunt: " & aCheck(22).ToString & "/" & aTotal(22).ToString
        If Not lblQuestBigPoes.Text = output Then lblQuestBigPoes.Text = output
        output = "Quest: Frogs: " & aCheck(23).ToString & "/" & aTotal(23).ToString
        If Not lblQuestFrogs.Text = output Then lblQuestFrogs.Text = output
        output = "Quest: Gold Skulltulas: " & aCheck(24).ToString & "/" & aTotal(24).ToString
        If Not lblQuestGoldSkulltulas.Text = output Then lblQuestGoldSkulltulas.Text = output
        output = "Quest: Masks: " & aCheck(25).ToString & "/" & aTotal(25).ToString
        If Not lblQuestMasks.Text = output Then lblQuestMasks.Text = output
    End Sub
    Private Sub updateLabelsDungeons()
        Dim aTotal(11) As Integer
        Dim aCheck(11) As Integer
        Dim countCheck As Boolean = False

        For i = 0 To 11
            aTotal(i) = 0
            aCheck(i) = 0
            For ii = 0 To aKeysDungeons(i).Length - 1
                With aKeysDungeons(i)(ii)
                    If .scan = True Then
                        countCheck = False
                        If .gs Then
                            If My.Settings.setSkulltula Then countCheck = True
                            'inc(aTotal(i))
                            'If .checked Then inc(aCheck(i))
                            'End If
                        ElseIf .cow Then
                            If My.Settings.setCow Then countCheck = True
                            'inc(aTotal(i))
                            'If .checked Then inc(aCheck(i))
                            'End If
                        ElseIf .scrub Then
                            If My.Settings.setScrub Then countCheck = True
                            'inc(aTotal(i))
                            'If .checked Then inc(aCheck(i))
                            'End If
                        Else
                            countCheck = True
                            'inc(aTotal(i))
                            'If .checked Then inc(aCheck(i))
                        End If

                        If countCheck Then
                            inc(aTotal(i))
                            If .checked Or .forced Then inc(aCheck(i))
                        End If
                    End If
                End With
            Next
        Next

        Dim output As String = String.Empty

        output = "Deku Tree: " & aCheck(0).ToString & "/" & aTotal(0).ToString
        If Not lblDekuTree.Text = output Then lblDekuTree.Text = output
        output = "Dodongo's Cavern: " & aCheck(1).ToString & "/" & aTotal(1).ToString
        If Not lblDodongosCavern.Text = output Then lblDodongosCavern.Text = output
        output = "Jabu-Jabu's Belly: " & aCheck(2).ToString & "/" & aTotal(2).ToString
        If Not lblJabuJabusBelly.Text = output Then lblJabuJabusBelly.Text = output
        output = "Forest Temple: " & aCheck(3).ToString & "/" & aTotal(3).ToString
        If Not lblForestTemple.Text = output Then lblForestTemple.Text = output
        output = "Fire Temple: " & aCheck(4).ToString & "/" & aTotal(4).ToString
        If Not lblFireTemple.Text = output Then lblFireTemple.Text = output
        output = "Water Temple: " & aCheck(5).ToString & "/" & aTotal(5).ToString
        If Not lblWaterTemple.Text = output Then lblWaterTemple.Text = output
        output = "Spirit Temple: " & aCheck(6).ToString & "/" & aTotal(6).ToString
        If Not lblSpiritTemple.Text = output Then lblSpiritTemple.Text = output
        output = "Shadow Temple: " & aCheck(7).ToString & "/" & aTotal(7).ToString
        If Not lblShadowTemple.Text = output Then lblShadowTemple.Text = output
        output = "Bottom of the Well: " & aCheck(8).ToString & "/" & aTotal(8).ToString
        If Not lblBottomOfTheWell.Text = output Then lblBottomOfTheWell.Text = output
        output = "Ice Cavern: " & aCheck(9).ToString & "/" & aTotal(9).ToString
        If Not lblIceCavern.Text = output Then lblIceCavern.Text = output
        output = "Gerudo Training Ground: " & aCheck(10).ToString & "/" & aTotal(10).ToString
        If Not lblGerudoTrainingGround.Text = output Then lblGerudoTrainingGround.Text = output
        output = "Ganon's Castle: " & aCheck(11).ToString & "/" & aTotal(11).ToString
        If Not lblGanonsCastle.Text = output Then lblGanonsCastle.Text = output
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
            Case "HF"
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
            Case "ZR"
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
            Case "QBPH"
                Return 22
            Case "QF"
                Return 23
            Case "QGS"
                Return 24
            Case "QM"
                Return 25
            Case "INV"
                Return 26
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
            Case "Ooutside Ganon's Castle"
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
        If emulator = "emuhawk" Or emulator = "rmg" Or emulator = "mupen64plus-gui" Then
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
                MessageBox.Show("goRead Problem: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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

    Private Sub lbtnScan_Click(sender As Object, e As EventArgs)
        goScan(False)
    End Sub

    Private Sub goScan(ByVal auto As Boolean)
        keepRunning = True
        readChestData()
        If Not keepRunning Then
            stopScanning()
            Exit Sub
        End If
        checkMQs()
        readChestData()
        updateItems()
        updateLabels()
        updateLabelsDungeons()
        If Not auto Then Exit Sub
        If tmrAutoScan.Enabled = False Then
            AutoScanToolStripMenuItem.Text = "Stop"
            tmrAutoScan.Enabled = True
            ' checkMQs()
        Else
            stopScanning()
        End If
    End Sub

    Private Sub lbtnAutoScan_Click(sender As Object, e As EventArgs)
        goScan(True)
    End Sub

    Private Sub tmrAutoScan_Tick(sender As Object, e As EventArgs) Handles tmrAutoScan.Tick
        If Not keepRunning Or emulator = String.Empty Then
            stopScanning()
            Exit Sub
        End If
        If zeldaCheck() = False Then
            stopScanning()
            Exit Sub
        End If
        readChestData()
        updateItems()
        updateLabels()
        updateLabelsDungeons()
    End Sub

    Private Sub lbtnReset_Click(sender As Object, e As EventArgs)
        stopScanning()
        For Each chk In pnlHidden.Controls.OfType(Of CheckBox)()
            chk.Checked = False
        Next
        rtbOutput.ResetText()
        lastRoomScan = 0
        populateLocations()
        'btnScan.Focus()
    End Sub

    Private Sub checkMQs()
        ' Update the MQ Dungeons
        Dim MQs As String = String.Empty
        Dim testSpot As Integer = &H40B6E0
        Dim hits As Integer = 0
        Dim allZeros As Integer = 0

        For i = 0 To aMQ.Length - 1
            aMQ(i) = False
        Next

        For i = 0 To 1
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
        Next

        updateMQs()
    End Sub

    Private Sub scanEmulator(Optional emuName As String = "rmg")
        Dim target As Process = Nothing
        Try
            target = Process.GetProcessesByName(emuName)(0)
            rtbOutput.AppendText(target.ProcessName & vbCrLf)
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                rtbOutput.AppendText(emuName & " not found!" & vbCrLf)
                'MessageBox.Show("BizHawk not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                rtbOutput.AppendText("Problem: " & ex.Message & vbCrLf)
            End If
            Return
        End Try
        Dim addressDLL As Int64 = 0
        For Each mo As ProcessModule In target.Modules
            'If LCase(mo.ModuleName) = "mupen64plus.dll" Then
            'addressDLL = mo.BaseAddress.ToInt64
            'Exit For
            'End If
            rtbOutput.AppendText(mo.ModuleName & ":" & Hex(mo.BaseAddress.ToInt64) & vbCrLf)
        Next
    End Sub

    Private Sub btnTheme_Click(sender As Object, e As EventArgs) Handles btnTheme.Click
        ' This is just my little debug testing button
        'Dim goRead As Integer = ReadMemory(Of Byte)(romAddrStart64 + &H11A5EC)
        Dim goRead As Integer = quickRead8(romAddrStart + &H11A5EC, emulator)
        rtbOutput.AppendText(Hex(goRead) & vbCrLf)
        Exit Sub
        For i = 0 To arrHigh.Length - 1
            rtbOutput.AppendText(i.ToString & ": " & arrLow(i).ToString & "/" & arrHigh(i).ToString & vbCrLf)
        Next
        Exit Sub
        For Each key In aKeysDungeons(2)
            rtbOutput.AppendText(key.loc & ": " & key.name & " = " & key.checked & vbCrLf)
        Next
        Exit Sub

        compressGUI()
    End Sub

    Private Sub updateLCX()
        ' LCX are the Label comboboxes I use for custom checkbox drawing. This updates how to draw them

        ' Set art to work with graphics, pen for single lines, and brushes for filled rectangles
        Dim art As Graphics
        Dim pn1 As Pen = New Pen(Me.ForeColor, 1)
        Dim brBack As Brush = New SolidBrush(Me.BackColor)
        Dim brFore As Brush = New SolidBrush(Me.ForeColor)

        ' Work with the Show Skulltula label
        With lcxShowSkulltulas
            .Refresh()

            ' Set up its graphics and draw the basic outline square
            art = .CreateGraphics
            art.DrawRectangle(pn1, 2, 0, 11, 11)

            If My.Settings.setSkulltula Then
                ' If the setting is TRUE, draw a filled square
                art.FillRectangle(brFore, 4, 2, 8, 8)
            Else
                ' If the setting is FALSE, empty out the square
                art.FillRectangle(brBack, 4, 2, 8, 8)
            End If
        End With

        ' Work with the Cow Shuffle label
        With lcxCowShuffle
            .Refresh()

            ' Set up its graphics and draw the basic outline square
            art = .CreateGraphics
            art.DrawRectangle(pn1, 2, 0, 11, 11)

            If My.Settings.setCow Then
                ' If the setting is TRUE, draw a filled square
                art.FillRectangle(brFore, 4, 2, 8, 8)
            Else
                ' If the setting is FALSE, empty out the square
                art.FillRectangle(brBack, 4, 2, 8, 8)
            End If
        End With

        ' Work with the Scrub Shuffle label
        With lcxScrubShuffle
            .Refresh()

            ' Set up its graphics and draw the basic outline square
            art = .CreateGraphics
            art.DrawRectangle(pn1, 2, 0, 11, 11)

            If My.Settings.setScrub Then
                ' If the setting is TRUE, draw a filled square
                art.FillRectangle(brFore, 4, 2, 8, 8)
            Else
                ' If the setting is FALSE, empty out the square
                art.FillRectangle(brBack, 4, 2, 8, 8)
            End If
        End With

        ' Work with the Shopsanity label
        With lcxShopsanity
            .Refresh()

            ' Set up its graphics and draw the basic outline square
            art = .CreateGraphics
            art.DrawRectangle(pn1, 2, 0, 11, 11)

            If My.Settings.setShop Then
                ' If the setting is TRUE, draw a filled square
                art.FillRectangle(brFore, 4, 2, 8, 8)
            Else
                ' If the setting is FALSE, empty out the square
                art.FillRectangle(brBack, 4, 2, 8, 8)
            End If
        End With
    End Sub

    Private Sub changeTheme(Optional theme As Byte = 0)
        Dim cBack As Color = Control.DefaultBackColor
        Dim cFore As Color = Control.DefaultForeColor
        Dim cpnlBack As Color = Control.DefaultBackColor
        Dim cpnlFore As Color = Control.DefaultForeColor
        Select Case theme
            Case 0  ' Light Mode
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


        'ddThemes.BackColor = cBack
        'ddThemes.ForeColor = cFore

        rtbOutput.BackColor = cBack
        rtbOutput.ForeColor = cFore

        mnuOptions.BackColor = cBack
        mnuOptions.ForeColor = cFore
        custMenu.highlight = cBlend
        custMenu.backColour = cBack
        custMenu.foreColour = cFore

        My.Settings.setTheme = theme
        My.Settings.Save()

        drawGraphics()
    End Sub

    Private Sub compressGUI()
        'Exit Sub
        pnlHidden.Visible = False
        Button2.Visible = False

        Me.Width = pnlDekuTree.Location.X + pnlDekuTree.Width + 22
        Me.Height = rtbOutput.Location.Y + rtbOutput.Height + 46
        btnTheme.Visible = False
    End Sub

    Private Sub displayChecks(ByVal area As String, Optional showChecked As Boolean = False)
        Dim displayName As String = String.Empty
        Dim sOut As String = String.Empty

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
            Case "QBPH"
                displayName = "Quest: Big Poe Hunt"
            Case "QF"
                displayName = "Quest: Frogs"
            Case "QGS"
                displayName = "Quest: Gold Skulltulas"
            Case "QM"
                displayName = "Quest: Masks"
            Case Else
                displayName = "Bad Area Code"
        End Select

        ' If right-clicked to show checked, not that it is found checks
        If showChecked Then displayName = displayName & " (Found)"

        ' Scan the area code for missed or found checks
        sOut = scanArea(area, showChecked)

        ' If the string is empty, nothing was found, note if no checks or complete checks
        If sOut = String.Empty Then sOut = "  " & IIf(showChecked, "None", "Complete!").ToString & vbCrLf

        ' If there is previous text in the output box, add a couple new lines for spacing
        If rtbOutput.TextLength > 0 Then displayName = vbCrLf & vbCrLf & displayName

        ' Display the output, cropping off the last new line code
        rtbOutput.AppendText(displayName & ":" & vbCrLf & sOut.Substring(0, sOut.Length - 2))
    End Sub
    Private Function scanArea(ByVal area As String, ByVal showChecked As Boolean) As String
        scanArea = String.Empty
        ' Prefix for adding indent and things like GS (Gold Skulltula) for other options
        Dim prefix As String = String.Empty
        Dim addCheck As Boolean = False

        For i = 0 To aKeys.Length - 1
            With aKeys(i)
                ' Stop if an empty key is found
                If .loc = String.Empty Then Exit For
                If .area = area Then
                    ' Determine the prefix
                    If .gs Then
                        prefix = "  GS: "
                    ElseIf .cow Then
                        prefix = "  Cow: "
                    ElseIf .scrub Then
                        prefix = "  Deku Scrub: "
                    ElseIf .shop Then
                        prefix = "  Shopsanity: "
                    Else
                        prefix = "  "
                    End If
                    ' So long as it is not checked and in good standing, add it to the output
                    If .scan = True Then
                        If .checked = showChecked Or (showChecked And .forced) Then
                            addCheck = False
                            If .gs Then
                                If My.Settings.setSkulltula Then addCheck = True
                            ElseIf .cow Then
                                If My.Settings.setCow Then addCheck = True
                            ElseIf .scrub Then
                                If My.Settings.setScrub Then addCheck = True
                            ElseIf .shop Then
                                If My.Settings.setShop Then addCheck = True
                            Else
                                addCheck = True
                            End If
                            ' Remove the forced checks from the unchecked list
                            If Not showChecked And .forced Then addCheck = False

                            ' Output the check and note if it is forced
                            If addCheck Then scanArea = scanArea & prefix & .name & IIf(.forced, " (Forced)", "").ToString & vbCrLf
                        End If
                    End If
                End If
            End With
        Next
    End Function
    Private Sub displayChecksDungeons(ByVal dungeon As Byte, Optional showChecked As Boolean = False)
        Dim displayName As String = String.Empty
        Dim sOut As String = String.Empty

        displayName = dungeonNumber2name(dungeon)

        ' Add 'MQ' to the display if it is a Master Quest dungeon
        If aMQ(dungeon) Then displayName = displayName & " MQ"

        ' If right-clicked to show checked, not that it is found checks
        If showChecked Then displayName = displayName & " (Found)"

        ' Scan the dungeon for missed or found checks
        sOut = scanDungeon(dungeon, showChecked)

        ' If the string is empty, nothing was found, note if no checks or complete checks
        If sOut = String.Empty Then sOut = "  " & IIf(showChecked, "None", "Complete!").ToString & vbCrLf

        ' If there is previous text in the output box, add a couple new lines for spacing
        If rtbOutput.TextLength > 0 Then displayName = vbCrLf & vbCrLf & displayName

        ' Display the output, cropping off the last new line code
        rtbOutput.AppendText(displayName & ":" & vbCrLf & sOut.Substring(0, sOut.Length - 2))
    End Sub
    Private Function scanDungeon(ByVal dungeon As Byte, ByVal showChecked As Boolean) As String
        scanDungeon = String.Empty
        Dim prefix As String = String.Empty
        Dim addCheck As Boolean = False
        ' Dim count As Integer = 0

        For i = 0 To aKeysDungeons(dungeon).Length - 1
            With aKeysDungeons(dungeon)(i)
                ' Stop if an empty key is found
                If .loc = String.Empty Then Exit For
                ' Determine the prefix
                If .gs Then
                    prefix = "  GS: "
                ElseIf .cow Then
                    prefix = "  Cow: "
                ElseIf .scrub Then
                    prefix = "  Deku Scrub: "
                ElseIf .shop Then
                    prefix = "  Shopsanity: "
                Else
                    prefix = "  "
                End If

                If .scan = True Then
                    If .checked = showChecked Or (showChecked And .forced) Then
                        addCheck = False
                        If .gs Then
                            If My.Settings.setSkulltula Then addCheck = True
                        ElseIf .cow Then
                            If My.Settings.setCow Then addCheck = True
                        ElseIf .scrub Then
                            If My.Settings.setScrub Then addCheck = True
                        ElseIf .shop Then
                            If My.Settings.setShop Then addCheck = True
                        Else
                            addCheck = True
                        End If
                        If Not showChecked And .forced Then addCheck = False

                        If addCheck Then scanDungeon = scanDungeon & prefix & .name & IIf(.forced, " (Forced)", "").ToString & vbCrLf
                    End If
                End If

            End With
        Next
        '    If DO_COUNT Then scanDungeon = scanDungeon & "Key Count: " & count.ToString & vbCrLf
    End Function

    Private Sub setupKeys()
        ' tK is short for thisKey
        Dim tK As Integer = 0

        For i = 0 To aKeys.Length - 1
            With aKeys(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
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
        makeKeysShoppes(tK)

        If DO_COUNT Then rtbOutput.AppendText("aKeys: " & tK.ToString & vbCrLf)
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
                    While strLoc.Length < 5
                        strLoc = "0" & strLoc
                    End While
                    bLoc = CByte(Mid(strLoc, 1, 3))
                    bVal = CByte(Mid(strLoc, 4, 2))
                    If bVal > arrHigh(bLoc) Then arrHigh(bLoc) = bVal
                    If bVal < arrLow(bLoc) Then arrLow(bLoc) = bVal
                End If
            End With
        Next
        For i As Byte = 0 To 11
            For Each key In aKeysDungeons(i)
                With key
                    If IsNumeric(.loc) Then
                        strLoc = .loc
                        While strLoc.Length < 4
                            strLoc = "0" & strLoc
                        End While
                        bLoc = CByte(Mid(strLoc, 1, 2))
                        bVal = CByte(Mid(strLoc, 3, 2))
                        If bVal > arrHigh(bLoc) Then
                            arrHigh(bLoc) = bVal
                        End If
                        If bVal < arrLow(bLoc) Then arrLow(bLoc) = bVal
                    End If
                End With
            Next
        Next
    End Sub

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
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6717"
            .area = "KF"
            .name = "Gift from Saria"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4612"
            .area = "KF"
            .name = "Storms Grotto Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8100"
            .area = "KF"
            .name = "Soft Soil Near Shoppe (Young)"
            .gs = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8101"
            .area = "KF"
            .name = "Behind Know-it-All Brother's House (N) (Young)"
            .gs = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8102"
            .area = "KF"
            .name = "Above the House of Twins (N) (Adult)"
            .gs = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "10024"
            .area = "KF"
            .name = "Link's House (Adult)"
            .cow = True
        End With
        inc(tK)
    End Sub
    Private Sub makeKeysLW(ByRef tK As Integer)
        ' Set up keys for the Lost Woods
        With aKeys(tK)
            .loc = "7202"
            .area = "LW"
            .name = "Deku Scrub Near Bridge (Young)"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "4620"
            .area = "LW"
            .name = "Near Shortcuts Grotto Chest"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "7203"
            .area = "LW"
            .name = "Deku Scrub in Grotto"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6813"
            .area = "LW"
            .name = "Target in Woods (Young)"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6807"
            .area = "LW"
            .name = "Ocarina Memory Game (Young)"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6806"
            .area = "LW"
            .name = "Skull Kid (Young)"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6814"
            .area = "LW"
            .name = "Deku Theatre Skull Mask (Young)"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "6815"
            .area = "LW"
            .name = "Deku Theatre Mask of Truth (Young)"
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8108"
            .area = "LW"
            .name = "Soft Soil Near Bridge (Young)"
            .gs = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8109"
            .area = "LW"
            .name = "Soft Soil Near Theatre (Young)"
            .gs = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "8110"
            .area = "LW"
            .name = "Theatre Bean Plant Ride (N) (Adult)"
            .gs = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9810"
            .area = "LW"
            .name = "Near Bridge (Young)"
            .scrub = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9801"
            .area = "LW"
            .name = "Outside Theatre Right (Young)"
            .scrub = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9802"
            .area = "LW"
            .name = "Outside Theatre Left (Young)"
            .scrub = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9311"
            .area = "LW"
            .name = "Grotto Left"
            .scrub = True
        End With
        inc(tK)
        With aKeys(tK)
            .loc = "9304"
            .area = "LW"
            .name = "Grotto Right"
            .scrub = True
        End With
        inc(tK)
    End Sub
    Private Sub makeKeysSFM(ByRef tk As Integer)
        ' Set up keys for the Sacred Forest Meadow
        With aKeys(tk)
            .loc = "4617"
            .area = "SFM"
            .name = "Wolfos Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6407"
            .area = "SFM"
            .name = "Song from Saria (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6400"
            .area = "SFM"
            .name = "Song from Sheik (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8111"
            .area = "SFM"
            .name = "On East Wall (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9009"
            .area = "SFM"
            .name = "Storms Grotto Front"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9008"
            .area = "SFM"
            .name = "Storms Grotto Back"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysHF(ByRef tk As Integer)
        ' Set up keys for Hyrule Field
        With aKeys(tk)
            .loc = "4600"
            .area = "HF"
            .name = "Near Market Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4602"
            .area = "HF"
            .name = "Southeast Grotto Chest"
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
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1901"
            .area = "HF"
            .name = "Tektite Grotto Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "103"
            .area = "HF"
            .name = "Get Ocarina of Time"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6625"
            .area = "HF"
            .name = "Song from Ocarina of Time"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8016"
            .area = "HF"
            .name = "Grotto Near Gerudo Valley"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8017"
            .area = "HF"
            .name = "Grotto Near Kakariko Village"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1925"
            .area = "HF"
            .name = "Grotto Near Gerudo Valley"
            .cow = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8803"
            .area = "HF"
            .name = "Inside Fence Grotto"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysLLR(ByRef tk As Integer)
        ' Set up keys for Lon Lon Ranch
        With aKeys(tk)
            .loc = "2101"
            .area = "LLR"
            .name = "Tower Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6818"
            .area = "LLR"
            .name = "Talon's Chickens (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6408"
            .area = "LLR"
            .name = "Song from Malon (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6208"
            .area = "LLR"
            .name = "Get Epona (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8024"
            .area = "LLR"
            .name = "Wall Near Tower (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8025"
            .area = "LLR"
            .name = "Behind Coral (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8026"
            .area = "LLR"
            .name = "On House Window (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8027"
            .area = "LLR"
            .name = "Tree Near House (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "10124"
            .area = "LLR"
            .name = "Stables Left"
            .cow = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "10125"
            .area = "LLR"
            .name = "Stables Right"
            .cow = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2125"
            .area = "LLR"
            .name = "Tower Left"
            .cow = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2124"
            .area = "LLR"
            .name = "Tower Right"
            .cow = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9601"
            .area = "LLR"
            .name = "Open Grotto Left (Young)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9604"
            .area = "LLR"
            .name = "Open Grotto Centre (Young)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9606"
            .area = "LLR"
            .name = "Open Grotto Right (Young)"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysMK(ByRef tk As Integer)
        ' Set up keys for the Market
        With aKeys(tk)
            .loc = "6829"
            .area = "MK"
            .name = "Shooting Gallery Reward (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6801"
            .area = "MK"
            .name = "Bombchu Bowling First Prize (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6802"
            .area = "MK"
            .name = "Bombchu Bowling Second Prize (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7201"
            .area = "MK"
            .name = "Lost Dog (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6811"
            .area = "MK"
            .name = "Treasure Chest Game (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8119"
            .area = "MK"
            .name = "Crate in Guard House (Young)"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysTT(ByRef tk As Integer)
        ' Set up keys for the Temple of Time
        With aKeys(tk)
            .loc = "6405"
            .area = "TT"
            .name = "Song from Shiek"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6720"
            .area = "TT"
            .name = "Light Arrows Cutscene"
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
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6409"
            .area = "HC"
            .name = "Song from Impa (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6809"
            .area = "HC"
            .name = "Great Fairy Fountain (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8117"
            .area = "HC"
            .name = "Inside Storms Grotto (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8118"
            .area = "HC"
            .name = "Tree Near Entrance (Young)"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysKV(ByRef tk As Integer)
        ' Set up keys for Kakariko Village
        With aKeys(tk)
            .loc = "4610"
            .area = "KV"
            .name = "Redead Grotto Chest"
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
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6928"
            .area = "KV"
            .name = "Talk to Anju (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6805"
            .area = "KV"
            .name = "Talk to Man on Roof"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1801"
            .area = "KV"
            .name = "Impa's House Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6830"
            .area = "KV"
            .name = "Shooting Gallery (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2001"
            .area = "KV"
            .name = "Windmill Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6411"
            .area = "KV"
            .name = "Song from Windmill (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6404"
            .area = "KV"
            .name = "Song from Shiek (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8201"
            .area = "KV"
            .name = "On House Near Death Mountain Trail (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8202"
            .area = "KV"
            .name = "Large Ladder (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8203"
            .area = "KV"
            .name = "Construction Site (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8204"
            .area = "KV"
            .name = "On Skulltula House (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8205"
            .area = "KV"
            .name = "Tree Near Entrance (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8206"
            .area = "KV"
            .name = "Above Impa's House (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "1824"
            .area = "KV"
            .name = "Inside Impa's House"
            .cow = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGY(ByRef tk As Integer)
        ' Set up keys for the Graveyard
        With aKeys(tk)
            .loc = "2208"
            .area = "GY"
            .name = "Dampe's Gravedigging Tour (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4800"
            .area = "GY"
            .name = "Shield Grave Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2204"
            .area = "GY"
            .name = "Ledge Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4700"
            .area = "GY"
            .name = "Heart Piece Grave Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4900"
            .area = "GY"
            .name = "Royal Family's Tomb Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6410"
            .area = "GY"
            .name = "Song from Royal Family's Tomb"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5000"
            .area = "GY"
            .name = "Hookshot Chest (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2007"
            .area = "GY"
            .name = "Dampe's Race Piece of Heart (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8200"
            .area = "GY"
            .name = "Soft Soil (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8207"
            .area = "GY"
            .name = "On South Wall (N) (Young)"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDMT(ByRef tk As Integer)
        ' Set up keys for Death Mountain Trail
        With aKeys(tk)
            .loc = "5801"
            .area = "DMT"
            .name = "Blast Wall Chest"
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
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4623"
            .area = "DMT"
            .name = "Storms Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6008"
            .area = "DMT"
            .name = "Help Biggoron (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8125"
            .area = "DMT"
            .name = "Soft Soil Near Dodongo's Cavern (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8126"
            .area = "DMT"
            .name = "Blast Wall Near Village (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8127"
            .area = "DMT"
            .name = "Rock Near Bomb Flower (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8128"
            .area = "DMT"
            .name = "Rock Near Climb Wall (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2824"
            .area = "DMT"
            .name = "In Grotto"
            .cow = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDMC(ByRef tk As Integer)
        ' Set up keys for Death Mountain Crater
        With aKeys(tk)
            .loc = "4626"
            .area = "DMC"
            .name = "Upper Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2902"
            .area = "DMC"
            .name = "Wall Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2908"
            .area = "DMC"
            .name = "Volcano Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "016"
            .area = "DMC"
            .name = "Great Fairy Fountain"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6401"
            .area = "DMC"
            .name = "Song from Sheik (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8124"
            .area = "DMC"
            .name = "Soft Soil Near Warp (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8131"
            .area = "DMC"
            .name = "Crate at Entrance (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9906"
            .area = "DMC"
            .name = "Near Ladder (Young)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9401"
            .area = "DMC"
            .name = "Grotto Left (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9404"
            .area = "DMC"
            .name = "Grotto Centre (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9406"
            .area = "DMC"
            .name = "Grotto Right (Adult)"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGC(ByRef tk As Integer)
        ' Set up keys for Goron City
        With aKeys(tk)
            .loc = "5900"
            .area = "GC"
            .name = "Maze Left Chest (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5902"
            .area = "GC"
            .name = "Maze Centre Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5901"
            .area = "GC"
            .name = "Maze Right Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7014"
            .area = "GC"
            .name = "Rolling Goron (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7025"
            .area = "GC"
            .name = "Rolling Goron (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "3031"
            .area = "GC"
            .name = "Pot Piece of Heart (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6306"
            .area = "GC"
            .name = "Darunia's Joy (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8129"
            .area = "GC"
            .name = "Rope Platform (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8130"
            .area = "GC"
            .name = "Crate in Maze Near Chests (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9501"
            .area = "GC"
            .name = "Open Grotto Left (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9504"
            .area = "GC"
            .name = "Open Grotto Centre (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9506"
            .area = "GC"
            .name = "Open Grotto Right (Adult)"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZR(ByRef tk As Integer)
        ' Set up keys for Zora's River
        With aKeys(tk)
            .loc = "2304"
            .area = "ZR"
            .name = "Near Open Grotto Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "4609"
            .area = "ZR"
            .name = "Open Grotto Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2311"
            .area = "ZR"
            .name = "Near Domain Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8208"
            .area = "ZR"
            .name = "On Ladder Near Zora's Domain (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8209"
            .area = "ZR"
            .name = "Tree Near Entrance (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8211"
            .area = "ZR"
            .name = "Wall Above Bridge (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8212"
            .area = "ZR"
            .name = "Wall Near Open Grotto (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8909"
            .area = "ZR"
            .name = "Storms Grotto Front"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8908"
            .area = "ZR"
            .name = "Storms Grotto Back"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZD(ByRef tk As Integer)
        ' Set up keys for Zora's Domain
        With aKeys(tk)
            .loc = "5300"
            .area = "ZD"
            .name = "Torches Chest (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6308"
            .area = "ZD"
            .name = "Diving Minigame (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "7108"
            .area = "ZD"
            .name = "Thaw Zora King (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8214"
            .area = "ZD"
            .name = "Frozen Waterfall Top (N) (Adult)"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysZF(ByRef tk As Integer)
        ' Set up keys for Zora's Fountain
        With aKeys(tk)
            .loc = "2501"
            .area = "ZF"
            .name = "Iceberg Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2520"
            .area = "ZF"
            .name = "Bottom Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6808"
            .area = "ZF"
            .name = "Great Fairy Fountain"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8210"
            .area = "ZF"
            .name = "Lower West Wall (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8213"
            .area = "ZF"
            .name = "Southeast Corner Silver Rock Passage (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8215"
            .area = "ZF"
            .name = "Southeast Corner Tree (Young)"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysLH(ByRef tk As Integer)
        ' Set up keys for Lake Hylia
        With aKeys(tk)
            .loc = "7316"
            .area = "LH"
            .name = "Scarecrow (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6512"
            .area = "LH"
            .name = "Scarecrow (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "211"
            .area = "LH"
            .name = "Underwater Item (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6110"
            .area = "LH"
            .name = "Big Fish (Young)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6111"
            .area = "LH"
            .name = "Big Fish (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2430"
            .area = "LH"
            .name = "Lab Tower Piece of Heart (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6800"
            .area = "LH"
            .name = "Lab Dive"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "5200"
            .area = "LH"
            .name = "Shoot the Sun (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8216"
            .area = "LH"
            .name = "Soft Soil Near Lab (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8217"
            .area = "LH"
            .name = "On Small Island Pillar (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8218"
            .area = "LH"
            .name = "Behind Lab (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8219"
            .area = "LH"
            .name = "Crate in Lab Pool (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8220"
            .area = "LH"
            .name = "On Tree Near Warp (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9101"
            .area = "LH"
            .name = "Grave Grotto Left (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9104"
            .area = "LH"
            .name = "Grave Grotto Centre (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9106"
            .area = "LH"
            .name = "Grave Grotto Right (Adult)"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGV(ByRef tk As Integer)
        ' Set up keys for Gerudo Valley
        With aKeys(tk)
            .loc = "2602"
            .area = "GV"
            .name = "Crate Piece of Heart"
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
            .name = "Chest Behind Rocks"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8224"
            .area = "GV"
            .name = "Soft Soil Near Cow (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8225"
            .area = "GV"
            .name = "Waterfall Near Entrance (N) (Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8226"
            .area = "GV"
            .name = "Pillar Near Tents (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8227"
            .area = "GV"
            .name = "Behind Tents (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "2624"
            .area = "GV"
            .name = "Near Soft Soil"
            .cow = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9209"
            .area = "GV"
            .name = "Storms Grotto Front (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9208"
            .area = "GV"
            .name = "Storms Grotto Back (Adult)"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysGF(ByRef tk As Integer)
        ' Set up keys for Gerudo's Fortress
        With aKeys(tk)
            .loc = "5600"
            .area = "GF"
            .name = "Chest on Top(Adult)"
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
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6831"
            .area = "GF"
            .name = "Archery 1500 points (Adult)"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6500"
            .area = "GF"
            .name = "Carpenter #1"
            .scan = False
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6501"
            .area = "GF"
            .name = "Carpenter #2"
            .scan = False
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6502"
            .area = "GF"
            .name = "Carpenter #3"
            .scan = False
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6503"
            .area = "GF"
            .name = "Carpenter #4"
            .scan = False
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8300"
            .area = "GF"
            .name = "Top Wall Near Scarecrow (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8301"
            .area = "GF"
            .name = "Far Archery Target (N) (Adult)"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysHW(ByRef tk As Integer)
        ' Set up keys for the Haunted Wasteland
        With aKeys(tk)
            .loc = "5700"
            .area = "HW"
            .name = "Structure Chest"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8309"
            .area = "HW"
            .name = "Structure Basement"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysDC(ByRef tk As Integer)
        ' Set up keys for the Desert Colossus
        With aKeys(tk)
            .loc = "6810"
            .area = "DC"
            .name = "Great Fairy Fountain"
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
            .name = "Piece of Heart"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8308"
            .area = "DC"
            .name = "Soft Soil By Temple Entrance(Young)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8310"
            .area = "DC"
            .name = "Bean Plant Ride Hill (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8311"
            .area = "DC"
            .name = "Southern Edge On Tree (N) (Adult)"
            .gs = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9709"
            .area = "DC"
            .name = "Grotto Front (Adult)"
            .scrub = True
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "9708"
            .area = "DC"
            .name = "Grotto Back (Adult)"
            .scrub = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysOGC(ByRef tk As Integer)
        ' Set up keys for Outside Ganon's Castle
        With aKeys(tk)
            .loc = "008"
            .area = "OGC"
            .name = "Great Fairy Fountain"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "8116"
            .area = "OGC"
            .name = "On Pillar (Adult)"
            .gs = True
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQBPH(ByRef tk As Integer)
        ' Set up keys for Quest: Big Poe Hunt
        With aKeys(tk)
            .loc = "124"
            .area = "QBPH"
            .name = "Big Poe #1: Near Castle Gate"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "123"
            .area = "QBPH"
            .name = "Big Poe #2: Near Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "122"
            .area = "QBPH"
            .name = "Big Poe #3: East of Castle"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "130"
            .area = "QBPH"
            .name = "Big Poe #4: Between Gerudo Valley and Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "131"
            .area = "QBPH"
            .name = "Big Poe #5: Near Gerudo Valley"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "128"
            .area = "QBPH"
            .name = "Big Poe #6: Southeast Field Near Path"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "129"
            .area = "QBPH"
            .name = "Big Poe #7: Southeast Field Near Rock"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "127"
            .area = "QBPH"
            .name = "Big Poe #8: Betweek Kokiri Forest and Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "126"
            .area = "QBPH"
            .name = "Big Poe #9: Wall East of Lon Lon Ranch"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "125"
            .area = "QBPH"
            .name = "Big Poe #10: Near Kakariko Village"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQF(ByRef tk As Integer)
        ' Set up keys for Quest: Frogs
        With aKeys(tk)
            .loc = "6706"
            .area = "QF"
            .name = "Play Song of Storms"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6705"
            .area = "QF"
            .name = "Play Song of Time"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6704"
            .area = "QF"
            .name = "Play Sarai's Song"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6703"
            .area = "QF"
            .name = "Play Sun's Song"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6702"
            .area = "QF"
            .name = "Play Epona's Song"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6701"
            .area = "QF"
            .name = "Play Zelda's Lullaby"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6700"
            .area = "QF"
            .name = "Ocarina Game"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQGS(ByRef tk As Integer)
        ' Set up keys for Quest: Gold Skulltula Rewards
        With aKeys(tk)
            .loc = "6710"
            .area = "QGS"
            .name = "10 Gold Skulltulas"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6711"
            .area = "QGS"
            .name = "20 Gold Skulltulas"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6712"
            .area = "QGS"
            .name = "30 Gold Skulltulas"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6713"
            .area = "QGS"
            .name = "40 Gold Skulltulas"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6714"
            .area = "QGS"
            .name = "50 Gold Skulltulas"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysQM(ByRef tk As Integer)
        ' Set up keys for the Quest: Masks
        With aKeys(tk)
            .loc = "6919"
            .area = "QM"
            .name = "Get Keaton Mask"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6908"
            .area = "QM"
            .name = "Sell Keaton Mask"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6920"
            .area = "QM"
            .name = "Get Skull Mask"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6909"
            .area = "QM"
            .name = "Sell Skull Mask"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6921"
            .area = "QM"
            .name = "Get Spooky Mask"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6910"
            .area = "QM"
            .name = "Sell Spooky Mask"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6922"
            .area = "QM"
            .name = "Get Bunny Hood"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6911"
            .area = "QM"
            .name = "Sell Bunny Hood"
        End With
        inc(tk)
        With aKeys(tk)
            .loc = "6926"
            .area = "QM"
            .name = "Get Mask of Truth"
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
            .name = "Gerudo's Card"
        End With
        inc(tk)
    End Sub
    Private Sub makeKeysShoppes(ByRef tk As Integer)
        ' Set up keys for Shopsanity
        With aKeys(tk)
            .loc = "10323"
            .area = "KF"
            .name = "Kokiri Shop #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10322"
            .area = "KF"
            .name = "Kokiri Shop #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10321"
            .area = "KF"
            .name = "Kokiri Shop #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10320"
            .area = "KF"
            .name = "Kokiri Shop #4"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10315"
            .area = "MK"
            .name = "Bazaar #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10314"
            .area = "MK"
            .name = "Bazaar #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10313"
            .area = "MK"
            .name = "Bazaar #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10312"
            .area = "MK"
            .name = "Bazaar #4"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10303"
            .area = "MK"
            .name = "Bombchu Shop #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10302"
            .area = "MK"
            .name = "Bombchu Shop #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10301"
            .area = "MK"
            .name = "Bombchu Shop #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10300"
            .area = "MK"
            .name = "Bombchu Shop #4"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10331"
            .area = "MK"
            .name = "Potion Shop #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10330"
            .area = "MK"
            .name = "Potion Shop #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10329"
            .area = "MK"
            .name = "Potion Shop #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10328"
            .area = "MK"
            .name = "Potion Shop #4"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10327"
            .area = "KV"
            .name = "Bazaar #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10326"
            .area = "KV"
            .name = "Bazaar #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10325"
            .area = "KV"
            .name = "Bazaar #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10324"
            .area = "KV"
            .name = "Bazaar #4"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10311"
            .area = "KV"
            .name = "Potion Shop #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10310"
            .area = "KV"
            .name = "Potion Shop #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10309"
            .area = "KV"
            .name = "Potion Shop #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10308"
            .area = "KV"
            .name = "Potion Shop #4"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10319"
            .area = "GC"
            .name = "Goron Shop #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10318"
            .area = "GC"
            .name = "Goron Shop #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10317"
            .area = "GC"
            .name = "Goron Shop #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10316"
            .area = "GC"
            .name = "Goron Shop #4"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10307"
            .area = "ZD"
            .name = "Zora Shop #1"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10306"
            .area = "ZD"
            .name = "Zora Shop #2"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10305"
            .area = "ZD"
            .name = "Zora Shop #3"
            .shop = True
        End With
        inc(tK)
        With aKeys(tk)
            .loc = "10304"
            .area = "ZD"
            .name = "Zora Shop #4"
            .shop = True
        End With
    End Sub
    Private Sub updateMQs()
        ' Set up the dungeons based on if they are Master Quest versions or not
        ' If debug option is on, run both normal and Master Quest version setups, then exit this sub
        If DO_COUNT Then
            makeKeysDekuTree(False)
            makeKeysDekuTree(True)
            makeKeysDodongosCavern(False)
            makeKeysDodongosCavern(True)
            makeKeysJabuJabusBelly(False)
            makeKeysJabuJabusBelly(True)
            makeKeysForestTemple(False)
            makeKeysForestTemple(True)
            makeKeysFireTemple(False)
            makeKeysFireTemple(True)
            makeKeysWaterTemple(False)
            makeKeysWaterTemple(True)
            makeKeysSpiritTemple(False)
            makeKeysSpiritTemple(True)
            makeKeysShadowTemple(False)
            makeKeysShadowTemple(True)
            makeKeysBottomOfTheWell(False)
            makeKeysBottomOfTheWell(True)
            makeKeysIceCavern(False)
            makeKeysIceCavern(True)
            makeKeysGerudoTrainingGround(False)
            makeKeysGerudoTrainingGround(True)
            makeKeysGanonsCastle(False)
            makeKeysGanonsCastle(True)
            Exit Sub
        End If

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
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(0)(CInt(IIf(isMQ, 12, 10)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(0)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(0).Length - 1
            aKeysDungeons(0)(i) = New keyCheck
            With aKeysDungeons(0)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        ' Start off with keys for both versions of the Deku Tree
        With aKeysDungeons(0)(tK)
            .loc = "3103"
            .area = "DT"
            .name = "Map Chest"
        End With

        inc(tK)
        If Not isMQ Then
            ' The non-Master Quest keys for the Deku Tree
            With aKeysDungeons(0)(tK)
                .loc = "3101"
                .area = "DT"
                .name = "Slingshot Room First Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3105"
                .area = "DT"
                .name = "Slingshot Room Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3102"
                .area = "DT"
                .name = "Compass Room First Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3106"
                .area = "DT"
                .name = "Compass Room Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3104"
                .area = "DT"
                .name = "Basement Chest"
            End With
        Else
            ' The Master Quest keys for the Deku Tree
            With aKeysDungeons(0)(tK)
                .loc = "3102"
                .area = "DT"
                .name = "Slingshot Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3106"
                .area = "DT"
                .name = "Slingshot Back Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3101"
                .area = "DT"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3104"
                .area = "DT"
                .name = "Basement Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3105"
                .area = "DT"
                .name = "Before Spinning Log Chest"
            End With
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "3100"
                .area = "DT"
                .name = "After Spinning Log Chest"
            End With
        End If

        ' End with keys for both versions of the Deku Tree
        inc(tK)
        With aKeysDungeons(0)(tK)
            .loc = "1031"
            .area = "DT"
            .name = "Queen Gohma"
        End With
        inc(tK)
        With aKeysDungeons(0)(tK)
            .loc = "7800"
            .area = "DT"
            .name = "Basement Back Room"
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(0)(tK)
            .loc = "7801"
            .area = "DT"
            .name = IIf(isMQ, "Lobby", "Basement Gate").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(0)(tK)
            .loc = "7802"
            .area = "DT"
            .name = IIf(isMQ, "Basement Graves Room", "Basement Vines").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(0)(tK)
            .loc = "7803"
            .area = "DT"
            .name = "Compass Room"
            .gs = True
        End With

        If isMQ Then
            ' Master Quest key for its Deku Scrub
            inc(tK)
            With aKeysDungeons(0)(tK)
                .loc = "8405"
                .area = "DT"
                .name = "Basement"
                .scrub = True
            End With
        End If

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Deku Tree (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysDodongosCavern(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(1) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(1)(CInt(IIf(isMQ, 16, 15)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(1)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(1).Length - 1
            aKeysDungeons(1)(i) = New keyCheck
            With aKeysDungeons(1)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        If Not isMQ Then
            ' The non-Master Quest keys to Dodongo's Cavern
            With aKeysDungeons(1)(tK)
                .loc = "3208"
                .area = "DDC"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3205"
                .area = "DDC"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3206"
                .area = "DDC"
                .name = "Bomb Flower Platform Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3204"
                .area = "DDC"
                .name = "Bomb Bag Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3210"
                .area = "DDC"
                .name = "End of Bridge Chest"
            End With
        Else
            ' The Master Quest keys to Dodongo's Cavern
            With aKeysDungeons(1)(tK)
                .loc = "3200"
                .area = "DDC"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3204"
                .area = "DDC"
                .name = "Bomb Bag Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3203"
                .area = "DDC"
                .name = "Torch Puzzle Room"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3202"
                .area = "DDC"
                .name = "Larvae Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3205"
                .area = "DDC"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "3201"
                .area = "DDC"
                .name = "Under Grave Chest"
            End With
        End If

        ' End with keys for both versions of Dodongo's Cavern
        inc(tK)
        With aKeysDungeons(1)(tK)
            .loc = "4400"
            .area = "DDC"
            .name = "Boss Room Chest"
        End With
        inc(tK)
        With aKeysDungeons(1)(tK)
            .loc = "1131"
            .area = "DDC"
            .name = "King Dodongo"
        End With
        inc(tK)
        With aKeysDungeons(1)(tK)
            .loc = "7808"
            .area = "DDC"
            .name = IIf(isMQ, "Back Area", "Vines Above Stairs").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(1)(tK)
            .loc = "7809"
            .area = "DDC"
            .name = IIf(isMQ, "Lizalfos Room", "Scarecrow").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(1)(tK)
            .loc = "7810"
            .area = "DDC"
            .name = IIf(isMQ, "Scrub Room", "Alcove Above Stairs").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(1)(tK)
            .loc = "7811"
            .area = "DDC"
            .name = IIf(isMQ, "Song of Time Block Room", "Back Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(1)(tK)
            .loc = "7812"
            .area = "DDC"
            .name = IIf(isMQ, "Larvae Room", "Side Room Near Lower Lizalfos").ToString
            .gs = True
        End With

        inc(tK)
        If Not isMQ Then
            ' The non-Master Quest keys for the Deku Scrubs in Dodongo's Cavern
            With aKeysDungeons(1)(tK)
                .loc = "8505"
                .area = "DDC"
                .name = "Lobby"
                .scrub = True
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "8502"
                .area = "DDC"
                .name = "Side Room Near Dodongos"
                .scrub = True
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "8501"
                .area = "DDC"
                .name = "Near Bomb Bag Left"
                .scrub = True
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "8504"
                .area = "DDC"
                .name = "Near Bomb Bag Right"
                .scrub = True
            End With
        Else
            ' The Master Quest keys for the Deku Scrubs in Dodongo's Cavern
            With aKeysDungeons(1)(tK)
                .loc = "8504"
                .area = "DDC"
                .name = "Lobby Front"
                .scrub = True
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "8502"
                .area = "DDC"
                .name = "Lobby Back"
                .scrub = True
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "8508"
                .area = "DDC"
                .name = "Side Room Near Lizalfos"
                .scrub = True
            End With
            inc(tK)
            With aKeysDungeons(1)(tK)
                .loc = "8505"
                .area = "DDC"
                .name = "Staircase Room"
                .scrub = True
            End With
        End If

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Dodongo's Cavern (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysJabuJabusBelly(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(2) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(2)(CInt(IIf(isMQ, 16, 8)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(2)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(2).Length - 1
            aKeysDungeons(2)(i) = New keyCheck
            With aKeysDungeons(2)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for Jabu-Jabu's Belly
            With aKeysDungeons(2)(tK)
                .loc = "3301"
                .area = "JB"
                .name = "Boomerang Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3302"
                .area = "JB"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3304"
                .area = "JB"
                .name = "Compass Chest"
            End With
        Else
            ' The Master Quest keys for Jabu-Jabu's Belly
            With aKeysDungeons(2)(tK)
                .loc = "3303"
                .area = "JB"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3305"
                .area = "JB"
                .name = "First Room Side Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3302"
                .area = "JB"
                .name = "Second Room Lower Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3300"
                .area = "JB"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3304"
                .area = "JB"
                .name = "Basement First Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3308"
                .area = "JB"
                .name = "Basement Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3301"
                .area = "JB"
                .name = "Boomerang Room First Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3306"
                .area = "JB"
                .name = "Boomerang Room Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3309"
                .area = "JB"
                .name = "Falling Like Like Room"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3307"
                .area = "JB"
                .name = "Second Room Upper Chest"
            End With
            inc(tK)
            With aKeysDungeons(2)(tK)
                .loc = "3310"
                .area = "JB"
                .name = "Near Boss Chest"
            End With
        End If
        inc(tK)
        With aKeysDungeons(2)(tK)
            .loc = "1231"
            .area = "JB"
            .name = "Barinade"
        End With

        ' End with keys for both versions of Jabu-Jabu's Belly
        inc(tK)
        With aKeysDungeons(2)(tK)
            .loc = "7816"
            .area = "JB"
            .name = IIf(isMQ, "Boomerang Chest Room", "Lobby Basement Lower").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(2)(tK)
            .loc = "7817"
            .area = "JB"
            .name = IIf(isMQ, "Near Boss", "Lobby Basement Upper").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(2)(tK)
            .loc = "7818"
            .area = "JB"
            .name = IIf(isMQ, "Tailpasaran Room", "Near Boss").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(2)(tK)
            .loc = "7819"
            .area = "JB"
            .name = IIf(isMQ, "Invisible Enemies Room", "Water Switch Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(2)(tK)
            If Not isMQ Then
                ' Non-Master Quest has a Deku Scrub in it
                .loc = "8601"
                .area = "JB"
                .name = "Through Water Passage"
                .scrub = True
            Else
                ' Master Quest has a Cow in it
                .loc = "10224"
                .area = "JB"
                .name = "Wobble Room"
                .cow = True
            End If
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Jabu-Jabu's Belly (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysForestTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(3) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0


        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(3)(CInt(IIf(isMQ, 17, 18)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(3)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(3).Length - 1
            aKeysDungeons(3)(i) = New keyCheck
            With aKeysDungeons(3)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        ' Start off with keys for both versions of the Forest Temple
        With aKeysDungeons(3)(tK)
            .loc = "3403"
            .area = "FOT"
            .name = "First Room Chest"
        End With
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "3400"
            .area = "FOT"
            .name = IIf(isMQ, "Wolfos Chest", "First Stalfos Chest").ToString
        End With

        inc(tK)
        If Not isMQ Then
            ' The non-Master Quest keys for Forest Temple
            With aKeysDungeons(3)(tK)
                .loc = "3405"
                .area = "FOT"
                .name = "Raised Island Courtyard Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3401"
                .area = "FOT"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3409"
                .area = "FOT"
                .name = "Well Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3404"
                .area = "FOT"
                .name = "Eye Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3414"
                .area = "FOT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3402"
                .area = "FOT"
                .name = "Floormaster Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3413"
                .area = "FOT"
                .name = "Red Poe Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3412"
                .area = "FOT"
                .name = "Bow Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3415"
                .area = "FOT"
                .name = "Blue Poe Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3407"
                .area = "FOT"
                .name = "Falling Ceiling Room Chest"
            End With
        Else
            ' The Master Quest keys for Forest Temple
            With aKeysDungeons(3)(tK)
                .loc = "3409"
                .area = "FOT"
                .name = "Well Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3401"
                .area = "FOT"
                .name = "Raised Island Courtyard Lower Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3405"
                .area = "FOT"
                .name = "Raised Island Courtyard Upper Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3414"
                .area = "FOT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3402"
                .area = "FOT"
                .name = "Redead Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3413"
                .area = "FOT"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3412"
                .area = "FOT"
                .name = "Bow Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3415"
                .area = "FOT"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(3)(tK)
                .loc = "3406"
                .area = "FOT"
                .name = "Falling Ceiling Room Chest"
            End With
        End If

        ' End with keys for both versions of the Forest Temple
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "3411"
            .area = "FOT"
            .name = "Basement Chest"
        End With
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "1331"
            .area = "FOT"
            .name = "Phantom Ganon"
        End With
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "7824"
            .area = "FOT"
            .name = "Raised Island Courtyard"
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "7825"
            .area = "FOT"
            .name = IIf(isMQ, "First Hallway", "First Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "7826"
            .area = "FOT"
            .name = "Level Island Courtyard"
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "7827"
            .area = "FOT"
            .name = IIf(isMQ, "Well", "Lobby").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(3)(tK)
            .loc = "7828"
            .area = "FOT"
            .name = IIf(isMQ, "Block Push Room", "Basement").ToString
            .gs = True
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Forest Temple (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysFireTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(4) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(4)(CInt(IIf(isMQ, 17, 19)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(4)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(4).Length - 1
            aKeysDungeons(4)(i) = New keyCheck
            With aKeysDungeons(4)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Fire Temple
            With aKeysDungeons(4)(tK)
                .loc = "3501"
                .area = "FIT"
                .name = "Near Boss Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3500"
                .area = "FIT"
                .name = "Flare Dancer Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3512"
                .area = "FIT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3504"
                .area = "FIT"
                .name = "Big Lava Room Lower Open Door Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3502"
                .area = "FIT"
                .name = "Big Lava Room Blocked Door Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3503"
                .area = "FIT"
                .name = "Boulder Maze Lower Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3508"
                .area = "FIT"
                .name = "Boulder Maze Side Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3510"
                .area = "FIT"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3506"
                .area = "FIT"
                .name = "Boulder Maze Upper Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3511"
                .area = "FIT"
                .name = "Boulder Maze Shortcut Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3513"
                .area = "FIT"
                .name = "Scarecrow Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3507"
                .area = "FIT"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3509"
                .area = "FIT"
                .name = "Highest Goron Chest"
            End With
        Else
            ' The Master Quest keys for the Fire Temple
            With aKeysDungeons(4)(tK)
                .loc = "3507"
                .area = "FIT"
                .name = "Near Boss Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3502"
                .area = "FIT"
                .name = "Map Room Side Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3500"
                .area = "FIT"
                .name = "Megaton Hammer Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3512"
                .area = "FIT"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3501"
                .area = "FIT"
                .name = "Big Lava Room Blocked Door Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3504"
                .area = "FIT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3503"
                .area = "FIT"
                .name = "Lizalfos Maze Lower Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3508"
                .area = "FIT"
                .name = "Lizalfos Maze Side Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3506"
                .area = "FIT"
                .name = "Lizalfos Maze Upper Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "3511"
                .area = "FIT"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(4)(tK)
                .loc = "328"
                .area = "FIT"
                .name = "Under Platform Key"
            End With
        End If

        ' End with keys for both versions of the Fire Temple
        inc(tK)
        With aKeysDungeons(4)(tK)
            .loc = "3505"
            .area = "FIT"
            .name = IIf(isMQ, "Chest on Fire", "Megaton Hammer Chest").ToString
        End With
        inc(tK)
        With aKeysDungeons(4)(tK)
            .loc = "1431"
            .area = "FIT"
            .name = "Volvagia"
        End With
        inc(tK)
        With aKeysDungeons(4)(tK)
            .loc = "7900"
            .area = "FIT"
            .name = IIf(isMQ, "Big Lava Room Open Door", "Song of Time Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(4)(tK)
            .loc = "7901"
            .area = "FIT"
            .name = IIf(isMQ, "Above Fire Wall Maze", "Boss Key Loop").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(4)(tK)
            .loc = "7902"
            .area = "FIT"
            .name = IIf(isMQ, "Skull On Fire", "Boulder Maze").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(4)(tK)
            .loc = "7903"
            .area = "FIT"
            .name = IIf(isMQ, "Fire Wall Maze Centre", "Scarecrow Top").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(4)(tK)
            .loc = "7904"
            .area = "FIT"
            .name = IIf(isMQ, "Fire Wall Maze Side Room", "Scarecrow Climb").ToString
            .gs = True
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Fire Temple (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysWaterTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old arra
        aMQOld(5) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(5)(CInt(IIf(isMQ, 11, 15)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(5)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(5).Length - 1
            aKeysDungeons(5)(i) = New keyCheck
            With aKeysDungeons(5)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        ' Start off with keys for both versions of the Water Temple
        With aKeysDungeons(5)(tK)
            .loc = "3602"
            .area = "WAT"
            .name = "Map Chest"
        End With
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "3600"
            .area = "WAT"
            .name = IIf(isMQ, "Longshot Chest", "Cracked Wall Chest").ToString
        End With
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "3601"
            .area = "WAT"
            .name = IIf(isMQ, "Compass Chest", "Torches Chest").ToString
        End With

        inc(tK)
        If Not isMQ Then
            ' The non-Master Quest keys for the Water Temple
            With aKeysDungeons(5)(tK)
                .loc = "3609"
                .area = "WAT"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "3608"
                .area = "WAT"
                .name = "Central Bow Target Chest"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "3606"
                .area = "WAT"
                .name = "Central Pillar Chest"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "3607"
                .area = "WAT"
                .name = "Longshot Chest"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "3603"
                .area = "WAT"
                .name = "River Chear"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "3610"
                .area = "WAT"
                .name = "Dragon Chest"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "3605"
                .area = "WAT"
                .name = "Boss Key Chest"
            End With
        Else
            ' The Master Quest keys for the Water Temple
            With aKeysDungeons(5)(tK)
                .loc = "3606"
                .area = "WAT"
                .name = "Central Pillar Chest"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "3605"
                .area = "WAT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(5)(tK)
                .loc = "401"
                .area = "WAT"
                .name = "Stalfos Room Key"
            End With
        End If

        ' End with keys for both versions of the Water Temple
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "1531"
            .area = "WAT"
            .name = "Morpha"
        End With
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "7908"
            .area = "WAT"
            .name = IIf(isMQ, "Lizalfos Hallway", "Behind Gate").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "7909"
            .area = "WAT"
            .name = IIf(isMQ, "River", "Falling Platform Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "7910"
            .area = "WAT"
            .name = IIf(isMQ, "Before Upper Water Switch", "Central Pillar").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "7911"
            .area = "WAT"
            .name = IIf(isMQ, "Freestanding Key Area", "Near Boss Key Chest").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(5)(tK)
            .loc = "7912"
            .area = "WAT"
            .name = IIf(isMQ, "Triple Wall Torch", "River").ToString
            .gs = True
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Water Temple (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysSpiritTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(6) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0


        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(6)(CInt(IIf(isMQ, 27, 24)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(6)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(6).Length - 1
            aKeysDungeons(6)(i) = New keyCheck
            With aKeysDungeons(6)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the  Spirit Temple
            With aKeysDungeons(6)(tK)
                .loc = "3708"
                .area = "SPT"
                .name = "Child Bridge Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3700"
                .area = "SPT"
                .name = "Child Early Torches Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3706"
                .area = "SPT"
                .name = "Child Climb North Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3712"
                .area = "SPT"
                .name = "Child Climb East Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "5511"
                .area = "SPT"
                .name = "Desert Colossus: Silver Gauntlets Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3701"
                .area = "SPT"
                .name = "Sun Block Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3703"
                .area = "SPT"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3704"
                .area = "SPT"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3707"
                .area = "SPT"
                .name = "Early Adult Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3713"
                .area = "SPT"
                .name = "First Mirror First Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3714"
                .area = "SPT"
                .name = "First Mirror Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3702"
                .area = "SPT"
                .name = "Statue Room Hand Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3715"
                .area = "SPT"
                .name = "Statue Room Northeast Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3705"
                .area = "SPT"
                .name = "Near Four Armos Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3721"
                .area = "SPT"
                .name = "Hallway Invisible Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3720"
                .area = "SPT"
                .name = "Hallway Invisible Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "5509"
                .area = "SPT"
                .name = "Desert Colossus: Mirror Shield Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3710"
                .area = "SPT"
                .name = "Boss Key Chest"
            End With
        Else
            ' The Master Quest keys for the Spirit Temple
            With aKeysDungeons(6)(tK)
                .loc = "3726"
                .area = "SPT"
                .name = "Entrance Front Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3730"
                .area = "SPT"
                .name = "Entrance Back Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3731"
                .area = "SPT"
                .name = "Entrance Back Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3708"
                .area = "SPT"
                .name = "Map Room Enemy Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3700"
                .area = "SPT"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3729"
                .area = "SPT"
                .name = "Child Hammer Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3728"
                .area = "SPT"
                .name = "Silver Block Hallway Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3703"
                .area = "SPT"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3701"
                .area = "SPT"
                .name = "Sun Block Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "5511"
                .area = "SPT"
                .name = "Desert Colossus: Silver Gauntlets Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3706"
                .area = "SPT"
                .name = "Child Climb North Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3712"
                .area = "SPT"
                .name = "Child Climb South Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3715"
                .area = "SPT"
                .name = "Statue Room Lullaby Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3702"
                .area = "SPT"
                .name = "Statue Room Invisible Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3704"
                .area = "SPT"
                .name = "Leever Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3707"
                .area = "SPT"
                .name = "Symphony Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3727"
                .area = "SPT"
                .name = "Entrance Front Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3725"
                .area = "SPT"
                .name = "Beamos Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3724"
                .area = "SPT"
                .name = "Chest Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "3705"
                .area = "SPT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(6)(tK)
                .loc = "5509"
                .area = "SPT"
                .name = "Desert Colossus: Mirror Shield Chest"
            End With
        End If

        ' End with keys for both versions of the Spirit Temple
        inc(tK)
        With aKeysDungeons(6)(tK)
            .loc = "3718"
            .area = "SPT"
            .name = IIf(isMQ, "Mirror Puzzle Invisible Chest", "Topmost Chest").ToString
        End With
        inc(tK)
        With aKeysDungeons(6)(tK)
            .loc = "1631"
            .area = "SPT"
            .name = "Twinrova"
        End With
        inc(tK)
        With aKeysDungeons(6)(tK)
            .loc = "7916"
            .area = "SPT"
            .name = IIf(isMQ, "Sun Block Room", "Hall After Sun Block Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(6)(tK)
            .loc = "7917"
            .area = "SPT"
            .name = IIf(isMQ, "Leever Room", "Boulder Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(6)(tK)
            .loc = "7918"
            .area = "SPT"
            .name = IIf(isMQ, "Nine Thrones Room West", "Lobby").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(6)(tK)
            .loc = "7919"
            .area = "SPT"
            .name = IIf(isMQ, "Symphony Room", "Sun on Floor Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(6)(tK)
            .loc = "7920"
            .area = "SPT"
            .name = IIf(isMQ, "Nine Thrones Room North", "Metal Fence").ToString
            .gs = True
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Spirit Temple (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysShadowTemple(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(7) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0


        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(7)(CInt(IIf(isMQ, 25, 22)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(7)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(7).Length - 1
            aKeysDungeons(7)(i) = New keyCheck
            With aKeysDungeons(7)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        ' Start off with keys for both versions of the Shadow Temple
        With aKeysDungeons(7)(tK)
            .loc = "3801"
            .area = "SHT"
            .name = IIf(isMQ, "Compass Chest", "Map Chest").ToString
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "3807"
            .area = "SHT"
            .name = "Hover Boots Chest"
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "3803"
            .area = "SHT"
            .name = IIf(isMQ, "Early Gibdos Chest", "Compass Chest").ToString
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "3802"
            .area = "SHT"
            .name = IIf(isMQ, "Map Chest", "Early Silver Rupee Chest").ToString
        End With

        inc(tK)
        If Not isMQ Then
            ' The non-Master Quest keys for the Shadow Temple
            With aKeysDungeons(7)(tK)
                .loc = "3812"
                .area = "SHT"
                .name = "Invisible Blades Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3822"
                .area = "SHT"
                .name = "Invisible Blades Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3805"
                .area = "SHT"
                .name = "Falling Spikes Lower Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3806"
                .area = "SHT"
                .name = "Falling Spikes Upper First Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3804"
                .area = "SHT"
                .name = "Falling Spikes Upper Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3809"
                .area = "SHT"
                .name = "Invisible Spikes Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "501"
                .area = "SHT"
                .name = "Skull Pot Key"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3821"
                .area = "SHT"
                .name = "Wind Hint Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3820"
                .area = "SHT"
                .name = "After Wind First Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3808"
                .area = "SHT"
                .name = "After Wind Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3811"
                .area = "SHT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3810"
                .area = "SHT"
                .name = "Spike Walls Left Chest"
            End With
        Else
            ' The Master Quest keys for the Shadow Temple
            With aKeysDungeons(7)(tK)
                .loc = "3814"
                .area = "SHT"
                .name = "Near Ship Invisible Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3812"
                .area = "SHT"
                .name = "Invisible Blades Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3822"
                .area = "SHT"
                .name = "Invisible Blades Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3815"
                .area = "SHT"
                .name = "Beamos Silver Rupee Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3805"
                .area = "SHT"
                .name = "Falling Spikes Lower Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3806"
                .area = "SHT"
                .name = "Falling Spikes Upper First Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3804"
                .area = "SHT"
                .name = "Falling Spikes Upper Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3809"
                .area = "SHT"
                .name = "Invisible Spikes Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3816"
                .area = "SHT"
                .name = "Stalfos Room Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3821"
                .area = "SHT"
                .name = "Wind Hint Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3820"
                .area = "SHT"
                .name = "After Wind First Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3808"
                .area = "SHT"
                .name = "After Wind Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3811"
                .area = "SHT"
                .name = "Boss Key Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "3810"
                .area = "SHT"
                .name = "Spike Walls Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(7)(tK)
                .loc = "506"
                .area = "SHT"
                .name = "Skull Pots Room Key"
            End With
        End If

        ' End with keys for both versions of the Shadow Temple
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "3813"
            .area = "SHT"
            .name = "Invisible Floormaster Chest"
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "1731"
            .area = "SHT"
            .name = "Bongo Bongo"
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "7924"
            .area = "SHT"
            .name = IIf(isMQ, "Wind Hint Room", "Single Giant Pot").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "7925"
            .area = "SHT"
            .name = "Falling Spikes Room"
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "7926"
            .area = "SHT"
            .name = IIf(isMQ, "Near Boss", "Triple Giant Pot").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "7927"
            .area = "SHT"
            .name = IIf(isMQ, "After Wind", "Like Like Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(7)(tK)
            .loc = "7928"
            .area = "SHT"
            .name = IIf(isMQ, "After Ship", "Near Ship").ToString
            .gs = True
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Shadow Temple (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysBottomOfTheWell(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(8) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(8)(CInt(IIf(isMQ, 7, 16)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(8)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(8).Length - 1
            aKeysDungeons(8)(i) = New keyCheck
            With aKeysDungeons(8)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Bottom of the Well
            With aKeysDungeons(8)(tK)
                .loc = "3908"
                .area = "BW"
                .name = "Front Left Fake Wall Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3905"
                .area = "BW"
                .name = "Right Bottom Fake Wall Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3901"
                .area = "BW"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3914"
                .area = "BW"
                .name = "Centre Skulltula Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3910"
                .area = "BW"
                .name = "Fire Keese Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3912"
                .area = "BW"
                .name = "Like Like Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "601"
                .area = "BW"
                .name = "Coffin Key"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3902"
                .area = "BW"
                .name = "Front Centre Bombable Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3904"
                .area = "BW"
                .name = "Back Left Bombable Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3907"
                .area = "BW"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3909"
                .area = "BW"
                .name = "Underwater Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3916"
                .area = "BW"
                .name = "Underwater Front Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3903"
                .area = "BW"
                .name = "Lens of Truth Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3920"
                .area = "BW"
                .name = "Invisible Chest"
            End With
        Else
            ' The Master Quest keys for the Bottom of the Well
            With aKeysDungeons(8)(tK)
                .loc = "3903"
                .area = "BW"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "601"
                .area = "BW"
                .name = "East Inner Room Key"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3902"
                .area = "BW"
                .name = "Compass Chest"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "602"
                .area = "BW"
                .name = "Dead Hand Key"
            End With
            inc(tK)
            With aKeysDungeons(8)(tK)
                .loc = "3901"
                .area = "BW"
                .name = "Lens of Truth Chest"
            End With
        End If

        ' Keys for both versions of the Bottom of the Well
        inc(tK)
        With aKeysDungeons(8)(tK)
            .loc = "8000"
            .area = "BW"
            .name = IIf(isMQ, "Basement", "Like Like Cage").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(8)(tK)
            .loc = "8001"
            .area = "BW"
            .name = IIf(isMQ, "West Inner Room", "East Inner Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(8)(tK)
            .loc = "8002"
            .area = "BW"
            .name = IIf(isMQ, "Coffin Room", "West Inner Room").ToString
            .gs = True
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Bottom of the Well (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysIceCavern(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(9) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys, nice that they are the same for both Master Quest and not
        ReDim aKeysDungeons(9)(7)

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(9)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(9).Length - 1
            aKeysDungeons(9)(i) = New keyCheck
            With aKeysDungeons(9)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Ice Cavern
            With aKeysDungeons(9)(tK)
                .loc = "4000"
                .area = "IC"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(9)(tK)
                .loc = "4001"
                .area = "IC"
                .name = "Compass Chest"
            End With
        Else
            ' The Master Quest keys for the Ice Carvern
            With aKeysDungeons(9)(tK)
                .loc = "4001"
                .area = "IC"
                .name = "Map Chest"
            End With
            inc(tK)
            With aKeysDungeons(9)(tK)
                .loc = "4000"
                .area = "IC"
                .name = "Compass Chest"
            End With
        End If

        ' Keys for both versions of the Ice Cavern
        inc(tK)
        With aKeysDungeons(9)(tK)
            .loc = "701"
            .area = "IC"
            .name = "Piece of Heart"
        End With
        inc(tK)
        With aKeysDungeons(9)(tK)
            .loc = "4002"
            .area = "IC"
            .name = "Iron Boots Chest"
        End With
        inc(tK)
        With aKeysDungeons(9)(tK)
            .loc = "6402"
            .area = "IC"
            .name = "Song from Sheik"
        End With
        inc(tK)
        With aKeysDungeons(9)(tK)
            .loc = "8008"
            .area = "IC"
            .name = IIf(isMQ, "Scarecrow", "Push Block Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(9)(tK)
            .loc = "8009"
            .area = "IC"
            .name = IIf(isMQ, "Red Ice", "Spinning Scythe Room").ToString
            .gs = True
        End With
        inc(tK)
        With aKeysDungeons(9)(tK)
            .loc = "8010"
            .area = "IC"
            .name = IIf(isMQ, "Ice Block", "Heart Piece Room").ToString
            .gs = True
        End With

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Ice Cavern (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysGerudoTrainingGround(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(10) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(10)(CInt(IIf(isMQ, 16, 21)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(10)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(10).Length - 1
            aKeysDungeons(10)(i) = New keyCheck
            With aKeysDungeons(10)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next

        If Not isMQ Then
            ' The non-Master Quest keys for the Gerudo Training Ground
            With aKeysDungeons(10)(tK)
                .loc = "4219"
                .area = "GTG"
                .name = "Lobby Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4207"
                .area = "GTG"
                .name = "Lobby Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4201"
                .area = "GTG"
                .name = "Beamos Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4200"
                .area = "GTG"
                .name = "Stalfos Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4217"
                .area = "GTG"
                .name = "Before Heavy Block Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4215"
                .area = "GTG"
                .name = "Heavy Block First Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4214"
                .area = "GTG"
                .name = "Heavy Block Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4220"
                .area = "GTG"
                .name = "Heaby Block Third Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4202"
                .area = "GTG"
                .name = "Heavy Block Fourth Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4203"
                .area = "GTG"
                .name = "Eye Statue Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4204"
                .area = "GTG"
                .name = "Near Scarecrow Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4218"
                .area = "GTG"
                .name = "Hammer Room Clear Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4216"
                .area = "GTG"
                .name = "Hammer Room Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "801"
                .area = "GTG"
                .name = "Lava Room Key"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4205"
                .area = "GTG"
                .name = "Maze Right First Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4208"
                .area = "GTG"
                .name = "Maze Right Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4213"
                .area = "GTG"
                .name = "Underwater Silver Rupee Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4211"
                .area = "GTG"
                .name = "Hidden Ceiling Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4206"
                .area = "GTG"
                .name = "Maze Path First Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4210"
                .area = "GTG"
                .name = "Maze Path Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4209"
                .area = "GTG"
                .name = "Maze Path Third Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4212"
                .area = "GTG"
                .name = "Maze Path Final Chest"
            End With
        Else
            ' The Master Quest keys for the Gerudo Training Ground
            With aKeysDungeons(10)(tK)
                .loc = "4219"
                .area = "GTG"
                .name = "Lobby Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4207"
                .area = "GTG"
                .name = "Lobby Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4200"
                .area = "GTG"
                .name = "First Iron Knuckle Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4217"
                .area = "GTG"
                .name = "Before Heavy Block Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4202"
                .area = "GTG"
                .name = "Heavy Block Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4203"
                .area = "GTG"
                .name = "Eye Statue Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4218"
                .area = "GTG"
                .name = "Second Iron Knuckle Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4214"
                .area = "GTG"
                .name = "Flame Circle Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4205"
                .area = "GTG"
                .name = "Maze Right First Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4208"
                .area = "GTG"
                .name = "Maze Right Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4201"
                .area = "GTG"
                .name = "Dinolfos Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4213"
                .area = "GTG"
                .name = "Underwater Silver Rupee Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4211"
                .area = "GTG"
                .name = "Hidden Ceiling Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4206"
                .area = "GTG"
                .name = "Maze Path First Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4209"
                .area = "GTG"
                .name = "Maze Path Third Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4210"
                .area = "GTG"
                .name = "Maze path Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(10)(tK)
                .loc = "4204"
                .area = "GTG"
                .name = "Ice Arrows Chest"
            End With
        End If

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Gerudo Training Ground (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
    End Sub
    Private Sub makeKeysGanonsCastle(Optional isMQ As Boolean = False)
        ' Store Master Quest setting into old array
        aMQOld(11) = isMQ

        ' Varaible used for counting up the entries
        Dim tK As Integer = 0

        ' Resize the array for the dungeon keys based on MQ setting
        ReDim aKeysDungeons(11)(CInt(IIf(isMQ, 18, 19)))

        ' If the debug option for counting is on, make array large to avoid errors
        If DO_COUNT Then ReDim aKeysDungeons(11)(100)

        ' Make sure each key has a base template
        For i = 0 To aKeysDungeons(11).Length - 1
            aKeysDungeons(11)(i) = New keyCheck
            With aKeysDungeons(11)(i)
                .loc = String.Empty
                .area = String.Empty
                .name = String.Empty
                .scan = True
                .checked = False
                .gs = False
                .cow = False
                .scrub = False
                .shop = False
                .forced = False
            End With
        Next


        If Not isMQ Then
            ' The non-Master Quest keys for Ganon's Castle
            With aKeysDungeons(11)(tK)
                .loc = "4309"
                .area = "IGC"
                .name = "Forest Trial Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4307"
                .area = "IGC"
                .name = "Water Trial Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4306"
                .area = "IGC"
                .name = "Water Trial Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4308"
                .area = "IGC"
                .name = "Shadow Trial Front Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4305"
                .area = "IGC"
                .name = "Shadow Trial Golden Gauntlets Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4312"
                .area = "IGC"
                .name = "Light Trial First Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4311"
                .area = "IGC"
                .name = "Light Trial Second Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4313"
                .area = "IGC"
                .name = "Light Trial Third Left Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4314"
                .area = "IGC"
                .name = "Light Trial First Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4310"
                .area = "IGC"
                .name = "Light Trial Second Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4315"
                .area = "IGC"
                .name = "Light Trial Third Right Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4316"
                .area = "IGC"
                .name = "Light Trial Invisible Enemies Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4317"
                .area = "IGC"
                .name = "Light Trail Lullaby Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4318"
                .area = "IGC"
                .name = "Spirit Trial Crystal Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4320"
                .area = "IGC"
                .name = "Spirit Trial Invisible Chest"
            End With
        Else
            ' The Master Quest keys for Ganon's Castle
            With aKeysDungeons(11)(tK)
                .loc = "901"
                .area = "IGC"
                .name = "Forest Trial Stalfos Room Key"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4302"
                .area = "IGC"
                .name = "Forest Trial Eye Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4303"
                .area = "IGC"
                .name = "Forest Trial Frozen Eye Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4301"
                .area = "IGC"
                .name = "Water Trial Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4300"
                .area = "IGC"
                .name = "Shadow Trial Bomb Flower Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4305"
                .area = "IGC"
                .name = "Shadow Trial Eye Switch Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4304"
                .area = "IGC"
                .name = "Light Trial Lullaby Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4310"
                .area = "IGC"
                .name = "Spirit Trial Front Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4320"
                .area = "IGC"
                .name = "Spirit Trial Invisible Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4306"
                .area = "IGC"
                .name = "Spirit Trial Golden Gauntlets Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4307"
                .area = "IGC"
                .name = "Spirit Trial Sun First Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4308"
                .area = "IGC"
                .name = "Spirit Trial Sun Second Chest"
            End With
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "4309"
                .area = "IGC"
                .name = "Spirit Trial Sun Third Chest"
            End With
        End If

        ' Keys for both versions of Ganon's Castle
        inc(tK)
        With aKeysDungeons(11)(tK)
            .loc = "4111"
            .area = "IGC"
            .name = "Ganon’s Tower Boss Key Chest"
        End With
        inc(tK)
        With aKeysDungeons(11)(tK)
            .loc = "8709"
            .area = "IGC"
            .name = "Secret Room Left"
            .scrub = True
        End With
        inc(tK)
        With aKeysDungeons(11)(tK)
            .loc = "8706"
            .area = "IGC"
            .name = "Secret Room Centre Left"
            .scrub = True
        End With
        inc(tK)
        With aKeysDungeons(11)(tK)
            .loc = "8704"
            .area = "IGC"
            .name = "Secret Room Centre Right"
            .scrub = True
        End With
        inc(tK)
        With aKeysDungeons(11)(tK)
            .loc = "8708"
            .area = "IGC"
            .name = "Secret Room Right"
            .scrub = True
        End With
        If isMQ Then
            inc(tK)
            With aKeysDungeons(11)(tK)
                .loc = "8701"
                .area = "IGC"
                .name = "Secret Room Front Right"
                .scrub = True
            End With
        End If

        ' If debug option is on, output the count
        If DO_COUNT Then rtbOutput.AppendText("Ganon's Castle (" & isMQ.ToString & "): " & tK.ToString & vbCrLf)
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
    Private Sub lblQuestBigPoes_MouseClick(sender As Object, e As MouseEventArgs) Handles lblQuestBigPoes.MouseClick
        displayChecks("QBPH", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlQuestBigPoes_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlQuestBigPoes.MouseClick
        displayChecks("QBPH", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub lblQuestFrogs_MouseClick(sender As Object, e As MouseEventArgs) Handles lblQuestFrogs.MouseClick
        displayChecks("QF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
    End Sub
    Private Sub pnlQuestFrogs_MouseClick(sender As Object, e As MouseEventArgs) Handles pnlQuestFrogs.MouseClick
        displayChecks("QF", CBool(IIf(e.Button = Windows.Forms.MouseButtons.Right, True, False)))
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

    Private Sub inc(ByRef value As Integer, Optional ByRef amount As Integer = 1)
        ' Small sub for shorthand increments
        value = value + amount
    End Sub

    Private Sub checkCarpenters()
        ' Checks that all 4 carpenters have been saved
        Dim count As Integer = 0

        ' Search for the 4 carpenter checks
        For Each key In aKeys
            With key
                Select Case .loc
                    ' For each check, increase the counter
                    Case "6500", "6501", "6502", "6503"
                        If .checked Then inc(count)
                End Select
            End With
        Next

        ' Find the key for the Membership Card
        For Each key In aKeys
            With key
                Select Case .loc
                    Case "CARD"
                        ' Set to false unless all 4 carpenter checks were found
                        .checked = CBool(IIf(count < 4, False, True))
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
                    .Visible = True
                Case 2
                    .Image = My.Resources.upgradeQuiver2
                    .Visible = True
                Case 3
                    .Image = My.Resources.upgradeQuiver3
                    .Visible = True
                Case Else
                    .Visible = False
            End Select
        End With

        With pbxBombBag
            Select Case upgrades(1)
                Case 1
                    .Image = My.Resources.upgradeBombBag1
                    .Visible = True
                Case 2
                    .Image = My.Resources.upgradeBombBag2
                    .Visible = True
                Case 3
                    .Image = My.Resources.upgradeBombBag3
                    .Visible = True
                Case Else
                    .Visible = False
            End Select
        End With

        With pbxGauntlet
            Select Case upgrades(2)
                Case 1
                    .Image = My.Resources.upgradeGoronsBracelet
                    .Visible = True
                Case 2
                    .Image = My.Resources.upgradeSilverGauntlets
                    .Visible = True
                Case 3
                    .Image = My.Resources.upgradeGoldenGauntlets
                    .Visible = True
                Case Else
                    .Visible = False
            End Select
        End With

        With pbxScale
            Select Case upgrades(3)
                Case 1
                    .Image = My.Resources.upgradeSilverScale
                    .Visible = True
                Case 2
                    .Image = My.Resources.upgradeGoldenScale
                    .Visible = True
                Case Else
                    .Visible = False
            End Select
        End With

        With pbxWallet
            Select Case upgrades(4)
                Case 1
                    .Image = My.Resources.upgradeWallet1
                    .Visible = True
                Case 2
                    .Image = My.Resources.upgradeWallet2
                    .Visible = True
                Case Else
                    .Visible = False
            End Select
        End With

        With pbxBulletBag
            Select Case upgrades(5)
                Case 1
                    .Image = My.Resources.upgradeBulletBag1
                    .Visible = True
                Case 2
                    .Image = My.Resources.upgradeBulletBag2
                    .Visible = True
                Case 3
                    .Image = My.Resources.upgradeBulletBag3
                    .Visible = True
                Case Else
                    .Visible = False
            End Select
        End With
    End Sub
    Private Sub checkQuestItems()
        pbxMedalForest.Visible = cb7700.Checked
        pbxMedalFire.Visible = cb7701.Checked
        pbxMedalWater.Visible = cb7702.Checked
        pbxMedalSpirit.Visible = cb7703.Checked
        pbxMedalShadow.Visible = cb7704.Checked
        pbxMedalLight.Visible = cb7705.Checked
        pbxMinuetOfForest.Visible = cb7706.Checked
        pbxBoleroOfFire.Visible = cb7707.Checked
        pbxSerenadeOfWater.Visible = cb7708.Checked
        pbxRequiemOfSpirit.Visible = cb7709.Checked
        pbxNocturneOfShadow.Visible = cb7710.Checked
        pbxPreludeOfLight.Visible = cb7711.Checked
        pbxZeldasLullaby.Visible = cb7712.Checked
        pbxEponasSong.Visible = cb7713.Checked
        pbxSaraisSong.Visible = cb7714.Checked
        pbxSunsSong.Visible = cb7715.Checked
        pbxSongOfTime.Visible = cb7716.Checked
        pbxSongOfStorms.Visible = cb7717.Checked
        pbxStoneKokiri.Visible = cb7718.Checked
        pbxStoneGoron.Visible = cb7719.Checked
        pbxStoneZora.Visible = cb7720.Checked
        pbxStoneOfAgony.Visible = cb7721.Checked
        pbxGerudosCard.Visible = cb7722.Checked
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
                rtbOutput.AppendText("Attachment Problem: " & ex.Message & vbCrLf)
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

    Private Sub attachToM64P()
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
                rtbOutput.AppendText("Attachment Problem: " & ex.Message & vbCrLf)
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
        addressDLL = addressDLL + &H29C95D8

        ' Attach to process and set it as the current emulator
        SetProcessName("mupen64plus-gui")
        emulator = "mupen64plus-gui"

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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'updateItems()
        'Exit Sub

        'attachToBizHawk()
        If emulator = "emuhawk" Or emulator = "rmg" Or emulator = "mupen64plus-gui" Then
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
        For Each pbx In pnlSongsMedals.Controls.OfType(Of PictureBox)()
            pbx.Visible = False
        Next
    End Sub
    Private Sub updateItems()
        If emulator = String.Empty Then Exit Sub
        If zeldaCheck() = False Then Exit Sub
        Dim items As String = String.Empty
        Dim temp As String = String.Empty
        Dim cTemp As Integer = 0
        Dim inventory(23) As PictureBox
        inventory(0) = pbx01
        inventory(1) = pbx02
        inventory(2) = pbx03
        inventory(3) = pbx04
        inventory(4) = pbx05
        inventory(5) = pbx06
        inventory(6) = pbx07
        inventory(7) = pbx08
        inventory(8) = pbx09
        inventory(9) = pbx10
        inventory(10) = pbx11
        inventory(11) = pbx12
        inventory(12) = pbx13
        inventory(13) = pbx14
        inventory(14) = pbx15
        inventory(15) = pbx16
        inventory(16) = pbx17
        inventory(17) = pbx18
        inventory(18) = pbx19
        inventory(19) = pbx20
        inventory(20) = pbx21
        inventory(21) = pbx22
        inventory(22) = pbx23
        inventory(23) = pbx24

        For i = &H11A644 To &H11A658 Step 4
            temp = Hex(goRead(i))
            While temp.Length < 8
                temp = "0" & temp
            End While
            items = items & temp
        Next

        For i = 0 To 23 ' 1 To items.Length Step 2
            temp = Mid(items, (i * 2) + 1, 2)
            cTemp = CInt("&H" & temp)
            With inventory(i)
                Select Case cTemp
                    Case 0 To 55
                        If Not .Visible Then .Visible = True
                    Case Else
                        .Visible = False
                End Select
                Select Case cTemp
                    Case 0
                        .Image = My.Resources.dekuStick
                    Case 1
                        .Image = My.Resources.dekuNut
                    Case 2
                        .Image = My.Resources.bombs
                    Case 3
                        .Image = My.Resources.bow
                    Case 4
                        .Image = My.Resources.fireArrow
                    Case 5
                        .Image = My.Resources.dinsFire
                    Case 6
                        .Image = My.Resources.fairySlingshot
                    Case 7
                        .Image = My.Resources.fairyOcarina
                    Case 8
                        .Image = My.Resources.ocarinaOfTime
                    Case 9
                        .Image = My.Resources.bombchus
                    Case 10
                        .Image = My.Resources.hookshot
                    Case 11
                        .Image = My.Resources.longshot
                    Case 12
                        .Image = My.Resources.iceArrows
                    Case 13
                        .Image = My.Resources.faroresWind
                    Case 14
                        .Image = My.Resources.boomerang
                    Case 15
                        .Image = My.Resources.lensOfTruth
                    Case 16
                        .Image = My.Resources.magicBeans
                    Case 17
                        .Image = My.Resources.megatonHammer
                    Case 18
                        .Image = My.Resources.lightArrows
                    Case 19
                        .Image = My.Resources.naryusLove
                    Case 20
                        .Image = My.Resources.bottleEmpty
                    Case 21
                        .Image = My.Resources.bottleRedPotion
                    Case 22
                        .Image = My.Resources.bottleGreenPotion
                    Case 23
                        .Image = My.Resources.bottleBluePotion
                    Case 24
                        .Image = My.Resources.bottleBottledFairy
                    Case 25
                        .Image = My.Resources.bottleFish
                    Case 26
                        .Image = My.Resources.bottleLonLonMilk
                    Case 27
                        .Image = My.Resources.bottleLetter
                    Case 28
                        .Image = My.Resources.bottleBlueFire
                    Case 29
                        .Image = My.Resources.bottleBug
                    Case 30
                        .Image = My.Resources.bottleBigPoe
                    Case 31
                        .Image = My.Resources.bottleLonLonMilkHalf
                    Case 32
                        .Image = My.Resources.bottlePoe
                    Case 33
                        .Image = My.Resources.youngWeirdEgg
                    Case 34
                        .Image = My.Resources.youngChicken
                    Case 35
                        .Image = My.Resources.youngZeldasLetter
                    Case 36
                        .Image = My.Resources.youngKeatonMask
                    Case 37
                        .Image = My.Resources.youngSkullMask
                    Case 38
                        .Image = My.Resources.youngSpookyMask
                    Case 39
                        .Image = My.Resources.youngBunnyHood
                    Case 40
                        .Image = My.Resources.youngGoronMask
                    Case 41
                        .Image = My.Resources.youngZoraMask
                    Case 42
                        .Image = My.Resources.youngGerudoMask
                    Case 43
                        .Image = My.Resources.youngMaskOfTruth
                    Case 44
                        .Image = My.Resources.youngSoldOut
                    Case 45
                        .Image = My.Resources.adultPocketEgg
                    Case 46
                        .Image = My.Resources.adultPocketCucco
                    Case 47
                        .Image = My.Resources.adultCojiro
                    Case 48
                        .Image = My.Resources.adultOddMushroom
                    Case 49
                        .Image = My.Resources.adultOddPotion
                    Case 50
                        .Image = My.Resources.adultPoachersSaw
                    Case 51
                        .Image = My.Resources.adultGoronsSwordBroken
                    Case 52
                        .Image = My.Resources.adultPrescription
                    Case 53
                        .Image = My.Resources.adultEyeballFrog
                    Case 54
                        .Image = My.Resources.adultEyeDrops
                    Case 55
                        .Image = My.Resources.adultClaimCheck
                    Case Else
                        .Image = My.Resources.emptySlot
                End Select
            End With
        Next
    End Sub

    Private Sub rtbOutput_KeyDown(sender As Object, e As KeyEventArgs) Handles rtbOutput.KeyDown
        ' Do not want to disable key inputs, as we want scrolling to work, so just supress keys
        e.SuppressKeyPress = True
    End Sub

    Private Sub rtbOutput_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles rtbOutput.MouseDoubleClick
        ' Get the double-clicked location
        Dim clickPos As Integer = rtbOutput.GetCharIndexFromPosition(e.Location)
        ' Get the line that location was located on
        Dim linePos As Integer = rtbOutput.GetLineFromCharIndex(clickPos)
        ' Read the line into a string
        Dim readLine As String = rtbOutput.Lines(linePos).ToString
        ' Empty string to store the area name into
        Dim readArea As String = String.Empty

        ' Step backwards from current line to the very beginning
        For i = linePos To 0 Step -1
            ' Read the current line's text
            readArea = rtbOutput.Lines(i).ToString
            ' All lines start with 2 spaces except the area names, stop on the first one found
            If Not Mid(readArea, 1, 2) = "  " Then Exit For
        Next
        ' If the found line was the same as the double-clicked line, or none was found, then exit
        If readArea = readLine Or readArea = String.Empty Then Exit Sub
        ' Remove any added things and trim up the area name
        readArea = Trim(readArea.Replace(":", "").Replace("MQ", "").Replace("(Found)", ""))
        ' Convert area name into the code used
        Dim areaCode As String = area2code(readArea)
        If getKeyInfo(Trim(readLine), areaCode) Then
            Dim newLine As String = readLine
            updateLabels()
            updateLabelsDungeons()
            btnFocus.Focus()
            If newLine.Contains("(Forced)") Then
                newLine = newLine.Replace(" (Forced)", "")
            Else
                newLine = newLine & " (Forced)"
            End If
            replaceRTB(linePos, newLine)
        End If
        'MsgBox(getKeyInfo(Trim(readLine), areaCode))
    End Sub
    Private Sub rtbOutput_TextChanged(sender As Object, e As EventArgs) Handles rtbOutput.TextChanged
        ' Whenever text is changed, scroll down to the bottom
        If doNotScroll Then
            doNotScroll = False
        Else
            With rtbOutput
                .SelectionStart = .TextLength
                .ScrollToCaret()
            End With
        End If

    End Sub
    Private Sub replaceRTB(ByVal line As Integer, ByVal newText As String)
        'Dim storeStart As Integer = rtbOutput.SelectionStart
        doNotScroll = True
        Dim newRTB As String = String.Empty
        For i = 0 To rtbOutput.Lines.Count - 1
            If i > 0 Then newRTB = newRTB & vbCrLf
            If Not i = line Then
                newRTB = newRTB & rtbOutput.Lines(i)
            Else
                newRTB = newRTB & newText
            End If
        Next
        rtbOutput.Text = newRTB
        'rtbOutput.SelectionStart = storeStart
    End Sub

    Private Sub drawGraphics()
        Dim art As Graphics
        Dim pnFore As Pen = New Pen(Me.ForeColor, 1)

        For Each pnl In Me.Controls.OfType(Of Panel)()
            With pnl
                If .Location.Y > 200 Then
                    .Refresh()
                    art = .CreateGraphics
                    art.DrawRectangle(pnFore, 0, 0, .Width - 1, .Height - 1)
                End If
            End With
        Next
        art = Me.CreateGraphics
        With rtbOutput
            .Refresh()
            art.DrawRectangle(pnFore, .Location.X - 1, .Location.Y - 1, .Width + 1, .Height + 1)
        End With

        For Each lbl In Me.Controls.OfType(Of Label)()
            With lbl
                .Refresh()
                .BackColor = Me.BackColor
                .ForeColor = Me.ForeColor
                If .TextAlign = ContentAlignment.MiddleCenter Then art.DrawRectangle(pnFore, .Location.X - 1, .Location.Y - 1, .Width + 1, .Height + 1)
            End With
        Next

        With lblHideScroll
            .Refresh()
            art = .CreateGraphics
            art.DrawLine(pnFore, 0, 0, 0, .Height - 1)
        End With
        updateLCX()
    End Sub

    Private Sub outputSong(ByVal title As String, ByVal notes As String)
        Dim outSong As String = title & ":" & vbCrLf & "  "
        If rtbOutput.TextLength > 0 Then outSong = vbCrLf & vbCrLf & outSong
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
        rtbOutput.AppendText(outSong)
    End Sub
    Private Sub pbxZeldasLullaby_Click(sender As Object, e As EventArgs) Handles pbxZeldasLullaby.Click
        outputSong("Zelda's Lullaby", "LURLUR")
    End Sub
    Private Sub pbxEponasSong_Click(sender As Object, e As EventArgs) Handles pbxEponasSong.Click
        outputSong("Epona's Song", "ULRULR")
    End Sub
    Private Sub pbxSaraisSong_Click(sender As Object, e As EventArgs) Handles pbxSaraisSong.Click
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

    Private Sub frmTrackerOfTime_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        btnFocus.Focus()
    End Sub
    Private Sub frmTrackerOfTime_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        drawGraphics()
    End Sub
    
    Private Sub lcxShowSkulltulas_Click(sender As Object, e As EventArgs) Handles lcxShowSkulltulas.Click
        ' Flip the Gold Skulltulas setting, save settings, then update labels and graphics
        My.Settings.setSkulltula = Not My.Settings.setSkulltula
        My.Settings.Save()
        updateLabels()
        updateLabelsDungeons()
        drawGraphics()
    End Sub

    Private Sub lcxCowShuffle_Click(sender As Object, e As EventArgs) Handles lcxCowShuffle.Click
        ' Flip the Cow Shuffle setting, save settings, then update labels and graphics
        My.Settings.setCow = Not My.Settings.setCow
        My.Settings.Save()
        updateLabels()
        updateLabelsDungeons()
        drawGraphics()
    End Sub

    Private Sub lcxScrubShuffle_Click(sender As Object, e As EventArgs) Handles lcxScrubShuffle.Click
        ' Flip the Deku Scrub Shuffle setting and save settings
        My.Settings.setScrub = Not My.Settings.setScrub
        My.Settings.Save()

        ' Variable for counting how many of the 3 Scrubs have been found
        Dim found As Integer = 0

        ' Step through each entry in aKeys and search for the 3 Scrubs
        For Each key In aKeys
            Select Case key.loc
                ' The 3 .loc codes for the 3 Scrubs
                Case "6827", "7202", "7203"
                    ' Add up a found, and set them to a flipped Scrub settings
                    found = found + 1
                    key.scan = Not My.Settings.setScrub
            End Select
            ' On 3 founds, quit the loop
            If found >= 3 Then Exit For
        Next

        ' Update labels and graphics
        updateLabels()
        updateLabelsDungeons()
        drawGraphics()
    End Sub

    Private Sub lcxShopsanity_Click(sender As Object, e As EventArgs) Handles lcxShopsanity.Click
        ' Flip the Shopsanity setting, save settings, then update labels and graphics
        My.Settings.setShop = Not My.Settings.setShop
        My.Settings.Save()
        updateLabels()
        updateLabelsDungeons()
        drawGraphics()
    End Sub


    Private Sub subMenu(ByVal strTheme As String)
        Dim valTheme As Byte = 0

        Select Case LCase(strTheme)
            Case "light mode"
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
                rtbOutput.AppendText("-- Theme Error: " & strTheme & vbCrLf)
                Exit Sub
        End Select

        subMenuCheck(valTheme)
        changeTheme(valTheme)
    End Sub

    Private Sub subMenuCheck(ByVal valTheme As Byte)
        LightModeToolStripMenuItem.Checked = False
        DarkModeToolStripMenuItem.Checked = False
        LavenderToolStripMenuItem.Checked = False
        MidnightToolStripMenuItem.Checked = False
        HotdogStandToolStripMenuItem.Checked = False
        TheHubToolStripMenuItem.Checked = False
        Select Case valTheme
            Case 0
                LightModeToolStripMenuItem.Checked = True
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
                rtbOutput.AppendText("-- Check Theme Error: " & valTheme.ToString & vbCrLf)
        End Select
    End Sub

    Private Sub LightModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LightModeToolStripMenuItem.Click
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
    Private Function getKeyInfo(ByVal name As String, ByVal area As String) As Boolean
        getKeyInfo = False
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
                        getKeyInfo = True
                        'displayChecks(.area, False)
                        Exit Function
                    End If
                End With
            Next
        Else
            ' Convert the dungeon into a number for which array to look unto
            Dim dunNum As Byte = CByte(area.Replace("DUN", ""))
            For j = 0 To aKeysDungeons(dunNum).Length - 1
                With aKeysDungeons(dunNum)(j)
                    If .name = name Then
                        .forced = Not .forced
                        getKeyInfo = True
                        'displayChecksDungeons(dunNum, False)
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
        goScan(True)
    End Sub

    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        stopScanning()
        For Each chk In pnlHidden.Controls.OfType(Of CheckBox)()
            chk.Checked = False
        Next
        rtbOutput.ResetText()
        lastRoomScan = 0
        populateLocations()
    End Sub
End Class



Public Class custMenu
    Inherits ToolStripProfessionalRenderer
    Shared Property highlight As Color = Color.Blue
    Shared Property backColour As Color = Color.White
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
                Case "scan", "auto scan", "stop", "reset", "themes"
                    e.Graphics.DrawRectangle(New Pen(foreColour, 1), New Rectangle(0, 0, .Width - 1, .Height - 1))
            End Select
            .ForeColor = foreColour
            .Font = font
        End With
    End Sub
End Class