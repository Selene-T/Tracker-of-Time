Module minimap
    ' Contains minimap functions, including ER

    Public Sub erExitArray(Optional justClear As Boolean = False, Optional fromHere As Byte = 0, Optional toHere As Byte = 37)
        ' Function to create the array of exits used in the tracker. Each map has a number of exits. This builds or clears out the
        ' exits for overworld maps and dungeons, currently not the individual interior scenes.

        With frmTrackerOfTime
            ' If just clear, it will just clear the arrays. fromHere and toHere declare the range

            ' Clear out the visited and exit arrays
            For i = fromHere To toHere
                .aVisited(i) = Not justClear
                For ii = 0 To 6
                    .aExitMap(i)(ii) = 255
                Next
            Next
            .aVisited(37) = True     ' Since GaC is not randomized yet, we will always enabled it
            If justClear Then Return

            ' Populates the exits with the default exits
            For i = fromHere To toHere
                Select Case i
                    Case 0
                        ' MK Entrance (27-29)
                        .aExitMap(0)(0) = 7      ' HF
                        .aExitMap(0)(1) = 10     ' MK
                    Case 1
                        ' MK Back Alley (30-31)
                        .aExitMap(1)(0) = 10     ' MK
                        .aExitMap(1)(1) = 10     ' MK
                    Case 2
                        ' MK (Young) (32-33)
                        .aExitMap(2)(0) = 13     ' HC
                        .aExitMap(2)(1) = 197    ' MK Entrance
                        .aExitMap(2)(2) = 11     ' Outside ToT
                        .aExitMap(2)(3) = 199    ' MK Back Alley (right)
                        .aExitMap(2)(4) = 198    ' MK Back Alley (left)
                    Case 3
                        ' MK (Adult) (34)
                        .aExitMap(3)(0) = 51     ' OGC
                        .aExitMap(3)(1) = 197    ' MK Entrance
                        .aExitMap(3)(2) = 11     ' Outside ToT
                    Case 4
                        ' Outside ToT (35-37)
                        .aExitMap(4)(0) = 200    ' ToT Front
                        .aExitMap(4)(1) = 10     ' MK
                    Case 5
                        ' ToT (67)
                        .aExitMap(5)(0) = 11     ' Outside ToT
                    Case 6
                        ' HF (81)
                        .aExitMap(6)(0) = 15     ' KV Main
                        .aExitMap(6)(1) = 4      ' LW Bridge
                        .aExitMap(6)(2) = 34     ' ZR Front
                        .aExitMap(6)(3) = 44     ' GV Hyrule Side
                        .aExitMap(6)(4) = 197    ' MK Entrance
                        .aExitMap(6)(5) = 9      ' LLR
                        .aExitMap(6)(6) = 42     ' LH Main
                    Case 7
                        ' KV (82)
                        .aExitMap(7)(1) = 20     ' DMT Lower
                        .aExitMap(7)(2) = 7      ' HF
                        .aExitMap(7)(3) = 18     ' GY Lower
                    Case 8
                        ' GY (83)
                        .aExitMap(8)(1) = 15     ' KV Main
                    Case 9
                        ' ZR (84)
                        .aExitMap(9)(0) = 7      ' HF
                        .aExitMap(9)(1) = 37     ' ZD Main
                        .aExitMap(9)(2) = 2      ' LW Front
                    Case 10
                        ' KF (85)
                        .aExitMap(10)(1) = 8      ' LW Between Bridge
                        .aExitMap(10)(2) = 2      ' LW Front
                    Case 11
                        ' SFM (86)
                        .aExitMap(11)(1) = 3     ' LW Behind Mido
                    Case 12
                        ' LH (87)
                        .aExitMap(12)(1) = 7     ' HF
                        .aExitMap(12)(2) = 37    ' ZD Main
                    Case 13
                        ' ZD (88)
                        .aExitMap(13)(0) = 40    ' ZF Main
                        .aExitMap(13)(1) = 36    ' ZR Behind Waterfall
                        .aExitMap(13)(2) = 42    ' LH Main
                    Case 14
                        ' ZF (89)
                        .aExitMap(14)(2) = 38    ' ZD Behind King
                    Case 15
                        ' GV (90)
                        .aExitMap(15)(0) = 42    ' LH Main
                        .aExitMap(15)(1) = 7     ' HF
                        .aExitMap(15)(2) = 46    ' GF Main
                    Case 16
                        ' LW (91)
                        .aExitMap(16)(0) = 5     ' SFM Main
                        .aExitMap(16)(1) = 0     ' KF Main
                        .aExitMap(16)(2) = 35    ' ZR Main
                        .aExitMap(16)(3) = 30    ' GC Shortcut
                        .aExitMap(16)(4) = 0     ' KF Main
                        .aExitMap(16)(5) = 7     ' HF
                    Case 17
                        ' DC (92)
                        .aExitMap(17)(1) = 49    ' HW Colossus Side
                    Case 18
                        ' GF (93)
                        .aExitMap(18)(1) = 45    ' GV Gerudo Side
                        .aExitMap(18)(2) = 48    ' HW Gerudo Side
                    Case 19
                        ' HW (94)
                        .aExitMap(19)(0) = 50    ' DC
                        .aExitMap(19)(1) = 47    ' GF Behind Gate
                    Case 20
                        ' HC (95)
                        .aExitMap(20)(0) = 10    ' MK
                    Case 21
                        ' DMT (96)
                        .aExitMap(21)(1) = 29    ' GC Main
                        .aExitMap(21)(2) = 17    ' KV Behind Gate
                        .aExitMap(21)(3) = 22    ' DMC Upper
                    Case 22
                        ' DMC (97)
                        .aExitMap(22)(1) = 31    ' GC Darunia
                        .aExitMap(22)(2) = 21    ' DMT Upper
                    Case 23
                        ' GC (98)
                        .aExitMap(23)(0) = 25    ' DMC Lower Local
                        .aExitMap(23)(1) = 20    ' DMT Lower
                        .aExitMap(23)(2) = 2     ' LW Front
                    Case 24
                        ' LLR (99)
                        .aExitMap(24)(0) = 7     ' HF
                    Case 25
                        ' OGC (100)
                        .aExitMap(25)(0) = 191   ' Ganon's Castle
                        .aExitMap(25)(1) = 10    ' MK
                    Case 26
                        ' Deku Tree
                        .aExitMap(10)(0) = 60
                        .aExitMap(26)(0) = 1
                    Case 27
                        ' Dodongo's Cavern
                        .aExitMap(21)(0) = 69
                        .aExitMap(27)(0) = 20
                    Case 28
                        ' Jabu-Jabu's Belly
                        .aExitMap(14)(0) = 79
                        .aExitMap(28)(0) = 40
                    Case 29
                        ' Forest Temple
                        .aExitMap(11)(0) = 88
                        .aExitMap(29)(0) = 6
                    Case 30
                        ' Fire Temple
                        .aExitMap(22)(0) = 108
                        .aExitMap(30)(0) = 28
                    Case 31
                        ' Water Temple
                        .aExitMap(12)(0) = 119
                        .aExitMap(31)(0) = 42
                    Case 32
                        ' Spirit Temple
                        .aExitMap(17)(0) = 132
                        .aExitMap(32)(0) = 50
                    Case 33
                        ' Shadow Temple
                        .aExitMap(8)(0) = 150
                        .aExitMap(33)(0) = 19
                    Case 34
                        ' BotW
                        .aExitMap(7)(0) = 174
                        .aExitMap(34)(0) = 15
                    Case 35
                        ' Ice Cavern
                        .aExitMap(14)(1) = 178
                        .aExitMap(35)(0) = 41
                    Case 36
                        ' Gerudo Training Ground
                        .aExitMap(18)(0) = 179
                        .aExitMap(36)(0) = 46
                    Case 37
                        ' Ganon's Castle
                        .aExitMap(25)(0) = 193
                        .aExitMap(37)(0) = 51
                        .aExitMap(37)(1) = 196
                End Select
            Next
        End With
    End Sub

    Private iOldER As Byte = 255
    Private iLastMinimap As Integer = 101

    Public Sub clearLastMinimap()
        ' Clears the last minimap by setting it into an unused range
        iLastMinimap = 101
    End Sub

    Public Function getER() As Byte
        ' Gets the ER settings from the save context

        With frmTrackerOfTime
            Dim iNewER As Byte = 0
            ' Overworld ER check
            If Not .aAddresses(18) = 0 Then
                If .goRead(.aAddresses(18), 1) = 1 Then iNewER += 1
            End If
            ' Dungeon ER check
            If Not .aAddresses(19) = 0 Then
                If .goRead(.aAddresses(19), 1) = 1 Then iNewER += 2
            End If
            If Not iNewER = iOldER Then
                ' If a change in ER is detected (basically first scan), clear the appropriate exits
                If iNewER Mod 2 = 1 Then erExitArray(True, 0, 25)
                If iNewER > 1 Then erExitArray(True, 26, 37)
            End If
            iOldER = iNewER
            scanER()

            Return iNewER
        End With
    End Function

    Private Sub scanER()
        ' Displays the minimap for the current location and if ER, scans the appropriate exits

        With frmTrackerOfTime
            ' ER scanning
            Dim isOverworld As Boolean = False
            Dim isDungeon As Boolean = False
            Dim doTrials As Boolean = False

            ' Split the ER value into separate checks
            If iOldER Mod 2 = 1 Then isOverworld = True
            If iOldER > 1 Then isDungeon = True

            ' For building the arrays for scenes with Dungeon entrances
            Dim ent As Byte = 0

            Dim locationCode As Integer = .getLocation
            Dim locationArray As Byte = 255
            Dim readExits(6) As Integer     ' Exits read for current map
            Dim aAlign(6) As Byte           ' Sets text alignment: 0 = Left | 1 = Centre | 2 = Right
            Dim lPoints As New List(Of Point)
            For i = 0 To aAlign.Length - 1
                aAlign(i) = 0
                readExits(i) = 0
            Next

            .getAge()

            Dim addrRoom As Integer = If(.isSoH, soh.gRoomAddr, &H1D8BEE)

            If locationCode <= 9 Then
                .iRoom = CByte(.goRead(addrRoom, 1))
            End If

            ' First load up the map, I keep this separated so that it does not matter for ER settings
            For i = 0 To 1
                Dim exitLoop = True
                Select Case locationCode
                    Case 0
                        Select Case .iRoom
                            Case 10, 12
                                .pbxMap.Image = My.Resources.mapDT4  ' 3F
                            Case 1, 2, 11
                                .pbxMap.Image = My.Resources.mapDT3  ' 2F
                            Case 0
                                .pbxMap.Image = My.Resources.mapDT2  ' 1F
                            Case 3 To 8
                                .pbxMap.Image = My.Resources.mapDT1  ' B1
                            Case 9
                                .pbxMap.Image = My.Resources.mapDT0  ' B2
                        End Select
                    Case 1
                        Select Case .iRoom
                            Case 5, 6, 9, 10, 12, 16, 17, 18
                                .pbxMap.Image = My.Resources.mapDDC1 ' 2F
                            Case 0 To 4, 7, 8, 11, 13 To 15
                                .pbxMap.Image = My.Resources.mapDDC0 ' 1F
                        End Select
                    Case 2
                        Select Case .iRoom
                            Case 0 To 2, 4 To 12
                                .pbxMap.Image = My.Resources.mapJB1  ' 1F
                            Case 3, 13 To 16
                                .pbxMap.Image = My.Resources.mapJB0  ' B1
                        End Select
                    Case 3
                        Select Case .iRoom
                            Case 10, 12 To 14, 19, 20, 23 To 26
                                .pbxMap.Image = My.Resources.mapFoT3 ' 2F
                            Case 0 To 8, 11, 15, 16, 18, 21, 22
                                .pbxMap.Image = My.Resources.mapFoT2 ' 1F
                            Case 9
                                .pbxMap.Image = My.Resources.mapFoT1 ' B1
                            Case 17
                                .pbxMap.Image = My.Resources.mapFoT0 ' B2
                        End Select
                    Case 4
                        Select Case .iRoom
                            Case 8, 30, 34, 35
                                .pbxMap.Image = My.Resources.mapFiT4 ' 5F
                            Case 7, 12 To 14, 27, 32, 33, 37
                                .pbxMap.Image = My.Resources.mapFiT3 ' 4F
                            Case 5, 9, 11, 16, 23 To 26, 28, 31
                                .pbxMap.Image = My.Resources.mapFiT2 ' 3F
                            Case 4, 10, 36
                                .pbxMap.Image = My.Resources.mapFiT1 ' 2F
                            Case 0 To 3, 15, 17 To 22
                                .pbxMap.Image = My.Resources.mapFiT0 ' 1F
                        End Select
                    Case 5
                        Select Case .iRoom
                            Case 0, 1, 4 To 7, 10, 11, 13, 17, 19, 20, 30, 31, 43
                                .pbxMap.Image = My.Resources.mapWaT3 ' 3F
                            Case 22, 25, 29, 32, 35, 39, 41
                                .pbxMap.Image = My.Resources.mapWaT2 ' 2F
                            Case 3, 8, 9, 12, 14 To 16, 18, 21, 23, 24, 26, 28, 33, 34, 36 To 38, 40, 42
                                .pbxMap.Image = My.Resources.mapWaT1 ' 1F
                            Case 2, 27
                                .pbxMap.Image = My.Resources.mapWaT0 ' B1
                        End Select
                    Case 6
                        Select Case .iRoom
                            Case 22, 24 To 26, 31
                                .pbxMap.Image = My.Resources.mapSpT3  ' 4F
                            Case 7 To 11, 16 To 21, 23, 29
                                .pbxMap.Image = My.Resources.mapSpT2  ' 3F
                            Case 5, 6, 28, 30
                                .pbxMap.Image = My.Resources.mapSpT1  ' 2F
                            Case 0 To 4, 12 To 15, 27
                                .pbxMap.Image = My.Resources.mapSpT0  ' 1F
                        End Select
                    Case 7
                        Select Case .iRoom
                            Case 0 To 2, 4
                                .pbxMap.Image = My.Resources.mapShT3 ' B1
                            Case 5 To 8
                                .pbxMap.Image = My.Resources.mapShT2 ' B2
                            Case 9, 16, 22
                                .pbxMap.Image = My.Resources.mapShT1 ' B3
                            Case 3, 10 To 15, 17 To 21, 23 To 26
                                .pbxMap.Image = My.Resources.mapShT0 ' B4
                        End Select
                    Case 8
                        Select Case .iRoom
                            Case 0 To 6
                                .pbxMap.Image = My.Resources.mapBotW2 ' B1
                            Case 7, 8
                                .pbxMap.Image = My.Resources.mapBotW1 ' B2
                            Case 9
                                .pbxMap.Image = My.Resources.mapBotW0 ' B3
                        End Select
                    Case 9
                        .pbxMap.Image = My.Resources.mapIC0  ' F1
                    Case 10
                        .pbxMap.Image = My.Resources.mapGAC0
                        doTrials = True
                        .iRoom = 0
                    Case 11
                        .pbxMap.Image = My.Resources.mapGTG    ' GTG
                    Case 13
                        .iRoom = getGanonMap()
                        Select Case .iRoom
                            Case 1
                                .pbxMap.Image = My.Resources.mapGAC1
                                doTrials = True
                            Case 2
                                .pbxMap.Image = My.Resources.mapGAC2
                            Case 3
                                .pbxMap.Image = My.Resources.mapGAC3
                            Case 4
                                .pbxMap.Image = My.Resources.mapGAC4
                            Case 5
                                .pbxMap.Image = My.Resources.mapGAC5
                            Case 6
                                .pbxMap.Image = My.Resources.mapGAC6
                            Case 7
                                .pbxMap.Image = My.Resources.mapGAC7
                            Case Else
                                .pbxMap.Image = My.Resources.mapGAC0
                                doTrials = True
                        End Select
                    Case 27 To 29
                        .pbxMap.Image = My.Resources.mapMKE
                        'If locationCode = 29 Then
                        '.pbxMap.Image = My.Resources.mapMKEntrance2
                        'Else
                        '.pbxMap.Image = My.Resources.mapMKEntrance
                        'End If
                    Case 30, 31
                        .pbxMap.Image = My.Resources.mapMKBA
                    Case 32, 33
                        .pbxMap.Image = My.Resources.mapMK
                    Case 34
                        .pbxMap.Image = My.Resources.mapMK2
                    Case 35 To 37
                        .pbxMap.Image = My.Resources.mapOToT
                        '.pbxMap.Image = My.Resources.mapToTOutside
                        'Case 37
                        '.pbxMap.Image = My.Resources.mapToTOutside2
                    Case 67
                        .pbxMap.Image = My.Resources.mapToT
                    Case 81
                        .pbxMap.Image = My.Resources.mapHF
                    Case 82
                        .pbxMap.Image = My.Resources.mapKV
                    Case 83
                        .pbxMap.Image = My.Resources.mapGY
                    Case 84
                        .pbxMap.Image = My.Resources.mapZR
                    Case 85
                        .pbxMap.Image = My.Resources.mapKF
                    Case 86
                        .pbxMap.Image = My.Resources.mapSFM
                    Case 87
                        .pbxMap.Image = My.Resources.mapLH
                    Case 88
                        .pbxMap.Image = My.Resources.mapZD
                    Case 89
                        .pbxMap.Image = My.Resources.mapZF
                    Case 90
                        .pbxMap.Image = My.Resources.mapGV
                    Case 91
                        .pbxMap.Image = My.Resources.mapLW
                    Case 92
                        .pbxMap.Image = My.Resources.mapDC
                    Case 93, 12
                        .pbxMap.Image = My.Resources.mapGF
                        locationCode = 93
                    Case 94
                        .pbxMap.Image = My.Resources.mapHW2
                    Case 95
                        .pbxMap.Image = My.Resources.mapHC
                    Case 96
                        .pbxMap.Image = My.Resources.mapDMT
                    Case 97
                        .pbxMap.Image = My.Resources.mapDMC
                    Case 98
                        .pbxMap.Image = My.Resources.mapGC
                    Case 99
                        .pbxMap.Image = My.Resources.mapLLR
                    Case 100
                        .pbxMap.Image = My.Resources.mapOGC
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
                .pbxTrialFire.Visible = False
                .pbxTrialForest.Visible = False
                .pbxTrialWater.Visible = False
                .pbxTrialSpirit.Visible = False
                .pbxTrialShadow.Visible = False
                .pbxTrialLight.Visible = False
            End If

            If iOldER = 0 Then Exit Sub

            ' Dungeon only checks
            If isDungeon Then
                Select Case locationCode
                    Case 0
                        readExits(0) = &H377116     ' KF from Deku Tree
                        If .aMQ(0) Then readExits(0) = &H3770B6 ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H377114     ' Queen Gohma from Deku Tree
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        locationArray = 26
                    Case 1
                        readExits(0) = &H36FABE     ' DMT from Dodongo's Cavern
                        If .aMQ(1) Then readExits(0) = &H36FA8E ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H36FABC     ' King Dodongo from Dodongo's Cavern
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        locationArray = 27
                    Case 2
                        readExits(0) = &H36F43E     ' ZD from Jabu-Jabu's Belly
                        If .aMQ(2) Then readExits(0) = &H36F40E ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H36F43C     ' Barinade from Jabu-Jabu's Belly
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        locationArray = 28
                    Case 3
                        readExits(0) = &H36ED12     ' SFM from Forest Temple
                        If .aMQ(3) Then readExits(0) = &H36ED12 ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H36ED10     ' Phantom Ganon from Forest Temple
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        locationArray = 29
                    Case 4
                        readExits(0) = &H36A3C6     ' DMC from Fire Temple
                        If .aMQ(4) Then readExits(0) = &H36A376 ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H36A3C4     ' Volvagia from Fire Temple
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        locationArray = 30
                    Case 5
                        readExits(0) = &H36EF76     ' LH from Water Temple
                        If .aMQ(5) Then readExits(0) = &H36EF36 ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H36EF74     ' Morpha from Temple
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        locationArray = 31
                    Case 6
                        readExits(0) = &H36B1EA     ' DC from Spirit Temple
                        If .aMQ(6) Then readExits(0) = &H36B12A ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H36B1E8     ' Twinrova from Spirit Temple
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        'readExits(2) = &H36B1EC     ' DC Statue Hand Right from Spirit Temple
                        'lPoints.Add(New Point(99, 332))
                        'aAlign(2) = 2
                        'readExits(3) = &H36B1EE     ' DC Statue Hand Left from Spirit Temple
                        'lPoints.Add(New Point(482, 357))
                        locationArray = 32
                    Case 7
                        readExits(0) = &H36C86E     ' GY from Shadow Temple
                        If .aMQ(7) Then readExits(0) = &H36C86E ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        readExits(1) = &H36C86C     ' Bongo Bongo from Shadow Temple
                        lPoints.Add(New Point(278, 3))
                        aAlign(1) = 1
                        locationArray = 33
                    Case 8
                        readExits(0) = &H3785C6     ' KV from BotW
                        If .aMQ(8) Then readExits(0) = &H378566 ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        locationArray = 34
                    Case 9
                        readExits(0) = &H37352E     ' ZF from Ice Cavern
                        If .aMQ(9) Then readExits(0) = &H37344E ' MQ version
                        lPoints.Add(New Point(278, 390))
                        aAlign(0) = 1
                        locationArray = 35
                    Case 10
                        'readExits(0) = &H374348     ' Ganon's Castle from Ganon's Tower
                        'lPoints.Add(New Point(278, 3))
                        'aAlign(0) = 1
                        readExits(1) = &H37434A     ' Ganon from Ganon's Tower
                        lPoints.Add(New Point(278, 390))
                        aAlign(1) = 1
                    Case 11
                        readExits(0) = &H373686     ' GF from GTG
                        If .aMQ(10) Then readExits(0) = &H373686 ' MQ version
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

            If iOldER Mod 2 = 1 Then
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
                        ent += 1
                        readExits(ent) = &H368A7A     ' HF from KV
                        lPoints.Add(New Point(57, 294))
                        aAlign(ent) = 2
                        ent += 1
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
                        ent += 1
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
                        ent += 1
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
                        ent += 1
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
                        ent += 1
                        readExits(ent) = &H365FEE     ' KV from DMT
                        lPoints.Add(New Point(259, 384))
                        aAlign(ent) = 1
                        ent += 1
                        readExits(ent) = &H365FF2     ' DMC from DMT
                        lPoints.Add(New Point(315, 3))
                        aAlign(ent) = 1
                        locationArray = 21
                    Case 97
                        readExits(ent) = &H374B98     ' GC from DMC
                        lPoints.Add(New Point(129, 192))
                        aAlign(ent) = 2
                        ent += 1
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
            .aVisited(locationArray) = True

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
                    exitCode = Hex(.goRead(readExits(i), 15))
                    .fixHex(exitCode, 3)
                    exitCode = exit2label(exitCode, iNewReach)


                    ' If aReachExit is not 255, set exit to the new iNewReach
                    If Not iNewReach = 255 Then
                        .aExitMap(locationArray)(i) = iNewReach
                    End If

                    ' Only run the display part if the panel is actually visable
                    If .pnlER.Visible Then
                        ' Convert the iNewReach into the location for aVisited
                        iVisited = zone2map(iNewReach)
                        doDisplay = True
                        If iVisited = 255 Then
                            doDisplay = True
                        Else
                            If Not .aVisited(iVisited) Then exitCode = "?"
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
                                    If ptX > (.pbxMap.Width / 2) Then
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

                            Graphics.FromImage(.pbxMap.Image).DrawString(exitCode, fontGS, New SolidBrush(Color.Black), ptX + 1, ptY + 1)
                            Graphics.FromImage(.pbxMap.Image).DrawString(exitCode, fontGS, New SolidBrush(Color.White), ptX, ptY)
                        End If
                    End If
                End If
            Next
        End With
    End Sub

    Private Sub updateTrials()
        ' Updates the visibility of the greyed out icons for Ganon's Castle Trials
        With frmTrackerOfTime
            .pbxTrialForest.Visible = .checkLoc("6611")
            .pbxTrialFire.Visible = .checkLoc("6614")
            .pbxTrialWater.Visible = .checkLoc("6612")
            .pbxTrialSpirit.Visible = .checkLoc("6629")
            .pbxTrialShadow.Visible = .checkLoc("6613")
            .pbxTrialLight.Visible = .checkLoc("6615")
        End With
    End Sub

    Private Function exit2label(ByVal exitCode As String, ByRef reachMap As Byte) As String
        ' Converts the games exit codes into text to display to the user

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
                If frmTrackerOfTime.isAdult Then
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

    Public Function zone2map(ByVal zone As Byte) As Byte
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
                If frmTrackerOfTime.isAdult Then
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

    Public Sub updateMiniMap()
        ' Updates the minimap with squares for each check

        With frmTrackerOfTime
            ' If the panel is not even visible, do not bother to do this code
            If Not .pnlER.Visible Then Exit Sub

            ' Clear everything
            Dim aIconLoc(30) As String
            Dim aIconPos As New List(Of Point)
            For i = 0 To aIconLoc.Length - 1
                aIconLoc(i) = String.Empty
                .aIconName(i) = String.Empty
            Next
            .lRegions.Clear()

            ' Check each minimap for which icons it should display, recording them by their .loc and adding location to a list
            Select Case iLastMinimap
                Case 0
                    If Not .aMQ(0) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    If Not .aMQ(1) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    If Not .aMQ(2) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    If Not .aMQ(3) Then
                        Select Case .iRoom
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
                                aIconLoc(8) = "7826"
                                aIconPos.Add(New Point(186, 87))
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
                        Select Case .iRoom
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
                    If Not .aMQ(4) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    If Not .aMQ(5) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    If Not .aMQ(6) Then
                        Select Case .iRoom
                            Case 22, 24 To 26, 31 ' 4F
                                aIconLoc(0) = "3710"
                                aIconPos.Add(New Point(322, 75))
                                aIconLoc(1) = "3718"
                                aIconPos.Add(New Point(197, 131))
                            Case 7 To 11, 16 To 21, 23, 29 ' 3F
                                aIconLoc(0) = "3701"
                                aIconPos.Add(New Point(136, 244))
                                aIconLoc(1) = "5511"
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
                                aIconPos.Add(New Point(108, 276))
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
                        Select Case .iRoom
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
                    If Not .aMQ(7) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    If Not .aMQ(8) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    If Not .aMQ(9) Then
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
                    If Not .aMQ(10) Then
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
                    If Not .aMQ(11) Then
                        Select Case .iRoom
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
                        Select Case .iRoom
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
                    aIconLoc(1) = .locSwap(12)
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
                    aIconLoc(6) = .locSwap(11)
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
                    aIconLoc(6) = .locSwap(18)
                    aIconPos.Add(New Point(315, 288))
                    aIconLoc(7) = "2001"
                    aIconPos.Add(New Point(412, 240))
                    aIconLoc(8) = "6411"
                    aIconPos.Add(New Point(428, 240))
                    aIconLoc(9) = .locSwap(7)
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
                    aIconLoc(0) = .locSwap(3)
                    aIconPos.Add(New Point(156, 203))
                    aIconLoc(1) = "4800"
                    aIconPos.Add(New Point(178, 214))
                    aIconLoc(2) = "2204"
                    aIconPos.Add(New Point(110, 160))
                    aIconLoc(3) = "4700"
                    aIconPos.Add(New Point(226, 231))
                    aIconLoc(4) = "4900"
                    aIconPos.Add(New Point(286, 201))
                    aIconLoc(5) = .locSwap(10)
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
                    aIconLoc(1) = .locSwap(8)
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
                    aIconLoc(7) = .locSwap(4)
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
                    aIconLoc(0) = .locSwap(13)
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
                    aIconLoc(7) = .locSwap(16)
                    aIconPos.Add(New Point(227, 118))
                    aIconLoc(8) = .locSwap(17)
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
                    aIconLoc(0) = .locSwap(15)
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
                    aIconLoc(3) = .locSwap(14)
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
                    aIconLoc(2) = .locSwap(2)
                    aIconPos.Add(New Point(294, 31))
                    aIconLoc(3) = "4623"
                    aIconPos.Add(New Point(288, 185))
                    aIconLoc(4) = .locSwap(5)
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
                    aIconLoc(3) = .locSwap(1)
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
                    aIconLoc(6) = .locSwap(6)
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
                    aIconLoc(2) = .locSwap(9)
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
                    aIconLoc(0) = .locSwap(0)
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
                        For Each thisKey In .aKeysOverworld.Where(Function(k As keyCheck) k.loc.Equals(aIconLoc(shutupLambda)))
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
                        For Each thiskey In .aKeysDungeons(ii).Where(Function(k As keyCheck) k.loc.Equals(aIconLoc(shutupLambda)))
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
                                    addCheck = CBool(IIf(.goldSkulltulas < 50, True, False))
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
                            frmTrackerOfTime.lRegions.Add(New Rectangle(.X, .Y, 15, 15))

                            Graphics.FromImage(frmTrackerOfTime.pbxMap.Image).DrawRectangle(Pens.Black, .X, .Y, 15, 15)
                            Graphics.FromImage(frmTrackerOfTime.pbxMap.Image).DrawRectangle(Pens.White, .X + 1, .Y + 1, 13, 13)

                            If Not key.forced And Not key.checked Then
                                fillColour = Color.Lime
                                Select Case frmTrackerOfTime.checkLogic(key.logic, key.zone)
                                    Case 0
                                        fillColour = Color.Red
                                    Case 1
                                        suffix = " (Y)"
                                    Case 2
                                        suffix = " (A)"
                                End Select
                                Graphics.FromImage(frmTrackerOfTime.pbxMap.Image).FillRectangle(New SolidBrush(fillColour), .X + 2, .Y + 2, 12, 12)
                            End If
                            frmTrackerOfTime.aIconName(i) = prefix & key.name & suffix
                        Else
                            ' Have to create an unreachable location to at least populate the list to make it match up
                            frmTrackerOfTime.lRegions.Add(New Rectangle(-2, -2, 1, 1))
                        End If
                    End With
                End If
            Next
            If iLastMinimap = 94 Then
                ' Slow down the slow scan, and speed up the fast scan
                .tmrAutoScan.Interval = 10000
                .tmrFastScan.Interval = 333
                wastelandPOS()
            Else
                ' Return them to normal
                .tmrAutoScan.Interval = 5000
                .tmrFastScan.Interval = 1000
            End If
        End With
    End Sub

    Private Function getGanonMap() As Byte
        ' Detects what area of Ganon's Castle you are in based on your position
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
        ' Detects what side of a line segment you are on based on how far from the points, main concern is either positive or negative
        lineSide = (pt2.X - pt1.X) * (ptTest.Y - pt1.Y) - (pt2.Y - pt1.Y) * (ptTest.X - pt1.X)
    End Function

    Private Sub wastelandPOS()
        ' Tracks player's position and rotation while in the Haunted Wasteland, and displays as a yellow arrow
        With frmTrackerOfTime
            ' Get player's position and convert that to a point on the minimap
            Dim linkPOS As Double() = getPosition()
            Dim coordX As Double = ((linkPOS(0) + 4550) / 8200) * 400 + 53
            Dim coordZ As Double = ((linkPOS(2) + 3750) / 8200) * 400 + 17
            Dim linkRot As Integer = 0

            ' SoH has to be read differently
            If .isSoH Then
                linkRot = CInt(GDATA(&H1730A, 2))
            Else
                linkRot = .goRead(&H1DAA74, 15)
            End If

            ' Create a point for both the head and the tail of the arrow.
            ' This is to have the player's location in the centre of the arrow, not at the tail end.
            Dim headA As Double = (((linkRot / 65535 * 360) - 90) * -1) * Math.PI / 180
            Dim tailA As Double = (((linkRot / 65535 * 360) + 90) * -1) * Math.PI / 180
            Dim headX As Double = (8 * Math.Cos(headA)) + coordX
            Dim headZ As Double = (8 * Math.Sin(headA)) + coordZ
            Dim tailX As Double = (8 * Math.Cos(tailA)) + coordX
            Dim tailZ As Double = (8 * Math.Sin(tailA)) + coordZ

            ' Create the yellow arrow and draw it on the minimap
            Dim p As New Pen(Color.Yellow, 5)
            p.EndCap = Drawing2D.LineCap.ArrowAnchor
            Graphics.FromImage(.pbxMap.Image).DrawLine(p, CInt(tailX), CInt(tailZ), CInt(headX), CInt(headZ))
        End With
    End Sub

    Private Function getPosition() As Double()
        ' Gets player's position
        With frmTrackerOfTime
            ' Grab values for XYZ
            Dim valX As String = String.Empty
            Dim valY As String = String.Empty
            Dim valZ As String = String.Empty

            If .isSoH Then
                valX = Convert.ToString(GDATA(&H17264), 2)
                valY = Convert.ToString(GDATA(&H17268), 2)
                valZ = Convert.ToString(GDATA(&H1726C), 2)
            Else
                valX = Convert.ToString(.goRead(&H1DAA54), 2)
                valY = Convert.ToString(.goRead(&H1DAA58), 2)
                valZ = Convert.ToString(.goRead(&H1DAA5C), 2)
            End If

            fixBinaryLength(valX, valY, valZ)

            ' Convert values into IEEE-754 floating points
            Dim coordX As Double = bin2float(valX)
            Dim coordY As Double = bin2float(valY)
            Dim coordZ As Double = bin2float(valZ)

            ' Return the whole array
            Return New Double() {coordX, coordY, coordZ}
        End With
    End Function

    Private Function name2loc(ByVal keyName As String, ByVal keyArea As String) As String
        name2loc = "0"
        If frmTrackerOfTime.firstRun Then Exit Function
        ' Checks for a specific key by name to get the loc

        ' Checks normal keys only because this is used only for shops
        For Each key In frmTrackerOfTime.aKeysOverworld.Where(Function(k As keyCheck) k.name.Equals(keyName))
            If key.area = keyArea Then Return key.loc
        Next
    End Function
End Module
