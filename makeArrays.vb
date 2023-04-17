Module makeArrays
    ' Creates most of the arrays used

    Public Sub makeArrLocationArray()
        With frmTrackerOfTime
            .arrLocation(0) = &H11AD18 + 4      ' DMC/DMT/OGC Great Fairy Fountain (Events)
            .arrLocation(1) = &H11AF80 + 4      ' Hyrule Field (Events) Big Poes Captured and Ocarina of Time
            .arrLocation(2) = &H11B028 + 4      ' Lake Hylia (Events) Ruto's Letter, Open Water Temple, and Bean Plant
            .arrLocation(3) = &H11A714 + 12     ' Fire Temple (Standing)
            .arrLocation(4) = &H11A730 + 12     ' Water Temple (Standing)
            .arrLocation(5) = &H11A768 + 12     ' Shadow Temple (Standing)
            .arrLocation(6) = &H11A784 + 12     ' Bottom of the Well (Standing)
            .arrLocation(7) = &H11A7A0 + 12     ' Ice Cavern (Standing)
            .arrLocation(8) = &H11A7D8 + 12     ' Gerudo Training Ground (Standing)
            .arrLocation(9) = &H11A810 + 12     ' Ganon’s Castle #1 (Standing)
            .arrLocation(10) = &H11A880 + 12    ' Deku Tree Boss Room (Standing)
            .arrLocation(11) = &H11A89C + 12    ' Dodongo's Cavern Boss Room (Standing)
            .arrLocation(12) = &H11A8B8 + 12    ' Jabu-Jabu's Belly Boss Room (Standing)
            .arrLocation(13) = &H11A8D4 + 12    ' Forest Temple Boss Room (Standing)
            .arrLocation(14) = &H11A8F0 + 12    ' Fire Temple Boss Room (Standing)
            .arrLocation(15) = &H11A90C + 12    ' Water Temple Boss Room (Standing)
            .arrLocation(16) = &H11A928 + 12    ' Spirit Temple Boss Room (Standing)
            .arrLocation(17) = &H11A944 + 12    ' Shadow Temple Boss Room (Standing)
            .arrLocation(18) = &H11ACA8 + 12    ' Impa’s House (Standing)
            .arrLocation(19) = &H11AD6C + 12    ' All Grottos (Standing)
            .arrLocation(20) = &H11AE84 + 12    ' Windmill / Dampe's Grave (Standing)
            .arrLocation(21) = &H11AEF4 + 12    ' Lon Lon Tower Item (Standing)
            .arrLocation(22) = &H11AFB8 + 12    ' Graveyard (Standing)
            .arrLocation(23) = &H11AFD4 + 12    ' Zora's River (Standing)
            .arrLocation(24) = &H11B028 + 12    ' Lake Hylia (Standing)
            .arrLocation(25) = &H11B060 + 12    ' Zora's Fountain (Standing)
            .arrLocation(26) = &H11B07C + 12    ' Gerudo Valley (Standing)
            .arrLocation(27) = &H11B0B4 + 12    ' Desert Colossus (Standing)
            .arrLocation(28) = &H11B124 + 12    ' Death Mountain Trail (Standing)
            .arrLocation(29) = &H11B140 + 12    ' Death Mountain Crater (Standing)
            .arrLocation(30) = &H11B15C + 12    ' Goron City (Standing)
            .arrLocation(31) = &H11A6A4         ' Deku Tree
            .arrLocation(32) = &H11A6C0         ' Dodongo's Cavern
            .arrLocation(33) = &H11A6DC         ' Jabu-Jabu's Belly
            .arrLocation(34) = &H11A6F8         ' Forest Temple
            .arrLocation(35) = &H11A714         ' Fire Temple
            .arrLocation(36) = &H11A730         ' Water Temple
            .arrLocation(37) = &H11A74C         ' Spirit Temple
            .arrLocation(38) = &H11A768         ' Shadow Temple
            .arrLocation(39) = &H11A784         ' Bottom of the Well
            .arrLocation(40) = &H11A7A0         ' Ice Cavern
            .arrLocation(41) = &H11A7BC         ' Ganon’s Castle #2
            .arrLocation(42) = &H11A7D8         ' Gerudo Training Ground
            .arrLocation(43) = &H11A810         ' Ganon’s Castle #1
            .arrLocation(44) = &H11A89C         ' Dodongo's Cavern Boss Room
            .arrLocation(45) = &H11AB04         ' Mido’s House
            .arrLocation(46) = &H11AD6C         ' All Grottos
            .arrLocation(47) = &H11AD88         ' Grave with Sun Song Chest
            .arrLocation(48) = &H11ADA4         ' Graveyard Under Grave
            .arrLocation(49) = &H11ADC0         ' Royal Grave
            .arrLocation(50) = &H11AE84         ' Windmill / Dampe
            .arrLocation(51) = &H11AFF0         ' Kokiri Forest
            .arrLocation(52) = &H11B028         ' Lake Hylia
            .arrLocation(53) = &H11B044         ' Zora's Domain
            .arrLocation(54) = &H11B07C         ' Gerudo Valley
            .arrLocation(55) = &H11B0B4         ' Desert Colossus
            .arrLocation(56) = &H11B0D0         ' Gerudo’s Fortress
            .arrLocation(57) = &H11B0EC         ' Haunted Wasteland
            .arrLocation(58) = &H11B124         ' Death Mountain Trail
            .arrLocation(59) = &H11B15C         ' Goron City
            .arrLocation(60) = &H11A640         ' *Biggoron Check
            .arrLocation(61) = &H11B490         ' *Big Fish
            .arrLocation(62) = &H11B4A4         ' *Events 1: Egg from Malon, Obtained Epona, Won Cow
            .arrLocation(63) = &H11B4A8         ' *Events 2: Zora Diving Game, Darunia’s Joy
            .arrLocation(64) = &H11B4AC         ' *Events 3: Zelda’s Letter, Song from Impa, Sun Song??, opened Temple of Time, Rainbow Bridge
            .arrLocation(65) = &H11B4B4         ' *Events 5: Scarecrow as Adult
            .arrLocation(66) = &H11B4B8         ' *Events 6: Song at Colossus, Trials
            .arrLocation(67) = &H11B4BC         ' *Events 7: Saria Gift, Skulltula trades, Barrier Lowered
            .arrLocation(68) = &H11B4C0         ' *Item Collect #1
            .arrLocation(69) = &H11B4C4         ' *Item Collection #2
            .arrLocation(70) = &H11B4E8         ' *Item: Rolling Goron as Young + Adult Link
            .arrLocation(71) = &H11B4EC         ' *Thaw Zora King
            .arrLocation(72) = &H11B4F8         ' *Items: 1st and 2nd Scrubs, Lost Dog
            .arrLocation(73) = &H11B894         ' *Scarecrow Song
            .arrLocation(74) = &H11A66C         ' *Equipment checks, figured this would be easier
            .arrLocation(75) = &H11A60C         ' *Check for Biggoron's Sword
            .arrLocation(76) = &H11A670         ' *Upgrades
            .arrLocation(77) = &H11A674         ' *Quest Items and Songs
            .arrLocation(78) = &H11B46C         ' **Gold Skulltulas 1
            .arrLocation(79) = &H11B470         ' **Gold Skulltulas 2
            .arrLocation(80) = &H11B474         ' **Gold Skulltulas 3
            .arrLocation(81) = &H11B478         ' **Gold Skulltulas 4
            .arrLocation(82) = &H11B47C         ' **Gold Skulltulas 5
            .arrLocation(83) = &H11B480         ' **Gold Skulltulas 6
            .arrLocation(84) = &H11A6B4         ' ***Scrub Shuffle (Deku Tree)
            .arrLocation(85) = &H11A6D0         ' ***Scrub Shuffle (Dodongo's Cavern)
            .arrLocation(86) = &H11A6EC         ' ***Scrub Shuffle (Jabu-Jabu's Belly)
            .arrLocation(87) = &H11A820         ' ***Scrub Shuffle (Ganon's Castle)
            .arrLocation(88) = &H11A874         ' ***Scrub Shuffle (Hyrule Field Grotto)
            .arrLocation(89) = &H11A900         ' ***Scrub Shuffle (Zora's River Grotto)
            .arrLocation(90) = &H11A954         ' ***Scrub Shuffle (Sacred Forest Meadow Grotto)
            .arrLocation(91) = &H11A970         ' ***Scrub Shuffle (Lake Hylia Grotto)
            .arrLocation(92) = &H11A98C         ' ***Scrub Shuffle (Gerudo Valley Grotto)
            .arrLocation(93) = &H11AA18         ' ***Scrub Shuffle (Lost Woods Grotto)
            .arrLocation(94) = &H11AA88         ' ***Scrub Shuffle (Death Mountain Crater Grotto)
            .arrLocation(95) = &H11AAC0         ' ***Scrub Shuffle (Goron City)
            .arrLocation(96) = &H11AADC         ' ***Scrub Shuffle (Lon Lon Ranch)
            .arrLocation(97) = &H11AAF8         ' ***Scrub Shuffle (Desert Colossus)
            .arrLocation(98) = &H11B0A8         ' ***Scrub Shuffle (Lost Woods)
            .arrLocation(99) = &H11AB84         ' *Shopsanity Checks
            .arrLocation(100) = &H11AC54 + 12   ' Link's House (Standing)
            .arrLocation(101) = &H11AC8C + 12   ' Lon Lon Ranch Stables (Standing)
            .arrLocation(102) = &H11A6DC + 12   ' Jabu-Jabu's Belly (Standing)
            .arrLocation(103) = &H11A6A4 + 4    ' Deku Tree (Events)
            .arrLocation(104) = &H11A6C0 + 4    ' Dodongo's Cavern (Events)
            .arrLocation(105) = &H11A6DC + 4    ' Jabu-Jabu's Belly (Events)
            .arrLocation(106) = &H11A6F8 + 4    ' Forest Temple (Events)
            .arrLocation(107) = &H11A714 + 4    ' Fire Temple (Events)
            .arrLocation(108) = &H11A730 + 4    ' Water Temple (Events)
            .arrLocation(109) = &H11A74C + 4    ' Spirit Temple (Events)
            .arrLocation(110) = &H11A768 + 4    ' Shadow Temple (Events)
            .arrLocation(111) = &H11A7A0 + 4    ' Ice Cavern (Events)
            .arrLocation(112) = &H11A7D8 + 4    ' Gerudo Training Ground (Events)
            .arrLocation(113) = &H11A810 + 4    ' Ganon’s Castle #1 (Events)
            .arrLocation(114) = &H11B0EC + 12   ' Haunted Wasteland (Standing)
            .arrLocation(115) = &H11B15C + 4    ' Goron City (Events)
            .arrLocation(116) = &H11AFD4 + 4    ' Zora's River (Events)
            .arrLocation(117) = &H11A784 + 4    ' BotW (Events)
            .arrLocation(118) = &H11B178        ' Lon Lon Ranch
            .arrLocation(119) = &H11B00C        ' Sacred Forest Meadow
            .arrLocation(120) = &H11ADDC        ' Shooting Gallery
            .arrLocation(121) = &H11AD50        ' GF DC/HC/ZF
            .arrLocation(122) = &H11AD18        ' DMC/DMT/OGC Great Fairy Fountain
            .arrLocation(123) = &H11ADF8        ' Temple of Time
            .arrLocation(124) = &H11AF80        ' Hyrule Field
            .arrLocation(125) = &H11AB90 + 12   ' KF Shop (Standing)
        End With
    End Sub

    Public Function makeLabelArray(ByVal whichArray As String) As Label()
        ' Label arrays
        Dim array() As Label
        Dim arraySize As Byte = 0

        ' Determine size based on array
        Select Case LCase(whichArray)
            Case "dungeon"
                arraySize = 11
            Case "overworld"
                arraySize = 23
            Case Else
                Return Nothing
        End Select

        ' Resize array
        ReDim array(arraySize)

        With frmTrackerOfTime
            Select Case LCase(whichArray)
                Case "dungeon"
                    ' Link the dungeon labels for displaying progress into an array
                    array(0) = .lblDekuTree
                    array(1) = .lblDodongosCavern
                    array(2) = .lblJabuJabusBelly
                    array(3) = .lblForestTemple
                    array(4) = .lblFireTemple
                    array(5) = .lblWaterTemple
                    array(6) = .lblSpiritTemple
                    array(7) = .lblShadowTemple
                    array(8) = .lblBottomOfTheWell
                    array(9) = .lblIceCavern
                    array(10) = .lblGerudoTrainingGround
                    array(11) = .lblGanonsCastle
                Case "overworld"
                    ' Link the labels for displaying progress into an array
                    array(0) = .lblKokiriForest
                    array(1) = .lblLostWoods
                    array(2) = .lblSacredForestMeadow
                    array(3) = .lblHyruleField
                    array(4) = .lblLonLonRanch
                    array(5) = .lblMarket
                    array(6) = .lblTempleOfTime
                    array(7) = .lblHyruleCastle
                    array(8) = .lblKakarikoVillage
                    array(9) = .lblGraveyard
                    array(10) = .lblDMTrail
                    array(11) = .lblDMCrater
                    array(12) = .lblGoronCity
                    array(13) = .lblZorasRiver
                    array(14) = .lblZorasDomain
                    array(15) = .lblZorasFountain
                    array(16) = .lblLakeHylia
                    array(17) = .lblGerudoValley
                    array(18) = .lblGerudoFortress
                    array(19) = .lblHauntedWasteland
                    array(20) = .lblDesertColossus
                    array(21) = .lblOutsideGanonsCastle
                    array(22) = .lblQuestGoldSkulltulas
                    array(23) = .lblQuestMasks
            End Select
        End With
        Return array
    End Function

    Public Function makePictureBoxArray(ByVal whichArray As String) As PictureBox()
        ' Label arrays
        Dim array() As PictureBox
        Dim arraySize As Byte = 0

        ' Determine size based on array
        Select Case LCase(whichArray)
            Case "bosskey"
                arraySize = 10
            Case "compass"
                arraySize = 9
            Case "equipment"
                arraySize = 20
            Case "inventory"
                arraySize = 23
            Case "map"
                arraySize = 9
            Case "quest"
                arraySize = 22
            Case "smallkey"
                arraySize = 7
            Case Else
                Return Nothing
        End Select

        ' Resize array
        ReDim array(arraySize)

        With frmTrackerOfTime
            Select Case LCase(whichArray)
                Case "bosskey"
                    ' Link the pictureboxes of each boss key into an array
                    array(3) = .pbxFoTBossKey
                    array(4) = .pbxFiTBossKey
                    array(5) = .pbxWTBossKey
                    array(6) = .pbxSpTBossKey
                    array(7) = .pbxShTBossKey
                    array(10) = .pbxGCBossKey
                Case "compass"
                    ' Link the pictureboxes of each compass into an array
                    array(0) = .pbxDTCompass
                    array(1) = .pbxDCCompass
                    array(2) = .pbxJBCompass
                    array(3) = .pbxFoTCompass
                    array(4) = .pbxFiTCompass
                    array(5) = .pbxWTCompass
                    array(6) = .pbxSpTCompass
                    array(7) = .pbxShTCompass
                    array(8) = .pbxBotWCompass
                    array(9) = .pbxICCompass
                Case "equipment"
                    ' Link the pictureboxes of each equipment item into an array
                    array(0) = .pbxQuiver
                    array(1) = .pbxBulletBag
                    array(2) = .pbxKokiriSword
                    array(3) = .pbxMasterSword
                    array(4) = .pbxBiggoronsSword
                    array(5) = .pbxWallet
                    array(6) = .pbxBombBag
                    array(7) = .pbxDekuShield
                    array(8) = .pbxHylianShield
                    array(9) = .pbxMirrorShield
                    array(10) = .pbxStoneOfAgony
                    array(11) = .pbxGauntlet
                    array(12) = .pbxKokiriTunic
                    array(13) = .pbxGoronTunic
                    array(14) = .pbxZoraTunic
                    array(15) = .pbxGerudosCard
                    array(16) = .pbxScale
                    array(17) = .pbxKokiriBoots
                    array(18) = .pbxIronBoots
                    array(19) = .pbxHoverBoots
                    array(20) = .pbxBrokenKnife
                Case "inventory"
                    ' Link the pictureboxes of each inventory item into an array
                    array(0) = .pbx01
                    array(1) = .pbx02
                    array(2) = .pbx03
                    array(3) = .pbx04
                    array(4) = .pbx05
                    array(5) = .pbx06
                    array(6) = .pbx07
                    array(7) = .pbx08
                    array(8) = .pbx09
                    array(9) = .pbx10
                    array(10) = .pbx11
                    array(11) = .pbx12
                    array(12) = .pbx13
                    array(13) = .pbx14
                    array(14) = .pbx15
                    array(15) = .pbx16
                    array(16) = .pbx17
                    array(17) = .pbx18
                    array(18) = .pbx19
                    array(19) = .pbx20
                    array(20) = .pbx21
                    array(21) = .pbx22
                    array(22) = .pbx23
                    array(23) = .pbx24
                Case "map"
                    ' Link the pictureboxes of each map into an array
                    array(0) = .pbxDTMap
                    array(1) = .pbxDCMap
                    array(2) = .pbxJBMap
                    array(3) = .pbxFotMap
                    array(4) = .pbxFiTMap
                    array(5) = .pbxWTMap
                    array(6) = .pbxSpTMap
                    array(7) = .pbxShTMap
                    array(8) = .pbxBotWMap
                    array(9) = .pbxICMap
                Case "quest"
                    ' Link the pictureboxes of each quest item into an array
                    array(0) = .pbxMedalForest
                    array(1) = .pbxMedalFire
                    array(2) = .pbxMedalWater
                    array(3) = .pbxMedalSpirit
                    array(4) = .pbxMedalShadow
                    array(5) = .pbxMedalLight
                    array(6) = .pbxMinuetOfForest
                    array(7) = .pbxBoleroOfFire
                    array(8) = .pbxSerenadeOfWater
                    array(9) = .pbxRequiemOfSpirit
                    array(10) = .pbxNocturneOfShadow
                    array(11) = .pbxPreludeOfLight
                    array(12) = .pbxZeldasLullaby
                    array(13) = .pbxEponasSong
                    array(14) = .pbxSariasSong
                    array(15) = .pbxSunsSong
                    array(16) = .pbxSongOfTime
                    array(17) = .pbxSongOfStorms
                    array(18) = .pbxStoneKokiri
                    array(19) = .pbxStoneGoron
                    array(20) = .pbxStoneZora
                    array(21) = .pbxStoneOfAgony
                    array(22) = .pbxGerudosCard
                Case "smallkey"
                    ' Link the pictureboxes of each small key into an array
                    array(0) = .pbxFoTSmallKey
                    array(1) = .pbxFiTSmallKey
                    array(2) = .pbxWTSmallKey
                    array(3) = .pbxSpTSmallKey
                    array(4) = .pbxShTSmallKey
                    array(5) = .pbxBotWSmallKey
                    array(6) = .pbxGTGSmallKey
                    array(7) = .pbxGCSmallKey
            End Select
        End With

        Return array
    End Function
End Module
