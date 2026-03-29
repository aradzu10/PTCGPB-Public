; ============================================================
; Globals
; ============================================================
global AllPokemonsCSV := A_ScriptDir . "\..\Resources\all_pokemon.csv"
global AllPokemons := ReadCSV(AllPokemonsCSV)

;-------------------------------------------------------------------------------
; TradeOrShareWithCard - Full flow: add friend, mark card, trade or share, clean up
;
; Parameters:
;   cardId     - ID of the card to trade or share
;   isShare    - True to share, false to trade
;   skipFriend - Skip AddFriends step if friend already added
;-------------------------------------------------------------------------------
TradeOrShareWithCard(cardId, isShare := false, skipFriend := false) {
    HomeAndMission(true)

    if (!skipFriend) {
        added := AddFriends()
        if (!added) {
            MsgBox, 16, Error, Could not add friend for trading/sharing.
            return
        }
        HomeAndMission(true)
    }

    card := FindPosition(cardId)
    CreateStatusMessage("Id: " . cardId . ", pokemon: " . card.name . ", rarity: " . card.rarity)

    MarkCardForSending(card.name, card.rarity, card.setIndex, card.isLastSet, card.cardRow, card.cardCol)

    if (isShare)
        TradeShareCard(card.name)
    else
        TradeSendCard(card.name)

    RemoveMarkAfterSend(card.name, card.rarity, card.setIndex, card.isLastSet, card.cardRow, card.cardCol)
    RemoveFriends()
}

;-------------------------------------------------------------------------------
; MarkCardForSending - Open Dex, find card, mark as wishlist and favorite
;
; Parameters:
;   pokemonName    - Name of the pokemon to find
;   rarity         - Card rarity (e.g. "1d", "3d")
;   setPosition    - 1-based index of the set
;   isLastSet      - True if this card is in the last set
;   rowPosition    - Row within the set grid
;   columnPosition - Column within the set grid
;-------------------------------------------------------------------------------
MarkCardForSending(pokemonName, rarity, setPosition, isLastSet, rowPosition, columnPosition) {
    OpenDexFullDisplay()

    SearchPokemon("Dex", pokemonName, rarity)

    SelectCard(setPosition, isLastSet, rowPosition, columnPosition)

    ; == Mark card to wishlist and favorite ==
    ; If wishlist isn't marked, click it
    if (!FindOrLoseImage(257, 67, 265, 77, , "Trade\DexAddToWishlist", 0, failSafeTime)) {
        FindImageAndClick(257, 67, 265, 77, , "Trade\DexAddToWishlist", 261, 77, 1000)
    }

    ; If favorite isn't marked, click it
    if (!FindOrLoseImage(227, 67, 230, 71, , "Trade\DexCardOnFavorite", 0, failSafeTime)) {
        FindImageAndClick(227, 67, 230, 71, , "Trade\DexCardOnFavorite", 227, 77, 1000)
    }

    ; Close card
    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Closing card`n(" . failSafeTime . "/90 seconds)")

        adbClick(139, 501)
        Delay(3)

        if (FindOrLoseImage(161, 225, 168, 234, , "Trade\CardDexDisplayAll", 0, failSafeTime)
        || FindOrLoseImage(259, 74, 275, 83, , "Trade\CardDexAfterCloseCard", 0, failSafeTime)) {
            Delay(3)
            break
        }
    }
}

;-------------------------------------------------------------------------------
; TradeSendCard - Navigate trade UI and send a card to the first friend
;
; Parameters:
;   pokemonName - Name of the pokemon card to trade
;-------------------------------------------------------------------------------
TradeSendCard(pokemonName) {
    ; Go to trade (do tutorial if needed)
    FindImageAndClick(120, 500, 155, 530, , "Social", 143, 518, 1000)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Opening trade menu`n(" . failSafeTime . "/90 seconds)")

        ; Click trade
        adbClick(207, 400)
        Delay(1)
        if (FindOrLoseImage(254, 155, 261, 160, , "Trade\TradeInsideTrade", 0, failSafeTime)) {
            break
        }

        ; Click on tutorial if exists
        adbClick(177, 446)
        Delay(1)
    }

    Loop {
        ; Click on trade card button
        FindImageAndClick(255, 156, 266, 161, , "Trade\TradeSelectFriendPage", 145, 437, 1000)
        Delay(1)

        found := false
        Loop 3 {
            ; Select first friend to trade with
            adbClick(218, 191)
            Delay(3)

            if (FindOrLoseImage(157, 163, 181, 180, , "Trade\TradeChooseCardPage", 0, 0)) {
                found := true
                Delay(3)
                break
            }

            Delay(3)
        }
        if (found) {
            break
        }

        ; Friend request probably not accepted yet — refresh by closing tab
        FindImageAndClick(254, 155, 261, 160, , "Trade\TradeInsideTrade", 139, 511, 1000)
        Delay(15)
    }

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Loading trade selection`n(" . failSafeTime . "/90 seconds)")

        Delay(3)
        if (FindOrLoseImage(157, 163, 181, 180, , "Trade\TradeChooseCardPage", 0, failSafeTime)) {
            Delay(3)
            break
        }

        ; In case of tutorial
        adbClick(138, 470)
    }

    SearchPokemon("Trade", pokemonName, , true, true)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Selecting card to trade`n(" . failSafeTime . "/90 seconds)")

        ; Click first card
        adbClick(55, 463)
        Delay(1)

        ; If card is in user wishlist, needs to click on different coordinates
        adbClick(55, 370)
        Delay(1)

        ; Click ok
        adbClick(153, 472)
        if (FindOrLoseImage(132, 444, 137, 453, , "Trade\TradeAfterCardChoose", 0, failSafeTime)) {
            break
        }
        Delay(3)
    }

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Confirming trade selection`n(" . failSafeTime . "/90 seconds)")

        ; Click ok for card trade
        adbClick(204, 464)
        Delay(1)

        ; Click ok 2
        adbClick(203, 376)
        Delay(1)

        ; Click ok 3
        adbClick(201, 394)
        Delay(1)

        ; Click ok 4
        adbClick(135, 450)
        Delay(1)

        if (FindOrLoseImage(255, 386, 259, 390, , "Trade\TradeRefresh", 0, failSafeTime)) {
            break
        }
    }

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Waiting for friend to send card`n(" . failSafeTime . "/300 seconds)")

        ; Refresh trade
        adbClick(224, 383)

        ; Wait for friend to send card
        if (FindOrLoseImage(99, 475, 103, 487, , "Trade\FriendSendTrade", 0, failSafeTime)) {
            break
        }
        Delay(10)
    }

    ; Click to accept trade
    FindImageAndClick(261, 308, 271, 315, , "Trade\ClickOnTrade", 205, 465, 1000)

    ; Click to confirm
    FindImageAndClick(234, 122, 243, 125, , "Trade\SwipeCard", 208, 371, 1000)

    ; Reduce speed to make sure swipe is registered
    GameSpeed(1)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Swiping trade card`n(" . failSafeTime . "/90 seconds)")

        ; Swipe card away
        adbSwipe("466 821 466 114 300")
        Delay(3)

        ; Click to proceed
        adbClick(143, 499)
        Delay(3)

        if (FindOrLoseImage(240, 103, 246, 136, , "Trade\TradeSwipeCardDone", 0, failSafeTime)) {
            break
        }

        Delay(3)
    }

    GameSpeed(3)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Finishing trade`n(" . failSafeTime . "/90 seconds)")

        ; Click skip cards
        adbClick(259, 512)
        Delay(0.2)

        ; Next
        adbClick(145, 496)

        if (FindOrLoseImage(142, 518, 147, 527, , "Trade\TradeComplete", 0, failSafeTime)) {
            break
        }
        Delay(3)
    }

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Completing trade`n(" . failSafeTime . "/90 seconds)")

        ; Thanks
        adbClick(167, 443)
        Delay(3)

        ; Close
        adbClick(146, 512)
        Delay(3)

        if (FindOrLoseImage(120, 500, 155, 530, , "Social", 0, failSafeTime)) {
            break
        }
    }
}

;-------------------------------------------------------------------------------
; TradeShareCard - Navigate share UI and share a card with the first friend
;
; Parameters:
;   pokemonName - Name of the pokemon card to share
;-------------------------------------------------------------------------------
TradeShareCard(pokemonName) {
    ; Go to trade (do tutorial if needed)
    FindImageAndClick(120, 500, 155, 530, , "Social", 143, 518, 1000)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Opening share menu`n(" . failSafeTime . "/90 seconds)")

        ; Click on share at social page
        adbClick(100, 400)
        Delay(3)
        if (FindOrLoseImage(22, 303, 31, 311, , "Trade\SharePage", 0, failSafeTime)) {
            break
        }

        ; Click on tutorial if exists
        adbClick(177, 446)
        Delay(3)
    }

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Selecting share partner`n(" . failSafeTime . "/90 seconds)")

        ; Click on share card button
        FindImageAndClick(255, 156, 266, 161, , "Trade\TradeSelectFriendPage", 147, 433, 1000)
        Delay(1)

        found := false
        Loop 3 {
            ; Select first friend to share with
            adbClick(218, 191)
            Delay(8)

            if (FindOrLoseImage(259, 136, 269, 156, , "Trade\ShareChooseCardPage", 0, failSafeTime)) {
                found := true
                Delay(3)
                break
            }

            Delay(3)
        }
        if (found) {
            break
        }

        ; Friend request probably not accepted yet — refresh by closing tab
        Delay(3)
        FindImageAndClick(22, 303, 31, 311, , "Trade\SharePage", 139, 511, 1000)
        Delay(10)
    }

    ; In case of tutorial
    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Handling share tutorial`n(" . failSafeTime . "/90 seconds)")

        if (FindOrLoseImage(259, 136, 269, 156, , "Trade\ShareChooseCardPage", 0, failSafeTime)) {
            Delay(3)
            break
        }

        adbClick(138, 470)
        Delay(3)
    }

    SearchPokemon("Share", pokemonName, , true, true)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Selecting card to share`n(" . failSafeTime . "/90 seconds)")

        ; Click first card
        adbClick(55, 463)
        Delay(1)

        ; If card is in user wishlist, needs to click on different coordinates
        adbClick(55, 370)
        Delay(1)

        ; Click ok
        adbClick(153, 472)
        if (FindOrLoseImage(207, 404, 262, 480, , "Trade\ShareCardSelected", 0, failSafeTime)) {
            break
        }
        Delay(3)
    }

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Confirming card share`n(" . failSafeTime . "/90 seconds)")

        ; Share
        adbClick(147, 460)
        Delay(3)

        ; Ok
        adbClick(196, 384)
        Delay(3)

        ; Ok
        adbClick(184, 387)
        Delay(3)

        if (FindOrLoseImage(247, 102, 253, 132, , "Trade\ShareCardSwipe", 0, failSafeTime)) {
            break
        }
        Delay(3)
    }

    ; Reduce speed to make sure swipe is registered
    GameSpeed(1)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Swiping share card`n(" . failSafeTime . "/90 seconds)")

        ; Swipe card away
        adbSwipe("466 800 466 114 300")
        Delay(3)

        if (FindOrLoseImage(234, 453, 271, 464, , "Trade\ShareCardDone", 0, failSafeTime)) {
            break
        }

        Delay(3)
    }

    GameSpeed(3)
}

;-------------------------------------------------------------------------------
; DexTutorial - Click through Dex tutorial screens until display-all is visible
;-------------------------------------------------------------------------------
DexTutorial() {
    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Handling Dex tutorial`n(" . failSafeTime . "/90 seconds)")

        ; Wait for display all, and click on button
        Loop 2 {
            adbClick(193, 193)
            Delay(5)
            if (FindOrLoseImage(141, 219, 147, 244, , "Trade\DexDisplayAll", 0, failSafeTime)) {
                return
            }
        }

        ; If found one of tutorial images
        if (FindOrLoseImage(182, 407, 188, 415, , "Trade\DexTutor1", 0, failSafeTime)
            || FindOrLoseImage(213, 442, 214, 451, , "Trade\DexTutor2", 0, failSafeTime)
        || FindOrLoseImage(220, 405, 223, 417, , "Trade\DexTutor3", 0, failSafeTime)) {

            Loop 3 {
                ; Click for dex tutorial ok buttons
                adbClick(166, 453)
                Delay(3)

                if (FindOrLoseImage(141, 219, 147, 244, , "Trade\DexDisplayAll", 0, failSafeTime)) {
                    break
                }

                ; Click for another dex tutorial ok buttons
                adbClick(167, 478)
                Delay(3)

                if (FindOrLoseImage(141, 219, 147, 244, , "Trade\DexDisplayAll", 0, failSafeTime)) {
                    break
                }

                ; Click for third dex tutorial ok buttons
                adbClick(169, 448)
                Delay(3)

                if (FindOrLoseImage(141, 219, 147, 244, , "Trade\DexDisplayAll", 0, failSafeTime)) {
                    break
                }
                Delay(3)
            }
        }
    }
}

;-------------------------------------------------------------------------------
; OpenDexFullDisplay - Navigate to Dex and ensure "display all" is turned off
;-------------------------------------------------------------------------------
OpenDexFullDisplay() {
    ; Click on dex
    FindImageAndClick(237, 192, 255, 198, , "Trade\DexPage", 92, 525, 1000)
    Delay(5)

    DexTutorial()

    ; Click turn off display, unless already off
    if (!FindOrLoseImage(249, 226, 255, 231, , "Trade\CardDexDisplayAllOff", 0, failSafeTime)) {
        FindImageAndClick(249, 226, 255, 231, , "Trade\CardDexDisplayAllOff", 247, 229, 1000)
        Delay(1)
    }
}

;-------------------------------------------------------------------------------
; SearchPokemon - Open search bar in Dex/Trade/Share and filter by name/rarity
;
; Parameters:
;   place       - "Dex", "Trade", or "Share"
;   pokemonName - Name to search for
;   rarity      - Optional rarity filter ("1d", "2d", "3d", "4d")
;   wishlist    - Apply wishlist filter if true
;   favorite    - Apply favorite filter if true
;-------------------------------------------------------------------------------
SearchPokemon(place, pokemonName, rarity := "", wishlist := false, favorite := false) {
    if (place = "Dex") {
        xClick := 252
        yClick := 187
    } else if (place = "Trade") {
        xClick := 252
        yClick := 150
    } else if (place = "Share") {
        xClick := 252
        yClick := 134
    } else {
        MsgBox, 16, Error, Invalid place for SearchPokemon: " . place
            return
    }

    ; Click on search
    FindImageAndClick(32, 149, 39, 154, , "Trade\CardDexSearch", xClick, yClick, 1000)
    Delay(1)

    ; Clear search if there is anything applied
    adbClick(220, 508)
    Delay(1)

    ; Click on search bar and type name
    FindImageAndClick(237, 506, 244, 511, , "Trade\DexSearchBar", 138, 148, 1000)
    adbInput("""" . pokemonName . """")
    Delay(1)

    if (rarity = "1d") {
        FindImageAndClick(50, 440, 224 - 221 + 50, 443, , "Trade\DexRaritySelected", 50, 440, 1000)
    } else if (rarity = "2d") {
        FindImageAndClick(110, 440, 224 - 221 + 110, 443, , "Trade\DexRaritySelected", 110, 440, 1000)
    } else if (rarity = "3d") {
        FindImageAndClick(180, 440, 224 - 221 + 180, 443, , "Trade\DexRaritySelected", 180, 440, 1000)
    } else if (rarity = "4d") {
        FindImageAndClick(235, 440, 224 - 221 + 235, 443, , "Trade\DexRaritySelected", 235, 440, 1000)
    }
    Delay(1)

    if (favorite) {
        FindImageAndClick(118, 244, 120, 248, , "Trade\SearchFavorite", 80, 237, 1000)
        Delay(1)
    }

    if (wishlist) {
        FindImageAndClick(104, 343, 106, 348, , "Trade\SearchWishlist", 77, 341, 1000)
        Delay(1)
    }

    if (place = "Dex") {
        ; Click ok and search
        FindImageAndClick(161, 225, 168, 234, , "Trade\CardDexDisplayAll", 142, 472, 1000)
    } else if (place = "Trade") {
        ; Click ok and search
        FindImageAndClick(157, 163, 181, 180, , "Trade\TradeChooseCardPage", 142, 472, 1000)
    } else if (place = "Share") {
        ; Click ok and search
        FindImageAndClick(259, 136, 269, 156, , "Trade\ShareChooseCardPage", 142, 472, 1000)
    }
    Delay(5)
}

;-------------------------------------------------------------------------------
; RemoveMarkAfterSend - Open Dex, find card, remove wishlist and favorite marks
;
; Parameters:
;   pokemonName    - Name of the pokemon
;   rarity         - Card rarity
;   setPosition    - 1-based index of the set
;   isLastSet      - True if this card is in the last set
;   rowPosition    - Row within the set grid
;   columnPosition - Column within the set grid
;-------------------------------------------------------------------------------
RemoveMarkAfterSend(pokemonName, rarity, setPosition, isLastSet, rowPosition, columnPosition) {
    OpenDexFullDisplay()

    SearchPokemon("Dex", pokemonName, rarity)

    SelectCard(setPosition, isLastSet, rowPosition, columnPosition)

    ; == Remove card from wishlist and favorite ==
    ; If wishlist isn't removed, click it
    if (!FindOrLoseImage(260, 68, 267, 76, , "Trade\DexRemoveFromWishlist", 0, failSafeTime)) {
        FindImageAndClick(260, 68, 267, 76, , "Trade\DexRemoveFromWishlist", 261, 77, 1000)
    }
    Delay(1)

    if (!FindOrLoseImage(258, 68, 268, 76, , "Trade\DexRemoveFromWishlist2", 0, failSafeTime)) {
        FindImageAndClick(258, 68, 268, 76, , "Trade\DexRemoveFromWishlist2", 258, 76, 1000)
    }
    Delay(1)

    ; If favorite isn't removed, click it
    if (!FindOrLoseImage(230, 69, 232, 74, , "Trade\DexRemoveFavorite", 0, failSafeTime)) {
        FindImageAndClick(230, 69, 232, 74, , "Trade\DexRemoveFavorite", 227, 77, 1000)
    }
    Delay(1)

    ; Close card
    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Closing card`n(" . failSafeTime . "/90 seconds)")

        adbClick(139, 501)
        Delay(3)

        if (FindOrLoseImage(161, 225, 168, 234, , "Trade\CardDexDisplayAll", 0, failSafeTime)
        || FindOrLoseImage(259, 74, 275, 83, , "Trade\CardDexAfterCloseCard", 0, failSafeTime)) {
            Delay(3)
            break
        }
    }
}

;-------------------------------------------------------------------------------
; SelectCard - Scroll to and click a specific card in the Dex
;
; Parameters:
;   setPosition    - 1-based index of the card's set
;   isLastSet      - True if this is the last set (uses special scroll logic)
;   rowPosition    - Row within the set grid (1-based)
;   columnPosition - Column within the set grid (1-based)
;-------------------------------------------------------------------------------
SelectCard(setPosition, isLastSet, rowPosition, columnPosition) {
    setPosition := setPosition - 1 ; convert to 0-based index

    ; TODO: handles more than 1 row in sets.
    cards_location_x := [40, 90, 140, 190, 240]
    if (!isLastSet) {
        ; Swipe to the correct set
        Loop % (setPosition // 2) {
            adbSwipe("19 728 19 410 3000")
            Delay(1)
        }

        ; Click the set and define Y coordinates based on odd/even position
        if Mod(setPosition, 2) {
            adbClick(138, 370)
            card_location_y := [470]
        } else {
            adbClick(59, 300)
            card_location_y := [375, 445]
        }
    } else {
        ; Scroll to the end of the page for the last set
        failSafe := A_TickCount
        failSafeTime := 0
        Loop {
            failSafeTime := (A_TickCount - failSafe) // 1000
            CreateStatusMessage("Scrolling to last set`n(" . failSafeTime . "/90 seconds)")

            adbSwipe("19 728 19 410 3000")
            if (FindOrLoseImage(269, 442, 273, 444, , "Trade\DexEndOfPage", 0, failSafeTime)) {
                break
            }
            Delay(1)
        }

        ; Click on set
        adbClick(100, 415)

        card_location_y := [475]
    }

    ; Execute the card click
    xClick := cards_location_x[columnPosition]
    yClick := card_location_y[rowPosition]
    Delay(5)

    failSafe := A_TickCount
    failSafeTime := 0
    Loop {
        failSafeTime := (A_TickCount - failSafe) // 1000
        CreateStatusMessage("Selecting card`n(" . failSafeTime . "/90 seconds)")

        adbClick(xClick, yClick)
        Delay(1)

        if (FindOrLoseImage(194, 73, 206, 83, , "Trade\DexCardOpened", 0, failSafeTime)
        || FindOrLoseImage(244, 431, 248, 442, , "Trade\DexCardOpened2", 0, failSafeTime)) {
            break
        }
    }
}

;-------------------------------------------------------------------------------
; GetPokeInfo - Get name, set code, card index, and rarity for a pokemon ID
;
; Parameters (ByRef):
;   pokemonId    - The ID to look up
;   outName      - Receives the pokemon name
;   outSetCode   - Receives the set code
;   outCardIndex - Receives the card index within the set
;   outRarity    - Receives the rarity string
;-------------------------------------------------------------------------------
GetPokeInfo(pokemonId, ByRef outName, ByRef outSetCode, ByRef outCardIndex, ByRef outRarity) {
    global AllPokemons
    outName := ""
    outSetCode := ""
    outCardIndex := ""
    outRarity := ""
    Loop, % AllPokemons.MaxIndex()
    {
        row := AllPokemons[A_Index]
        if (row["id"] = pokemonId) {
            outName := row["pokemon_name"]
            outSetCode := row["set_code"]
            outCardIndex := row["card_index"]
            outRarity := row["rarity"]
            return
        }
    }
}

;-------------------------------------------------------------------------------
; FindPosition - Find Dex grid position for a pokemon ID
;
; Returns object: {name, rarity, setIndex, isLastSet, cardRow, cardCol}
;
; Parameters:
;   targetID - The pokemon ID to locate
;-------------------------------------------------------------------------------
FindPosition(targetID) {
    global AllPokemons

    GetPokeInfo(targetID, targetName, targetSetCode, targetCardIdx, targetRarity)
    if (targetName = "")
        return ""

    ; Collect all pokemons with same name and rarity, track unique sets
    matches := []
    setLookup := {}
    setStr := ""

    Loop, % AllPokemons.MaxIndex()
    {
        row := AllPokemons[A_Index]
        if (!InStr(row["pokemon_name"], targetName))
            Continue

        if (row["rarity"] != targetRarity)
            Continue

        matches.Push({id: row["id"], set_code: row["set_code"], card_index: row["card_index"]})

        sc := row["set_code"]
        if (!setLookup.HasKey(sc)) {
            setLookup[sc] := 1
            setStr .= sc . "`n"
        }
    }

    ; Sort sets descending
    Sort, setStr, R
    sets := []
    Loop, Parse, setStr, `n, `r
        if (A_LoopField != "")
        sets.Push(A_LoopField)

    ; Find index of our set in sorted list
    setIndex := 0
    Loop, % sets.MaxIndex()
    {
        if (sets[A_Index] = targetSetCode) {
            setIndex := A_Index
            Break
        }
    }

    ; Collect cards in our set, build sortable string: zero-padded idx|id
    scStr := ""
    Loop, % matches.MaxIndex()
    {
        m := matches[A_Index]
        if (m.set_code = targetSetCode)
            scStr .= Format("{:06}", m.card_index) . "|" . m.id . "`n"
    }

    ; Sort by card_index ascending
    Sort, scStr
    setCards := []
    Loop, Parse, scStr, `n, `r
    {
        if (A_LoopField = "")
            Continue
        parts := StrSplit(A_LoopField, "|")
        setCards.Push({id: parts[2], idx: parts[1] + 0})
    }

    ; Find card position in sorted set
    cardPos := 0
    Loop, % setCards.MaxIndex()
    {
        if (setCards[A_Index].id = targetID) {
            cardPos := A_Index
            Break
        }
    }

    ; Row and column (5 per row, 1-based)
    if (cardPos > 0) {
        cardRow := Ceil(cardPos / 5)
        cardCol := Mod(cardPos - 1, 5) + 1
    } else {
        cardRow := 0
        cardCol := 0
    }

    ; Last set check (only if more than 2 sets)
    setCount := sets.MaxIndex()
    isLastSet := (setCount > 2 && setIndex = setCount)

    return {name: targetName, rarity: targetRarity, setIndex: setIndex, isLastSet: isLastSet, cardRow: cardRow, cardCol: cardCol}
}

;-------------------------------------------------------------------------------
; GameSpeed - Set emulator game speed (only active when setSpeed = 3)
;
; Parameters:
;   newSpeed - Target speed: 1 (normal) or 3 (fast)
;-------------------------------------------------------------------------------
GameSpeed(newSpeed) {
    global setSpeed

    if (setSpeed != 3)
        return

    ; Open menu
    FindImageAndClick(158, 252, 177, 259, , "speedmodMenu", 18, 109, 2000)
    if (newSpeed = 1) {
        FindImageAndClick(20, 170, 24, 174, , "One", 21, 172)
    } else if (newSpeed = 3) {
        FindImageAndClick(187, 168, 191, 174, , "Three", 187, 172)
    }
    ; Minimize menu
    adbClick_wbb(51, 297)
}
