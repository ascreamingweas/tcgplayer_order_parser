import mtg_packing_slip_organizer as organizer


def test_white_border_is_premium():
    assert organizer.treatment_bucket("White Border") == "Premium"
    assert organizer.treatment_bucket("Foil Etched") == "Regular"
    assert organizer.treatment_bucket(None) == "Regular"


def test_white_border_from_pull_sheet_lands_in_premium():
    pull = ('Product Line,Product Name,Condition,Number,Set,Rarity,Quantity,Main Photo URL,Set Release Date,SkuId,Order Quantity\n'
            '"Magic","Llanowar Elves (White Border)","Near Mint","1","Test Set","C","1","","","1","AAAA0000-111111-0000A:1"\n')
    cards, _ = organizer.parse_pull_sheet(pull)
    assert cards[0].variant == "White Border"
    assert organizer.treatment_bucket(cards[0].variant, cards[0].set_name) == "Premium"
