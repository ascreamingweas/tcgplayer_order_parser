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


def test_retro_frame_artifacts_set_is_premium():
    assert organizer.treatment_bucket(None, "The Brothers' War: Retro Frame Artifacts") == "Premium"
    assert organizer.treatment_bucket(None, "The Brothers' War") == "Regular"


def test_special_rarity_is_premium_but_not_tokens():
    assert organizer.treatment_bucket(None, "Time Spiral: Remastered", "S", "Talrand, Sky Summoner") == "Premium"
    assert organizer.treatment_bucket(None, "Time Spiral: Remastered", "R", "Extirpate") == "Regular"
    assert organizer.treatment_bucket(None, "Time Spiral: Remastered", "S", "Ape Token") == "Regular"


def test_bonus_sheet_set_type_is_premium():
    # Any Scryfall "masterpiece" (bonus sheet) printing, even from a set not listed by name.
    assert organizer.treatment_bucket(None, "Some Future Set: Bonus Sheet", "R", "X", "masterpiece") == "Premium"
    assert organizer.treatment_bucket(None, "Outlaws of Thunder Junction", "R", "X", "expansion") == "Regular"
    # Fallback by name when the lookup didn't return a set type.
    assert organizer.treatment_bucket(None, "Outlaws of Thunder Junction: Breaking News", "R", "X") == "Premium"


def test_innistrad_double_feature_is_premium():
    assert organizer.treatment_bucket(None, "Innistrad: Double Feature", "R", "Sigarda's Splendor", "draft_innovation") == "Premium"
    assert organizer.treatment_bucket(None, "Innistrad: Midnight Hunt", "R", "X", "expansion") == "Regular"


def test_future_sight_chip_matches_retro_frame():
    assert organizer.get_variant_style("Future Sight")[0] == organizer.get_variant_style("Retro Frame")[0]
    assert organizer.get_variant_style("Future Sight")[1] == "Future Sight"


def _card(**kw):
    base = dict(quantity=1, set_name="S", card_name="C", collector_number="1", rarity="R", condition="Near Mint",
                is_foil=False, price=0.0, total_price=0.0)
    return organizer.Card(**{**base, **kw})


def test_market_price_follows_the_finish():
    prices = {"usd": "1.50", "usd_foil": "4.00", "usd_etched": "9.00"}
    assert organizer.market_price(_card(prices=prices)) == 1.5
    assert organizer.market_price(_card(prices=prices, is_foil=True)) == 4.0
    assert organizer.market_price(_card(prices=prices, is_foil=True, variant="Foil Etched")) == 9.0
    assert organizer.market_price(_card(prices={"usd_foil": "3.00"})) == 3.0
    assert organizer.market_price(_card()) is None


def test_variant_badges_cover_every_treatment():
    assert [label for _, label in organizer.variant_badges("Promo / Extended Art")] == ["Promo", "Extended Art"]
    assert organizer.variant_badges(None) == []
