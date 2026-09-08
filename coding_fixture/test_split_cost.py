from split_cost import total_cost


def test_total_cost_charges_venue_fee_once():
    # 2명, 1인당 5,000원, 장소비 100,000원(고정, 1회만) -> 10,000 + 100,000 = 110,000
    assert total_cost(participant_count=2, per_person_cost=5000, venue_fee=100000) == 110000
