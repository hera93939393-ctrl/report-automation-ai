"""split_cost.py — 코딩 시나리오 실습용 결함 코드. 참가자별 비용 합계에
장소 고정비용(인원과 무관하게 1회만 청구)을 더한 총 비용을 계산한다.

의도된 결함: venue_fee가 participant_count만큼 곱해지고 있다(고정비용인데
인원수만큼 중복 청구됨)."""


def total_cost(participant_count: int, per_person_cost: int, venue_fee: int) -> int:
    return participant_count * per_person_cost + participant_count * venue_fee
