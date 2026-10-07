# Expected result: cowork-freight-invoice-check

| Invoice | Carrier | TO | Line | Status | Codes | Amount |
|---|---|---|---|---|---|---|
| NW-1042 | N100 | TO-27-0101 | 1 | RELEASED | - | 2,477.52 USD |
| NW-1042 | N100 | TO-27-0101 | 2 | RELEASED | - | 1,360.80 USD |
| NW-1043 | N100 | TO-27-0106 | 1 | BLOCKED | E03 | 2,116.80 USD |
| FL-7710 | F200 | TO-27-0102 | 1 | RELEASED | - | 4,054.80 USD |
| FL-7711 | F200 | TO-27-0102 | 1 | BLOCKED | E02 | 4,054.80 USD |
| AC-8821 | A300 | TO-27-0103 | 1 | RELEASED | - | 3,747.25 EUR |
| AC-8822 | A300 | TO-27-0103 | 2 | BLOCKED | E07 | 3,710.32 EUR |
| LW-3901 | L400 | TO-27-0104 | 1 | BLOCKED | E05 | 1,576.86 EUR |
| PC-5520-A | P500 | TO-27-0105 | 1 | BLOCKED | E04 | 2,006.40 USD |
| PC-5520-B | P500 | - | - | BLOCKED | E01 | 2,006.40 USD |

| Currency | Released | Blocked | Total |
|---|---|---|---|
| USD | 7,893.12 | 10,184.40 | 18,077.52 |
| EUR | 3,747.25 | 5,287.18 | 9,034.43 |

Decision cases for the user: NW-1043 (rate +5.9%), FL-7711 (duplicate), AC-8822 (weight +6.8% over POD), LW-3901 (carrier blocked), PC-5520-A (POD pending), PC-5520-B (no TO number, do not guess).
