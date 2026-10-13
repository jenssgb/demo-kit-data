# Servicecenter operating rules

Fictional Contoso workflow. Scenario clock: 2026-10-14 09:00 Europe/Berlin.

Use the same case ID and shipment ID throughout. Operations owner: Teresa Sac; portal/IT owner: Billie Vester. Rank SLA deadlines first, then customer impact. No ETA or delivery guarantee without evidence. A carrier's DELIVERED flag is not a signed POD. Required POD keys: shipment, customer, delivery date, receiver acknowledgement. Allowed states: NEW, WAITING_PROOF, RESPONSE_READY, WAITING_IT, CLOSED. Customer acceptance is required to close a case. SC-103 was accepted by Northwind on 2026-10-13 at 16:00; preserve its closure. Use dated updates over old register fields; record sources and next action. Send only after review and specific approval.
