"""Text content for build-cowork-usecases.py (policy, contracts, transcript, runbook).

All content is fictional: Contoso (distribution center Fargo), Northwind, Fabrikam, Adatum, Litware and
Proseware are Microsoft demo companies. People are the real users of the CDX demo tenant.
"""

# ---------------------------------------------------------------- freight audit policy
POLICY_TITLE = "Freight Audit Policy"
POLICY_SUBTITLE = "Contoso Distribution Center Fargo - Three-way matching of carrier invoices"
POLICY_NOTE = "Synthetic demo data - fictional organization."
POLICY = [
    ("1. Purpose and scope", [
        "This policy defines the evidence, tolerances and codes needed before a carrier invoice may be released for payment. "
        "It applies to every carrier invoice that cites a transport order (TO). The conclusion for each invoice line is "
        "taken from three sources: the approved transport order line, the proof of delivery (POD) and the carrier invoice.",
    ]),
    ("2. Matching keys", [
        "2.1 An invoice line is matched on exactly four keys: carrier ID, transport order number, transport order line and lane code.",
        "2.2 Never guess a missing or unclear key. Carrier names, amounts, weights or dates must not be used to find a match.",
        "2.3 If a match is not unambiguous, stop. The line is blocked with exception code E01 and goes to the AP owner for a decision.",
    ]),
    ("3. Tolerances", [
        "3.1 Rate: the invoiced rate per 100 kg must equal the agreed rate on the transport order line. A difference of more than 2.0 percent is E03.",
        "3.2 Fuel surcharge: must equal the percentage agreed on the transport order line. Any other percentage is E03.",
        "3.3 Weight: the invoiced chargeable weight may exceed the delivered weight on the POD by at most 3.0 percent. More is E07.",
        "3.4 Tax: the tax rate must equal the rate of the tax code on the transport order line. Otherwise E08.",
        "3.5 Currency: the invoice currency must equal the carrier's default currency. Otherwise E06. Do not convert currencies; report each currency separately.",
    ]),
    ("4. Controls", [
        "4.1 Duplicate screening: the same carrier, transport order and line must not be paid twice. A second invoice for an already invoiced line is E02, whatever its invoice number or date.",
        "4.2 Delivery: a line is payable only when the POD status is Confirmed. Pending or missing is E04.",
        "4.3 Carrier status: invoices from a carrier with status Blocked are E05, even if everything else matches.",
    ]),
    ("5. Status and exception codes", [
        "5.1 Allowed statuses per invoice line: RELEASED and BLOCKED. No other status may be used.",
        "5.2 A line is RELEASED only if no exception applies. Otherwise it is BLOCKED with every exception code that applies.",
        "5.3 Allowed exception codes: E01 no unambiguous match; E02 duplicate; E03 rate or fuel surcharge variance; E04 delivery not confirmed; "
        "E05 carrier blocked; E06 currency mismatch; E07 weight variance; E08 tax mismatch.",
    ]),
    ("6. Reporting", [
        "6.1 The control overview lists every invoice line with status, exception codes, invoiced amount and the evidence (TO line, POD, carrier master).",
        "6.2 Per currency, released amount plus blocked amount must equal the total invoiced amount.",
        "6.3 Blocked lines get a factual query to the carrier: what was found, which document it comes from, what is needed. No accusations, no commitments to pay.",
        "6.4 Nothing is released in the payment system and nothing is sent without approval of the AP owner.",
    ]),
]

# ---------------------------------------------------------------- carrier contracts
CONTRACTS = {
    "northwind": {
        "file": "contract_northwind_freight.pdf",
        "carrier": "Northwind Freight",
        "address": "Northwind Freight Inc., 410 Harbor Way, Minneapolis, MN",
        "clauses": [
            ("1. Term and renewal", "The agreement runs for 36 months from the effective date. It renews automatically for further periods of 36 months unless a party gives written notice at least 12 months before the end of the current period."),
            ("2. Rates", "Base rate for the lane Fargo-Minneapolis is USD 18.50 per 100 kg for less-than-truckload. Rates for other lanes follow Schedule A and are 6 to 9 percent below the market index."),
            ("3. Fuel surcharge", "A fuel surcharge applies as published by the carrier each week. The surcharge is not capped."),
            ("4. Payment terms", "Invoices are payable within 60 days of receipt of a correct invoice."),
            ("5. Price adjustment", "Rates increase by a fixed 2.0 percent on each anniversary of the effective date."),
            ("6. Service levels", "The carrier targets an on-time delivery rate of 94 percent per month. If the rate falls below 92 percent, the customer receives a service credit of 2 percent of that month's invoiced freight."),
            ("7. Liability and claims", "Liability for loss or damage is limited to USD 2.00 per kg of the affected goods. The carrier is not liable for indirect or consequential loss. Claims must be filed within 10 days of delivery."),
            ("8. Data protection", "Shipment and consignee data is processed in the United States. The carrier may use subprocessors without prior notice. Personal data breaches are reported within 7 days."),
            ("9. Termination and exit", "The customer may terminate for convenience only after month 24, with 6 months' written notice and a termination fee of 25 percent of the committed annual volume still outstanding. The customer commits to tender at least 90 percent of its Fargo-Minneapolis volume to the carrier."),
        ],
    },
    "fabrikam": {
        "file": "contract_fabrikam_logistics.pdf",
        "carrier": "Fabrikam Logistics",
        "address": "Fabrikam Logistics LLC, 77 Rail Yard Road, Chicago, IL",
        "clauses": [
            ("1. Term and renewal", "The agreement runs for 24 months from the effective date and then renews for periods of 12 months. Either party may object to a renewal with 90 days' written notice."),
            ("2. Rates", "Base rate for the lane Fargo-Minneapolis is USD 19.90 per 100 kg for less-than-truckload. Other lanes follow Schedule A and are at market index."),
            ("3. Fuel surcharge", "A fuel surcharge applies as published weekly. The surcharge is capped at 12 percent of the freight charge."),
            ("4. Payment terms", "Invoices are payable within 30 days of receipt of a correct invoice."),
            ("5. Price adjustment", "Rates may be adjusted once per year by the change in the consumer price index, but by no more than 3.0 percent."),
            ("6. Service levels", "The carrier commits to an on-time delivery rate of 98 percent per month. If the rate falls below 96 percent, the customer receives a service credit of 5 percent of that month's invoiced freight."),
            ("7. Liability and claims", "Liability for loss or damage is limited to the invoice value of the affected goods, up to USD 50,000 per claim. Claims must be filed within 30 days of delivery."),
            ("8. Data protection", "Shipment and consignee data is processed in the United States and in the European Union. The carrier maintains a list of subprocessors and informs the customer 30 days before any change. Personal data breaches are reported within 72 hours."),
            ("9. Termination and exit", "The customer may terminate for convenience at any time with 90 days' written notice and without a fee. There is no minimum volume commitment."),
        ],
    },
    "adatum": {
        "file": "contract_adatum_cargo.pdf",
        "carrier": "Adatum Cargo",
        "address": "Adatum Cargo B.V., Waalhaven 12, Rotterdam, Netherlands",
        "clauses": [
            ("1. Term and renewal", "The agreement runs for 24 months from the effective date and then renews for periods of 24 months unless a party gives written notice at least 6 months before the end of the current period."),
            ("2. Rates", "Base rate for the lane Fargo-Minneapolis is USD 19.20 per 100 kg for less-than-truckload. Other lanes follow Schedule A."),
            ("3. Fuel surcharge", "A fuel surcharge applies as published monthly. The surcharge is not capped."),
            ("4. Payment terms", "Invoices are payable within 45 days of receipt of a correct invoice."),
            ("5. Price adjustment", "Rates are adjusted on each anniversary by the change in the consumer price index plus 2.0 percentage points. There is no upper limit."),
            ("6. Service levels", "The carrier targets an on-time delivery rate of 96 percent per month. If the rate falls below 94 percent, the customer receives a service credit of 3 percent of that month's invoiced freight."),
            ("7. Liability and claims", "Liability for loss or damage is limited to USD 6.00 per kg of the affected goods. Claims must be filed within 21 days of delivery."),
            ("8. Data protection", "Shipment and consignee data may be processed by subprocessors in any country without prior notice. Personal data breaches are reported within 14 days."),
            ("9. Termination and exit", "The customer may terminate for convenience with 180 days' written notice and a fee of 10 percent of the preceding twelve months' spend."),
        ],
    },
}

# ---------------------------------------------------------------- steering workshop transcript
TRANSCRIPT_TITLE = "Fargo Expansion - Steering Workshop - Meeting Transcript"
TRANSCRIPT_META = [
    "Date: 24 March 2027  |  Duration: 09:00 - 12:30  |  Location: Contoso Fargo, Conference Room 2",
    "Participants: Sydney Mattos (chair, regional operations), Teresa Sac (logistics Fargo), Vance DeLeon (construction), "
    "Sonia Rees (safety and training), Billie Vester (IT).",
]
TRANSCRIPT = [
    ("Sydney", "Good morning. We have three and a half hours to settle the plan for the new wing. I want clear decisions today, and I want to know what is still open when we leave. Teresa, start with the carriers."),
    ("Teresa", "Dock 3 closes on April 27. From then on we lose two inbound doors for about eight weeks. Fabrikam can use Gate 4 from April 30. For the three days in between they need a slot at Dock 2, seven to nine in the morning."),
    ("Sydney", "Can we do that at Dock 2?"),
    ("Vance", "Yes, Dock 2 is free in that window. No construction traffic there."),
    ("Sydney", "Good. Decision: Fabrikam gets the Dock 2 slot from April 27 to April 29 and uses Gate 4 from April 30. Anyone against?"),
    ("Sonia", "No, that works."),
    ("Teresa", "Next point. Should we move the Dock 3 closure by two weeks? Some of us wanted more time for the carriers."),
    ("Vance", "I would not. The cable trays are already ordered for the closure date. Moving it by two weeks pushes electrical to the end of June and kills go-live."),
    ("Sydney", "Then we keep April 27. The trade-off is eight weeks with fewer doors against a go-live that stays on June 22. Decision made."),
    ("Teresa", "Northwind arrives at 5:30 in the morning. The vendor window starts at 6:00. If they move to 6:15 they lose their slot in Minneapolis."),
    ("Sonia", "I cannot sign off on trucks on site before the window without supervision. The lighting at Gate 4 is not good enough before 6."),
    ("Vance", "We could put up temporary floodlights. Two units, rented, about two thousand dollars for the month."),
    ("Sonia", "With floodlights and one marshal, I could accept a trial. Not an open-ended exception."),
    ("Sydney", "Then: a four-week trial from April 30, Northwind may arrive from 5:30 with floodlights and a marshal. Sonia reviews it on May 28 and decides whether it continues."),
    ("Sonia", "Agreed. I will review on May 28."),
    ("Teresa", "Saturday shifts for wave 1 are approved for May 2 and May 9. The problem is May 9. I still need four forklift drivers."),
    ("Sydney", "Where do we get them?"),
    ("Teresa", "Either temp agency or from the Minneapolis site. I do not know yet which is cheaper."),
    ("Sydney", "Okay, that stays open. We cannot decide without numbers. Teresa, get quotes."),
    ("Teresa", "I can try. I do not have a date for it yet."),
    ("Sydney", "Let's leave the date out for now."),
    ("Teresa", "Wave 1 is the A items only. Looking at the data from the last quarter, I think we should add the B items. Otherwise wave 2 is overloaded."),
    ("Sydney", "Then wave 1 covers A and B items. Everyone fine?"),
    ("Vance", "Fine for construction."),
    ("Sonia", "Safety is fine as long as the mezzanine is not part of wave 1."),
    ("Sydney", "The mezzanine stays out of wave 1. Decision: wave 1 is A and B items, no mezzanine."),
    ("Sydney", "One idea: we could run extra Sunday shifts to catch up. Teresa?"),
    ("Teresa", "Possible in principle."),
    ("Sonia", "Careful, our labor agreement limits Sunday work. Without the works council that does not work."),
    ("Sydney", "Right. Scrap the Sunday idea, it was only a thought. Do not minute it as a decision."),
    ("Vance", "Update from the building site. Electrical in the new wing will be finished on May 22, not May 15. That uses up a week of the buffer. Go-live on June 22 is still possible."),
    ("Sydney", "Is there anything that would make you say June 22 is not safe?"),
    ("Vance", "If the sprinkler test slips beyond June 1. At the moment the inspector has not confirmed a date."),
    ("Sonia", "I would be uneasy if the test is not done two weeks before go-live."),
    ("Sydney", "Then we need a go/no-go point. Decision: go/no-go for June 22 on June 8, chaired by me. Vance brings the sprinkler test result."),
    ("Vance", "I will ask the inspector again this week. I cannot promise a date from him."),
    ("Billie", "IT point: the scanners in the new wing. The Wi-Fi survey can only be done when the electrical work is complete, so not before May 22."),
    ("Sydney", "When do the scanners go live then?"),
    ("Billie", "June 1 is realistic. May 15 is not possible any more."),
    ("Sydney", "Decision: scanner go-live in the new wing moves from May 15 to June 1."),
    ("Billie", "I need the Wi-Fi survey as a task, though. I am not sure who owns it. Could be facilities."),
    ("Vance", "Not us, we only deliver the cabling."),
    ("Sydney", "Open then: who runs the Wi-Fi survey, and when. We need an owner next week."),
    ("Teresa", "Temporary staff for the peak after go-live: how much budget do we have?"),
    ("Sydney", "I would like to hold at 120 thousand dollars. But finance has not confirmed. Treat it as an open item, I will clarify."),
    ("Teresa", "Tailspin Toys is asking about the Dock 3 closure. Someone should tell them what changes for deliveries."),
    ("Sydney", "Yes. Who does it and when, we should settle that. Let's note it as open."),
    ("Sonia", "Safety training for the new wing: all forklift drivers need the new traffic rules before May 31. I will schedule sessions in the week of May 17."),
    ("Sydney", "Good: Sonia runs the safety training in the week of May 17, all drivers complete by May 31."),
    ("Sydney", "Let me summarize what I heard. I will send a note. Thanks everyone."),
]

# ---------------------------------------------------------------- disruption runbook
RUNBOOK_TITLE = "Fargo Outbound Disruption Playbook"
RUNBOOK_SUBTITLE = "Contoso Distribution Center Fargo - Response rules for incidents on the dock and gate"
RUNBOOK_NOTE = "Synthetic demo data - fictional organization."
RUNBOOK = [
    ("1. Severity levels", [
        "SEV-1: safety risk to people, or the whole outbound operation is stopped. Immediate call to the site lead.",
        "SEV-2: a gate, dock or system is down or degraded, trucks are delayed, no safety impact. First status within 30 minutes.",
        "SEV-3: local disruption with a workaround. Status at the end of the shift.",
    ]),
    ("2. Roles", [
        "Incident lead: Teresa Sac, logistics Fargo.",
        "IT on call (scanners, Wi-Fi, gate systems): Billie Vester.",
        "Safety: Sonia Rees. Informed on every SEV-1 and on every SEV-2 that involves trucks on site.",
        "Construction interface (new wing, Dock 3): Vance DeLeon.",
        "Regional operations: Sydney Mattos. Informed on every SEV-2 and SEV-1.",
    ]),
    ("3. Communication rules", [
        "A status update names impact, known facts, open questions and the time of the next update. Facts that are not confirmed are marked as open.",
        "Internal updates go to the logistics team channel and to the stakeholders by email. External customers are informed only after approval by the incident lead.",
        "Update cadence for SEV-2: every 45 minutes until the incident is resolved.",
        "Nothing is sent or posted without the approval of the incident lead.",
    ]),
    ("4. Customer commitments", [
        "Orders with a customer cutoff in the next two hours and a service penalty are escalated first (see the affected shipments list).",
        "Tailspin Toys and Woodgrove Bank have contractual on-time commitments of 95 percent per month.",
    ]),
    ("5. After the incident", [
        "A short review within two working days: timeline, impact in shipments, what worked, what to change.",
    ]),
]

TRIGGER_EN = ("[COWORK-DEMO] SEV-2: The gate scanners at Fargo Gate 4 have been down since 06:40. Trucks are queuing "
              "outside, no safety impact reported. Next update at 07:30.")
TRIGGER_DE = ("[COWORK-DEMO] SEV-2: Die Tor-Scanner an Fargo Tor 4 sind seit 06:40 ausgefallen. Die Lkw stauen sich "
              "vor dem Tor, keine Sicherheitsauswirkungen gemeldet. Nächstes Update um 07:30.")
