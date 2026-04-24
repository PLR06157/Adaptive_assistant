# Adaptive OTC Transformation Offer — Combined Adaptive + Emagia Timeline Description

## Purpose of this file
This file consolidates the current **combined Adaptive + Emagia** timeline logic for the OTC transformation offer for an **Emagia-enabled process transformation**.

It is intended to preserve:
- the agreed Adaptive activity structure
- the agreed Emagia pilot and rollout structure
- the sequencing and overlap logic across both parties
- the current week-based estimates
- the additional assumptions introduced from the Emagia pricing proposal
- the conflicts that were identified and resolved
- the remaining items that are still intentionally left open

This file reflects the latest working version after the most recent clarifications.

---

## Working scope and framing

### Current scope focus
At this stage, the timeline covers a **combined joint timeline** where:
- **Adaptive activities** and
- **Emagia activities**

are shown together in one integrated model.

### Timeline scope included for Adaptive
The Adaptive timeline currently includes:
- diagnostic and transformation design
- governance
- change management
- rollout support
- testing coordination

### Timeline scope included for Emagia
The Emagia timeline currently includes:
- pilot implementation for selected functionalities in one pilot country
- rollout implementation waves by functionality
- cutover and go-live readiness for pilot and each rollout wave
- stabilization and hypercare governance for pilot and each rollout wave

### Timeline endpoint
The combined timeline runs through:
- continued support through stabilization of the final rollout wave

### Timeline design principle
The timeline is being built first as a **high-level timeline**.

The total duration should **not** be assumed upfront. It should instead be **derived from the activities, timing, overlaps, implementation durations, and people involved**.

### Date handling principle
The updated timeline should be **agnostic to the start date**.

This means:
- all activities should show **week-based durations and offset logic**
- the chart can later be re-based to any chosen start date
- once the start date is defined, the corresponding calendar dates can then be populated

### Presentation principle
The timeline is intended to be presented in a **Gantt-chart style** with:
- separate swimlanes
- one swimlane per Emagia stream
- the existing Adaptive activity structure preserved in the same overall visual logic

---

## Rollout model

### Overall rollout shape
The rollout should be shown as:
- one pilot for a selected country first
- then scale-up preparation
- then multiple rollout waves by Emagia functionality

### Pilot structure
The pilot should be structured as:
- one selected country
- two separate functional pilot tracks with shared pilot governance logic

Specifically:
- **Cash Application Pilot**
- **Collections Pilot**

### Pilot country
- pilot country: **to be confirmed**

### Rollout structure after pilot
After pilot completion, the rollout should be shown as **multiple rollout waves**.

Each rollout wave represents **one Emagia functionality**.

### Rollout country coverage
Each rollout wave should represent:
- **all countries including the pilot country**

This means the pilot country is included again in the rollout-wave representation.

---

## Initial staffing concept captured

The following initial idea was captured earlier, but it is **not yet validated** and should not be treated as final:
- Senior Transformation Director — 0.5 FTE
- OTC SME / Business Analyst — 2–3 FTEs
- Organizational Change Manager — 1 FTE

The agreed logic is that final staffing should be **derived from activities and timing**, not fixed upfront.

### Updated staffing implication from Emagia rollout
Adaptive is expected to continue supporting the overall transformation through the Emagia rollout waves.

However:
- **no additional Adaptive task bars should be added** for the rollout waves in the timeline
- the impact of this continued support should instead be reflected later in the **final FTE calculation**

---

## Accepted activity groups

### Adaptive activity groups
The following Adaptive activity groups remain accepted for the high-level timeline:

1. **mobilization and program setup**
2. **current-state diagnostic**
3. **process and operating model assessment**
4. **future-state OTC design**
5. **pilot scope and country definition**
6. **business requirements and design authority**
7. **governance and decision-making setup**
8. **change impact and stakeholder alignment**
9. **pilot readiness planning**
10. **testing coordination and business validation**
11. **cutover and go-live readiness**
12. **stabilization and hypercare governance**
13. **scale-up preparation and roadmap**

At this point, they are still kept as separate activity groups and are **not yet merged** for executive presentation.

### Emagia pilot streams
The following Emagia pilot streams are accepted:
- **Cash Application Pilot**
- **Collections Pilot**

### Emagia rollout waves
The following Emagia rollout waves are accepted using simplified executive labels:
1. **Cash Application Rollout**
2. **Collections Rollout**
3. **Disputes / Deductions / Reporting Rollout**
4. **E-Invoicing Rollout**
5. **Credit Management Rollout**
6. **Order Management Rollout**

### Emagia post-implementation gates
The following Emagia post-implementation gates are accepted:
- **pilot cutover and go-live readiness**
- **pilot stabilization and hypercare governance**
- **rollout cutover and go-live readiness** for each rollout wave
- **rollout stabilization and hypercare governance** for each rollout wave

---

## Structural timeline decisions

### Pre-pilot Adaptive activities
The following Adaptive activity groups should be completed before the first pilot:
1. mobilization and program setup
2. current-state diagnostic
3. process and operating model assessment
4. future-state OTC design
5. pilot scope and country definition
6. business requirements and design authority
7. governance and decision-making setup
8. change impact and stakeholder alignment

### Placement of 5
- **5. pilot scope and country definition** should be placed **between 2 and 3**.

### Formal start point of the timeline
- the timeline starts **after formal kick-off**

### Start point of Emagia pilot work
- Emagia pilot work starts with the start of **9. pilot readiness planning**

### Main joint testing stage for the pilot
- **10. testing coordination and business validation** remains the **main joint testing stage for the pilot**
- no separate Emagia pilot testing activity needs to be added as a distinct standalone timeline bar at this stage

### Shared pilot go-live logic
The pilot tracks should converge into:
- **one shared pilot go-live point**

This shared pilot go-live point should occur:
- **after the pilot cutover and go-live readiness block**
- **only after both pilot implementations are complete**

### Post-pilot gating before rollout
Before the first rollout wave starts, the following must occur in sequence:
1. both pilot implementations are complete
2. shared pilot cutover and go-live readiness is completed
3. shared pilot go-live is reached
4. pilot stabilization and hypercare governance is completed
5. **13. scale-up preparation and roadmap** is completed

### Rollout start gate
The first rollout wave may start only:
- **after pilot stabilization is fully finished**
- **after scale-up preparation and roadmap is fully finished**

### Rollout wave sequencing structure
The rollout is shown as six functional waves in this order:
1. Cash Application Rollout
2. Collections Rollout
3. Disputes / Deductions / Reporting Rollout
4. E-Invoicing Rollout
5. Credit Management Rollout
6. Order Management Rollout

---

## Sequencing and overlap logic

### Confirmed Adaptive start and end logic
- **3. process and operating model assessment** starts **after 5**
- **6. business requirements and design authority** starts **during 3**
- **4. future-state OTC design** starts **during the final part of 3**
- **6** ends **at the same time as 4**
- **7. governance and decision-making setup** starts **during 1** and runs through **11**
- **8. change impact and stakeholder alignment** starts **during 3** and runs through **11**
- **9. pilot readiness planning** can run fully in parallel with **8**
- **9** ends at pilot-start readiness for the combined pilot structure
- **10. testing coordination and business validation** starts **immediately after 9**
- **10** ends **before 11**
- **11. cutover and go-live readiness** includes post-go-live command-center setup
- **12. stabilization and hypercare governance** starts **immediately after 11**
- **13. scale-up preparation and roadmap** starts after the pilot stage and must be completed before the first rollout wave begins

### Stage-gate treatment
All Adaptive activity groups are treated as **stage-gates**.

### Overlap allowed in Adaptive logic
Overlap is allowed for:
- 6
- 7
- 8
- 9

### Emagia pilot overlap logic
- **Cash Application Pilot** starts with the start of **9. pilot readiness planning**
- **Collections Pilot** starts after **1/3 of the Cash Application Pilot duration** has elapsed
- the pilot tracks are separate streams but share one overall pilot completion logic
- the shared pilot cutover starts **immediately after the later of the two pilot implementations finishes**
- the shared pilot go-live marker is placed **after the pilot cutover and go-live readiness block**
- pilot stabilization starts **immediately after the shared pilot go-live**
- pilot cutover and pilot stabilization are each shown as **single shared blocks**, not split by functionality

### Emagia rollout overlap logic
- **Cash Application Rollout** starts first after the required gates are completed
- each next rollout wave starts after **1/3 of the previous wave’s implementation duration only**
- this **1/3 rule applies only to the implementation duration**, not to cutover or stabilization
- later rollout waves **can overlap** with the previous wave’s cutover and stabilization activities

### Emagia rollout cutover and stabilization logic
For each rollout wave:
- **cutover and go-live readiness** is shown as **additional time after the implementation duration**
- **stabilization and hypercare governance** is shown as **additional time after cutover**
- rollout go-live markers are **not required to be shown at this stage**

### Updated clarification on 13
Earlier, 13 had been shown as ending at the same time as 12.
This was later corrected.

The currently preserved logic is:
- **13 starts during 11** in the earlier Adaptive-only logic
- in the updated combined model, **13 must finish before the first rollout wave starts**
- **13 is no longer forced to end at the same time as 12**

---

## Scope clarifications for later-stage activities

### 10. testing coordination and business validation
This covers:
- business / UAT coordination
- the main joint testing stage for the pilot

This does **not** cover:
- SIT support from Adaptive side
- a separately drawn Emagia testing bar for the pilot at this stage

### 11. cutover and go-live readiness
This includes:
- business go-live readiness
- deployment coordination
- post-go-live command-center setup

### 12. stabilization and hypercare governance
This includes:
- pilot stabilization
- early KPI tracking
- lessons learned capture for future scale-up

### Emagia pilot cutover and go-live readiness
This includes:
- final pilot deployment preparation
- shared pilot cutover across both pilot streams
- readiness for the shared pilot go-live point

### Emagia pilot stabilization and hypercare governance
This includes:
- immediate post-pilot stabilization after shared pilot go-live
- hypercare for the pilot country / pilot scope
- closure of the pilot before scale-up begins

### Emagia rollout cutover and go-live readiness
For each rollout wave this includes:
- wave-specific deployment readiness
- cutover planning and execution
- go-live readiness activities

### Emagia rollout stabilization and hypercare governance
For each rollout wave this includes:
- post-wave stabilization
- hypercare governance
- early performance monitoring during the wave close-out period

---

## Relative duration logic captured earlier

The earlier relative duration logic captured was:

### Pre-pilot ranking
- longest pre-pilot: **4. future-state OTC design**
- second-longest pre-pilot: **6. business requirements and design authority**
- third-longest pre-pilot: **3. process and operating model assessment**
- fourth-longest pre-pilot: **8. change impact and stakeholder alignment**

### Post-pilot ranking in earlier Adaptive-only logic
- longest post-pilot: **12. stabilization and hypercare governance**
- second-longest post-pilot: **10. testing coordination and business validation**

### Relative categories captured earlier
- 1 short
- 2 medium
- 3 long
- 4 clearly longer than 3
- 5 short
- 6 slightly shorter than 4
- 7 medium
- 8 medium
- 9 short
- 10 long and clearly longer than 12
- 11 medium
- 12 medium
- 13 medium

### Additional relative comparisons captured earlier
- 3 clearly longer than 2
- 2 slightly longer than 8
- 12 slightly longer than 13
- 13 same length as 11
- 7 same length as 11
- 9 shorter than 7
- 1 same length as 9
- 5 same length as 1

### Note on preserved earlier relative logic
The earlier relative comparisons are preserved for reference where still useful.

However, where a direct newer week-based decision has been made, the newer specific decision should take precedence.

---

## Current week-based estimates

The following week-based model is the current working version.

### Locked from direct answers
- **3. process and operating model assessment** — **2 weeks**
- **4. future-state OTC design** — **3 weeks**
- **6. business requirements and design authority** — **2 weeks**
- **7. governance and decision-making setup** — **2 weeks**

### Additional clarified estimate
- **2. current-state diagnostic** — **1–2 weeks**

### Corrected item
The previous estimate of:
- **8. change impact and stakeholder alignment = 10 weeks**

was explicitly identified for correction.

Therefore:
- the previous 10-week estimate is superseded / rejected
- **8. change impact and stakeholder alignment** is now set to **12 weeks**

### Current Adaptive working estimates
1. **mobilization and program setup** — **1 week**
2. **current-state diagnostic** — **1–2 weeks**
3. **process and operating model assessment** — **2 weeks**
4. **future-state OTC design** — **3 weeks**
5. **pilot scope and country definition** — **1 week**
6. **business requirements and design authority** — **2 weeks**
7. **governance and decision-making setup** — **2 weeks**
8. **change impact and stakeholder alignment** — **12 weeks**
9. **pilot readiness planning** — **1 week**
10. **testing coordination and business validation** — **4 weeks**
11. **cutover and go-live readiness** — **2 weeks**
12. **stabilization and hypercare governance** — **3 weeks**
13. **scale-up preparation and roadmap** — **2 weeks**

### Emagia pilot implementation durations
The pilot should be shown as a **shorter country-specific implementation** than the full rollout timing.

The agreed pilot durations are:
- **Cash Application Pilot** — **9 weeks**
- **Collections Pilot** — **8 weeks**

### Emagia rollout implementation durations
The rollout durations should follow the Emagia pricing proposal where available.

The agreed rollout implementation durations are:
- **Cash Application Rollout** — **17 weeks**
- **Collections Rollout** — **15 weeks**
- **Disputes / Deductions / Reporting Rollout** — **15 weeks**
- **E-Invoicing Rollout** — **8 weeks**
- **Credit Management Rollout** — **15 weeks**
- **Order Management Rollout** — **15 weeks**

### Emagia cutover and stabilization durations
For the pilot and for each rollout wave, the agreed additional durations are:
- **cutover and go-live readiness** — **1.5 weeks**
- **stabilization and hypercare governance** — **2.5 weeks**

These are shown as **additional time after implementation**, not included inside the implementation durations.

---

## Conflict log and resolution status

### Conflict 1: 2 vs 8
An inconsistency emerged because:
- earlier relative logic said **2 should be slightly longer than 8**
- later, **8** was estimated at **10 weeks**
- meanwhile **2** was clarified as **around 1–2 weeks**

#### Resolution status
This conflict was resolved by deciding that:
- the **10-week estimate for 8** should be corrected
- **2** remains at **around 1–2 weeks**
- **8** is now set to **12 weeks** in the updated working model

So the current position is:
- **2 remains valid as 1–2 weeks**
- **8 is now updated to 12 weeks**

### Conflict 2: 13 end point vs 12
An inconsistency emerged because the earlier model contained both:
- **13 same length as 11**
- **13 ends at the same time as 12**
- **12 slightly longer than 13**

These could not all hold true together.

#### Resolution status
This conflict was resolved by deciding that:
- the rule **“13 ends at the same time as 12”** should be corrected

So the current position is:
- **13 same length as 11** is still preserved only as an earlier relative reference
- **13 is no longer forced to end at the same time as 12**
- in the combined model, **13 must be completed before the first rollout wave starts**

---

## Current logical flow of the timeline

### Adaptive pre-pilot flow
1. **mobilization and program setup** starts after formal kick-off
2. **governance and decision-making setup** starts during mobilization
3. **current-state diagnostic** follows
4. **pilot scope and country definition** is placed between diagnostic and assessment
5. **process and operating model assessment** starts after pilot scope definition
6. **change impact and stakeholder alignment** starts during process assessment
7. **business requirements and design authority** starts during process assessment
8. **future-state OTC design** starts during the final part of process assessment
9. **business requirements and design authority** ends at the same time as future-state design
10. **pilot readiness planning** runs in parallel with the ongoing change / stakeholder work and provides the start point for Emagia pilot activity
11. **testing coordination and business validation** acts as the main joint testing stage for the pilot

### Emagia pilot flow
12. **Cash Application Pilot** starts with the start of pilot readiness planning
13. **Collections Pilot** starts after 1/3 of the Cash Application Pilot duration has elapsed
14. both pilot streams continue as separate swimlanes
15. the pilot waits for completion of both pilot implementations
16. shared **pilot cutover and go-live readiness** starts immediately after the later pilot implementation finishes
17. shared **pilot go-live** occurs after the pilot cutover block
18. shared **pilot stabilization and hypercare governance** starts immediately after the shared pilot go-live

### Transition to rollout
19. **scale-up preparation and roadmap** is shown after the pilot stage and must finish before the first rollout wave starts
20. the first rollout wave starts only after both:
   - pilot stabilization is fully complete
   - scale-up preparation and roadmap is complete

### Emagia rollout flow
21. **Cash Application Rollout** starts first
22. **Collections Rollout** starts after 1/3 of the Cash Application Rollout implementation duration
23. **Disputes / Deductions / Reporting Rollout** starts after 1/3 of the Collections Rollout implementation duration
24. **E-Invoicing Rollout** starts after 1/3 of the Disputes / Deductions / Reporting Rollout implementation duration
25. **Credit Management Rollout** starts after 1/3 of the E-Invoicing Rollout implementation duration
26. **Order Management Rollout** starts after 1/3 of the Credit Management Rollout implementation duration
27. each rollout wave has its own **cutover and go-live readiness** after implementation
28. each rollout wave then has its own **stabilization and hypercare governance** after cutover
29. later rollout waves may overlap with the previous wave’s cutover and stabilization activities

### Adaptive support logic during rollout
30. Adaptive continues supporting the overall transformation through the rollout period
31. however, no new Adaptive rollout task bars are added in the timeline at this stage
32. the corresponding impact is intended to be reflected later in the final staffing / FTE model

---

## Items still unresolved

The following points remain open and should be completed in the next iteration:

1. final choice within the **1–2 week range** for **2. current-state diagnostic**
2. confirmation of whether any other provisional Adaptive week estimates should be adjusted
3. total overall **combined Adaptive + Emagia timeline duration** derived from the final week model and overlaps
4. final staffing / FTE model derived from activity volume and overlaps, including ongoing Adaptive support through rollout
5. later decision on whether some activity groups should be merged for executive presentation
6. final re-basing of the chart once a formal start date is chosen

---

## Preservation note
This file is designed to preserve:
- prior decisions
- corrected logic
- superseded assumptions
- newly added Emagia agreements
- unresolved items still visible for later refinement
- the requirement to keep the timeline start-date agnostic until the final calendar base date is chosen

The goal is to ensure that no captured timeline information is lost before the model is translated into a final executive timeline and staffing view.
