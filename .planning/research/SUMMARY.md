# Project Research Summary

**Project:** Kallkritik: Konspirationsteorier och AI
**Domain:** Undervisningsmoment (gymnasiet, Samhallskunskap 3) - pedagogisk design, inte mjukvara
**Researched:** 2026-03-10
**Confidence:** HIGH

## Executive Summary

This is a 7-lesson teaching module (+exam) on source criticism, conspiracy theories, and AI-generated content for Swedish upper secondary students. The research base is strong and points clearly in one direction: inoculation (prebunking) is the primary strategy, not debunking. The Uppsala University study (2025, n=459 Swedish gymnasiet students) proves that isolated interventions have zero lasting effect - meaning every lesson must include retrieval practice and the SIFT method must be practiced repeatedly across contexts, not taught once. The pedagogical stack is Rosenshine's explicit instruction (6-fas model already in use) combined with inoculation theory from Cambridge and the SIFT lateral reading method from Stanford.

The module's architecture follows a Foundation-Application-Synthesis arc across three phases. The critical transition is Lesson 3 to Lesson 4, where scaffolding drops and students must apply tools independently. The devil's advocate review and architecture research both flag this as the point where the module succeeds or fails. A gateway check (L3 exit ticket) must determine whether students are ready for independent application. The existing change plan (replacing L5 debate with written analysis) is architecturally sound - it breaks oral-oral fatigue and creates a written rehearsal before summative assessment.

The dominant risk is not pedagogical failure but psychological harm: treating conspiracy theories with ridicule triggers identity-protective cognition and backfire effects in students who hold or are sympathetic to these beliefs (Jerome 2024, n=7691 teachers). Four critical pitfalls - the ridicule trap, cynicism overcorrection, exposure effect, and identity-threatening confrontation - must be actively prevented from Lesson 1 onward. The module's strongest differentiator is the AI-conspiracy intersection (Lesson 3) and the inoculation-through-creation approach (AI lab in Lesson 1), neither of which exist in comparable Swedish teaching modules. The main gap is C-level differentiation and the need for current (2025-2026) Swedish-language examples.

## Key Findings

### Recommended Stack

The "stack" is pedagogical, not technical. Three core frameworks form the foundation, each with strong evidence bases. See [STACK.md](STACK.md) for full details.

**Core frameworks:**
- **Rosenshine's Principles / 6-fas model**: Lesson structure and sequencing - already implemented in /planera-moment, scaffolding d=0.82
- **Inoculation Theory (Prebunking)**: Primary strategy for conspiracy theory resilience - medium-to-large effect sizes vs. small effects for debunking (van der Linden et al. 2020, Cambridge lab)
- **SIFT Method (Lateral Reading)**: Practical source evaluation - teaches what professional fact-checkers actually do, replaces outdated checklist approaches like CRAAP
- **UNESCO Conspiracy Theory Teacher Guide**: Classroom strategies for respectful engagement - the most comprehensive practitioner guide available
- **Gy11/Gy25 Curriculum Alignment**: Direct mandate for "kallkritisk granskning" in digital and other forms

**Critical research constraint:** Uppsala study proves isolated interventions fail. This means spaced retrieval practice across all 7 lessons is non-negotiable.

### Expected Features

See [FEATURES.md](FEATURES.md) for full landscape including dependency graph.

**Must have (table stakes):**
- Four source criticism questions + SIFT/lateral reading integrated
- Hands-on exercises with current (2025-2026) real-world examples
- Progression from tool-building (L1-3) to application (L4-7)
- Formative assessment pipeline (exit tickets feeding retrieval practice)
- Conspiracy theories treated with intellectual respect (not dismissal)
- Clear summative assessment criteria communicated from Lesson 1
- E/C/A differentiation (C-level currently missing - must be added)
- Harvard referencing introduced early (L2) and practiced throughout
- Swedish context in examples (not just American/English cases)

**Should have (differentiators):**
- Inoculation through AI content creation (AI lab, "switch sides")
- Metacognitive reflection ("When have YOU believed something without checking?")
- Psychological mechanisms behind conspiracy belief (not just "which ones exist")
- AI + conspiracy theory intersection (how AI amplifies conspiracy spread)
- Scaffolding fade with explicit per-student triggers
- Varied activity formats beyond EPA (gallery walks, speed-dating rotation, silent discussions)

**Defer:**
- Bad News game in classroom (AI lab provides similar effect, simpler)
- Extensive psychology deep-dive on conspiratorial thinking (brief intro in L2 sufficient)
- Digital tool deep-dives (Deepware Scanner, Video Authenticator - tools change too fast)

### Architecture Approach

The module follows a three-phase arc: Foundation (L1-3), Application (L4-5), Synthesis (L6-7+exam). See [ARCHITECTURE.md](ARCHITECTURE.md) for full component boundaries and dependency graph.

**Major components:**
1. **Foundation Phase (L1-3)** - Build the analytical toolkit: source criticism questions, SIFT, conspiracy mechanisms, Harvard referencing, AI-conspiracy intersection. High scaffolding. Theory/practice ratio moves from 40/60 to 30/70.
2. **Application Phase (L4-5)** - Use tools with decreasing support: seminar (oral) then written perspective analysis. Scaffolding available on request but not distributed by default. Theory/practice ratio 20/80 to 10/90.
3. **Synthesis Phase (L6-7+exam)** - Demonstrate mastery: summative written assessment, reflection/self-assessment, seminar exam. Minimal to no scaffolding. Practice-dominant.
4. **Formative Pipeline** - Exit tickets every lesson feed into next lesson's retrieval practice and serve as early warning system for the two summative endpoints (L6 written, L8 seminar).

**Key architectural pattern:** Build order should follow backward design - define assessment criteria and summative tasks (L6, L8) BEFORE building instruction (L2-L5). The momentplan is the master document everything depends on.

### Critical Pitfalls

See [PITFALLS.md](PITFALLS.md) for all 11 pitfalls with detection strategies.

1. **The Ridicule Trap** - Treating conspiracy theories as stupid triggers identity-protective cognition and destroys intellectual honesty. Prevention: "granska, inte avfarda" as mantra, steel-manning, mixed examples including verified conspiracies. Must be established in L1.
2. **Cynicism Overcorrection** - Over-emphasizing "question everything" produces epistemically paralyzed students who trust nothing. Prevention: always pair deconstruction with reconstruction ("So where CAN we find reliable information?"). Highest risk in L3-5.
3. **Exposure Effect** - Presenting conspiracy theories in detail can inadvertently spread them. Prevention: prebunking before exposure (order matters), focus on structural patterns not specific content, pair all conspiratorial content with analytical framework simultaneously.
4. **Identity-Threatening Confrontation** - Forcing public confrontation with believing students. Prevention: third-person framing only ("Why might someone find this convincing?"), never ask personal beliefs, private follow-up plan. 40% of teachers encounter this.
5. **AI Demo Decay** - AI examples become outdated within months. Prevention: refresh all AI examples before each teaching cycle, build around principles not specific examples.

## Implications for Roadmap

Based on research, the build order follows backward design and the dependency graph from ARCHITECTURE.md.

### Phase 1: Structural Foundation
**Rationale:** Everything depends on the master plan and assessment criteria being defined first. Backward design (define the destination before building the journey) is the strongest pattern from the architecture research.
**Delivers:** Updated momentplan reflecting all three devil's advocate changes, E/C/A assessment criteria matrix with the missing C-level added, and the epistemic culture guidelines (ridicule trap prevention).
**Addresses:** Differentiation gap (C-level), summative assessment clarity, conspiracy engagement norms.
**Avoids:** Pitfall 1 (ridicule trap) by codifying respectful engagement norms from the start; Pitfall 8 (seminar equity) by designing assessment criteria before lesson plans.

### Phase 2: Bookends (L1 + L6 + L8)
**Rationale:** Define the opening experience and the assessment endpoints before filling in the middle. L1 sets the epistemic tone for the entire module. L6 and L8 define what mastery looks like, which determines what L2-L5 must teach.
**Delivers:** L1 lesson plan with current 2025-2026 examples, L6 summative writing task with rubric, L8 seminar exam structure with equity safeguards.
**Addresses:** Table stakes (current examples, clear assessment), differentiators (AI lab/inoculation in L1).
**Avoids:** Pitfall 3 (exposure effect) by establishing inoculation framework in L1 before any conspiracy content; Pitfall 8 (seminar equity) by building equity into exam design.

### Phase 3: Foundation Lessons (L2 + L3)
**Rationale:** These are the toolkit-building lessons that depend on knowing both L1's framework and L6's assessment requirements. L2 introduces conspiracy mechanics and Harvard referencing. L3 is the critical integration lesson combining AI + conspiracy analysis.
**Delivers:** L2 and L3 lesson plans, Harvard referencing spiral (intro in L2, practice in L3 exit ticket), gateway check mechanism for the L3-L4 transition.
**Addresses:** Table stakes (SIFT integration, Harvard referencing, hands-on exercises), differentiators (AI+conspiracy intersection, psychological mechanisms).
**Avoids:** Pitfall 6 (false equivalence) by explicitly teaching misinformation/disinformation taxonomy in L3; Pitfall 7 (framework overload) by limiting to 2-3 core frameworks practiced repeatedly.

### Phase 4: Application Lessons (L4 + L5)
**Rationale:** These depend on all foundation content being finalized. L4 seminar uses L3's content. L5 written analysis builds on L4's arguments and serves as dress rehearsal for L6.
**Delivers:** L4 seminar plan with structured equity safeguards, L5 written perspective analysis task, scaffolding fade specifications per lesson.
**Addresses:** Differentiators (perspective analysis with assigned roles, scaffolding fade with explicit triggers, varied activity formats).
**Avoids:** Pitfall 2 (cynicism overcorrection) by pairing deconstruction with reconstruction in L4-5; Pitfall 4 (identity threat) by maintaining third-person framing in seminar.

### Phase 5: Closing (L7 + Momentoversikt)
**Rationale:** L7 references the entire module journey and cannot be written until all other lessons exist. The HTML student overview depends on all lessons being finalized.
**Delivers:** L7 reflection/self-assessment plan with specific metacognitive prompts, momentoversikt HTML for students.
**Addresses:** Differentiators (metacognitive reflection), table stakes (student-facing overview).
**Avoids:** Pitfall 10 (ignoring emotional dimension) by including genuine reflection on belief change, not superficial "we learned a lot."

### Phase Ordering Rationale

- **Backward design drives the order:** Assessment (L6, L8) defined before instruction (L2-L5) ensures alignment. This is the strongest recommendation from the architecture research.
- **Dependency graph respected:** Momentplan -> Assessment criteria -> Bookends -> Foundation -> Application -> Closing. No phase creates materials that depend on undefined components.
- **Pitfall prevention is front-loaded:** The epistemic culture (how to treat conspiracy theories respectfully) and the inoculation framework (prebunking before exposure) must be established before any conspiracy content appears.
- **The L3-L4 gateway is the critical joint:** If the gateway check shows students are not ready for independent application, L4 must have adjustable scaffolding levels. This must be designed into L4 during Phase 4, informed by the L3 exit ticket design from Phase 3.

### Research Flags

Phases likely needing deeper research during planning:
- **Phase 2 (L1):** Needs current 2025-2026 Swedish-language AI and conspiracy examples. The existing examples may be outdated. Requires web research for fresh material.
- **Phase 3 (L3):** The AI+conspiracy intersection lesson is the module's strongest differentiator but also the most novel - few existing models to follow. Needs careful design and possibly example testing with colleagues.
- **Phase 4 (L4 seminar):** Seminar equity design for conspiracy-sensitive topics is not well-documented in Swedish pedagogical literature. May need adaptation from UNESCO guidelines.

Phases with standard patterns (skip research-phase):
- **Phase 1 (Structural Foundation):** Momentplan updates and E/C/A matrices are well-established patterns in Swedish curriculum design.
- **Phase 5 (Closing):** Reflection and self-assessment lessons follow well-documented Rosenshine patterns.

## Confidence Assessment

| Area | Confidence | Notes |
|------|------------|-------|
| Stack | HIGH | Multiple meta-analyses, peer-reviewed research, directly relevant Swedish population study (Uppsala 2025). Inoculation > debunking is settled science. |
| Features | HIGH | Table stakes well-defined by curriculum (Gy11/Gy25) and research. Differentiators supported by Cambridge inoculation research and UNESCO guidelines. Devil's advocate review already validated feature set. |
| Architecture | HIGH | Three-phase arc follows established Rosenshine scaffolding principles. Backward design is standard practice. The formative pipeline is well-documented. |
| Pitfalls | HIGH | Critical pitfalls (ridicule trap, cynicism, exposure, identity threat) are documented across multiple peer-reviewed sources (Jerome 2024 n=7691, UNESCO 2022, Caulfield, danah boyd). Prevention strategies are concrete and actionable. |

**Overall confidence:** HIGH

### Gaps to Address

- **C-level differentiation:** The devil's advocate review identified that only E-level (support) and A-level (challenge) are specified. C-level expectations and scaffolding must be explicitly defined during Phase 1.
- **Current Swedish examples (2025-2026):** All AI-generated content examples and several conspiracy theory cases need refreshing. The "Pope in puffer jacket" (2023) is already dated. Need fresh Swedish-context examples at time of building.
- **AI demo shelf life:** AI capabilities change rapidly. Materials will need a refresh protocol - review all AI examples before each teaching cycle. This is an ongoing maintenance concern, not a one-time fix.
- **Seminar equity in conspiracy-sensitive context:** How to run an assessed oral seminar on conspiracy theories without triggering identity-protective cognition in believing students is under-documented. UNESCO guidelines provide principles but not Swedish classroom-specific protocols.
- **OECD AI Literacy Framework:** Still in draft (final version 2026). May need revisiting when finalized, but current scope (AI as source to evaluate, not technology to understand) is unlikely to change.

## Sources

### Primary (HIGH confidence)
- Jerome, L. et al. (2024). "Combatting conspiracies in the classroom." BERJ. n=7691.
- Nygren, T. et al. (2025). Uppsala University study on source criticism interventions. n=459 Swedish students.
- van der Linden, S. et al. (2020). Prebunking interventions. Harvard Kennedy School Misinformation Review.
- UNESCO (2022). "Addressing conspiracy theories: what teachers need to know."
- Rosenshine, B. (2012). "Principles of Instruction." American Educator.
- Caulfield, M. SIFT Method / Lateral Reading. Stanford History Education Group.
- Skolverket. Samhallskunskap amnesplan (Gy11/Gy25).
- Skolinspektionen (2024). Kallkritik granskningsrapport.
- danah boyd. "You Think You Want Media Literacy... Do You?" Data & Society.
- Bad News i svenska klassrum (2024). n=516 gymnasieelever. Taylor & Francis.

### Secondary (MEDIUM confidence)
- OECD/EC (2025). AI Literacy Framework (draft, final 2026).
- Bezalel (2022). "Of Conspiracy and Controversy - A Pedagogy of Conspiracy Theories."
- Swami et al. (2014) + 2024 replication failure on analytic thinking mechanism.
- Edutopia: Critical Literacy in the Age of AI.
- Digitala lektioner: Granska konspirationer.

### Tertiary (LOW confidence)
- None identified. The research base for this domain is unusually strong.

---
*Research completed: 2026-03-10*
*Ready for roadmap: yes*
