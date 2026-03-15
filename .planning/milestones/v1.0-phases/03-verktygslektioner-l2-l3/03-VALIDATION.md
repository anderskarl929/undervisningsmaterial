---
phase: 3
slug: verktygslektioner-l2-l3
status: draft
nyquist_compliant: false
wave_0_complete: false
created: 2026-03-12
---

# Phase 3 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | Manuell granskning + automatiserad filvalidering |
| **Config file** | Ingen - docx-validering + innehållsgranskning |
| **Quick run command** | `python resources/office-scripts/validate.py [fil].docx` |
| **Full suite command** | Validera alla .docx + .pptx + manuell granskning mot success criteria |
| **Estimated runtime** | ~10 seconds (filvalidering) + manuell granskning |

---

## Sampling Rate

- **After every task commit:** `python resources/office-scripts/validate.py [fil].docx` + manuell granskning
- **After every plan wave:** Alla filer validerade + inbördes koherens kontrollerad
- **Before `/gsd:verify-work`:** Full suite must be green
- **Max feedback latency:** 15 seconds

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|-----------|-------------------|-------------|--------|
| 03-01-01 | 01 | 1 | KRIT-01 | manual-only | Granska att lektion-2.docx har grundfråge-sektion + övningsaktivitet | N/A Wave 0 | ⬜ pending |
| 03-01-02 | 01 | 1 | KRIT-02 | manual-only | Granska att lektion-2.docx har SIFT-introduktion + elevövning | N/A Wave 0 | ⬜ pending |
| 03-01-03 | 01 | 1 | KRIT-03 | manual-only | Granska att lektion-2.docx har Harvard-referering intro | N/A Wave 0 | ⬜ pending |
| 03-01-04 | 01 | 1 | KRIT-04 | manual-only | Granska att lektion-2.docx nämner omvänd bildsök/Google Lens | N/A Wave 0 | ⬜ pending |
| 03-02-01 | 02 | 1 | KONSP-01 | manual-only | Granska att lektion-3.docx stationsblad inkluderar mekanismer | N/A Wave 0 | ⬜ pending |
| 03-02-02 | 02 | 1 | KONSP-04 | manual-only | Granska att lektion-3.docx har mekanismfokus i stationsarbete | N/A Wave 0 | ⬜ pending |
| 03-02-03 | 02 | 1 | AI-01 | manual-only | Granska stationsblad för explicit källkritisk fråga | N/A Wave 0 | ⬜ pending |
| 03-02-04 | 02 | 1 | KRIT-03 | manual-only | Granska att L3 exit ticket har Harvard-refereringsövning | N/A Wave 0 | ⬜ pending |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

*Existing infrastructure covers all phase requirements.* docx-generering och pptx-generering är etablerade i projektet med befintliga scripts och skills.

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| L2 introducerar de fyra grundfrågorna med hands-on övning | KRIT-01 | Pedagogiskt innehåll kräver mänsklig granskning | Öppna lektion-2.docx, verifiera att grundfrågorna presenteras + att övningsaktivitet finns |
| L2 introducerar SIFT/lateral reading med praktisk demo | KRIT-02 | Pedagogiskt innehåll kräver mänsklig granskning | Öppna lektion-2.docx, verifiera SIFT-stegen + lateral reading-demo |
| L2 intro Harvard + L3 exit ticket övar Harvard | KRIT-03 | Pedagogiskt innehåll kräver mänsklig granskning | Verifiera Harvard-intro i L2 + Harvard-del i L3 exit ticket |
| L2 visar omvänd bildsök som del av SIFT | KRIT-04 | Pedagogiskt innehåll kräver mänsklig granskning | Öppna lektion-2.docx, sök efter omvänd bildsök/Google Lens |
| L3 stationer analyserar psykologiska mekanismer | KONSP-01 | Pedagogiskt innehåll kräver mänsklig granskning | Öppna lektion-3.docx, verifiera att stationsbladen inkluderar mekanismer |
| L3 huvudfokus: AI förstärker konspirationsspridning | KONSP-04 | Pedagogiskt innehåll kräver mänsklig granskning | Verifiera att mekanismfokus genomsyrar stationsarbetet |
| L3 stationer har explicit SIFT/grundfråge-tillämpning på AI-innehåll | AI-01 | Pedagogiskt innehåll kräver mänsklig granskning | Verifiera att varje stationsblad har explicit källkritisk fråga |

---

## Validation Sign-Off

- [ ] All tasks have `<automated>` verify or Wave 0 dependencies
- [ ] Sampling continuity: no 3 consecutive tasks without automated verify
- [ ] Wave 0 covers all MISSING references
- [ ] No watch-mode flags
- [ ] Feedback latency < 15s
- [ ] `nyquist_compliant: true` set in frontmatter

**Approval:** pending
