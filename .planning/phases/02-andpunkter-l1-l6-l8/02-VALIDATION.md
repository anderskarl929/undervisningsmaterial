---
phase: 2
slug: andpunkter-l1-l6-l8
status: draft
nyquist_compliant: false
wave_0_complete: false
created: 2026-03-12
---

# Phase 2 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | Manuell granskning + automatiserad filvalidering |
| **Config file** | Ingen - docx-validering + innehallsgranskning |
| **Quick run command** | `node generate-lektion-X.js && ls -la lektion-X.docx` |
| **Full suite command** | Validera alla tre .docx + manuell granskning mot success criteria |
| **Estimated runtime** | ~10 seconds (generering), manuell granskning ~5 min |

---

## Sampling Rate

- **After every task commit:** Kör generate-script och verifiera att .docx skapas korrekt
- **After every plan wave:** Alla genererade .docx validerade + inbördes koherens kontrollerad
- **Before `/gsd:verify-work`:** Full suite must be green
- **Max feedback latency:** 10 seconds (automatiserad), 5 min (manuell)

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|-----------|-------------------|-------------|--------|
| 02-01-01 | 01 | 1 | KONSP-02 | manual-only | Granska att lektion-1.docx har AI-labb-aktivitet med skapande + reflektion | N/A Wave 0 | ⬜ pending |
| 02-01-02 | 01 | 1 | KONSP-03 | manual-only | Granska att lektion-1.docx refererar aktuella 2025-2026 case | N/A Wave 0 | ⬜ pending |
| 02-02-01 | 02 | 1 | BED-02 | manual-only | Granska att lektion-6.docx har summativ skrivuppgift med E/C/A-rubrik | N/A Wave 0 | ⬜ pending |
| 02-02-02 | 02 | 1 | BED-02 | manual-only | Granska att lektion-8.docx har tilldelade roller, tredjepersons-framing, privat uppföljningsplan | N/A Wave 0 | ⬜ pending |
| 02-02-03 | 02 | 1 | KONSP-03 | manual-only | Granska att alla tre lektionsplaner refererar återkommande case | N/A Wave 0 | ⬜ pending |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

Existing infrastructure covers all phase requirements — docx-generering är etablerat med befintliga generate-scripts som mall.

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| L1 innehåller AI-labb med inokulering | KONSP-02 | Pedagogiskt innehåll kräver manuell bedömning | Öppna lektion-1.docx, verifiera att fas 4 innehåller AI-labb-aktivitet med elevernas eget skapande + reflektionsfas |
| Aktuella 2025-2026 exempel | KONSP-03 | Aktualitet kräver manuell verifiering | Kontrollera att alla tre lektionsplaner refererar minst ett aktuellt svenskt/internationellt case |
| L6 summativ skrivuppgift | BED-02 | Bedömningskvalitet kräver manuell granskning | Öppna lektion-6.docx, verifiera att rubrik tillåter E/C/A-differentiering |
| L8 equity-safeguards | BED-02 | Pedagogisk säkerhet kräver manuell granskning | Öppna lektion-8.docx, verifiera tilldelade roller, tredjepersons-framing, privat uppföljningsplan |

---

## Validation Sign-Off

- [ ] All tasks have `<automated>` verify or Wave 0 dependencies
- [ ] Sampling continuity: no 3 consecutive tasks without automated verify
- [ ] Wave 0 covers all MISSING references
- [ ] No watch-mode flags
- [ ] Feedback latency < 10s (automated)
- [ ] `nyquist_compliant: true` set in frontmatter

**Approval:** pending
