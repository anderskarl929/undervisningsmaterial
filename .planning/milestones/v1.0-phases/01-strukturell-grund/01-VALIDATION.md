---
phase: 1
slug: strukturell-grund
status: draft
nyquist_compliant: false
wave_0_complete: false
created: 2026-03-12
---

# Phase 1 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | Manuell granskning (pedagogiskt innehåll) |
| **Config file** | Ingen - dokumentgranskningsbaserad validering |
| **Quick run command** | `cat [fil] | head -50` (snabbkoll av format) |
| **Full suite command** | Manuell granskning mot success criteria |
| **Estimated runtime** | ~2 minuter per dokument |

---

## Sampling Rate

- **After every task commit:** Snabbkoll att skapade .md-filer har rätt format och frontmatter
- **After every plan wave:** Alla dokument koherenta med varandra
- **Before `/gsd:verify-work`:** Alla success criteria verifierade
- **Max feedback latency:** Omedelbar (manuell granskning)

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|-----------|-------------------|-------------|--------|
| 1-01-01 | 01 | 1 | BED-01 | manual-only | Granska att exit-ticket-mall.md har alla 8 lektioner med fråga, lärandemål och retrieval review-koppling | N/A Wave 0 | ⬜ pending |
| 1-01-02 | 01 | 1 | BED-02 | manual-only | Granska att bedömningskriterier.md har explicit E/C/A för muntlig examination + equity-sektion | N/A Wave 0 | ⬜ pending |
| 1-01-03 | 01 | 1 | BED-03 | manual-only | Granska att tredjepersons-framing.md har definition + minst 4 exempelformuleringar | N/A Wave 0 | ⬜ pending |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

*Existing infrastructure covers all phase requirements. Fas 1 producerar Markdown-dokument, ingen testkod behövs.*

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| Momentplan med 8 lektioner, mål, innehåll, progression | BED-01 | Pedagogiskt innehåll kräver mänsklig bedömning | Verifiera att momentplan.md listar alla 8 lektioner med tydliga mål, innehåll och progressionslogik |
| E/C/A-kriterier med explicit C-nivå | BED-02 | Bedömningskvalitet kräver ämnesexpertis | Verifiera att E/C/A finns för både skriftlig och muntlig examination, att C-nivå är explicit |
| Exit ticket-mall kopplad till retrieval practice | BED-01 | Pedagogisk koherens kräver mänsklig granskning | Verifiera att varje exit ticket har fråga + lärandemål + koppling till nästa lektions retrieval review |
| Tredjepersons-framing med exempelformuleringar | BED-03 | Tonalitet och ämneskänslighet kräver lärarbedömning | Verifiera att dokumentet har definition, minst 4 konkreta exempelformuleringar, och riktlinjer för känsliga ämnen |

---

## Validation Sign-Off

- [ ] All tasks have manual verification instructions
- [ ] Sampling continuity: varje dokument granskas vid skapande
- [ ] Wave 0 covers all MISSING references
- [ ] No watch-mode flags
- [ ] Feedback latency: omedelbar
- [ ] `nyquist_compliant: true` set in frontmatter

**Approval:** pending
