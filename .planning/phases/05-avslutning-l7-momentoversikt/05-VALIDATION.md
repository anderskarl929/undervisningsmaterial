---
phase: 5
slug: avslutning-l7-momentoversikt
status: draft
nyquist_compliant: false
wave_0_complete: false
created: 2026-03-15
---

# Phase 5 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | Manual validation (document generation) |
| **Config file** | none |
| **Quick run command** | `node generate-lektion-7.js && python resources/office-scripts/validate.py lektion-7.docx` |
| **Full suite command** | Validate all .docx + open HTML in browser + verify CSV encoding |
| **Estimated runtime** | ~10 seconds |

---

## Sampling Rate

- **After every task commit:** Validate .docx after generation, open HTML in browser
- **After every plan wave:** All 7 .docx present, HTML renders, CSV imports correctly
- **Before `/gsd:verify-work`:** Full suite must be green
- **Max feedback latency:** 10 seconds

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|-----------|-------------------|-------------|--------|
| 05-01-01 | 01 | 1 | MAT-01 | smoke | `ls lektion-7.docx` | ❌ W0 | ⬜ pending |
| 05-01-02 | 01 | 1 | MAT-01 | unit | `python resources/office-scripts/validate.py lektion-7.docx` | ✅ | ⬜ pending |
| 05-02-01 | 02 | 1 | MAT-03 | smoke | `test -f momentoversikt.html` | ❌ W0 | ⬜ pending |
| 05-02-02 | 02 | 1 | MAT-03 | manual-only | Visual: all 8 lessons listed in HTML | N/A | ⬜ pending |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

*Existing infrastructure covers all phase requirements.*

- validate.py exists for .docx validation
- docx npm library installed from prior phases
- `/html-momentoversikt` skill available

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| Momentöversikt has all 8 lessons listed | MAT-03 | Visual content check | Open HTML, verify all L1-L8 appear with titles |
| L7 metacognitive prompts are specific | MAT-01 | Content quality judgment | Read L7 .docx, verify prompts are concrete not generic |
| CSV imports correctly | MAT-01 | Requires external app | Import CSV in survey platform, verify questions render |

---

## Validation Sign-Off

- [ ] All tasks have `<automated>` verify or Wave 0 dependencies
- [ ] Sampling continuity: no 3 consecutive tasks without automated verify
- [ ] Wave 0 covers all MISSING references
- [ ] No watch-mode flags
- [ ] Feedback latency < 10s
- [ ] `nyquist_compliant: true` set in frontmatter

**Approval:** pending
