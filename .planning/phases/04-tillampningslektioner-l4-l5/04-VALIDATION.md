---
phase: 4
slug: tillampningslektioner-l4-l5
status: draft
nyquist_compliant: false
wave_0_complete: false
created: 2026-03-15
---

# Phase 4 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | Node.js script execution |
| **Config file** | package.json i projektmappen |
| **Quick run command** | `node generate-lektion-4.js && node generate-lektion-5.js` |
| **Full suite command** | `node generate-lektion-4.js && node generate-lektion-5.js` |
| **Estimated runtime** | ~10 seconds |

---

## Sampling Rate

- **After every task commit:** Run `node generate-lektion-4.js` eller `node generate-lektion-5.js`
- **After every plan wave:** Run `node generate-lektion-4.js && node generate-lektion-5.js`
- **Before `/gsd:verify-work`:** Full suite must be green
- **Max feedback latency:** 10 seconds

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|-----------|-------------------|-------------|--------|
| 04-01-01 | 01 | 1 | AI-02 | manual-only | Granska .docx-innehall efter generering | N/A | ⬜ pending |
| 04-02-01 | 02 | 1 | MAT-02 | smoke | `ls lektion-4.pptx lektion-5.pptx` | ❌ W0 | ⬜ pending |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

- [ ] `lektion-4.pptx` — ny fil som ska skapas via pptxgenjs
- [ ] `lektion-5.pptx` — ny fil som ska skapas via pptxgenjs
- [ ] Revidering av `generate-lektion-4.js` (framatkoppling, scaffolding-bilaga)
- [ ] Revidering av `generate-lektion-5.js` (ta bort kamratrespons, lagga till lararfeedback-instruktion)

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| Metakognitiv reflektion i exit tickets och larpar-atergivning | AI-02 | Innehallsvalidering kraver mannisklig bedomning | Granska .docx for exit ticket-fragar och larpar-instruktioner |
| Scaffolding tillganglig pa begaran men inte utdelad | AI-02 | Pedagogisk designval, ej automatiserbart | Verifiera att scaffolding-bilaga finns separat, ej inbaddad i huvuduppgift |

---

## Validation Sign-Off

- [ ] All tasks have `<automated>` verify or Wave 0 dependencies
- [ ] Sampling continuity: no 3 consecutive tasks without automated verify
- [ ] Wave 0 covers all MISSING references
- [ ] No watch-mode flags
- [ ] Feedback latency < 10s
- [ ] `nyquist_compliant: true` set in frontmatter

**Approval:** pending
