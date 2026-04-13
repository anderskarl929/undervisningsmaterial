---
created: 2026-04-13T05:17:43.006Z
title: Installera superpowers-pluginen
area: tooling
files: []
---

## Problem

Vill installera Claude Code-pluginen `superpowers` från `claude-plugins-official`-marketplacen. Försökte köra `/plugin install superpowers@claude-plugins-official` via Remote Control men det gick inte — slash-kommandot `/plugin` är inte tillgängligt över Remote Control. Måste köras lokalt i en Claude Code-session direkt på maskinen.

## Solution

Kör i en lokal Claude Code-session (inte via Remote Control):

```
/plugin install superpowers@claude-plugins-official
```

Om marketplacen inte redan är tillagd kan det krävas ett `/plugin marketplace add` först. Verifiera efter installation med `/plugin` att superpowers syns som installerad.
