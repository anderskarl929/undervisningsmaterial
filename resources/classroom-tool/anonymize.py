"""Per-kurs alias-mapping. Persisterar userId -> "Elev N" i aliases.json.

Filen innehåller endast Google-userId och pseudonymer — inga riktiga namn.
Den är gitignored. Numreringen är stabil: en gång tilldelat behåller eleven
sitt nummer för all framtid i den kursen.
"""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ALIASES_PATH = ROOT / "aliases.json"


def _load_all() -> dict[str, dict[str, str]]:
    if not ALIASES_PATH.exists():
        return {}
    return json.loads(ALIASES_PATH.read_text(encoding="utf-8"))


def _save_all(data: dict[str, dict[str, str]]) -> None:
    ALIASES_PATH.write_text(
        json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True),
        encoding="utf-8",
    )


class CourseAliases:
    """Alias-mapping för en specifik kurs. Lagras i `aliases.json` under nyckel = course_id."""

    def __init__(self, course_id: str) -> None:
        self.course_id = course_id
        self._all = _load_all()
        self._mapping: dict[str, str] = self._all.get(course_id, {})
        self._dirty = False

    def alias_for(self, user_id: str) -> str:
        if user_id in self._mapping:
            return self._mapping[user_id]
        used = {int(v.split()[1]) for v in self._mapping.values() if v.startswith("Elev ")}
        next_num = max(used, default=0) + 1
        alias = f"Elev {next_num}"
        self._mapping[user_id] = alias
        self._dirty = True
        return alias

    def known_user_ids(self) -> set[str]:
        return set(self._mapping.keys())

    def user_id_for_alias(self, alias: str) -> str | None:
        for uid, a in self._mapping.items():
            if a == alias:
                return uid
        return None

    def all_aliases_sorted(self) -> list[tuple[str, str]]:
        """Returns [(alias, user_id), ...] sorterat på alias-nummer."""
        items = sorted(self._mapping.items(), key=lambda kv: int(kv[1].split()[1]))
        return [(alias, uid) for uid, alias in items]

    def save(self) -> None:
        if not self._dirty:
            return
        self._all[self.course_id] = self._mapping
        _save_all(self._all)
        self._dirty = False
