from __future__ import annotations

import sys
import re
from datetime import datetime
from pathlib import Path

from core import load_config, safe_update, log_event


def resolve_source_path(raw_path: str, use_date_template: bool, date_format: str) -> str:
    if not raw_path:
        return raw_path
    if not use_date_template:
        return raw_path

    fmt = date_format or "%m%d"
    today_str = datetime.now().strftime(fmt)

    if "{date" in raw_path:
        def replace(match: re.Match[str]) -> str:
            fmt_override = match.group(1) or fmt
            return datetime.now().strftime(fmt_override)

        return re.sub(r"\{date(?::([^}]+))?\}", replace, raw_path)

    name = Path(raw_path).name
    pattern = rf"(\d{{{len(today_str)}}})(?!.*\d)"
    new_name = re.sub(pattern, today_str, name, count=1)
    if new_name != name:
        return str(Path(raw_path).with_name(new_name))
    return raw_path


def main() -> int:
    cfg = load_config()
    raw_path = str(cfg.get("source_path", "")).strip()
    if not raw_path:
        log_event("schedule skipped: empty source_path")
        return 1

    encoding = str(cfg.get("encoding", "auto"))
    use_date_template = bool(cfg.get("use_date_template", False))
    date_format = str(cfg.get("date_format", "%m%d"))
    resolved = resolve_source_path(raw_path, use_date_template, date_format)

    df, err = safe_update(resolved, encoding, "schedule")
    if err:
        log_event(f"schedule failed: {err}")
        return 2
    if df is None:
        log_event("schedule failed: no data returned")
        return 3
    log_event(f"schedule ok: {resolved}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
