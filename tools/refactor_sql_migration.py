from __future__ import annotations

import heapq
import re
from collections import defaultdict
from pathlib import Path


SOURCE = Path(r"F:\Source Code\SatriahMain\AllScripts.original.sql")
OUTPUT = Path(r"F:\Source Code\SatriahMain\AllScripts.cleaned.sql")

GO_RE = re.compile(r"^\s*GO\s*$", re.IGNORECASE)
HEADER_RE = re.compile(
    r"^\s*(CREATE(?:\s+OR\s+ALTER)?|ALTER)\s+(FUNCTION|PROCEDURE|PROC|VIEW)\s+(.+)$",
    re.IGNORECASE,
)
REQUIRED_SETUP_PATTERNS = (
    "IF COL_LENGTH('dbo.TblTempItemAging', 'QtyBalance') IS NULL",
    "IF COL_LENGTH('dbo.TblTempItemAging', 'BalanceValue') IS NULL",
    "IF COL_LENGTH('dbo.TblTempItemAging', 'AvgCost') IS NULL",
    "IF COL_LENGTH('dbo.TblOptions','IgnoreSameDayCountsWhileCosting') IS NULL",
    "IF COL_LENGTH('dbo.TblOptions','CostStartingGard') IS NULL",
    "IF COL_LENGTH('dbo.TblOptions','TreatUncountedItemsAsZeroQty') IS NULL",
    "IF COL_LENGTH('dbo.TblOptions','ExcludeInventoryAdjustmentsFromCost') IS NULL",
    "IF COL_LENGTH('dbo.ACCOUNTS','Level') IS NULL",
    "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UX_ACCOUNTS_Code' AND object_id = OBJECT_ID('dbo.ACCOUNTS'))",
    "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_ACCOUNTS_Parent' AND object_id = OBJECT_ID('dbo.ACCOUNTS'))",
    "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_ACCOUNTS_Level' AND object_id = OBJECT_ID('dbo.ACCOUNTS'))",
    "UPDATE A\nSET A.[Level] = H.ParentDepth",
)
MANUAL_SETUP_BLOCKS = (
    "/* 0) عمود Level (لو مش موجود) */\n"
    "IF COL_LENGTH('dbo.ACCOUNTS','Level') IS NULL\n"
    "BEGIN\n"
    "    ALTER TABLE dbo.ACCOUNTS ADD [Level] INT NULL;\n"
    "END",
)


def normalize_ident(value: str) -> str:
    return value.strip().strip("[]")


def split_name(header_tail: str) -> tuple[str, str] | None:
    match = re.match(r"((?:\[[^\]]+\]|\w+)\.)?(\[[^\]]+\]|\w+)", header_tail.strip())
    if not match:
        return None
    schema = normalize_ident(match.group(1)[:-1] if match.group(1) else "dbo")
    name = normalize_ident(match.group(2))
    return schema, name


def read_source() -> list[str]:
    return SOURCE.read_text(encoding="utf-16", errors="replace").splitlines()


def split_batches(lines: list[str]) -> list[str]:
    batches: list[str] = []
    current: list[str] = []
    for line in lines:
        if GO_RE.match(line):
            text = "\n".join(current).strip()
            if text:
                batches.append(text)
            current = []
        else:
            current.append(line.rstrip())
    text = "\n".join(current).strip()
    if text:
        batches.append(text)
    return batches


def extract_required_setup(lines: list[str]) -> str:
    selected: list[str] = []
    seen: set[str] = set()
    for batch in split_batches(lines):
        batch_lines = batch.splitlines()
        if any(HEADER_RE.match(line) for line in batch_lines):
            continue
        if not any(
            any(pattern in line for line in batch_lines) or pattern in batch
            for pattern in REQUIRED_SETUP_PATTERNS
        ):
            continue
        normalized = batch.strip()
        if normalized in seen:
            continue
        seen.add(normalized)
        selected.append(normalized)
    for block in MANUAL_SETUP_BLOCKS:
        if block not in seen:
            insert_at = next((i for i, value in enumerate(selected) if "UX_ACCOUNTS_Code" in value), len(selected))
            selected.insert(insert_at, block)
            seen.add(block)
    return "\n\nGO\n\n".join(selected).strip()


def extract_objects(lines: list[str]) -> list[dict]:
    objects: list[dict] = []
    idx = 0
    while idx < len(lines):
        match = HEADER_RE.match(lines[idx])
        if not match:
            idx += 1
            continue

        end = idx + 1
        while end < len(lines) and not GO_RE.match(lines[end]):
            end += 1

        parsed = split_name(match.group(3))
        if parsed is None:
            idx = end + 1
            continue

        schema, name = parsed
        object_type = match.group(2).upper()
        if object_type == "PROC":
            object_type = "PROCEDURE"

        body = "\n".join(lines[idx:end]).rstrip()
        if name == "QryItemsTransactionsTotals":
            marker = "/* ========================================================="
            marker_pos = body.find(marker)
            if marker_pos != -1:
                body = body[:marker_pos].rstrip()
        if name == "sp_CreateIssueVoucherJE":
            body = body.replace(
                "    IF OBJECT_ID('dbo.DOUBLE_ENTRY_VOUCHERS','U') IS NOT NULL\n"
                "        DELETE FROM dbo.DOUBLE_ENTRY_VOUCHERS WHERE Notes_ID=@NoteID;\n"
                "    IF OBJECT_ID('dbo.DOUBLE_ENTREY_VOUCHERS','U') IS NOT NULL\n"
                "        DELETE FROM dbo.DOUBLE_ENTREY_VOUCHERS WHERE Notes_ID=@NoteID;",
                "    IF OBJECT_ID('dbo.DOUBLE_ENTRY_VOUCHERS','U') IS NOT NULL\n"
                "       AND COL_LENGTH('dbo.DOUBLE_ENTRY_VOUCHERS', 'Notes_ID') IS NOT NULL\n"
                "    BEGIN\n"
                "        EXEC sys.sp_executesql\n"
                "            N'DELETE FROM dbo.DOUBLE_ENTRY_VOUCHERS WHERE Notes_ID = @NoteID;',\n"
                "            N'@NoteID BIGINT',\n"
                "            @NoteID = @NoteID;\n"
                "    END;\n"
                "    IF OBJECT_ID('dbo.DOUBLE_ENTREY_VOUCHERS','U') IS NOT NULL\n"
                "       AND COL_LENGTH('dbo.DOUBLE_ENTREY_VOUCHERS', 'Notes_ID') IS NOT NULL\n"
                "    BEGIN\n"
                "        EXEC sys.sp_executesql\n"
                "            N'DELETE FROM dbo.DOUBLE_ENTREY_VOUCHERS WHERE Notes_ID = @NoteID;',\n"
                "            N'@NoteID BIGINT',\n"
                "            @NoteID = @NoteID;\n"
                "    END;"
            )

        body = re.sub(
            r"^\s*(CREATE(?:\s+OR\s+ALTER)?|ALTER)\s+(FUNCTION|PROCEDURE|PROC|VIEW)\s+((?:\[[^\]]+\]|\w+)\.)?(\[[^\]]+\]|\w+)",
            f"CREATE {object_type} dbo.{name}",
            body,
            count=1,
            flags=re.IGNORECASE,
        )

        body_upper = body.upper()
        if object_type == "FUNCTION":
            if re.search(r"RETURNS\s+@\w+\s+TABLE", body_upper):
                function_kind = "tvf_multi"
            elif re.search(r"RETURNS\s+TABLE", body_upper):
                function_kind = "tvf_inline"
            else:
                function_kind = "scalar"
        else:
            function_kind = ""

        objects.append(
            {
                "schema": schema,
                "name": name,
                "type": object_type,
                "function_kind": function_kind,
                "body": body.strip(),
                "start_line": idx + 1,
            }
        )
        idx = end + 1
    return objects


def keep_last_definitions(objects: list[dict]) -> list[dict]:
    last_by_name: dict[tuple[str, str], dict] = {}
    for obj in objects:
        last_by_name[(obj["schema"].lower(), obj["name"].lower())] = obj
    return sorted(last_by_name.values(), key=lambda obj: obj["start_line"])


def category_priority(obj: dict) -> tuple[int, str]:
    if obj["type"] == "FUNCTION" and obj["function_kind"] == "scalar":
        return 0, "scalar"
    if obj["type"] == "FUNCTION" and obj["function_kind"].startswith("tvf"):
        return 1, "tvf"
    if obj["type"] == "VIEW":
        return 2, "view"
    return 3, "procedure"


def find_dependencies(objects: list[dict]) -> dict[str, set[str]]:
    keys = {full_name(obj): obj for obj in objects}
    dependencies = {name: set() for name in keys}

    for name, obj in keys.items():
        body_lower = obj["body"].lower()
        for other_name in keys:
            if other_name == name:
                continue
            other_simple = other_name.split(".", 1)[1]
            patterns = [
                rf"\b{re.escape(other_name)}\b",
                rf"\b{re.escape(other_simple)}\b",
            ]
            if any(re.search(pattern, body_lower) for pattern in patterns):
                dependencies[name].add(other_name)

    # Enforce the explicit requirement even if textual detection misses it.
    anchor = "dbo.qryitemstransactionstotals"
    if anchor in dependencies:
        for forced in (
            "dbo.rpt_itemaging",
            "dbo.fn_getitemcostweighted",
        ):
            if forced in dependencies:
                dependencies[forced].add(anchor)

    return dependencies


def topo_sort(objects: list[dict]) -> list[dict]:
    keys = {full_name(obj): obj for obj in objects}
    dependencies = find_dependencies(objects)
    reverse: dict[str, set[str]] = defaultdict(set)
    indegree = {name: 0 for name in keys}

    for name, deps in dependencies.items():
        indegree[name] = len(deps)
        for dep in deps:
            reverse[dep].add(name)

    def sort_key(name: str) -> tuple[int, int, str]:
        obj = keys[name]
        return (
            category_priority(obj)[0],
            obj["start_line"],
            name,
        )

    queue = [sort_key(name) + (name,) for name, degree in indegree.items() if degree == 0]
    heapq.heapify(queue)
    ordered: list[str] = []

    while queue:
        _, _, _, name = heapq.heappop(queue)
        ordered.append(name)
        for follower in reverse.get(name, set()):
            indegree[follower] -= 1
            if indegree[follower] == 0:
                heapq.heappush(queue, sort_key(follower) + (follower,))

    if len(ordered) != len(keys):
        remaining = sorted((name for name in keys if name not in ordered), key=sort_key)
        ordered.extend(remaining)

    return [keys[name] for name in ordered]


def full_name(obj: dict) -> str:
    return f"{obj['schema'].lower()}.{obj['name'].lower()}"


def drop_statement(obj: dict) -> str:
    if obj["type"] == "FUNCTION":
        return (
            f"IF OBJECT_ID(N'dbo.{obj['name']}', N'IF') IS NOT NULL DROP FUNCTION dbo.{obj['name']};\n"
            f"IF OBJECT_ID(N'dbo.{obj['name']}', N'TF') IS NOT NULL DROP FUNCTION dbo.{obj['name']};\n"
            f"IF OBJECT_ID(N'dbo.{obj['name']}', N'FN') IS NOT NULL DROP FUNCTION dbo.{obj['name']};"
        )
    elif obj["type"] == "PROCEDURE":
        kind = "P"
        noun = "PROCEDURE"
    else:
        kind = "V"
        noun = "VIEW"
    return (
        f"IF OBJECT_ID('dbo.{obj['name']}','{kind}') IS NOT NULL\n"
        f"DROP {noun} dbo.{obj['name']}"
    )


def render(objects: list[dict], preamble: str) -> str:
    parts: list[str] = []
    if preamble:
        parts.append(preamble)

    for obj in objects:
        parts.append(drop_statement(obj))
        parts.append("GO")
        parts.append(obj["body"])
        parts.append("GO")

    return "\n\n".join(parts).strip() + "\n"


def main() -> None:
    lines = read_source()
    preamble = extract_required_setup(lines)
    objects = keep_last_definitions(extract_objects(lines))
    ordered = topo_sort(objects)
    OUTPUT.write_text(render(ordered, preamble), encoding="utf-8")


if __name__ == "__main__":
    main()
