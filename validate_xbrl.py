"""
Validate an XBRL instance with Arelle, auto-selecting the taxonomy package
from the local Taxonomy folder based on the instance schemaRef entry point.

Usage:
  py -3 validate_xbrl.py --xbrl <file.xbrl>
  py -3 validate_xbrl.py --instance <file.xbrl>
  py -3 validate_xbrl.py <file.xbrl>
  py -3 validate_xbrl.py
  py -3 validate_xbrl.py --xbrl XBRL --taxonomy-dir Taxonomy --output-dir Output
  py -3 validate_xbrl.py --xbrl-diff FILE_A FILE_B
  py -3 validate_xbrl.py --taxonomy-diff OLD.zip NEW.zip

When --instance is omitted, all XBRL/iXBRL/.zip files in the XBRL folder are processed.
For each file the matching taxonomy package is selected from Taxonomy/, and
outputs are written under Output/<XBRL-file-stem>/:

  Output/<fileStem>/
    XBRLXray_Report.html    — open this to browse HTML tables
    Validation_Log.html     — readable validation log
    Validation_Log.xml      — Arelle XML validation log (source for HTML details)
    HTML_Report/            — table HTML files + index.html
    _extracted/             — present when input was a .zip (extracted XBRL)

Large filings may be supplied as .zip; the tool extracts the XBRL/iXBRL member automatically.

Taxonomy differences (--taxonomy-diff) compare two taxonomy packages and write
Output/TaxonomyDiff_<old>_vs_<new>/TaxonomyDiff_Report.html.
"""

from __future__ import annotations

import argparse
import html
import re
import shutil
import sys
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from arelle.api.Session import Session
from arelle.RuntimeOptions import RuntimeOptions


TABLE_ID_PREFIXES = ("boe_t", "eba_t", "srb_t")
TP_NS = {"tp": "http://xbrl.org/2016/taxonomy-package"}
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_TAXONOMY_DIR = SCRIPT_DIR / "Taxonomy"
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "Output"
DEFAULT_XBRL_DIR = SCRIPT_DIR / "XBRL"
INSTANCE_SUFFIXES = {".xbrl", ".xml", ".xhtml", ".html"}
ZIP_SUFFIX = ".zip"
_RC_CODE_RE = re.compile(r"[RC]\d{4}")
_EMPTY_AXIS_PROF: dict = {}
_MEMBER_DISPLAY_CACHE: dict[tuple[int, str, str], dict] = {}
INPUT_SUFFIXES = INSTANCE_SUFFIXES | {ZIP_SUFFIX}
SCHEMA_REF_HREF_RE = re.compile(
    r"""(?:xlink:)?href\s*=\s*["']([^"']+)["']""",
    re.IGNORECASE,
)


@dataclass
class TaxonomyEntryPoint:
    name: str
    description: str
    href: str


@dataclass
class TaxonomyPackageInfo:
    path: Path
    name: str
    version: str
    identifier: str = ""
    publisher: str = ""
    publication_date: str = ""
    entry_points: list[TaxonomyEntryPoint] = field(default_factory=list)

    @property
    def display(self) -> str:
        ver = self.version or "(unknown version)"
        name = self.name or self.path.name
        return f"{name} (version {ver})"


@dataclass
class TaxonomySelection:
    package: TaxonomyPackageInfo
    schema_ref: str
    matched_entry_point: TaxonomyEntryPoint | None = None
    companion_packages: list[Path] = field(default_factory=list)

    @property
    def summary(self) -> str:
        ep = self.matched_entry_point
        ep_txt = ""
        if ep is not None:
            bits = [p for p in (ep.name, ep.description) if p]
            ep_txt = f"; entry point: {' — '.join(bits)}" if bits else ""
        companions = ""
        if self.companion_packages:
            companions = "; companions: " + ", ".join(
                p.name for p in self.companion_packages
            )
        return (
            f"{self.package.display}; package file: {self.package.path.name}"
            f"{ep_txt}{companions}; schemaRef: {self.schema_ref}"
        )

    @property
    def arelle_packages(self) -> list[str]:
        """Ordered package paths for Arelle RuntimeOptions.packages."""
        paths = [str(self.package.path.resolve())]
        for p in self.companion_packages:
            paths.append(str(p.resolve()))
        return paths


@dataclass
class CellCoverage:
    """Report cell population vs taxonomy table cells for this filing."""

    tables: int = 0
    total_cells: int = 0
    populated_cells: int = 0
    instance_facts: int = 0
    per_table: list[dict] = field(default_factory=list)

    @property
    def empty_cells(self) -> int:
        return max(self.total_cells - self.populated_cells, 0)

    @property
    def coverage_pct(self) -> float:
        if self.total_cells <= 0:
            return 0.0
        return round(100.0 * self.populated_cells / self.total_cells, 1)

    @property
    def summary_line(self) -> str:
        return (
            f"Report cells: {self.total_cells}; populated: {self.populated_cells}; "
            f"empty/not populated: {self.empty_cells}; coverage: {self.coverage_pct}% "
            f"(instance facts: {self.instance_facts}; tables: {self.tables})"
        )


@dataclass
class ValidationStats:
    """Validation rule and message-level outcome summary for one XBRL file."""

    rules_total: int = 0
    rules_checked: int = 0
    rules_passed: int = 0
    rules_failed: int = 0
    evaluations_passed: int = 0
    evaluations_failed: int = 0
    errors: int = 0
    warnings: int = 0
    infos: int = 0

    @property
    def summary_line(self) -> str:
        return (
            f"Rules total: {self.rules_total}; checked: {self.rules_checked}; "
            f"passed: {self.rules_passed}; failed: {self.rules_failed}; "
            f"errors: {self.errors}; warnings: {self.warnings}"
        )


@dataclass
class FactDiffItem:
    """One fact identity involved in an XBRLDiff result."""

    key: tuple
    concept: str
    value_a: str | None = None
    value_b: str | None = None
    unit: str = ""
    dims: str = ""


@dataclass
class XbrlDiffResult:
    """Report A vs Report B fact/validation comparison."""

    label_a: str
    label_b: str
    facts_a: int = 0
    facts_b: int = 0
    errors_a: int = 0
    errors_b: int = 0
    warnings_a: int = 0
    warnings_b: int = 0
    added: list[FactDiffItem] = field(default_factory=list)
    removed: list[FactDiffItem] = field(default_factory=list)
    changed: list[FactDiffItem] = field(default_factory=list)
    unchanged: int = 0
    added_keys: set = field(default_factory=set)
    removed_keys: set = field(default_factory=set)
    changed_keys: set = field(default_factory=set)

    @property
    def summary_line(self) -> str:
        return (
            f"Added: {len(self.added)}; Removed: {len(self.removed)}; "
            f"Changed: {len(self.changed)}; Unchanged: {self.unchanged}"
        )


def _fact_value_text(fact) -> str:
    """Exact lexical value text for fact comparison."""
    if getattr(fact, "isNil", False):
        return "(nil)"
    val = getattr(fact, "value", None)
    return "" if val is None else str(val)


def _fact_context_key(context) -> tuple:
    if context is None:
        return ()
    entity = getattr(context, "entityIdentifier", None) or (None, None)
    if getattr(context, "isInstantPeriod", False):
        period = ("I", str(getattr(context, "instantDatetime", None)))
    elif getattr(context, "isStartEndPeriod", False):
        period = (
            "D",
            str(getattr(context, "startDatetime", None)),
            str(getattr(context, "endDatetime", None)),
        )
    elif getattr(context, "isForeverPeriod", False):
        period = ("F",)
    else:
        period = ("?",)
    dims: list[tuple[str, str]] = []
    for dim_qn, dim_val in sorted(
        (getattr(context, "qnameDims", None) or {}).items(),
        key=lambda kv: str(kv[0]),
    ):
        if getattr(dim_val, "isExplicit", False):
            mem = str(getattr(dim_val, "memberQname", None) or "")
        else:
            typed = getattr(dim_val, "typedMember", None)
            mem = str(
                getattr(typed, "xValue", None)
                or getattr(typed, "text", None)
                or typed
                or ""
            )
        dims.append((str(dim_qn), mem))
    return (entity, period, tuple(dims))


def _fact_unit_key(unit) -> tuple | str:
    if unit is None:
        return ()
    measures = getattr(unit, "measures", None)
    if measures is not None:
        try:
            nums, dens = measures
            return (
                tuple(str(m) for m in (nums or ())),
                tuple(str(m) for m in (dens or ())),
            )
        except Exception:
            pass
    return str(getattr(unit, "id", None) or unit)


def fact_identity_key(fact) -> tuple:
    """
    Identity for XBRLDiff: concept + context (entity/period/dims) + unit.
    Same key with different value => Changed; key only in B => Added; only in A => Removed.
    """
    concept = str(getattr(fact, "qname", None) or "")
    return (
        concept,
        _fact_context_key(getattr(fact, "context", None)),
        _fact_unit_key(getattr(fact, "unit", None)),
    )


def build_fact_value_map(model) -> dict[tuple, str]:
    """Map fact identity -> lexical value (last duplicate wins)."""
    out: dict[tuple, str] = {}
    for fact in getattr(model, "facts", None) or ():
        concept = getattr(fact, "concept", None)
        if concept is not None and getattr(concept, "isTuple", False):
            continue
        if getattr(fact, "isTuple", False):
            continue
        out[fact_identity_key(fact)] = _fact_value_text(fact)
    return out


def _fact_diff_item(
    key: tuple,
    value_a: str | None,
    value_b: str | None,
) -> FactDiffItem:
    concept = str(key[0]) if key else ""
    unit = str(key[2]) if len(key) > 2 else ""
    dims = ""
    if len(key) > 1 and isinstance(key[1], tuple) and len(key[1]) > 2:
        dim_pairs = key[1][2]
        if dim_pairs:
            dims = "; ".join(f"{d}={m}" for d, m in dim_pairs)
    return FactDiffItem(
        key=key,
        concept=concept,
        value_a=value_a,
        value_b=value_b,
        unit=unit,
        dims=dims,
    )


def diff_fact_maps(
    map_a: dict[tuple, str],
    map_b: dict[tuple, str],
    *,
    label_a: str,
    label_b: str,
    facts_a: int,
    facts_b: int,
    stats_a: ValidationStats,
    stats_b: ValidationStats,
) -> XbrlDiffResult:
    """Compare two fact maps and validation stats into an XbrlDiffResult."""
    keys_a = set(map_a)
    keys_b = set(map_b)
    added_keys = keys_b - keys_a
    removed_keys = keys_a - keys_b
    common = keys_a & keys_b
    changed_keys = {k for k in common if map_a[k] != map_b[k]}
    unchanged = len(common) - len(changed_keys)

    return XbrlDiffResult(
        label_a=label_a,
        label_b=label_b,
        facts_a=facts_a,
        facts_b=facts_b,
        errors_a=stats_a.errors,
        errors_b=stats_b.errors,
        warnings_a=stats_a.warnings,
        warnings_b=stats_b.warnings,
        added=[_fact_diff_item(k, None, map_b[k]) for k in sorted(added_keys, key=str)],
        removed=[_fact_diff_item(k, map_a[k], None) for k in sorted(removed_keys, key=str)],
        changed=[
            _fact_diff_item(k, map_a[k], map_b[k]) for k in sorted(changed_keys, key=str)
        ],
        unchanged=unchanged,
        added_keys=added_keys,
        removed_keys=removed_keys,
        changed_keys=changed_keys,
    )


def _layout_cell_primary_fact(cell):
    """Return the first fact object on a layout body cell, if any."""
    for item in getattr(cell, "facts", None) or ():
        if isinstance(item, (tuple, list)) and item:
            return item[0]
        return item
    return None


def _cell_diff_statuses_from_view(
    view,
    added_keys: set,
    changed_keys: set,
) -> list[str | None]:
    """
    Per HTML td.cell status in Arelle body order: 'added', 'changed', or None.
    """
    statuses: list[str | None] = []
    append = statuses.append
    is_surrogate = lambda c: getattr(c, "isOpenAspectEntrySurrogate", False)

    for lyt_table in _lyt_tables_from_view(view):
        z_bodies = getattr(lyt_table, "lytMdlBodyChildren", None) or ()
        if not z_bodies:
            continue
        for y_body in z_bodies[0].lytMdlBodyChildren or ():
            for x_body in getattr(y_body, "lytMdlBodyChildren", None) or ():
                for cell in getattr(x_body, "lytMdlBodyChildren", None) or ():
                    if is_surrogate(cell):
                        continue
                    status = None
                    fact = _layout_cell_primary_fact(cell)
                    if fact is not None:
                        key = fact_identity_key(fact)
                        if key in added_keys:
                            status = "added"
                        elif key in changed_keys:
                            status = "changed"
                    append(status)
    return statuses


def collect_validation_stats(model, log_xml: str = "") -> ValidationStats:
    """
    Summarise formula/assertion outcomes and log message levels for a model.
    Assertion dict in model.errors maps id -> (satisfied, notSatisfied, ...).
    """
    import logging

    stats = ValidationStats()

    assertion_maps: list[dict] = []
    for item in getattr(model, "errors", None) or []:
        if isinstance(item, dict):
            assertion_maps.append(item)

    # Prefer Arelle's aggregated assertionResults dict when present.
    merged: dict = {}
    for amap in assertion_maps:
        merged.update(amap)

    if not merged:
        # Fallback: read counts directly from variable set objects.
        try:
            from arelle.ModelFormulaObject import (
                ModelVariableSetAssertion,
                ModelConsistencyAssertion,
            )
        except ImportError:
            ModelVariableSetAssertion = ()  # type: ignore[assignment,misc]
            ModelConsistencyAssertion = ()  # type: ignore[assignment,misc]
        for vs in getattr(model, "modelVariableSets", None) or []:
            if ModelVariableSetAssertion and isinstance(vs, ModelVariableSetAssertion):
                aid = getattr(vs, "id", None) or str(vs)
                merged[aid] = (
                    int(getattr(vs, "countSatisfied", 0) or 0),
                    int(getattr(vs, "countNotSatisfied", 0) or 0),
                )
        for ca in getattr(model, "modelConsistencyAssertions", None) or []:
            if ModelConsistencyAssertion and isinstance(ca, ModelConsistencyAssertion):
                aid = getattr(ca, "id", None) or str(ca)
                merged[aid] = (
                    int(getattr(ca, "countSatisfied", 0) or 0),
                    int(getattr(ca, "countNotSatisfied", 0) or 0),
                )

    stats.rules_total = len(merged)
    for _rule_id, counts in merged.items():
        if not isinstance(counts, tuple) or len(counts) < 2:
            continue
        sat = int(counts[0] or 0)
        unsat = int(counts[1] or 0)
        stats.evaluations_passed += sat
        stats.evaluations_failed += unsat
        if sat + unsat > 0:
            stats.rules_checked += 1
            if unsat > 0:
                stats.rules_failed += 1
            else:
                stats.rules_passed += 1

    # Message severity counts from Arelle logger, with XML log fallback.
    log_count = getattr(model, "logCount", None) or {}
    if isinstance(log_count, dict) and log_count:
        stats.errors = int(log_count.get(logging.ERROR, 0) or log_count.get(40, 0) or 0)
        stats.errors += int(log_count.get(logging.CRITICAL, 0) or log_count.get(50, 0) or 0)
        stats.warnings = int(log_count.get(logging.WARNING, 0) or log_count.get(30, 0) or 0)
        stats.infos = int(log_count.get(logging.INFO, 0) or log_count.get(20, 0) or 0)
    elif log_xml:
        level_counts = _count_log_levels(log_xml)
        stats.errors = level_counts.get("error", 0) + level_counts.get("critical", 0)
        stats.warnings = level_counts.get("warning", 0)
        stats.infos = level_counts.get("info", 0)

    return stats


def _count_log_levels(log_xml: str) -> dict[str, int]:
    """Count Arelle XML log entries by level."""
    from collections import Counter

    text = (log_xml or "").strip()
    if not text:
        return {}
    # Normalise possibly broken/annotated logs into a single <log> root.
    body = text
    if body.lstrip().lower().startswith("<log"):
        body = re.sub(r"^<\s*log[^>]*>", "", body.lstrip(), count=1, flags=re.IGNORECASE)
        body = re.sub(r"</\s*log\s*>", "", body, flags=re.IGNORECASE)
    wrapped = f"<log>{body}</log>"
    try:
        root = ET.fromstring(wrapped)
    except ET.ParseError:
        # Last resort: regex count level="..."
        counts: Counter[str] = Counter()
        for m in re.finditer(r'\blevel\s*=\s*["\']([^"\']+)["\']', text, flags=re.I):
            counts[m.group(1).casefold()] += 1
        return dict(counts)

    counts = Counter(
        (entry.attrib.get("level") or "").casefold() for entry in root.findall("entry")
    )
    return dict(counts)


def default_taxonomy_dir() -> Path:
    return DEFAULT_TAXONOMY_DIR


def default_output_dir() -> Path:
    return DEFAULT_OUTPUT_DIR


def default_xbrl_dir() -> Path:
    return DEFAULT_XBRL_DIR


def resolve_xbrl_input(
    raw: str,
    *,
    xbrl_dir: Path,
    label: str = "Instance",
) -> Path:
    """
    Resolve an XBRL/iXBRL/.zip file argument.

    Bare filenames (no directory) are looked up in ``xbrl_dir`` (default:
    <script>/XBRL/). Full/relative paths with a directory are used as given.
    """
    path = Path(raw).expanduser()
    # Bare name only — no drive/parent folders in the user string.
    if path.parent == Path(".") and not path.is_absolute():
        resolved = (xbrl_dir / path.name).resolve()
        if not resolved.is_file():
            raise FileNotFoundError(
                f"{label} not found in XBRL folder: {resolved}"
            )
        return resolved

    resolved = path.resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"{label} not found: {resolved}")
    return resolved


def resolve_xbrl_diff_input(raw: str, *, xbrl_dir: Path, side: str) -> Path:
    """Resolve an --xbrl-diff file argument (bare names → XBRL folder)."""
    return resolve_xbrl_input(raw, xbrl_dir=xbrl_dir, label=f"XBRLDiff file {side}")


def ensure_output_dir(output_dir: Path) -> Path:
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def extract_schema_refs(instance_path: Path) -> list[str]:
    """Return schemaRef href entry points from an XBRL/iXBRL instance."""
    text = instance_path.read_text(encoding="utf-8", errors="replace")
    refs: list[str] = []

    # Prefer XML parse when well-formed; fall back to regex for large/odd files.
    try:
        # Only need the root + early children; full parse is fine for typical instances.
        root = ET.parse(instance_path).getroot()
        for el in root.iter():
            tag = el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""
            if tag != "schemaRef":
                continue
            href = (
                el.attrib.get("{http://www.w3.org/1999/xlink}href")
                or el.attrib.get("href")
                or ""
            ).strip()
            if href:
                refs.append(href)
    except ET.ParseError:
        refs = []

    if not refs:
        # Capture schemaRef blocks only (avoid other xlink:href noise).
        for block in re.findall(r"<[^>]*schemaRef[^>]*>", text, flags=re.IGNORECASE):
            m = SCHEMA_REF_HREF_RE.search(block)
            if m:
                refs.append(m.group(1).strip())

    # Preserve order, unique
    seen: set[str] = set()
    unique: list[str] = []
    for href in refs:
        if href not in seen:
            seen.add(href)
            unique.append(href)
    return unique


def _normalize_href(href: str) -> str:
    return href.strip().replace("\\", "/").rstrip("/")


def _load_taxonomy_package_info(zip_path: Path) -> TaxonomyPackageInfo | None:
    try:
        with zipfile.ZipFile(zip_path) as zf:
            meta_names = [
                n
                for n in zf.namelist()
                if n.replace("\\", "/").endswith("META-INF/taxonomyPackage.xml")
            ]
            if not meta_names:
                return None
            root = ET.fromstring(zf.read(meta_names[0]))
    except (OSError, zipfile.BadZipFile, ET.ParseError):
        return None

    entry_points: list[TaxonomyEntryPoint] = []
    for ep_el in root.findall("tp:entryPoints/tp:entryPoint", TP_NS):
        name = (ep_el.findtext("tp:name", default="", namespaces=TP_NS) or "").strip()
        desc = (
            ep_el.findtext("tp:description", default="", namespaces=TP_NS) or ""
        ).strip()
        for doc in ep_el.findall("tp:entryPointDocument", TP_NS):
            href = (doc.get("href") or "").strip()
            if href:
                entry_points.append(
                    TaxonomyEntryPoint(name=name, description=desc, href=href)
                )

    return TaxonomyPackageInfo(
        path=zip_path.resolve(),
        name=(root.findtext("tp:name", default="", namespaces=TP_NS) or "").strip()
        or zip_path.stem,
        version=(root.findtext("tp:version", default="", namespaces=TP_NS) or "").strip(),
        identifier=(
            root.findtext("tp:identifier", default="", namespaces=TP_NS) or ""
        ).strip(),
        publisher=(
            root.findtext("tp:publisher", default="", namespaces=TP_NS) or ""
        ).strip(),
        publication_date=(
            root.findtext("tp:publicationDate", default="", namespaces=TP_NS) or ""
        ).strip(),
        entry_points=entry_points,
    )


def discover_taxonomy_packages(taxonomy_dir: Path) -> list[TaxonomyPackageInfo]:
    if not taxonomy_dir.is_dir():
        raise FileNotFoundError(
            f"Taxonomy folder not found: {taxonomy_dir}. "
            "Place taxonomy package .zip files in this folder."
        )
    packages: list[TaxonomyPackageInfo] = []
    for zip_path in sorted(taxonomy_dir.glob("*.zip")):
        info = _load_taxonomy_package_info(zip_path)
        if info is not None:
            packages.append(info)
    if not packages:
        raise FileNotFoundError(
            f"No taxonomy packages (.zip with META-INF/taxonomyPackage.xml) "
            f"found in {taxonomy_dir}"
        )
    return packages


def select_taxonomy_for_instance(
    instance_path: Path,
    taxonomy_dir: Path,
) -> TaxonomySelection:
    """
    Match instance schemaRef entry point(s) to a taxonomy package in taxonomy_dir.
    Prefers exact entryPointDocument href match; if several match, newest version wins.
    """
    schema_refs = extract_schema_refs(instance_path)
    if not schema_refs:
        raise ValueError(
            f"No schemaRef entry point found in instance: {instance_path}"
        )

    packages = discover_taxonomy_packages(taxonomy_dir)
    matches: list[tuple[TaxonomyPackageInfo, TaxonomyEntryPoint, str]] = []

    for schema_ref in schema_refs:
        want = _normalize_href(schema_ref).casefold()
        for pkg in packages:
            for ep in pkg.entry_points:
                have = _normalize_href(ep.href).casefold()
                if have == want:
                    matches.append((pkg, ep, schema_ref))

    if not matches:
        # Fallback: match by module file name + date segment when present.
        for schema_ref in schema_refs:
            parts = _normalize_href(schema_ref).casefold().split("/")
            module = parts[-1] if parts else ""
            date_seg = next(
                (p for p in parts if re.fullmatch(r"\d{4}-\d{2}-\d{2}", p)),
                "",
            )
            for pkg in packages:
                for ep in pkg.entry_points:
                    ep_parts = _normalize_href(ep.href).casefold().split("/")
                    ep_module = ep_parts[-1] if ep_parts else ""
                    if module and ep_module == module:
                        if date_seg and date_seg not in ep_parts:
                            continue
                        matches.append((pkg, ep, schema_ref))

    if not matches:
        available = []
        for pkg in packages:
            available.append(
                f"  - {pkg.path.name}: {pkg.display} "
                f"({len(pkg.entry_points)} entry points)"
            )
        raise ValueError(
            "No taxonomy package in the Taxonomy folder matches the instance "
            f"entry point(s): {', '.join(schema_refs)}\n"
            "Available packages:\n" + "\n".join(available)
        )

    # Prefer higher version / later publication date when multiple packages match.
    def sort_key(item: tuple[TaxonomyPackageInfo, TaxonomyEntryPoint, str]):
        pkg, _ep, _ref = item
        return (pkg.version, pkg.publication_date, pkg.path.name)

    pkg, ep, schema_ref = sorted(matches, key=sort_key)[-1]
    selection = TaxonomySelection(
        package=pkg, schema_ref=schema_ref, matched_entry_point=ep
    )
    selection.companion_packages = find_companion_taxonomy_packages(pkg, packages)
    return selection


def _taxonomy_version_family(version: str) -> str:
    """Return major.minor family, e.g. 3.2.0.0 -> 3.2."""
    parts = re.findall(r"\d+", version or "")
    if len(parts) >= 2:
        return f"{parts[0]}.{parts[1]}"
    return (version or "").strip()


def _looks_like_reporting_framework(pkg: TaxonomyPackageInfo) -> bool:
    blob = " ".join(
        [
            pkg.name or "",
            pkg.path.name,
            pkg.identifier or "",
        ]
    ).casefold()
    return (
        "reporting" in blob
        or "framework" in blob
        or ("fws" in blob and "eba" in blob)
    ) and "dictionary" not in blob


def _looks_like_dictionary(pkg: TaxonomyPackageInfo) -> bool:
    blob = " ".join(
        [
            pkg.name or "",
            pkg.path.name,
            pkg.identifier or "",
        ]
    ).casefold()
    return "dictionary" in blob or (
        "dict" in blob and "eba" in blob and "reporting" not in blob
    )


def find_companion_taxonomy_packages(
    primary: TaxonomyPackageInfo,
    packages: list[TaxonomyPackageInfo] | None = None,
    *,
    taxonomy_dir: Path | None = None,
) -> list[Path]:
    """
    For EBA-style reporting frameworks, also load the matching Dictionary zip.

    Reporting frameworks alone do not contain metrics/dimensions; without the
    Dictionary package, facts fail to bind and HTML report cells stay empty.
    """
    if not _looks_like_reporting_framework(primary):
        return []

    if packages is None:
        packages = discover_taxonomy_packages(taxonomy_dir or default_taxonomy_dir())

    family = _taxonomy_version_family(primary.version)
    dict_pkgs = [
        p
        for p in packages
        if p.path.resolve() != primary.path.resolve() and _looks_like_dictionary(p)
    ]
    if not dict_pkgs:
        return []

    # Prefer same version family (e.g. 3.2.*), then exact version, then name.
    def sort_key(p: TaxonomyPackageInfo):
        same_family = (
            1 if family and _taxonomy_version_family(p.version) == family else 0
        )
        exact = 1 if (primary.version and p.version == primary.version) else 0
        return (same_family, exact, p.version, p.publication_date, p.path.name)

    chosen = sorted(dict_pkgs, key=sort_key)[-1]
    # If we have a family filter and nothing matches, still take best dictionary
    # only when at least one shares the family — otherwise warn via empty.
    if family:
        same = [
            p
            for p in dict_pkgs
            if _taxonomy_version_family(p.version) == family
        ]
        if same:
            chosen = sorted(same, key=sort_key)[-1]
        else:
            return []

    return [chosen.path]


@contextmanager
def _exact_xbrl_fact_display():
    """
    Arelle's ModelFact.effectiveValue rounds numerics to the fact's decimals
    attribute (e.g. 3801 with decimals="-3" becomes "4,000"). For HTML reports
    we want the exact lexical value as delivered in the XBRL instance.
    """
    from arelle.ModelInstanceObject import ModelFact
    from arelle.XmlValidateConst import VALID

    original = ModelFact.effectiveValue

    def exact_effective_value(self) -> str | None:
        concept = self.concept
        if concept is None or concept.isTuple:
            return None
        if self.isNil:
            return "(nil)"
        try:
            if concept.isFraction:
                if self.xValid >= VALID:
                    return str(self.xValue)
                return "/".join(self.fractionValue)
            val = self.value
            if concept.isNumeric:
                # Exact instance text — do not round by decimals/precision.
                return "" if val is None else str(val)
            if len(val) == 0:
                return "(reported)"
            if self.xValid >= VALID and isinstance(self.xValue, str):
                return self.xValue
            return val
        except Exception as ex:
            return str(ex)

    ModelFact.effectiveValue = property(exact_effective_value)  # type: ignore[assignment,method-assign]
    try:
        yield
    finally:
        ModelFact.effectiveValue = original  # type: ignore[method-assign]


def _summarize_model_errors(model) -> tuple[list[str], list[str]]:
    """Return (error_codes, unsatisfied_assertions)."""
    error_codes: list[str] = []
    unsatisfied: list[str] = []
    for item in getattr(model, "errors", None) or []:
        if isinstance(item, dict):
            for code, counts in item.items():
                # Arelle formula counts: (satisfied, notSatisfied, ...)
                if isinstance(counts, tuple) and len(counts) >= 2 and counts[1]:
                    unsatisfied.append(f"{code} notSatisfied={counts[1]}")
        else:
            error_codes.append(str(item))
    return error_codes, unsatisfied


def instance_output_dir(instance_path: Path, output_dir: Path) -> Path:
    """
    Per-instance folder under Output, named after the XBRL file stem.
    Example: Output/CC.20260601_20260807135250000/
    """
    stem = instance_path.stem.strip().rstrip(". ")
    if not stem:
        stem = "instance"
    # Windows-invalid characters in folder names
    for ch in '<>:"/\\|?*':
        stem = stem.replace(ch, "_")
    return ensure_output_dir(output_dir / stem)


def validation_log_xml_path(instance_dir: Path) -> Path:
    return instance_dir / "Validation_Log.xml"


def validation_log_html_path(instance_dir: Path) -> Path:
    return instance_dir / "Validation_Log.html"


def annotate_validation_xml_log(
    log_xml: str,
    selection: TaxonomySelection,
) -> str:
    """Insert an INFO entry describing the taxonomy applied for validation."""
    pkg = selection.package
    ep = selection.matched_entry_point
    ep_name = ep.name if ep else ""
    ep_desc = ep.description if ep else ""
    msg = (
        f"[INFO] [taxonomyApplied] Taxonomy applied for validation: {pkg.display}. "
        f"Package file: {pkg.path.name}. "
        f"Publisher: {pkg.publisher or 'n/a'}. "
        f"Publication date: {pkg.publication_date or 'n/a'}. "
        f"Identifier: {pkg.identifier or 'n/a'}. "
        f"Matched entry point: {ep_name or 'n/a'}"
        + (f" ({ep_desc})" if ep_desc else "")
        + f". Instance schemaRef: {selection.schema_ref}."
    )
    entry = (
        f'<entry code="taxonomyApplied" level="info">\n'
        f" <message>{html.escape(msg)}</message>\n"
        f' <ref href="{html.escape(str(pkg.path))}"/>'
        f'<property name="taxonomyName" value="{html.escape(pkg.name)}"/>'
        f'<property name="taxonomyVersion" value="{html.escape(pkg.version)}"/>'
        f'<property name="taxonomyPackage" value="{html.escape(pkg.path.name)}"/>'
        f'<property name="schemaRef" value="{html.escape(selection.schema_ref)}"/>'
        f"</entry>\n\n"
    )
    stripped = log_xml.lstrip()
    if stripped.startswith("<log"):
        m = re.match(r"(<\s*log[^>]*>)", stripped, flags=re.IGNORECASE)
        if m:
            start = m.end()
            return stripped[:start] + "\n\n" + entry + stripped[start:]
    return entry + log_xml


def html_report_dir(instance_dir: Path) -> Path:
    return instance_dir / "HTML_Report"


def _zip_member_priority(name: str) -> tuple:
    """Lower tuple sorts first when choosing an instance inside a zip."""
    norm = name.replace("\\", "/")
    base = Path(norm).name
    suffix = Path(base).suffix.casefold()
    depth = norm.count("/")
    # Prefer .xbrl, then iXBRL-ish, then .xml
    ext_rank = {
        ".xbrl": 0,
        ".xhtml": 1,
        ".html": 2,
        ".xml": 3,
    }.get(suffix, 9)
    return (ext_rank, depth, len(norm), norm.casefold())


def _is_taxonomy_package_zip(zip_path: Path) -> bool:
    try:
        with zipfile.ZipFile(zip_path) as zf:
            return any(
                n.replace("\\", "/").endswith("META-INF/taxonomyPackage.xml")
                for n in zf.namelist()
            )
    except (OSError, zipfile.BadZipFile):
        return False


def choose_instance_member(zip_path: Path) -> str:
    """Pick the best XBRL/iXBRL member path from a zip archive."""
    with zipfile.ZipFile(zip_path) as zf:
        members = []
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename.replace("\\", "/")
            base = Path(name).name
            if base.startswith(".") or name.startswith("__MACOSX/"):
                continue
            if Path(base).suffix.casefold() not in INSTANCE_SUFFIXES:
                continue
            members.append(name)
    if not members:
        raise ValueError(
            f"No XBRL/iXBRL file found inside zip: {zip_path.name}. "
            f"Expected a member with extension: {', '.join(sorted(INSTANCE_SUFFIXES))}"
        )
    return sorted(members, key=_zip_member_priority)[0]


def extract_instance_from_zip(zip_path: Path, extract_dir: Path) -> Path:
    """
    Extract the XBRL/iXBRL instance from a compressed zip into extract_dir.
    Returns the path to the extracted instance file.
    """
    if _is_taxonomy_package_zip(zip_path):
        # Avoid treating taxonomy packages dropped into XBRL/ as filings.
        # Still allow if an .xbrl member exists alongside taxonomy metadata.
        pass

    member = choose_instance_member(zip_path)
    if extract_dir.exists():
        shutil.rmtree(extract_dir)
    extract_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path) as zf:
        # Extract only the chosen member (and flatten to extract_dir root).
        target_name = Path(member.replace("\\", "/")).name
        target_path = extract_dir / target_name
        with zf.open(member) as src, open(target_path, "wb") as dst:
            shutil.copyfileobj(src, dst)

    if not target_path.is_file() or target_path.stat().st_size == 0:
        raise RuntimeError(f"Failed to extract instance from zip: {zip_path}")
    return target_path


def prepare_instance_file(source_path: Path, instance_dir: Path) -> Path:
    """
    Resolve a user-supplied instance path to a concrete XBRL/iXBRL file.
    If source_path is a .zip, extract its XBRL member under instance_dir/_extracted.
    """
    if source_path.suffix.casefold() != ZIP_SUFFIX:
        return source_path

    print(f"Compressed input: {source_path}")
    extracted = extract_instance_from_zip(source_path, instance_dir / "_extracted")
    print(f"Extracted XBRL:   {extracted.name} (from {source_path.name})")
    return extracted


def discover_instance_files(xbrl_dir: Path) -> list[Path]:
    """Return XBRL/iXBRL/.zip files in xbrl_dir (non-recursive), sorted by name."""
    if not xbrl_dir.is_dir():
        raise FileNotFoundError(
            f"XBRL folder not found: {xbrl_dir}. "
            "Place instance .xbrl or .zip files in this folder, or pass --instance."
        )
    files = [
        p
        for p in sorted(xbrl_dir.iterdir(), key=lambda x: x.name.casefold())
        if p.is_file() and p.suffix.casefold() in INPUT_SUFFIXES
    ]
    # Skip taxonomy packages mistakenly placed in XBRL/
    filtered: list[Path] = []
    for path in files:
        if path.suffix.casefold() == ZIP_SUFFIX and _is_taxonomy_package_zip(path):
            try:
                choose_instance_member(path)
            except ValueError:
                print(
                    f"Skipping taxonomy package in XBRL folder (not an instance zip): "
                    f"{path.name}"
                )
                continue
        filtered.append(path)
    if not filtered:
        raise FileNotFoundError(
            f"No XBRL/iXBRL/.zip instance files found in {xbrl_dir}. "
            f"Expected extensions: {', '.join(sorted(INPUT_SUFFIXES))}"
        )
    return filtered


def _table_render_name(table_id: str) -> str:
    name = table_id or ""
    for prefix in TABLE_ID_PREFIXES:
        if name.startswith(prefix):
            return name[len(prefix) :]
    return name


def _walk_group_tables(model) -> list:
    from arelle import XbrlConst
    from arelle.ModelRenderingObject import DefnMdlTable

    group = model.relationshipSet(XbrlConst.euGroupTable)
    tables: list = []

    def walk(obj) -> None:
        if obj is None:
            return
        if isinstance(obj, DefnMdlTable):
            tables.append(obj)
        for rel in group.fromModelObject(obj):
            walk(rel.toModelObject)

    for root in group.rootConcepts:
        walk(root)
    return tables


def _iter_layout_data_cells(lyt_table):
    """Yield non-surrogate body data cells from a rendered layout table."""
    for z_cell in getattr(lyt_table, "lytMdlBodyChildren", None) or []:
        for y_cell in getattr(z_cell, "lytMdlBodyChildren", None) or []:
            for x_cell in getattr(y_cell, "lytMdlBodyChildren", None) or []:
                children = getattr(x_cell, "lytMdlBodyChildren", None) or []
                if any(getattr(c, "isOpenAspectEntrySurrogate", False) for c in children):
                    continue
                for cell in children:
                    if getattr(cell, "isOpenAspectEntrySurrogate", False):
                        continue
                    yield cell


def _layout_table_view(model, render_name: str, *, lang: str = "en"):
    """
    Layout one taxonomy table once and return the Arelle view.

    Callers should reuse this view for HTML render (sourceView=), Z-axis
    slices, cell tooltips, and coverage instead of calling layoutTable again.
    """
    from arelle.ViewFileRenderedGrid import ViewRenderedGrid
    from arelle.rendering.RenderingLayout import layoutTable

    view = ViewRenderedGrid(model, None, lang, "")
    try:
        layoutTable(view, render_name)
    except Exception:
        return None
    if not getattr(view, "lytMdlTblMdl", None):
        return None
    return view


def _lyt_tables_from_view(view) -> list:
    table_sets = getattr(view.lytMdlTblMdl, "lytMdlTableSets", None) or ()
    if not table_sets:
        return []
    tables = getattr(table_sets[0], "lytMdlTables", None)
    return tables if tables else []


def _count_cells_in_layout_view(view) -> tuple[int, int]:
    """Return (total_cells, populated_cells) for a laid-out table view."""
    t_total = 0
    t_pop = 0
    for lyt_table in _lyt_tables_from_view(view):
        for cell in _iter_layout_data_cells(lyt_table):
            t_total += 1
            if getattr(cell, "facts", None):
                t_pop += 1
    return t_total, t_pop


def compute_cell_coverage(model, *, lang: str = "en") -> CellCoverage:
    """
    Count taxonomy report cells vs cells populated with instance facts.
    Empty cells are not validated for those dimensions/metrics because no
    value was reported.
    """
    from arelle.rendering import RenderingEvaluator

    RenderingEvaluator.init(model)
    coverage = CellCoverage(instance_facts=len(getattr(model, "facts", None) or []))

    for model_table in _walk_group_tables(model):
        table_id = model_table.id or "table"
        label = model_table.genLabel(lang=lang, strip=True) or table_id
        render_name = _table_render_name(table_id)
        view = _layout_table_view(model, render_name, lang=lang)
        if view is None:
            continue

        t_total, t_pop = _count_cells_in_layout_view(view)
        coverage.tables += 1
        coverage.total_cells += t_total
        coverage.populated_cells += t_pop
        coverage.per_table.append(
            {
                "table_id": table_id,
                "label": label,
                "total_cells": t_total,
                "populated_cells": t_pop,
                "empty_cells": max(t_total - t_pop, 0),
                "coverage_pct": (
                    round(100.0 * t_pop / t_total, 1) if t_total else 0.0
                ),
            }
        )

    return coverage


def _member_display(model, member_qname_str: str, *, lang: str = "en") -> dict:
    """
    Resolve a Z-axis domain member into display fields.
    Examples:
      eba_CU:GBP          -> code=GBP, label=Pound Sterling
      boe_eba_CU:x9007    -> code=Other, label=Currencies other than...
      boe_eba_MC:x9188    -> code=Options and warrants [Statistics], label=same
      boe_eba_MC:x9245    -> code=Futures/Forwards, label=same
    """
    from arelle.ModelValue import QName

    text = str(member_qname_str or "").strip()
    cache_key = (id(model), text, lang)
    cached = _MEMBER_DISPLAY_CACHE.get(cache_key)
    if cached is not None:
        # Callers may mutate the returned dict; give them a shallow copy.
        return cached.copy()

    prefix = ""
    local = text
    if ":" in text:
        prefix, local = text.split(":", 1)

    ns = model.prefixedNamespaces.get(prefix) if prefix else None
    full_label = ""
    if ns:
        concept = model.qnameConcepts.get(QName(prefix, ns, local))
        if concept is not None:
            full_label = (
                concept.label(preferredLabel=None, lang=lang, strip=True)
                or concept.label(strip=True)
                or ""
            )

    prefix_l = prefix.casefold()
    is_currency_member = (
        prefix_l.endswith("_cu")
        or prefix_l == "eba_cu"
        or prefix_l.endswith(":cu")
        or "eba_cu" in prefix_l
        or prefix_l.endswith("_cud")
    )
    iso_currencies = {"GBP", "EUR", "USD", "CHF", "JPY"}

    # Primary chip text shown in the HTML banner / menu.
    if local.upper() in iso_currencies:
        code = local.upper()
    elif is_currency_member and local.lower().startswith("x"):
        # Only currency-domain "other" members collapse to Other.
        code = "Other"
    elif full_label:
        code = full_label
    else:
        code = local

    result = {
        "qname": text,
        "code": code,
        "label": full_label or code,
        "prefix": prefix,
        "local": local,
        "is_currency": "1" if is_currency_member else "0",
    }
    _MEMBER_DISPLAY_CACHE[cache_key] = result
    return result.copy()


def _axis_kind(slices: list[dict]) -> str:
    """Classify Z-axis for wording: currency | instrument | generic."""
    if not slices:
        return "generic"
    aspects = " ".join(str(s.get("aspect", "")) for s in slices).casefold()
    labels = " ".join(str(s.get("axis_label", "")) for s in slices).casefold()
    qnames = " ".join(str(s.get("qname", "")) for s in slices).casefold()
    if (
        "cud" in aspects
        or "currency" in labels
        or "currency" in aspects
        or any(s.get("is_currency") == "1" for s in slices)
    ):
        return "currency"
    if (
        ":mc" in aspects
        or aspects.endswith("mc")
        or " instrument" in f" {labels}"
        or labels.strip() == "instrument"
        or "eba_mc" in qnames
        or "boe_eba_mc" in qnames
    ):
        return "instrument"
    return "generic"


def _select_z_axis_slices(positions: list[list[dict]]) -> list[dict]:
    """
    Pick one banner member per Z sheet.

    A sheet can carry more than one Z constraint. PRA110 significant-currency
    tables pin eba_dim:CCA to eba_CA:x1 on every sheet and open eba_dim:CUS
    once per currency. Those sheets are the eba_CU values (EUR, GBP, …).
    eba_CA:x1 is the conversion approach, not a currency sheet.
    """
    if not positions:
        return []

    def _values(pred) -> list[tuple]:
        return [
            tuple(m.get("qname") for m in members if pred(m))
            for members in positions
        ]

    def _varies(pred) -> bool:
        vals = _values(pred)
        return any(vals) and len(set(vals)) > 1

    def _is_currency(member: dict) -> bool:
        return member.get("is_currency") == "1"

    if _varies(_is_currency):
        pred = _is_currency
    else:
        aspects: list[str] = []
        for members in positions:
            for member in members:
                aspect = str(member.get("aspect") or "")
                if aspect and aspect not in aspects:
                    aspects.append(aspect)
        varying = [
            aspect
            for aspect in aspects
            if _varies(lambda member, aspect=aspect: str(member.get("aspect") or "") == aspect)
        ]
        if varying:
            chosen_aspect = varying[-1]
            pred = lambda member, aspect=chosen_aspect: str(member.get("aspect") or "") == aspect
        elif any(_is_currency(member) for members in positions for member in members):
            # One currency, plus a constant filter such as eba_CA:x1.
            pred = _is_currency
        else:
            pred = lambda member: True

    slices: list[dict] = []
    for members in positions:
        chosen = [member for member in members if pred(member)] or list(members)
        if chosen:
            slices.append(chosen[-1])
    return slices


def _z_axis_slices_from_view(model, view, *, lang: str = "en") -> list[dict]:
    """
    Return one Z-axis slice per rendered sheet (not one per header constraint).

    Header spans are expanded so a breakdown that covers every sheet, such as
    eba_CA:x1, stays on those sheets instead of becoming an extra banner.
    """
    tables = _lyt_tables_from_view(view)
    if not tables:
        return []

    positions: list[list[dict]] = []
    for prof in _axis_leaf_profiles(tables[0], "z"):
        axis_labels = prof.get("aspect_axis_labels") or {}
        members: list[dict] = []
        seen: set[tuple[str, str]] = set()
        for aspect, value in (prof.get("constraints") or {}).items():
            if not value:
                continue
            text = str(value).strip()
            if not text:
                continue
            key = (str(aspect), text)
            if key in seen:
                continue
            seen.add(key)
            member = _member_display(model, text, lang=lang)
            member["axis_label"] = axis_labels.get(aspect) or "Z-axis"
            member["aspect"] = str(aspect)
            members.append(member)
        if members:
            positions.append(members)

    return _select_z_axis_slices(positions)


def _z_axis_slices(model, render_name: str, *, lang: str = "en") -> list[dict]:
    """
    Return ordered Z-axis member slices for a table (empty if none).
    Works for Currency (CUD), Instrument (MC / MC804), and other Z aspects.
    """
    view = _layout_table_view(model, render_name, lang=lang)
    if view is None:
        return []
    return _z_axis_slices_from_view(model, view, lang=lang)


Z_AXIS_BANNER_CSS = """
.z-axis-banner {
  background: linear-gradient(90deg, #0f3d6e 0%, #1f4e78 55%, #2f6aa3 100%);
  color: #fff;
  border: 2px solid #0b2f54;
  border-radius: 6px;
  padding: 10px 14px;
  margin: 10px 0 6px;
  font-family: Arial, sans-serif;
}
.z-axis-banner .z-axis-code {
  display: inline-block;
  font-size: 16px;
  font-weight: bold;
  letter-spacing: 0.02em;
  background: #f4c542;
  color: #1a1a1a;
  border-radius: 4px;
  padding: 3px 10px;
  margin-right: 10px;
  max-width: 70%;
  vertical-align: middle;
}
.z-axis-banner .z-axis-name {
  font-size: 14px;
  font-weight: bold;
  margin-right: 12px;
}
.z-axis-banner .z-axis-qname {
  font-size: 12px;
  opacity: 0.95;
  font-family: Consolas, monospace;
  background: rgba(255,255,255,0.15);
  padding: 2px 8px;
  border-radius: 3px;
}
.z-axis-banner .z-axis-axis {
  display: block;
  margin-top: 6px;
  font-size: 12px;
  opacity: 0.9;
}
.z-axis-summary {
  font-family: Arial, sans-serif;
  background: #fff8e1;
  border: 1px solid #e0c36a;
  border-radius: 6px;
  padding: 8px 12px;
  margin: 8px 0 12px;
  color: #333;
}
.z-axis-summary strong { color: #1f4e78; }
.z-axis-chip {
  display: inline-block;
  margin: 2px 4px 2px 0;
  padding: 2px 8px;
  border-radius: 999px;
  background: #1f4e78;
  color: #fff;
  font-size: 12px;
  font-weight: bold;
}
"""


# Keep wide/tall taxonomy tables readable: viewport scrollport so BOTH
# horizontal and vertical scrollbars stay visible (not only at table bottom).
REPORT_SCROLL_CSS_START = "/* xbrlxray-scroll-css-start */"
REPORT_SCROLL_CSS_END = "/* xbrlxray-scroll-css-end */"
REPORT_SCROLL_CSS = f"""
{REPORT_SCROLL_CSS_START}
html, body {{
  margin: 0;
  padding: 0;
  height: 100%;
  width: 100%;
  overflow: hidden !important;
}}
.xbrlxray-scroll {{
  box-sizing: border-box;
  display: block;
  position: relative;
  width: 100%;
  height: 100%;
  max-width: 100%;
  max-height: 100%;
  overflow: auto !important; /* x + y scrollbars on the visible viewport */
  -webkit-overflow-scrolling: touch;
  border: 0;
  background: #fff;
  padding: 8px;
}}
.xbrlxray-scroll > table,
.xbrlxray-scroll table {{
  background: #fff;
  width: max-content !important;
  max-width: none !important;
  min-width: 100%;
  table-layout: auto !important;
  border-collapse: collapse;
}}
.xbrlxray-scroll th,
.xbrlxray-scroll td {{
  max-width: none !important;
  white-space: nowrap;
  vertical-align: middle;
}}
.xbrlxray-scroll th.yAxisHdr,
.xbrlxray-scroll th.yAxisHdrWithLeg,
.xbrlxray-scroll th.yAxisHdrAbstract,
.xbrlxray-scroll th.yAxisHdrWithChildrenFirst,
.xbrlxray-scroll th.yAxisHdrAbstractChildrenFirst,
.xbrlxray-scroll th.yAxisTopSpanLeg,
.xbrlxray-scroll th.yAxisSpanLeg,
.xbrlxray-scroll th.tableHdr {{
  white-space: normal;
  min-width: 6em;
  max-width: 28em !important;
}}
{REPORT_SCROLL_CSS_END}
"""


def _inject_report_style(
    head,
    css_text: str,
    *,
    replace_marked: bool = False,
    mark_start: str | None = None,
    mark_end: str | None = None,
) -> None:
    """Append CSS to the report document <head> (create <style> if needed)."""
    from lxml import etree

    if head is None:
        return
    style = None
    for child in head:
        tag = child.tag.lower() if isinstance(child.tag, str) else ""
        if tag.endswith("style"):
            style = child
            break
    if style is None:
        style = etree.SubElement(head, "style")
        style.set("type", "text/css")
        style.text = css_text
        return

    existing = style.text or ""
    start_mark = mark_start or REPORT_SCROLL_CSS_START
    end_mark = mark_end or REPORT_SCROLL_CSS_END
    if replace_marked and start_mark in existing:
        # Replace previous marked CSS block so re-runs update cleanly.
        start = existing.find(start_mark)
        end = existing.find(end_mark)
        if start >= 0 and end >= 0:
            end += len(end_mark)
            style.text = existing[:start].rstrip() + "\n" + css_text + existing[end:]
            return
        style.text = existing + "\n" + css_text
        return

    if css_text.strip() and css_text.strip() in existing:
        return
    style.text = existing + "\n" + css_text


def _apply_scroll_wrap(body, head) -> None:
    """Wrap body content in .xbrlxray-scroll and inject scroll CSS (in-memory)."""
    from lxml import etree

    _inject_report_style(head, REPORT_SCROLL_CSS, replace_marked=True)

    if any(
        isinstance(child.tag, str)
        and child.tag.lower().endswith("div")
        and (child.get("class") or "") == "xbrlxray-scroll"
        for child in body
    ):
        return

    wrapper = etree.Element("div")
    wrapper.set("class", "xbrlxray-scroll")
    for child in list(body):
        body.remove(child)
        wrapper.append(child)
    body.append(wrapper)


def _apply_cell_tips(
    root,
    body,
    head,
    tips: list[str] | None = None,
    diff_statuses: list[str | None] | None = None,
) -> None:
    """Attach data-cell-tip / CSS / JS / diff highlight classes onto td.cell nodes."""
    tips = tips or []
    diff_statuses = diff_statuses or []
    if not tips and not diff_statuses:
        return

    cells = [
        el
        for el in body.iter()
        if isinstance(getattr(el, "tag", None), str)
        and el.tag.lower().endswith("td")
        and "cell" in (el.get("class") or "").split()
    ]
    if not cells:
        return

    diff_lines = {
        "added": "Diff: ADDED (present in Report B only)",
        "changed": "Diff: CHANGED (value differs from Report A)",
        "removed": "Diff: REMOVED (was in Report A)",
    }

    for i, cell in enumerate(cells):
        tip = tips[i] if i < len(tips) else ""
        status = diff_statuses[i] if i < len(diff_statuses) else None
        if status:
            classes = (cell.get("class") or "cell").split()
            marker = f"diff-{status}"
            if marker not in classes:
                classes.append(marker)
            cell.set("class", " ".join(classes))
            extra = diff_lines.get(status, f"Diff: {status}")
            tip = f"{tip}\n{extra}" if tip else extra
        if not tip:
            continue
        cell.set("data-cell-tip", tip)
        cell.set("tabindex", "0")
        cell.set("aria-label", tip.replace("\n", "; "))
        attrib = cell.attrib
        if "title" in attrib:
            del attrib["title"]

    _inject_report_style(
        head,
        CELL_TIP_CSS,
        replace_marked=True,
        mark_start=CELL_TIP_CSS_START,
        mark_end=CELL_TIP_CSS_END,
    )
    if any(diff_statuses):
        _inject_report_style(head, DIFF_CELL_CSS)
    _inject_report_script(root, CELL_TIP_JS)


def _apply_z_axis_banners(body, head, slices: list[dict], table_title: str) -> None:
    """Insert Z-axis summary/banners into the outer table (in-memory)."""
    from lxml import etree

    if not slices:
        return

    _inject_report_style(head, Z_AXIS_BANNER_CSS)

    outer_tables = body.findall(".//table")
    if not outer_tables:
        return
    outer = outer_tables[0]

    title_row = None
    slice_rows: list = []
    for tr in list(outer):
        if not isinstance(tr.tag, str) or not tr.tag.lower().endswith("tr"):
            continue
        inner_table = tr.find(".//table")
        if inner_table is not None:
            slice_rows.append(tr)
        elif title_row is None:
            title_row = tr

    if not slice_rows:
        return

    kind = _axis_kind(slices)
    axis_label = (slices[0].get("axis_label") or "Z-axis").strip()
    aspect = str(slices[0].get("aspect") or "").strip()
    if kind == "currency":
        summary_title = "Currencies in this report: "
        kind_phrase = (
            axis_label
            if axis_label.casefold() not in {"", "z-axis", "sheets"}
            else "Currency"
        )
    elif kind == "instrument":
        summary_title = "Instruments in this report: "
        kind_phrase = "Instrument (derivative product type)"
    else:
        summary_title = f"{axis_label} values in this report: "
        kind_phrase = axis_label

    summary = etree.Element("tr")
    summary_td = etree.SubElement(summary, "td")
    summary_td.set("colspan", "99")
    summary_div = etree.SubElement(summary_td, "div")
    summary_div.set("class", "z-axis-summary")
    strong = etree.SubElement(summary_div, "strong")
    strong.text = summary_title
    for slice_info in slices:
        chip = etree.SubElement(summary_div, "span")
        chip.set("class", "z-axis-chip")
        chip.text = slice_info["code"]
        chip.tail = " "
    note = etree.SubElement(summary_div, "div")
    note.set("style", "margin-top:6px;font-size:12px;")
    aspect_note = f" ({aspect})" if aspect else ""
    note.text = (
        f"Table {table_title} — each section below is one {kind_phrase}"
        f"{aspect_note}. "
        f"Axis: {axis_label}."
    )

    if title_row is not None:
        title_row.addnext(summary)
    elif slice_rows:
        slice_rows[0].addprevious(summary)
    else:
        outer.insert(0, summary)

    for index, tr in enumerate(slice_rows):
        if index >= len(slices):
            break
        info = slices[index]
        banner_tr = etree.Element("tr")
        banner_td = etree.SubElement(banner_tr, "td")
        banner_td.set("colspan", "99")
        banner = etree.SubElement(banner_td, "div")
        banner.set("class", "z-axis-banner")
        safe_id = "".join(ch if ch.isalnum() else "_" for ch in info["local"] or info["code"])
        banner.set("id", f"z-axis-{safe_id}-{index + 1}")

        code_el = etree.SubElement(banner, "span")
        code_el.set("class", "z-axis-code")
        code_el.text = info["code"]

        if info["label"] and info["label"] != info["code"]:
            name_el = etree.SubElement(banner, "span")
            name_el.set("class", "z-axis-name")
            name_el.text = info["label"]

        qn_el = etree.SubElement(banner, "span")
        qn_el.set("class", "z-axis-qname")
        qn_el.text = info["qname"]

        axis_el = etree.SubElement(banner, "span")
        axis_el.set("class", "z-axis-axis")
        axis_el.text = (
            f"{info.get('axis_label', axis_label)} — "
            f"this grid shows values for {info['code']} only"
        )

        tr.addprevious(banner_tr)


def _enhance_html_report(
    html_path: Path,
    *,
    slices: list[dict] | None = None,
    tips: list[str] | None = None,
    diff_statuses: list[str | None] | None = None,
    table_title: str = "",
) -> None:
    """
    Single parse/write HTML enhance pass: Z banners, cell tips, diff
    highlights, and scrolling.
    """
    if not html_path.is_file():
        return

    from lxml import etree

    parser = etree.HTMLParser()
    doc = etree.parse(str(html_path), parser)
    root = doc.getroot()
    body = root.find(".//body")
    if body is None:
        return
    head = root.find(".//head")

    if slices:
        _apply_z_axis_banners(body, head, slices, table_title)
    if tips or diff_statuses:
        _apply_cell_tips(root, body, head, tips, diff_statuses)
    _apply_scroll_wrap(body, head)

    doc.write(
        str(html_path),
        encoding="utf-8",
        xml_declaration=False,
        method="html",
    )


def _enhance_html_report_scrolling(html_path: Path) -> None:
    """Backward-compatible wrapper: scroll enhance only."""
    _enhance_html_report(html_path)


def _enhance_html_with_cell_tooltips(html_path: Path, tips: list[str]) -> None:
    """Backward-compatible wrapper: cell-tip enhance only."""
    _enhance_html_report(html_path, tips=tips)


def _enhance_html_with_z_axis_highlights(
    html_path: Path,
    slices: list[dict],
    *,
    table_title: str,
) -> None:
    """Backward-compatible wrapper: Z-axis enhance only."""
    _enhance_html_report(html_path, slices=slices, table_title=table_title)


CELL_TIP_CSS_START = "/* xbrlxray-cell-tip-css-start */"
CELL_TIP_CSS_END = "/* xbrlxray-cell-tip-css-end */"
CELL_TIP_CSS = f"""
{CELL_TIP_CSS_START}
td.cell[data-cell-tip] {{
  cursor: help;
}}
td.cell[data-cell-tip]:focus {{
  outline: 2px solid #1f4e78;
  outline-offset: -2px;
}}
#xbrlxray-cell-tip {{
  display: none;
  position: absolute;
  z-index: 99999;
  width: max-content;
  min-width: 28em;
  max-width: min(90vw, 72em);
  max-height: 75vh;
  overflow: auto;
  white-space: pre; /* keep each tip line intact; no mid-dimension wrapping */
  background: #ffffff;
  color: #111111;
  border: 1px solid #222222;
  border-radius: 2px;
  padding: 8px 12px;
  font: 12px/1.45 Consolas, "Courier New", monospace;
  box-shadow: 0 4px 12px rgba(0,0,0,.18);
  pointer-events: none;
}}
#xbrlxray-cell-tip.visible {{
  display: block;
}}
{CELL_TIP_CSS_END}
"""

DIFF_CELL_CSS = """
td.cell.diff-changed {
  background: #fff3cd !important;
  box-shadow: inset 0 0 0 2px #e6a700;
}
td.cell.diff-added {
  background: #d4edda !important;
  box-shadow: inset 0 0 0 2px #28a745;
}
td.cell.diff-removed {
  background: #f8d7da !important;
  box-shadow: inset 0 0 0 2px #dc3545;
}
.diff-legend {
  font-family: Arial, sans-serif;
  font-size: 12px;
  margin: 8px 0 12px;
  padding: 8px 10px;
  background: #f7fafc;
  border: 1px solid #d9e2ec;
  border-radius: 4px;
}
.diff-legend span {
  display: inline-block;
  margin-right: 14px;
  padding: 2px 8px;
  border-radius: 3px;
}
.diff-legend .swatch-changed { background: #fff3cd; border: 1px solid #e6a700; }
.diff-legend .swatch-added { background: #d4edda; border: 1px solid #28a745; }
"""

CELL_TIP_JS = """
(function () {
  if (window.__xbrlxrayCellTipBound) return;
  window.__xbrlxrayCellTipBound = true;

  var tip = null;
  var active = null;

  function ensureTip() {
    if (tip && tip.parentNode) return tip;
    tip = document.getElementById("xbrlxray-cell-tip");
    if (!tip) {
      tip = document.createElement("div");
      tip.id = "xbrlxray-cell-tip";
      tip.setAttribute("role", "tooltip");
    }
    // Prefer the scrollport so body { overflow:hidden } cannot clip the tip.
    var parent =
      document.querySelector(".xbrlxray-scroll") ||
      document.body ||
      document.documentElement;
    if (tip.parentNode !== parent) parent.appendChild(tip);
    return tip;
  }

  function place(ev) {
    var node = ensureTip();
    var parent = node.parentNode;
    var parentRect = parent.getBoundingClientRect
      ? parent.getBoundingClientRect()
      : { left: 0, top: 0 };
    var scrollLeft = parent.scrollLeft || 0;
    var scrollTop = parent.scrollTop || 0;
    var clientX = 12;
    var clientY = 12;
    if (ev && ev.clientX != null) {
      clientX = ev.clientX + 14;
      clientY = ev.clientY + 16;
    } else if (active) {
      var r = active.getBoundingClientRect();
      clientX = r.left + 8;
      clientY = r.bottom + 8;
    }
    var x = clientX - parentRect.left + scrollLeft;
    var y = clientY - parentRect.top + scrollTop;
    var w = node.offsetWidth || 0;
    var h = node.offsetHeight || 0;
    var viewW = parent.clientWidth || window.innerWidth;
    var viewH = parent.clientHeight || window.innerHeight;
    if (clientX - parentRect.left + w > viewW - 8) {
      x = scrollLeft + viewW - w - 8;
    }
    if (clientY - parentRect.top + h > viewH - 8) {
      y = scrollTop + viewH - h - 8;
    }
    node.style.left = Math.max(scrollLeft + 8, x) + "px";
    node.style.top = Math.max(scrollTop + 8, y) + "px";
  }

  function show(el, ev) {
    var text = el.getAttribute("data-cell-tip");
    if (!text) return;
    active = el;
    var node = ensureTip();
    node.textContent = text;
    node.classList.add("visible");
    place(ev);
  }

  function hide() {
    active = null;
    if (tip) tip.classList.remove("visible");
  }

  function cellFrom(target) {
    var el = target;
    if (el && el.nodeType === 3) el = el.parentElement;
    if (!el || !el.closest) return null;
    return el.closest("td.cell[data-cell-tip]");
  }

  function bind() {
    ensureTip();
    document.addEventListener("mouseover", function (e) {
      var el = cellFrom(e.target);
      if (el) show(el, e);
    });
    document.addEventListener("mousemove", function (e) {
      if (active) place(e);
    });
    document.addEventListener("mouseout", function (e) {
      var el = cellFrom(e.target);
      if (!el) return;
      var rel = e.relatedTarget;
      if (rel && el.contains(rel)) return;
      if (document.activeElement === el) return;
      hide();
    });
    document.addEventListener("focusin", function (e) {
      var el = cellFrom(e.target);
      if (el) show(el, null);
    });
    document.addEventListener("focusout", function (e) {
      var el = cellFrom(e.target);
      if (!el) return;
      var rel = e.relatedTarget;
      if (rel && el.contains(rel)) return;
      hide();
    });
    document.addEventListener("click", function (e) {
      var el = cellFrom(e.target);
      if (el) show(el, e);
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") hide();
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bind);
  } else {
    bind();
  }
})();
"""


def _header_cell_label_bits(cell) -> tuple[str | None, list[str]]:
    """Extract RC code and human labels from a layout header cell."""
    rc_code: str | None = None
    labels: list[str] = []
    labels_seen: set[str] = set()
    fullmatch = _RC_CODE_RE.fullmatch
    for lab in getattr(cell, "labels", None) or ():
        if isinstance(lab, (tuple, list)):
            text = str(lab[0] or "").strip()
            role = lab[1] if len(lab) > 1 else None
        else:
            text = str(lab or "").strip()
            role = None
        if not text:
            continue
        if role == "rc-code" or fullmatch(text):
            rc_code = text
        elif text not in labels_seen:
            labels_seen.add(text)
            labels.append(text)
    return rc_code, labels


def _header_cell_constraints(cell) -> dict:
    """Non-null aspect constraints from a layout header cell."""
    out: dict = {}
    for c in getattr(cell, "lytMdlConstraints", None) or ():
        value = getattr(c, "value", None)
        aspect = getattr(c, "aspect", None)
        if value is not None and aspect is not None:
            out[aspect] = value
    return out


def _axis_leaf_profiles(lyt_table, axis: str) -> list[dict]:
    """
    Build one profile per leaf position on an axis by expanding header spans
    and merging labels/constraints from outer → inner header rows.
    """
    hdrs = lyt_table.lytMdlAxisHeaders(axis)
    if hdrs is None:
        return []

    is_surrogate = lambda c: getattr(c, "isOpenAspectEntrySurrogate", False)
    rows: list[list] = []
    for group in hdrs.lytMdlGroups:
        axis_label = (getattr(group, "label", None) or "").strip() or "Z-axis"
        for header in group.lytMdlHeaders:
            cells = getattr(header, "lytMdlCells", None) or ()
            if cells and all(is_surrogate(c) for c in cells):
                continue
            expanded: list = []
            expand = expanded.extend
            for cell in cells:
                if is_surrogate(cell):
                    continue
                span = int(getattr(cell, "span", 1) or 1)
                expand(((cell, axis_label),) * max(span, 1))
            if expanded:
                rows.append(expanded)

    if not rows:
        return []

    n = max(map(len, rows))
    label_bits = _header_cell_label_bits
    cell_cons = _header_cell_constraints
    profiles: list[dict] = []
    profiles_append = profiles.append
    for i in range(n):
        constraints: dict = {}
        aspect_axis_labels: dict = {}
        labels: list[str] = []
        labels_seen: set[str] = set()
        rc_code: str | None = None
        for row in rows:
            if i >= len(row):
                continue
            cell, axis_label = row[i]
            cell_rc, cell_labels = label_bits(cell)
            if cell_rc:
                rc_code = cell_rc
            for lab in cell_labels:
                if lab not in labels_seen:
                    labels_seen.add(lab)
                    labels.append(lab)
            cons = cell_cons(cell)
            if cons:
                constraints.update(cons)
                for aspect in cons:
                    aspect_axis_labels[aspect] = axis_label
        profiles_append(
            {
                "rc": rc_code,
                "labels": labels,
                "constraints": constraints,
                "aspect_axis_labels": aspect_axis_labels,
            }
        )
    return profiles


# Concept aspect id used when formatting cell tooltips (hoisted; avoid per-cell import).
try:
    from arelle.ModelFormulaObject import Aspect as _ArelleAspect

    _CONCEPT_ASPECT = _ArelleAspect.CONCEPT
except Exception:  # pragma: no cover - Arelle always present at runtime
    _CONCEPT_ASPECT = 2


def _format_cell_tooltip(
    y_prof: dict | None,
    x_prof: dict | None,
    z_prof: dict | None = None,
) -> str:
    """Human-readable expected metric/dimensions for one report cell."""
    if y_prof is None:
        y_prof = _EMPTY_AXIS_PROF
    if x_prof is None:
        x_prof = _EMPTY_AXIS_PROF
    if z_prof is None:
        z_prof = _EMPTY_AXIS_PROF

    lines: list[str] = []
    append = lines.append
    r_code = y_prof.get("rc")
    c_code = x_prof.get("rc")
    if r_code or c_code:
        append(
            "Cell: ["
            + ", ".join(p for p in (r_code, c_code) if p)
            + "]"
        )
    y_labels = y_prof.get("labels")
    if y_labels:
        append("Row: " + " / ".join(y_labels))
    x_labels = x_prof.get("labels")
    if x_labels:
        append("Column: " + " / ".join(x_labels))
    z_labels = z_prof.get("labels")
    if z_labels:
        z_rc = z_prof.get("rc")
        z_lab = " / ".join(z_labels)
        append(f"Z: {z_rc} — {z_lab}" if z_rc else f"Z: {z_lab}")
    elif z_prof.get("rc"):
        append(f"Z: {z_prof['rc']}")

    merged: dict = {}
    for prof in (z_prof, y_prof, x_prof):
        cons = prof.get("constraints")
        if cons:
            merged.update(cons)

    metric = merged.pop(_CONCEPT_ASPECT, None)
    if metric is None:
        metric = merged.pop(2, None)
    if metric is not None:
        append(f"Metric: {metric}")

    if merged:
        append("Dimensions:")
        lines.extend(
            f"  {aspect} = {value}"
            for aspect, value in sorted(merged.items(), key=lambda kv: str(kv[0]))
        )

    return "\n".join(lines)


def _table_cell_tooltips_from_view(view) -> list[str]:
    """
    Expected metric/dimension tooltip text for each HTML td.cell, in the same
    order Arelle emits body cells (including empty cells).
    """
    tables = _lyt_tables_from_view(view)
    if not tables:
        return []

    tips: list[str] = []
    tips_append = tips.append
    format_tip = _format_cell_tooltip
    empty = _EMPTY_AXIS_PROF
    is_surrogate = lambda c: getattr(c, "isOpenAspectEntrySurrogate", False)

    for lyt_table in tables:
        x_profs = _axis_leaf_profiles(lyt_table, "x")
        y_profs = _axis_leaf_profiles(lyt_table, "y")
        z_profs = _axis_leaf_profiles(lyt_table, "z")
        x_n = len(x_profs)
        y_n = len(y_profs)
        z_n = len(z_profs)
        z_bodies = getattr(lyt_table, "lytMdlBodyChildren", None) or ()
        if not z_bodies:
            continue
        # Arelle emits one HTML grid per child under the primary Z body node.
        for zi, y_body in enumerate(z_bodies[0].lytMdlBodyChildren or ()):
            z_prof = z_profs[zi] if zi < z_n else empty
            for yi, x_body in enumerate(getattr(y_body, "lytMdlBodyChildren", None) or ()):
                y_prof = y_profs[yi] if yi < y_n else empty
                for xi, cell in enumerate(
                    getattr(x_body, "lytMdlBodyChildren", None) or ()
                ):
                    if is_surrogate(cell):
                        continue
                    x_prof = x_profs[xi] if xi < x_n else empty
                    tip = format_tip(y_prof, x_prof, z_prof)
                    tips_append(tip)
    return tips


def _strip_rc_prefix(code: str) -> str:
    """R0512 / C0010 / r0070 -> 0512 / 0010 / 0070."""
    text = (code or "").strip()
    if len(text) >= 2 and text[0] in "RrCc":
        return text[1:]
    return text


def _cell_detail_from_profiles(y_prof: dict | None, x_prof: dict | None) -> str:
    """Format layout leaf profiles as row|col cell detail (e.g. 0512|0010)."""
    y_prof = y_prof or {}
    x_prof = x_prof or {}
    row = _strip_rc_prefix(str(y_prof.get("rc") or ""))
    col = _strip_rc_prefix(str(x_prof.get("rc") or ""))
    if row or col:
        return f"{row}|{col}"
    return ""


def _index_layout_facts_to_cells(view, lookup: dict[str, str]) -> None:
    """
    Index facts on a laid-out table into lookup keys:
      "{qname}|{contextId}" -> "0512|0010"
      "line:{sourceline}"   -> "0512|0010"
    """
    empty = _EMPTY_AXIS_PROF
    is_surrogate = lambda c: getattr(c, "isOpenAspectEntrySurrogate", False)

    for lyt_table in _lyt_tables_from_view(view):
        x_profs = _axis_leaf_profiles(lyt_table, "x")
        y_profs = _axis_leaf_profiles(lyt_table, "y")
        x_n = len(x_profs)
        y_n = len(y_profs)
        z_bodies = getattr(lyt_table, "lytMdlBodyChildren", None) or ()
        if not z_bodies:
            continue
        for y_body in z_bodies[0].lytMdlBodyChildren or ():
            for yi, x_body in enumerate(getattr(y_body, "lytMdlBodyChildren", None) or ()):
                y_prof = y_profs[yi] if yi < y_n else empty
                for xi, cell in enumerate(
                    getattr(x_body, "lytMdlBodyChildren", None) or ()
                ):
                    if is_surrogate(cell):
                        continue
                    x_prof = x_profs[xi] if xi < x_n else empty
                    detail = _cell_detail_from_profiles(y_prof, x_prof)
                    if not detail:
                        continue
                    for item in getattr(cell, "facts", None) or ():
                        fact = item[0] if isinstance(item, (tuple, list)) and item else item
                        if fact is None:
                            continue
                        qn = str(getattr(fact, "qname", None) or "")
                        ctx = getattr(getattr(fact, "context", None), "id", None) or ""
                        if qn:
                            lookup[f"{qn}|{ctx}"] = detail
                        line = getattr(fact, "sourceline", None)
                        if line:
                            lookup[f"line:{line}"] = detail


def build_fact_cell_lookup(model, *, lang: str = "en") -> dict[str, str]:
    """Build fact→cell-detail map for all taxonomy tables (layouts each table)."""
    from arelle.rendering import RenderingEvaluator

    RenderingEvaluator.init(model)
    lookup: dict[str, str] = {}
    for model_table in _walk_group_tables(model):
        table_id = model_table.id or "table"
        render_name = _table_render_name(table_id)
        view = _layout_table_view(model, render_name, lang=lang)
        if view is None:
            continue
        _index_layout_facts_to_cells(view, lookup)
    return lookup


def enrich_errors_with_cell_detail(
    errors: list[dict],
    cell_lookup: dict[str, str] | None,
) -> None:
    """Attach fact['cell'] (e.g. 0512|0010) using the layout fact→cell map."""
    if not cell_lookup:
        for err in errors:
            for fact in err.get("facts") or []:
                fact.setdefault("cell", "")
        return

    for err in errors:
        for fact in err.get("facts") or []:
            qn = fact.get("qname") or ""
            ctx = fact.get("contextRef") or ""
            line = fact.get("sourceLine") or ""
            cell = ""
            if qn:
                cell = cell_lookup.get(f"{qn}|{ctx}", "")
            if not cell and line:
                cell = cell_lookup.get(f"line:{line}", "")
            fact["cell"] = cell


def _table_cell_tooltips(model, render_name: str, *, lang: str = "en") -> list[str]:
    """Layout a table and return per-cell tooltip strings (convenience wrapper)."""
    view = _layout_table_view(model, render_name, lang=lang)
    if view is None:
        return []
    return _table_cell_tooltips_from_view(view)


def _inject_report_script(root, script_text: str) -> None:
    """
    Install/replace the cell-tooltip script at the end of <body>.

    Must not run from <head> before body exists — that previously aborted
    setup and left only the help cursor with no visible tip.
    """
    from lxml import etree

    if root is None or not script_text.strip():
        return
    marker = "window.__xbrlxrayCellTipBound"
    # Remove any prior copies (head or body) so re-runs stay clean.
    for el in list(root.iter()):
        tag = getattr(el, "tag", None)
        if not isinstance(tag, str) or not tag.lower().endswith("script"):
            continue
        if marker in (el.text or ""):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    body = root.find(".//body")
    parent = body if body is not None else root
    script = etree.SubElement(parent, "script")
    script.set("type", "text/javascript")
    script.text = script_text


def render_html_report(
    model,
    report_dir: Path,
    *,
    lang: str = "en",
    outer_index_path: Path | None = None,
    validation_log_href: str | None = None,
    coverage_out: CellCoverage | None = None,
    diff: XbrlDiffResult | None = None,
    diff_report_href: str | None = None,
    cell_lookup_out: dict[str, str] | None = None,
) -> Path:
    """
    Render taxonomy tables filled with instance values to HTML.
    Multi-value Z-axis tables get highlighted section banners
    (currency, instrument/MC804, etc.).

    Layouts each table once and reuses that layout for HTML output, Z-axis
    slices, cell tooltips, and optional coverage counts. HTML post-processing
    (Z banners, tips, scrolling) is a single parse/write pass per file.

    When ``diff`` is set (XBRLDiff Report B), added/changed cells are highlighted.

    Writes index.html inside report_dir. When outer_index_path is set, also
    writes XBRLXray_Report.html next to validation logs with links into
    HTML_Report/ so the user can open it without entering the report folder.
    Returns path to report_dir/index.html.
    """
    from arelle.rendering import RenderingEvaluator
    from arelle.ViewFileRenderedGrid import viewRenderedGrid

    report_dir.mkdir(parents=True, exist_ok=True)
    index_path = report_dir / "index.html"

    RenderingEvaluator.init(model)
    tables = _walk_group_tables(model)
    menu_items: list[tuple[str, str, str, str]] = []
    added_keys = diff.added_keys if diff is not None else set()
    changed_keys = diff.changed_keys if diff is not None else set()

    # Render with exact XBRL lexical values (Arelle otherwise rounds by decimals).
    with _exact_xbrl_fact_display():
        for model_table in tables:
            table_id = model_table.id or "table"
            label = model_table.genLabel(lang=lang, strip=True) or table_id
            render_name = _table_render_name(table_id)
            out_file = report_dir / f"{table_id}.html"

            # One layoutTable for HTML + Z + tips (+ coverage).
            layout_view = _layout_table_view(model, render_name, lang=lang)
            if layout_view is None:
                continue

            viewRenderedGrid(
                model,
                str(out_file),
                lang=lang,
                table=render_name,
                sourceView=layout_view,
                cssExtras=(
                    "table {background:#fff; width:max-content; max-width:none;} "
                    "th,td {max-width:none !important;}"
                ),
            )

            if coverage_out is not None:
                t_total, t_pop = _count_cells_in_layout_view(layout_view)
                coverage_out.tables += 1
                coverage_out.total_cells += t_total
                coverage_out.populated_cells += t_pop
                coverage_out.per_table.append(
                    {
                        "table_id": table_id,
                        "label": label,
                        "total_cells": t_total,
                        "populated_cells": t_pop,
                        "empty_cells": max(t_total - t_pop, 0),
                        "coverage_pct": (
                            round(100.0 * t_pop / t_total, 1) if t_total else 0.0
                        ),
                    }
                )

            z_slices = _z_axis_slices_from_view(model, layout_view, lang=lang)
            if z_slices:
                kind = _axis_kind(z_slices)
                prefix = (
                    "Currencies"
                    if kind == "currency"
                    else "Instruments"
                    if kind == "instrument"
                    else (z_slices[0].get("axis_label") or "Z-axis")
                )
                z_note = f"{prefix}: " + ", ".join(
                    f"{s['code']} ({s['qname']})" for s in z_slices
                )
            else:
                z_note = ""

            cell_tips = _table_cell_tooltips_from_view(layout_view)
            if cell_lookup_out is not None:
                _index_layout_facts_to_cells(layout_view, cell_lookup_out)
            diff_statuses = None
            if diff is not None:
                diff_statuses = _cell_diff_statuses_from_view(
                    layout_view, added_keys, changed_keys
                )

            # One HTML parse/write: Z banners + cell tips + diff + scroll.
            _enhance_html_report(
                out_file,
                slices=z_slices or None,
                tips=cell_tips or None,
                diff_statuses=diff_statuses,
                table_title=f"{label} ({table_id})",
            )

            menu_items.append((table_id, label, out_file.name, z_note))

    def build_index(*, href_prefix: str, include_validation_link: bool) -> str:
        menu_links = "\n".join(
            (
                f'      <li><a href="{html.escape(href_prefix + file_name)}" '
                f'target="reportFrame">{html.escape(label)}</a>'
                f"<div class='table-id'>{html.escape(table_id)}</div>"
                + (
                    f"<div class='currency-note'>{html.escape(z_note)}</div>"
                    if z_note
                    else ""
                )
                + "</li>"
            )
            for table_id, label, file_name, z_note in menu_items
        )
        first_href = (href_prefix + menu_items[0][2]) if menu_items else ""
        validation_link = ""
        if include_validation_link and validation_log_href:
            validation_link = (
                f'    <p><a class="val-link" href="{html.escape(validation_log_href)}">'
                f"Open Validation Log</a></p>\n"
            )
        if include_validation_link and diff_report_href:
            validation_link += (
                f'    <p><a class="val-link" href="{html.escape(diff_report_href)}">'
                f"Open XBRLDiff Report</a></p>\n"
            )
        diff_banner = ""
        if diff is not None:
            diff_banner = (
                f'    <p>XBRLDiff — Report B vs Report A: {html.escape(diff.summary_line)}</p>\n'
                f'    <p><span style="background:#fff3cd;padding:2px 6px;margin-right:8px;">'
                f"Changed</span>"
                f'<span style="background:#d4edda;padding:2px 6px;">Added</span>'
                f" cells are highlighted in tables.</p>\n"
            )
        title = "XBRL HTML Report"
        if diff is not None:
            title = "XBRL HTML Report (XBRLDiff — Report B)"
        return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>{html.escape(title)} — {html.escape(Path(getattr(model, 'uri', 'report')).name)}</title>
  <style>
    html, body {{ margin:0; padding:0; height:100%; font-family: Arial, sans-serif; }}
    #header {{ background:#1f4e78; color:#fff; padding:10px 16px; }}
    #header h1 {{ margin:0; font-size:18px; }}
    #header p {{ margin:4px 0 0; font-size:12px; opacity:.9; }}
    #header a.val-link {{ color:#ffe082; font-weight:bold; }}
    #layout {{ display:flex; height:calc(100vh - 58px); }}
    #menu {{ width:340px; border-right:1px solid #cfd8e3; overflow:auto; background:#f7fafc; }}
    #menu h2 {{ font-size:14px; margin:12px 14px 8px; color:#1f4e78; }}
    #menu ul {{ list-style:none; margin:0; padding:0 8px 16px; }}
    #menu li {{ margin:4px 0; }}
    #menu a {{ display:block; padding:8px 10px; color:#1f4e78; text-decoration:none;
               border-radius:4px; background:#fff; border:1px solid #d9e2ec; }}
    #menu a:hover {{ background:#e8f1f8; }}
    #menu .table-id {{ font-size:11px; color:#666; margin:2px 10px 4px; }}
    #menu .currency-note {{ font-size:11px; color:#8a5a00; margin:0 10px 8px;
                            background:#fff8e1; border:1px solid #e0c36a;
                            border-radius:4px; padding:4px 6px; }}
    #content {{ flex:1; min-width:0; overflow:hidden; }}
    iframe {{ border:0; width:100%; height:100%; overflow:auto; }}
  </style>
</head>
<body>
  <div id="header">
    <h1>{html.escape(title)}</h1>
    <p>Instance: {html.escape(str(getattr(model, 'uri', '')))}</p>
{diff_banner}{validation_link}  </div>
  <div id="layout">
    <div id="menu">
      <h2>Tables ({len(menu_items)})</h2>
      <ul>
{menu_links}
      </ul>
    </div>
    <div id="content">
      <iframe name="reportFrame" src="{html.escape(first_href)}" title="Rendered table"></iframe>
    </div>
  </div>
</body>
</html>
"""

    index_path.write_text(build_index(href_prefix="", include_validation_link=False), encoding="utf-8")
    if outer_index_path is not None:
        outer_index_path.parent.mkdir(parents=True, exist_ok=True)
        outer_index_path.write_text(
            build_index(href_prefix="HTML_Report/", include_validation_link=True),
            encoding="utf-8",
        )
    return index_path


STRUCTURAL_CODE_PREFIXES = (
    "xbrldie:",
    "xbrl.",
    "xbrldt.",
    "xmlSchema:",
    "xml.",
    "arelle:",
    "ix11.",
    "ix:",
    "xbrlte:",
    "lrr.",
    "utre.",
)

# Ordered labels for Validation Log root-cause summary (non-zero only shown).
ROOT_CAUSE_LABELS = (
    "Missing dimension",
    "Invalid taxonomy concept",
    "Incorrect value",
    "Missing mandatory fact",
    "Calculation mismatch",
    "Other",
)


def _ref_properties(ref_el: ET.Element) -> dict[str, str]:
    props: dict[str, str] = {}
    for prop in ref_el.findall("property"):
        name = prop.attrib.get("name")
        if not name:
            continue
        props[name] = prop.attrib.get("value", "")
        # nested entity under contextRef
        for child in prop.findall("property"):
            child_name = child.attrib.get("name")
            if child_name:
                props[f"{name}.{child_name}"] = child.attrib.get("value", "")
    return props


def _ref_dimensions(ref_el: ET.Element) -> list[dict[str, str]]:
    """Extract dimension/member pairs from a fact ref's contextRef properties."""
    dims: list[dict[str, str]] = []
    for prop in ref_el.findall("property"):
        if prop.attrib.get("name") != "contextRef":
            continue
        for child in prop.findall("property"):
            if child.attrib.get("name") != "dimensions":
                continue
            for dim in child.findall("property"):
                axis = (dim.attrib.get("name") or "").strip()
                member = (dim.attrib.get("value") or "").strip()
                if axis:
                    dims.append({"axis": axis, "member": member})
    return dims


def normalize_log_xml(log_xml: str) -> str:
    """Collapse duplicate Arelle log roots / XML declarations into one <log>."""
    text = (log_xml or "").strip()
    if not text:
        return "<log>\n</log>\n"
    text = re.sub(r"<\?xml[^?]*\?>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"</?\s*log[^>]*>", "", text, flags=re.IGNORECASE)
    return f"<log>\n{text.strip()}\n</log>\n"


def is_structural_error(err: dict) -> bool:
    """
    True for taxonomy/schema/dimension structure problems.
    Formula / DQ assertion failures (message:*) stay in the rule section.
    """
    code = err.get("code") or ""
    if code.startswith("message:"):
        return False
    code_l = code.casefold()
    if any(code_l.startswith(p.casefold()) for p in STRUCTURAL_CODE_PREFIXES):
        return True
    msg = f"{err.get('message') or ''} {err.get('formula_message') or ''}".casefold()
    hints = (
        "dimensionally not valid",
        "primaryitem",
        "is not a valid concept",
        "not found in the dts",
        "not found in taxonomy",
        "schema validation",
        "invalid dimension",
        "hypercube",
        "has no taxonomy",
        "unit of measure",
    )
    if (err.get("level") or "").casefold() in {"error", "critical"} and any(
        h in msg for h in hints
    ):
        return True
    # Spec-style ERROR codes (prefix:Name) that are not formula messages.
    if (
        (err.get("level") or "").casefold() in {"error", "critical"}
        and ":" in code
        and not code.startswith("message:")
    ):
        return True
    return False


def classify_error_root_cause(err: dict) -> str:
    """
    Map one validation finding to a root-cause bucket for the summary tree.
    """
    code = err.get("code") or ""
    code_l = code.casefold()
    msg = " ".join(
        str(err.get(k) or "")
        for k in (
            "message",
            "formula_message",
            "display_line",
            "rule_test",
            "header",
            "issue",
        )
    ).casefold()

    # Calculation linkbase / summation inconsistencies
    if (
        "calculation" in code_l
        or "xbrl.5.2" in code_l
        or "summation" in msg
        or "calculation inconsistency" in msg
        or ("calculation" in msg and ("mismatch" in msg or "inconsisten" in msg))
    ):
        return "Calculation mismatch"

    # Dimension / hypercube structure
    if (
        "xbrldie" in code_l
        or "xbrldt" in code_l
        or "dimensionally" in msg
        or "hypercube" in msg
        or "primaryitemdimensionallyinvalid" in code_l
        or "missing dimension" in msg
        or "incomplete dimension" in msg
        or ("dimension" in code_l and err.get("category") == "structural")
    ):
        return "Missing dimension"

    # Concept not in taxonomy / unknown element
    if (
        "not a valid concept" in msg
        or "not found in the dts" in msg
        or "not found in taxonomy" in msg
        or "unknown element" in msg
        or "has no taxonomy" in msg
        or (
            ("xmlschema" in code_l or "schema validation" in msg)
            and ("element" in msg or "attribute" in msg or "concept" in msg)
        )
    ):
        return "Invalid taxonomy concept"

    # Existence / mandatory reporting
    if (
        "existenceassertion" in code_l
        or "existence assertion" in msg
        or "missing fact" in msg
        or "missing mandatory" in msg
        or "mandatory fact" in msg
        or "required fact" in msg
        or ("existence" in code_l and "message:" in code_l)
        or ("not reported" in msg and ("must" in msg or "required" in msg))
    ):
        return "Missing mandatory fact"

    # Formula / value assertions (most message:* rules)
    if code.startswith("message:") or err.get("category") == "rule":
        if "existence" in code_l or "existence" in msg:
            return "Missing mandatory fact"
        return "Incorrect value"

    if err.get("category") == "structural":
        if "schema" in code_l or "xmlschema" in code_l:
            return "Invalid taxonomy concept"
        return "Missing dimension"

    return "Other"


def build_root_cause_analysis(errors: list[dict]) -> list[tuple[str, int]]:
    """Return ordered (root_cause, count) pairs for non-zero buckets."""
    from collections import Counter

    counts: Counter[str] = Counter()
    for err in errors:
        cause = err.get("root_cause") or classify_error_root_cause(err)
        counts[cause] += 1

    ordered: list[tuple[str, int]] = []
    seen: set[str] = set()
    for label in ROOT_CAUSE_LABELS:
        n = counts.get(label, 0)
        if n:
            ordered.append((label, n))
            seen.add(label)
    for label, n in counts.most_common():
        if label not in seen and n:
            ordered.append((label, n))
    return ordered


def format_root_cause_tree(items: list[tuple[str, int]]) -> str:
    """ASCII tree like: ├── 42 errors → Missing dimension"""
    if not items:
        return "No ERROR-level findings to classify."
    lines: list[str] = []
    last = len(items) - 1
    for i, (label, n) in enumerate(items):
        branch = "└──" if i == last else "├──"
        noun = "error" if n == 1 else "errors"
        lines.append(f"{branch} {n} {noun} -> {label}")
    return "\n".join(lines)


def describe_structural_issue(err: dict) -> tuple[str, str]:
    """Return (issue explanation, troubleshooting guidance)."""
    code = err.get("code") or ""
    msg = (err.get("message") or err.get("formula_message") or "").casefold()
    fact = (err.get("facts") or [{}])[0]
    qname = fact.get("qname") or err.get("fact_qname") or "unknown metric"
    context = fact.get("contextRef") or err.get("context_id") or "unknown context"
    line = fact.get("sourceLine") or "n/a"
    dims = err.get("dimensions") or fact.get("dimensions") or []
    if isinstance(dims, list) and dims and isinstance(dims[0], dict):
        dim_txt = ", ".join(
            f"{d.get('axis', '?')}={d.get('member', '?')}" for d in dims
        )
    else:
        dim_txt = str(dims) if dims else "(no dimensions listed)"

    if (
        "PrimaryItemDimensionallyInvalidError" in code
        or "dimensionally not valid" in msg
    ):
        issue = (
            f"Metric/concept {qname} reported in context {context} "
            f"(instance line {line}) is dimensionally invalid against the "
            f"applied taxonomy. The dimension/member combination used for this "
            f"fact is not allowed for that metric (hypercube / table axis mismatch). "
            f"Dimensions on the fact: {dim_txt}."
        )
        tips = (
            "1. Open the XBRL instance at the reported line and inspect the fact "
            f"({qname}) and its contextRef ({context}).\n"
            "2. Compare each dimension member with the taxonomy table definition "
            "for this form (row / column / Z-axis). Remove or correct members that "
            "are not permitted for this metric.\n"
            "3. Confirm the metric QName exists in the taxonomy package shown in "
            "the Taxonomy section above (wrong taxonomy version often causes this).\n"
            "4. Re-map the cell in the reporting template to the correct metric and "
            "dimensions, then regenerate the XBRL.\n"
            "5. Re-run XBRLXray after fixing; structural errors should clear before "
            "formula/DQ rules are considered reliable for that cell."
        )
        return issue, tips

    if "xmlschema" in code.casefold() or "schema validation" in msg:
        issue = (
            f"XML Schema / XBRL structural validation failed ({code or 'schema error'}). "
            "The instance document does not conform to the schema for the applied "
            "taxonomy (element, attribute, or type mismatch)."
        )
        tips = (
            "1. Read the full error detail for the exact element/attribute.\n"
            "2. Check namespaces and schemaRef match the taxonomy package entry point.\n"
            "3. Ensure no custom/unknown tags or attributes were introduced by the "
            "generator.\n"
            "4. Regenerate the file from a template built for this taxonomy version."
        )
        return issue, tips

    if any(
        k in msg
        for k in (
            "not a valid concept",
            "not found in the dts",
            "not found in taxonomy",
        )
    ) or "unknown" in msg and "concept" in msg:
        issue = (
            f"A concept/metric used in the instance ({qname}) is not part of the "
            f"applied taxonomy DTS. This usually means a wrong metric name, namespace, "
            f"or taxonomy version."
        )
        tips = (
            "1. Search the taxonomy package for the metric QName.\n"
            "2. Align the instance schemaRef / entry point with the form being filed.\n"
            "3. Update the reporting mapping to the metric defined for this taxonomy "
            "version and regenerate the XBRL."
        )
        return issue, tips

    issue = (
        f"Structural XBRL/taxonomy error {code or '(see message)'} for "
        f"{qname} / context {context} (line {line}). The instance uses a metric "
        f"or dimension combination that does not match the applied taxonomy."
    )
    tips = (
        "1. Review the full error detail below for the precise failure.\n"
        "2. Check metric QName, context dimensions, and units against the taxonomy.\n"
        "3. Confirm the correct taxonomy package/version was selected "
        "(see Taxonomy section).\n"
        "4. Correct the instance (or reporting mapping) and re-validate."
    )
    return issue, tips


def parse_validation_errors(log_xml: str) -> list[dict]:
    """Parse Arelle XML log into structured error dicts."""
    if not log_xml or not log_xml.strip():
        return []

    text = normalize_log_xml(log_xml)
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        return []

    errors: list[dict] = []
    seen: set[tuple[str, str, str]] = set()
    for entry in root.findall("entry"):
        level = (entry.attrib.get("level") or "").casefold()
        code = entry.attrib.get("code") or ""
        if level not in {"error", "critical"} and not code.startswith("message:"):
            continue
        if level not in {"error", "critical"} and "error" not in level:
            if not code.startswith("message:"):
                continue

        message_el = entry.find("message")
        message_text = (message_el.text or "").strip() if message_el is not None else ""
        header = ""
        formula_message = ""
        fact_qname = ""
        context_id = ""
        if message_el is not None:
            header = (message_el.attrib.get("label") or "").strip()
            formula_message = (message_el.attrib.get("formulaMessage") or "").strip()
            fact_qname = (message_el.attrib.get("fact") or "").strip()
            context_id = (message_el.attrib.get("contextID") or "").strip()

        rule_id = code.removeprefix("message:") if code else ""
        rule_test = ""
        rule_href = ""
        facts: list[dict[str, str]] = []
        dimensions: list[dict[str, str]] = []

        for ref in entry.findall("ref"):
            href = ref.attrib.get("href") or ""
            props = _ref_properties(ref)
            ref_dims = _ref_dimensions(ref)
            if "test" in props or href.endswith(f"#{rule_id}") or "/val/" in href:
                rule_test = props.get("test") or rule_test
                rule_href = href or rule_href
                if not rule_id:
                    rule_id = props.get("id") or props.get("label") or rule_id
            if props.get("QName") or props.get("name") or props.get("value"):
                dim_txt = "; ".join(
                    f"{d['axis']}={d['member']}" for d in ref_dims
                )
                facts.append(
                    {
                        "label": props.get("label", ""),
                        "qname": props.get("QName") or props.get("name", ""),
                        "value": props.get("value", ""),
                        "contextRef": props.get("contextRef", ""),
                        "entity": props.get("contextRef.entity", ""),
                        "period": props.get("instant")
                        or props.get("contextRef.instant", ""),
                        "href": href,
                        "sourceLine": ref.attrib.get("sourceLine", ""),
                        "dimensions": dim_txt,
                        "cell": "",
                    }
                )
                if ref_dims and not dimensions:
                    dimensions = ref_dims

        if not header:
            if formula_message.startswith("[") and "]" in formula_message:
                header = formula_message.split("]", 1)[0].strip("[]")
            else:
                header = rule_id or code or "Validation error"

        assertion_detail = formula_message
        if (not assertion_detail) or assertion_detail.endswith("..."):
            assertion_detail = message_text or formula_message

        display_line = message_text or assertion_detail or code
        dedupe_key = (code, display_line, facts[0].get("sourceLine", "") if facts else "")
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        err = {
            "code": code,
            "rule_id": rule_id,
            "level": level or "error",
            "header": header,
            "formula_message": assertion_detail,
            "message": message_text,
            "display_line": display_line,
            "rule_test": rule_test,
            "rule_href": rule_href,
            "facts": facts,
            "fact_qname": fact_qname,
            "context_id": context_id,
            "dimensions": dimensions,
        }
        err["category"] = "structural" if is_structural_error(err) else "rule"
        if err["category"] == "structural":
            issue, tips = describe_structural_issue(err)
            err["issue"] = issue
            err["troubleshooting"] = tips
        err["root_cause"] = classify_error_root_cause(err)
        errors.append(err)
    return errors


def write_readable_validation_log(
    errors: list[dict],
    output_path: Path,
    *,
    instance_path: Path,
    taxonomy_selection: TaxonomySelection,
    coverage: CellCoverage | None = None,
    stats: ValidationStats | None = None,
) -> None:
    """Write HTML validation log with one section per error (header + rule)."""
    pkg = taxonomy_selection.package
    ep = taxonomy_selection.matched_entry_point
    ep_label = ""
    if ep is not None:
        bits = [p for p in (ep.name, ep.description) if p]
        ep_label = " — ".join(bits) if bits else ep.href

    stats_block = ""
    if stats is not None:
        stats_block = f"""
  <div class="stats">
    <div><span class="k">Validation rules (total)</span> {stats.rules_total}</div>
    <div><span class="k">Rules checked</span> {stats.rules_checked}</div>
    <div><span class="k">Rules passed</span> {stats.rules_passed}</div>
    <div><span class="k">Rules failed</span> {stats.rules_failed}</div>
    <div><span class="k">Evaluations passed</span> {stats.evaluations_passed}</div>
    <div><span class="k">Evaluations failed</span> {stats.evaluations_failed}</div>
    <div><span class="k">Errors</span> {stats.errors}</div>
    <div><span class="k">Warnings</span> {stats.warnings}</div>
    <p class="note">
      Rules checked are validation assertions that fired at least once for this filing.
      Passed = checked with no failures; Failed = checked with one or more unsatisfied evaluations.
      Errors/Warnings are validation message counts for this XBRL file.
    </p>
  </div>
"""

    coverage_block = ""
    if coverage is not None:
        table_rows = ""
        for row in coverage.per_table:
            table_rows += (
                "<tr>"
                f"<td>{html.escape(row['label'])}</td>"
                f"<td>{html.escape(row['table_id'])}</td>"
                f"<td>{row['total_cells']}</td>"
                f"<td>{row['populated_cells']}</td>"
                f"<td>{row['empty_cells']}</td>"
                f"<td>{row['coverage_pct']}%</td>"
                "</tr>"
            )
        if not table_rows:
            table_rows = "<tr><td colspan='6'>(no report tables found)</td></tr>"
        coverage_block = f"""
  <div class="coverage">
    <div><span class="k">Cell coverage</span> {coverage.coverage_pct}%</div>
    <div><span class="k">Report cells (taxonomy tables)</span> {coverage.total_cells}</div>
    <div><span class="k">Populated by user/instance</span> {coverage.populated_cells}</div>
    <div><span class="k">Empty / not populated</span> {coverage.empty_cells}</div>
    <div><span class="k">Instance facts</span> {coverage.instance_facts}</div>
    <div><span class="k">Tables counted</span> {coverage.tables}</div>
    <p class="note">
      Validation checks the <strong>{coverage.populated_cells}</strong> populated cells
      (reported facts). The remaining <strong>{coverage.empty_cells}</strong> cells are
      not validated for those dimensions/metrics because they were not populated in
      the XBRL instance.
    </p>
    <table class="coverage-table">
      <thead>
        <tr>
          <th>Table</th><th>ID</th><th>Report cells</th>
          <th>Populated</th><th>Empty</th><th>Coverage</th>
        </tr>
      </thead>
      <tbody>
        {table_rows}
      </tbody>
    </table>
  </div>
"""

    rca_items = build_root_cause_analysis(errors)
    rca_tree = format_root_cause_tree(rca_items)
    root_cause_block = f"""
  <div class="root-cause">
    <h2 class="rca-title">Error Root-Cause Analysis</h2>
    <pre class="rca-tree">{html.escape(rca_tree)}</pre>
  </div>
"""

    structural_errors = [e for e in errors if e.get("category") == "structural"]
    rule_errors = [e for e in errors if e.get("category") != "structural"]

    def _fact_rows(err: dict, include_dims: bool = False) -> str:
        rows = ""
        for fact in err["facts"]:
            rows += (
                "<tr>"
                f"<td>{html.escape(fact.get('label', ''))}</td>"
                f"<td>{html.escape(fact.get('qname', ''))}</td>"
                f"<td>{html.escape(fact.get('value', ''))}</td>"
                f"<td>{html.escape(fact.get('contextRef', ''))}</td>"
                f"<td>{html.escape(fact.get('cell', '') or '—')}</td>"
                f"<td>{html.escape(fact.get('sourceLine', ''))}</td>"
            )
            if include_dims:
                rows += f"<td>{html.escape(fact.get('dimensions', ''))}</td>"
            rows += "</tr>"
        colspan = 7 if include_dims else 6
        return rows or f"<tr><td colspan='{colspan}'>(no fact references)</td></tr>"

    structural_blocks: list[str] = []
    for index, err in enumerate(structural_errors, start=1):
        dim_rows = ""
        for dim in err.get("dimensions") or []:
            dim_rows += (
                "<tr>"
                f"<td>{html.escape(dim.get('axis', ''))}</td>"
                f"<td>{html.escape(dim.get('member', ''))}</td>"
                "</tr>"
            )
        dim_table = ""
        if dim_rows:
            dim_table = f"""
  <h3>Dimensions on fact</h3>
  <table>
    <thead><tr><th>Axis / dimension</th><th>Member</th></tr></thead>
    <tbody>{dim_rows}</tbody>
  </table>
"""
        structural_blocks.append(
            f"""
<section class="error-block structural">
  <h2>Structural Error {index}: {html.escape(err['header'] or err['code'])}</h2>
  <div class="meta">
    <div><span class="k">Error code</span> {html.escape(err['code'])}</div>
    <div><span class="k">Level</span> {html.escape(err['level'].upper())}</div>
    <div><span class="k">Root cause</span> {html.escape(err.get('root_cause') or classify_error_root_cause(err))}</div>
  </div>

  <h3>Error detail</h3>
  <pre class="formula">{html.escape(err.get('display_line') or err.get('message') or '(none)')}</pre>

  <h3>What is the issue?</h3>
  <p class="issue">{html.escape(err.get('issue') or '')}</p>

  <h3>Possible troubleshooting</h3>
  <pre class="tips">{html.escape(err.get('troubleshooting') or '')}</pre>
{dim_table}
  <h3>Affected fact(s)</h3>
  <table>
    <thead>
      <tr>
        <th>Label</th><th>QName</th><th>Value</th>
        <th>Context</th><th>Cell Detail</th><th>Line</th><th>Dimensions</th>
      </tr>
    </thead>
    <tbody>
      {_fact_rows(err, include_dims=True)}
    </tbody>
  </table>
</section>
"""
        )

    rule_blocks: list[str] = []
    for index, err in enumerate(rule_errors, start=1):
        rule_blocks.append(
            f"""
<section class="error-block">
  <h2>Error {index}: {html.escape(err['header'])}</h2>
  <div class="meta">
    <div><span class="k">Rule ID</span> {html.escape(err['rule_id'] or err['code'])}</div>
    <div><span class="k">Level</span> {html.escape(err['level'].upper())}</div>
    <div><span class="k">Root cause</span> {html.escape(err.get('root_cause') or classify_error_root_cause(err))}</div>
  </div>

  <h3>Validation rule</h3>
  <pre class="rule">{html.escape(err['rule_test'] or '(rule expression not available)')}</pre>
  {f'<p class="href">Rule source: {html.escape(err["rule_href"])}</p>' if err.get('rule_href') else ''}

  <h3>Assertion / formula message</h3>
  <pre class="formula">{html.escape(err.get('display_line') or err['formula_message'] or err['message'] or '(none)')}</pre>

  <h3>Affected fact(s)</h3>
  <table>
    <thead>
      <tr>
        <th>Label</th><th>QName</th><th>Value</th>
        <th>Context</th><th>Cell Detail</th><th>Line</th>
      </tr>
    </thead>
    <tbody>
      {_fact_rows(err)}
    </tbody>
  </table>
</section>
"""
        )

    parts: list[str] = []
    if structural_blocks:
        parts.append(
            f'<h1 class="section-title structural-title">Structural Error'
            f"{'s' if len(structural_blocks) != 1 else ''}"
            f" ({len(structural_blocks)})</h1>"
        )
        parts.append(
            "<p class='section-note'>Taxonomy / dimension / metric structure problems "
            "(fact not allowed for the reported axes). Fix these before relying on "
            "formula validation results for the same cells.</p>"
        )
        parts.extend(structural_blocks)
    if rule_blocks:
        parts.append(
            f'<h1 class="section-title">Validation rule failures'
            f" ({len(rule_blocks)})</h1>"
        )
        parts.extend(rule_blocks)
    if not parts:
        parts.append("<p class='pass'>No ERROR-level validation findings.</p>")
    body = "\n".join(parts)

    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Validation Log — {html.escape(instance_path.name)}</title>
  <style>
    body {{ font-family: Verdana, Geneva, sans-serif; font-size: 14px; color:#222;
           max-width: 980px; margin: 24px auto; padding: 0 16px 40px; line-height:1.45; }}
    h1 {{ color:#1f4e78; font-size:22px; }}
    h1.section-title {{ font-size:20px; margin-top:28px; }}
    h1.structural-title {{ color:#7a1f1f; }}
    h2 {{ color:#8a1f1f; font-size:16px; margin:0 0 10px; border-bottom:1px solid #e2b6b6;
          padding-bottom:6px; }}
    h3 {{ color:#1f4e78; font-size:14px; margin:16px 0 6px; }}
    .taxonomy {{ background:#eef7ee; border:1px solid #b7d7b7; border-radius:6px;
                 padding:12px 14px; margin:12px 0 22px; }}
    .stats {{ background:#e8f1fb; border:1px solid #9bb8d9; border-radius:6px;
              padding:12px 14px; margin:0 0 22px; }}
    .stats .note {{ margin:10px 0 0; font-size:13px; color:#444; }}
    .coverage {{ background:#fff8e1; border:1px solid #e0c36a; border-radius:6px;
                 padding:12px 14px; margin:0 0 22px; }}
    .coverage .note {{ margin:10px 0 12px; font-size:13px; color:#444; }}
    .coverage-table {{ margin-top:8px; }}
    .root-cause {{ background:#f4f0fa; border:1px solid #b8a4d4; border-radius:6px;
                   padding:12px 14px; margin:0 0 22px; }}
    .root-cause .rca-title {{ color:#4a2f7a; font-size:16px; margin:0 0 10px;
                             border-bottom:1px solid #d4c4ea; padding-bottom:6px; }}
    pre.rca-tree {{ white-space:pre; margin:0; background:#fff; border:1px solid #d4c4ea;
                    border-radius:4px; padding:12px 14px; font-family: Consolas, monospace;
                    font-size:13px; line-height:1.55; color:#2a1f3d; }}
    .error-block {{ background:#fff8f8; border:1px solid #e2b6b6; border-radius:6px;
                    padding:14px 16px; margin:0 0 18px; }}
    .error-block.structural {{ background:#fff4ef; border-color:#d4926a; }}
    .section-note {{ color:#555; margin:0 0 14px; }}
    .issue {{ background:#fff; border-left:4px solid #c45c26; padding:8px 12px; }}
    .meta {{ display:grid; grid-template-columns:1fr 1fr; gap:6px 16px; margin-bottom:8px; }}
    .k {{ font-weight:bold; color:#555; margin-right:8px; }}
    pre.rule, pre.formula, pre.tips {{ white-space:pre-wrap; word-break:break-word; background:#fff;
                             border:1px solid #e5d0d0; border-radius:4px; padding:10px 12px;
                             font-family: Consolas, monospace; font-size:12.5px; }}
    pre.tips {{ font-family: Verdana, Geneva, sans-serif; font-size:13px; border-color:#d4926a; }}
    table {{ border-collapse:collapse; width:100%; font-size:12.5px; }}
    th, td {{ border:1px solid #d9e2ec; padding:6px 8px; text-align:left; vertical-align:top; }}
    th {{ background:#f3f6f9; }}
    .href {{ font-size:12px; color:#555; word-break:break-all; }}
    .pass {{ color:#2e7d4f; font-weight:bold; }}
  </style>
</head>
<body>
  <h1>Validation Log</h1>
  <div class="taxonomy">
    <div><span class="k">Instance</span> {html.escape(str(instance_path))}</div>
    <div><span class="k">Taxonomy applied</span> {html.escape(pkg.display)}</div>
    <div><span class="k">Taxonomy name</span> {html.escape(pkg.name)}</div>
    <div><span class="k">Taxonomy version</span> {html.escape(pkg.version or 'n/a')}</div>
    <div><span class="k">Package file</span> {html.escape(pkg.path.name)}</div>
    <div><span class="k">Package path</span> {html.escape(str(pkg.path))}</div>
    <div><span class="k">Companion packages</span> {html.escape(', '.join(p.name for p in taxonomy_selection.companion_packages) or 'n/a')}</div>
    <div><span class="k">Publisher</span> {html.escape(pkg.publisher or 'n/a')}</div>
    <div><span class="k">Publication date</span> {html.escape(pkg.publication_date or 'n/a')}</div>
    <div><span class="k">Matched entry point</span> {html.escape(ep_label or 'n/a')}</div>
    <div><span class="k">Instance schemaRef</span> {html.escape(taxonomy_selection.schema_ref)}</div>
  </div>
{stats_block}
{coverage_block}
{root_cause_block}
  {body}
</body>
</html>
"""
    output_path.write_text(doc, encoding="utf-8")


def print_readable_errors(errors: list[dict]) -> None:
    print("\n===== Validation errors (readable) =====")
    if not errors:
        print("No ERROR-level findings.")
        return

    rca = format_root_cause_tree(build_root_cause_analysis(errors))
    print("\n----- Error Root-Cause Analysis -----")
    try:
        print(rca)
    except UnicodeEncodeError:
        print(
            rca.replace("├──", "|--")
            .replace("└──", "`--")
            .replace("→", "->")
            .replace("->", "->")
        )

    structural = [e for e in errors if e.get("category") == "structural"]
    rules = [e for e in errors if e.get("category") != "structural"]

    if structural:
        print("\n----- Structural Error -----")
        for index, err in enumerate(structural, start=1):
            print()
            print("=" * 72)
            print(f"STRUCTURAL ERROR {index}: {err['header'] or err['code']}")
            print("-" * 72)
            print(f"Error code: {err['code']}")
            print("Error detail:")
            print(f"  {err.get('display_line') or err.get('message')}")
            print("What is the issue?")
            print(f"  {err.get('issue')}")
            print("Possible troubleshooting:")
            for tip_line in (err.get("troubleshooting") or "").splitlines():
                print(f"  {tip_line}")
            if err["facts"]:
                print("Affected fact(s):")
                for fact in err["facts"]:
                    print(
                        f"  - {fact.get('label') or fact.get('qname')}: "
                        f"value={fact.get('value')!r}, "
                        f"cell={fact.get('cell') or '—'}, "
                        f"context={fact.get('contextRef')}, line={fact.get('sourceLine')}"
                    )
                    if fact.get("dimensions"):
                        print(f"    dimensions: {fact['dimensions']}")

    if rules:
        print("\n----- Validation rule failures -----")
        for index, err in enumerate(rules, start=1):
            print()
            print("=" * 72)
            print(f"ERROR {index}: {err['header']}")
            print("-" * 72)
            print(f"Rule ID:  {err['rule_id'] or err['code']}")
            print("Validation rule:")
            print(f"  {err['rule_test'] or '(rule expression not available)'}")
            print("Assertion message:")
            print(f"  {err.get('display_line') or err.get('formula_message') or err.get('message')}")
            if err["facts"]:
                print("Affected fact(s):")
                for fact in err["facts"]:
                    print(
                        f"  - {fact.get('label') or fact.get('qname')}: "
                        f"value={fact.get('value')!r}, "
                        f"cell={fact.get('cell') or '—'}, "
                        f"context={fact.get('contextRef')}, line={fact.get('sourceLine')}"
                    )
    print("=" * 72)


def xbrl_diff_output_dir(path_a: Path, path_b: Path, output_dir: Path) -> Path:
    """Output folder: Output/XBRLDiff_<stemA>_vs_<stemB>/."""

    def _safe_stem(path: Path) -> str:
        stem = path.stem.strip().rstrip(". ")
        if not stem:
            stem = "instance"
        for ch in '<>:"/\\|?*':
            stem = stem.replace(ch, "_")
        return stem

    return ensure_output_dir(
        output_dir / f"XBRLDiff_{_safe_stem(path_a)}_vs_{_safe_stem(path_b)}"
    )


def write_xbrl_diff_report(
    diff: XbrlDiffResult,
    out_path: Path,
    *,
    report_b_href: str | None = None,
    log_a_href: str | None = None,
    log_b_href: str | None = None,
    max_rows: int = 200,
) -> Path:
    """Write XBRLDiff_Report.html with A vs B summary and B-detail fact changes."""

    def _rows(items: list[FactDiffItem], kind: str) -> str:
        if not items:
            return f"<tr><td colspan='5'><em>None</em></td></tr>"
        lines = []
        for item in items[:max_rows]:
            va = html.escape(item.value_a if item.value_a is not None else "—")
            vb = html.escape(item.value_b if item.value_b is not None else "—")
            lines.append(
                "<tr>"
                f"<td>{html.escape(kind)}</td>"
                f"<td>{html.escape(item.concept)}</td>"
                f"<td>{va}</td>"
                f"<td>{vb}</td>"
                f"<td>{html.escape(item.dims or '—')}</td>"
                "</tr>"
            )
        if len(items) > max_rows:
            lines.append(
                f"<tr><td colspan='5'><em>… and {len(items) - max_rows} more</em></td></tr>"
            )
        return "\n".join(lines)

    links = []
    if report_b_href:
        links.append(f'<a href="{html.escape(report_b_href)}">Open Report B HTML tables</a>')
    if log_a_href:
        links.append(f'<a href="{html.escape(log_a_href)}">Validation Log A</a>')
    if log_b_href:
        links.append(f'<a href="{html.escape(log_b_href)}">Validation Log B</a>')
    links_html = " · ".join(links) if links else ""

    body = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>XBRLDiff — {html.escape(diff.label_a)} vs {html.escape(diff.label_b)}</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; color: #222; }}
    h1 {{ color: #1f4e78; margin-bottom: 6px; }}
    h2 {{ color: #1f4e78; margin-top: 28px; }}
    .compare {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; max-width: 900px; }}
    .card {{ border: 1px solid #cfd8e3; border-radius: 6px; padding: 14px 16px; background: #f7fafc; }}
    .card h3 {{ margin: 0 0 10px; color: #1f4e78; }}
    .card .metric {{ font-size: 15px; margin: 6px 0; }}
    .card .metric strong {{ display: inline-block; min-width: 7em; }}
    .detail {{ background: #fff8e1; border: 1px solid #e0c36a; border-radius: 6px;
               padding: 12px 16px; max-width: 900px; }}
    .detail .metric {{ font-size: 15px; margin: 4px 0; font-family: Consolas, monospace; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 13px; }}
    th, td {{ border: 1px solid #d9e2ec; padding: 6px 8px; text-align: left; vertical-align: top; }}
    th {{ background: #1f4e78; color: #fff; }}
    tr:nth-child(even) {{ background: #f7fafc; }}
    .links {{ margin: 12px 0 20px; }}
    .links a {{ color: #1f4e78; font-weight: bold; margin-right: 12px; }}
  </style>
</head>
<body>
  <h1>XBRLDiff</h1>
  <p>Compare Report A vs Report B (fact identity = concept + context + unit).</p>
  <div class="links">{links_html}</div>

  <div class="compare">
    <div class="card">
      <h3>Report A</h3>
      <div class="metric"><strong>File:</strong> {html.escape(diff.label_a)}</div>
      <div class="metric"><strong>Facts:</strong> {diff.facts_a:,}</div>
      <div class="metric"><strong>Errors:</strong> {diff.errors_a:,}</div>
      <div class="metric"><strong>Warnings:</strong> {diff.warnings_a:,}</div>
    </div>
    <div class="card">
      <h3>Report B</h3>
      <div class="metric"><strong>File:</strong> {html.escape(diff.label_b)}</div>
      <div class="metric"><strong>Facts:</strong> {diff.facts_b:,}</div>
      <div class="metric"><strong>Errors:</strong> {diff.errors_b:,}</div>
      <div class="metric"><strong>Warnings:</strong> {diff.warnings_b:,}</div>
    </div>
  </div>

  <h2>Detail in 2nd XBRL (Report B vs A)</h2>
  <div class="detail">
    <div class="metric">Added Facts       {len(diff.added):>8,}</div>
    <div class="metric">Removed Facts     {len(diff.removed):>8,}</div>
    <div class="metric">Changed Facts     {len(diff.changed):>8,}</div>
    <div class="metric">Unchanged         {diff.unchanged:>8,}</div>
  </div>

  <h2>Changed facts</h2>
  <table>
    <thead><tr><th>Type</th><th>Concept</th><th>Value A</th><th>Value B</th><th>Dimensions</th></tr></thead>
    <tbody>
{_rows(diff.changed, "Changed")}
    </tbody>
  </table>

  <h2>Added facts (in B only)</h2>
  <table>
    <thead><tr><th>Type</th><th>Concept</th><th>Value A</th><th>Value B</th><th>Dimensions</th></tr></thead>
    <tbody>
{_rows(diff.added, "Added")}
    </tbody>
  </table>

  <h2>Removed facts (in A only)</h2>
  <table>
    <thead><tr><th>Type</th><th>Concept</th><th>Value A</th><th>Value B</th><th>Dimensions</th></tr></thead>
    <tbody>
{_rows(diff.removed, "Removed")}
    </tbody>
  </table>
</body>
</html>
"""
    out_path.write_text(body, encoding="utf-8")
    return out_path


def validate_xbrl_diff(
    path_a: Path,
    path_b: Path,
    *,
    taxonomy_dir: Path | None = None,
    output_dir: Path | None = None,
    online: bool = False,
    render_html: bool = True,
) -> int:
    """
    Compare two XBRL instances (Report A vs Report B).

    Writes XBRLDiff_Report.html, validation logs for both sides, and (by default)
    Report B HTML tables with added/changed cells highlighted.
    """
    path_a = path_a.resolve()
    path_b = path_b.resolve()
    taxonomy_dir = (taxonomy_dir or default_taxonomy_dir()).resolve()
    output_dir = ensure_output_dir(output_dir or default_output_dir())

    for path, label in ((path_a, "A"), (path_b, "B")):
        if not path.is_file():
            raise FileNotFoundError(f"XBRLDiff file {label} not found: {path}")
        if path.suffix.casefold() not in INPUT_SUFFIXES:
            raise ValueError(
                f"Unsupported XBRLDiff file {label} type: {path.suffix}. "
                f"Expected: {', '.join(sorted(INPUT_SUFFIXES))}"
            )

    diff_dir = xbrl_diff_output_dir(path_a, path_b, output_dir)
    side_a_dir = ensure_output_dir(diff_dir / "Report_A")
    side_b_dir = ensure_output_dir(diff_dir / "Report_B")
    instance_a = prepare_instance_file(path_a, side_a_dir)
    instance_b = prepare_instance_file(path_b, side_b_dir)

    print(f"XBRLDiff Report A: {path_a}")
    print(f"XBRLDiff Report B: {path_b}")
    print(f"Output dir:        {diff_dir}")

    selection_a = select_taxonomy_for_instance(instance_a, taxonomy_dir)
    selection_b = select_taxonomy_for_instance(instance_b, taxonomy_dir)
    print(f"Taxonomy A: {selection_a.package.display}")
    print(f"Taxonomy B: {selection_b.package.display}")

    def _run_side(instance_path: Path, selection: TaxonomySelection):
        options = RuntimeOptions(
            entrypointFile=str(instance_path),
            packages=selection.arelle_packages,
            validate=True,
            internetConnectivity="online" if online else "offline",
            keepOpen=True,
            logFile="logToBuffer",
            logFormat="[%(levelname)s] [%(messageCode)s] %(message)s - %(file)s",
            logRefObjectProperties=True,
        )
        with Session() as session:
            session.run(options)
            models = session.get_models()
            try:
                log_xml = session.get_logs("xml")
            except ValueError:
                log_xml = ""
            if not models:
                raise RuntimeError(f"No model loaded for {instance_path}")
            model = models[0]
            stats = collect_validation_stats(model, log_xml)
            fact_map = build_fact_value_map(model)
            fact_count = len(getattr(model, "facts", None) or [])
            cell_lookup = build_fact_cell_lookup(model)
            # Return pure data — model is invalid after session closes.
            return {
                "log_xml": log_xml,
                "stats": stats,
                "fact_map": fact_map,
                "fact_count": fact_count,
                "selection": selection,
                "cell_lookup": cell_lookup,
            }

    print("Loading Report A…")
    side_a = _run_side(instance_a, selection_a)
    print(f"  A facts={side_a['fact_count']} errors={side_a['stats'].errors} "
          f"warnings={side_a['stats'].warnings}")

    # Report B kept open for HTML render with highlights.
    print("Loading Report B…")
    options_b = RuntimeOptions(
        entrypointFile=str(instance_b),
        packages=selection_b.arelle_packages,
        validate=True,
        internetConnectivity="online" if online else "offline",
        keepOpen=True,
        logFile="logToBuffer",
        logFormat="[%(levelname)s] [%(messageCode)s] %(message)s - %(file)s",
        logRefObjectProperties=True,
    )

    with Session() as session_b:
        session_b.run(options_b)
        models_b = session_b.get_models()
        try:
            log_b = session_b.get_logs("xml")
        except ValueError:
            log_b = ""
        if not models_b:
            raise RuntimeError(f"No model loaded for {instance_b}")
        model_b = models_b[0]
        stats_b = collect_validation_stats(model_b, log_b)
        map_b = build_fact_value_map(model_b)
        facts_b = len(getattr(model_b, "facts", None) or [])
        print(f"  B facts={facts_b} errors={stats_b.errors} warnings={stats_b.warnings}")

        diff = diff_fact_maps(
            side_a["fact_map"],
            map_b,
            label_a=path_a.name,
            label_b=path_b.name,
            facts_a=side_a["fact_count"],
            facts_b=facts_b,
            stats_a=side_a["stats"],
            stats_b=stats_b,
        )
        print(f"XBRLDiff detail: {diff.summary_line}")

        # Validation logs
        log_a_xml = side_a_dir / "Validation_Log.xml"
        log_a_html = side_a_dir / "Validation_Log.html"
        log_b_xml = side_b_dir / "Validation_Log.xml"
        log_b_html = side_b_dir / "Validation_Log.html"

        annotated_a = annotate_validation_xml_log(
            normalize_log_xml(side_a["log_xml"]), side_a["selection"]
        )
        log_a_xml.write_text(annotated_a, encoding="utf-8")
        errors_a = parse_validation_errors(annotated_a)
        enrich_errors_with_cell_detail(errors_a, side_a.get("cell_lookup"))
        write_readable_validation_log(
            errors_a,
            log_a_html,
            instance_path=instance_a,
            taxonomy_selection=side_a["selection"],
            coverage=None,
            stats=side_a["stats"],
        )

        annotated_b = annotate_validation_xml_log(normalize_log_xml(log_b), selection_b)
        log_b_xml.write_text(annotated_b, encoding="utf-8")

        report_dir = html_report_dir(side_b_dir)
        outer_index = side_b_dir / "XBRLXray_Report.html"
        coverage = None
        cell_lookup_b: dict[str, str] = {}
        if render_html:
            print("Rendering Report B HTML with change highlights…")
            coverage = CellCoverage(instance_facts=facts_b)
            render_html_report(
                model_b,
                report_dir,
                outer_index_path=outer_index,
                validation_log_href="Validation_Log.html",
                coverage_out=coverage,
                diff=diff,
                diff_report_href="../XBRLDiff_Report.html",
                cell_lookup_out=cell_lookup_b,
            )
            # Also expose a top-level shortcut next to the diff report.
            top_report = diff_dir / "XBRLXray_Report.html"
            top_report.write_text(
                outer_index.read_text(encoding="utf-8").replace(
                    'href="HTML_Report/', 'href="Report_B/HTML_Report/'
                ).replace(
                    'href="Validation_Log.html"',
                    'href="Report_B/Validation_Log.html"',
                ).replace(
                    'href="../XBRLDiff_Report.html"',
                    'href="XBRLDiff_Report.html"',
                ),
                encoding="utf-8",
            )
        else:
            cell_lookup_b = build_fact_cell_lookup(model_b)

        errors_b = parse_validation_errors(annotated_b)
        enrich_errors_with_cell_detail(errors_b, cell_lookup_b)
        write_readable_validation_log(
            errors_b,
            log_b_html,
            instance_path=instance_b,
            taxonomy_selection=selection_b,
            coverage=coverage,
            stats=stats_b,
        )

        diff_report = write_xbrl_diff_report(
            diff,
            diff_dir / "XBRLDiff_Report.html",
            report_b_href="Report_B/XBRLXray_Report.html" if render_html else None,
            log_a_href="Report_A/Validation_Log.html",
            log_b_href="Report_B/Validation_Log.html",
        )

    print("\n===== XBRLDiff Result =====")
    print(f"Report A: {diff.label_a}")
    print(f"  Facts: {diff.facts_a:,}  Errors: {diff.errors_a:,}  Warnings: {diff.warnings_a:,}")
    print(f"Report B: {diff.label_b}")
    print(f"  Facts: {diff.facts_b:,}  Errors: {diff.errors_b:,}  Warnings: {diff.warnings_b:,}")
    print("Detail in 2nd XBRL:")
    print(f"  Added Facts     {len(diff.added):,}")
    print(f"  Removed Facts   {len(diff.removed):,}")
    print(f"  Changed Facts   {len(diff.changed):,}")
    print(f"  Unchanged       {diff.unchanged:,}")
    print(f"Diff report: {diff_report}")
    if render_html:
        print(f"Report B HTML: {outer_index}")
        if coverage is not None:
            print(f"Cell coverage: {coverage.summary_line}")

    # Exit 0 even when facts differ — this is a compare tool, not a pass/fail gate.
    # Return 1 only if either side had validation errors.
    if diff.errors_a or diff.errors_b:
        return 1
    return 0


# ---------------------------------------------------------------------------
# Taxonomy Differences (--taxonomy-diff OLD NEW)
# ---------------------------------------------------------------------------


@dataclass
class TaxonomyItemDiff:
    """One taxonomy artifact involved in a TaxonomyDiff result."""

    key: str
    kind: str  # Added | Removed | Changed
    detail_a: str = ""
    detail_b: str = ""
    note: str = ""


@dataclass
class TaxonomyDiffResult:
    """Old taxonomy vs new taxonomy comparison."""

    label_a: str
    label_b: str
    version_a: str = ""
    version_b: str = ""
    path_a: str = ""
    path_b: str = ""
    concepts_a: int = 0
    concepts_b: int = 0
    dimensions_a: int = 0
    dimensions_b: int = 0
    rules_a: int = 0
    rules_b: int = 0
    entry_points_a: int = 0
    entry_points_b: int = 0
    layouts_a: int = 0
    layouts_b: int = 0
    concepts_added: list[TaxonomyItemDiff] = field(default_factory=list)
    concepts_removed: list[TaxonomyItemDiff] = field(default_factory=list)
    concepts_changed: list[TaxonomyItemDiff] = field(default_factory=list)
    dimensions_added: list[TaxonomyItemDiff] = field(default_factory=list)
    dimensions_removed: list[TaxonomyItemDiff] = field(default_factory=list)
    rules_added: list[TaxonomyItemDiff] = field(default_factory=list)
    rules_removed: list[TaxonomyItemDiff] = field(default_factory=list)
    rules_changed: list[TaxonomyItemDiff] = field(default_factory=list)
    entry_points_added: list[TaxonomyItemDiff] = field(default_factory=list)
    entry_points_removed: list[TaxonomyItemDiff] = field(default_factory=list)
    entry_points_changed: list[TaxonomyItemDiff] = field(default_factory=list)
    layouts_added: list[TaxonomyItemDiff] = field(default_factory=list)
    layouts_removed: list[TaxonomyItemDiff] = field(default_factory=list)
    layouts_changed: list[TaxonomyItemDiff] = field(default_factory=list)

    @property
    def report_layout_changed(self) -> bool:
        return bool(
            self.layouts_added or self.layouts_removed or self.layouts_changed
        )

    @property
    def summary_line(self) -> str:
        layout = "Yes" if self.report_layout_changed else "No"
        return (
            f"Concepts +{len(self.concepts_added)}/-{len(self.concepts_removed)}/"
            f"~{len(self.concepts_changed)}; "
            f"Dimensions +{len(self.dimensions_added)}/-{len(self.dimensions_removed)}; "
            f"Rules +{len(self.rules_added)}/-{len(self.rules_removed)}/"
            f"~{len(self.rules_changed)}; "
            f"EntryPoints +{len(self.entry_points_added)}/-{len(self.entry_points_removed)}/"
            f"~{len(self.entry_points_changed)}; "
            f"Layouts changed={layout}"
        )


def resolve_taxonomy_input(
    raw: str,
    *,
    taxonomy_dir: Path,
    label: str = "Taxonomy package",
) -> Path:
    """
    Resolve a taxonomy package .zip argument.

    Bare filenames (no directory) are looked up in ``taxonomy_dir``
    (default: <script>/Taxonomy/). Full/relative paths with a directory are
    used as given.
    """
    path = Path(raw).expanduser()
    if path.parent == Path(".") and not path.is_absolute():
        resolved = (taxonomy_dir / path.name).resolve()
        if not resolved.is_file():
            raise FileNotFoundError(
                f"{label} not found in Taxonomy folder: {resolved}"
            )
        return resolved

    resolved = path.resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"{label} not found: {resolved}")
    return resolved


def pick_taxonomy_all_entry_point(
    info: TaxonomyPackageInfo,
) -> TaxonomyEntryPoint:
    """Prefer the package 'all' entry point (newest dated href if several)."""
    all_eps = [ep for ep in info.entry_points if (ep.name or "").casefold() == "all"]
    if not all_eps:
        all_eps = [
            ep
            for ep in info.entry_points
            if ep.href.casefold().endswith("/mod/all.xsd")
            or "/mod/all." in ep.href.casefold()
        ]
    if not all_eps and info.entry_points:
        all_eps = list(info.entry_points)

    if not all_eps:
        raise ValueError(
            f"No entry points found in taxonomy package: {info.path.name}"
        )

    def _ep_key(ep: TaxonomyEntryPoint) -> str:
        m = re.search(r"(\d{4}-\d{2}-\d{2})", ep.href or "")
        return m.group(1) if m else ""

    all_eps.sort(key=_ep_key, reverse=True)
    return all_eps[0]


def _taxonomy_concept_label(concept) -> str:
    try:
        return (concept.label(lang="en", strip=True) or "").strip()
    except Exception:
        return ""


def _taxonomy_concept_record(concept) -> dict[str, str]:
    return {
        "type": str(getattr(concept, "typeQname", None) or ""),
        "period": str(getattr(concept, "periodType", None) or ""),
        "balance": str(getattr(concept, "balance", None) or ""),
        "abstract": "true" if getattr(concept, "isAbstract", False) else "false",
        "subst": str(getattr(concept, "substitutionGroupQname", None) or ""),
        "label": _taxonomy_concept_label(concept),
    }


def _taxonomy_dim_record(concept) -> dict[str, str]:
    kind = "typed" if getattr(concept, "isTypedDimension", False) else "explicit"
    if getattr(concept, "isTypedDimension", False) and getattr(
        concept, "isExplicitDimension", False
    ):
        kind = "typed+explicit"
    elif not getattr(concept, "isTypedDimension", False) and not getattr(
        concept, "isExplicitDimension", False
    ):
        kind = "dimension"
    return {
        "kind": kind,
        "label": _taxonomy_concept_label(concept),
    }


def _taxonomy_rule_record(assertion) -> dict[str, str]:
    msg = ""
    try:
        msg = (assertion.genLabel(lang="en", strip=True) or "").strip()
    except Exception:
        msg = ""
    if not msg:
        try:
            msg = (assertion.message(satisfied=False, lang="en") or "").strip()
        except Exception:
            msg = ""
    return {
        "qname": str(getattr(assertion, "qname", None) or ""),
        "test": str(getattr(assertion, "test", None) or ""),
        "message": msg,
        "aspectModel": str(getattr(assertion, "aspectModel", None) or ""),
    }


def _fmt_record(rec: dict[str, str]) -> str:
    parts: list[str] = []
    for k, v in rec.items():
        if not v:
            continue
        text = v
        if k == "structure" and len(text) > 220:
            text = text[:220] + "…"
        parts.append(f"{k}={text}")
    return "; ".join(parts) if parts else "—"


_TAXONOMY_DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")


def _normalize_taxonomy_href(href: str) -> str:
    """Strip dated path segments so version bumps don't mark every EP changed."""
    return _TAXONOMY_DATE_RE.sub("{date}", href or "")


def inventory_entry_points(info: TaxonomyPackageInfo) -> dict[str, dict[str, str]]:
    """Entry points from package META-INF, keyed by entry-point name (casefold)."""
    eps: dict[str, dict[str, str]] = {}
    for ep in info.entry_points:
        name = (ep.name or "").strip()
        if not name and ep.href:
            name = Path(ep.href).stem
        if not name:
            continue
        key = name.casefold()
        href = ep.href or ""
        # Do not store raw name casing — BOE packages flip AD/ad between versions.
        eps[key] = {
            "description": (ep.description or "").strip(),
            "module": Path(href).name if href else "",
            "href_norm": _normalize_taxonomy_href(href),
        }
    return eps


def _taxonomy_layout_record(model, table) -> dict[str, str]:
    """
    Structural fingerprint of a report table layout.

    Uses axis direction + definition-node labels/types (not regenerated node
    ids, which change every BOE taxonomy release even when layout is unchanged).
    """
    from arelle import XbrlConst

    label = ""
    try:
        label = (table.genLabel(lang="en", strip=True) or "").strip()
    except Exception:
        label = ""

    tb = model.relationshipSet(XbrlConst.tableBreakdown)
    bt = model.relationshipSet(XbrlConst.tableBreakdownTree)
    sub = model.relationshipSet(XbrlConst.tableDefinitionNodeSubtree)
    parts: list[str] = []

    def _node_label(node) -> str:
        try:
            return (node.genLabel(lang="en", strip=True) or "").strip()
        except Exception:
            return ""

    def _node_tag(node) -> str:
        return str(
            getattr(node, "localName", None)
            or getattr(getattr(node, "qname", None), "localName", None)
            or type(node).__name__
        )

    def _child_sort_key(node):
        return (_node_tag(node), _node_label(node))

    def walk_node(node, depth: int) -> None:
        tag = _node_tag(node)
        nlabel = _node_label(node)
        flags: list[str] = []
        for attr in ("aspectModel", "merge", "abstract", "parentChildOrder"):
            val = getattr(node, attr, None)
            if val is not None and str(val) != "":
                flags.append(f"{attr}={val}")
        parts.append(f"{depth}:{tag}:{nlabel}:{';'.join(flags)}")
        kids = []
        for rel in bt.fromModelObject(node) or ():
            child = rel.toModelObject
            if child is not None:
                kids.append(child)
        for rel in sub.fromModelObject(node) or ():
            child = rel.toModelObject
            if child is not None:
                kids.append(child)
        for child in sorted(kids, key=_child_sort_key):
            walk_node(child, depth + 1)

    breakdowns = list(tb.fromModelObject(table) or ())

    def _breakdown_key(rel) -> tuple:
        axis = rel.toModelObject
        axis_dir = str(
            getattr(rel, "axis", None) or getattr(axis, "axis", None) or rel.get("axis") or ""
        )
        return (axis_dir, _node_label(axis) if axis is not None else "")

    breakdowns.sort(key=_breakdown_key)
    for rel in breakdowns:
        axis = rel.toModelObject
        if axis is None:
            continue
        axis_dir = str(
            getattr(rel, "axis", None) or getattr(axis, "axis", None) or rel.get("axis") or ""
        )
        parts.append(f"AXIS:{axis_dir}:{_node_label(axis)}")
        walk_node(axis, 1)

    return {
        "label": label,
        "breakdowns": str(len(breakdowns)),
        "structure": "|".join(parts),
    }


def inventory_report_layouts(model) -> dict[str, dict[str, str]]:
    """Report tables keyed by table id (layout structure fingerprint)."""
    layouts: dict[str, dict[str, str]] = {}
    for table in _walk_group_tables(model):
        tid = str(getattr(table, "id", None) or getattr(table, "qname", None) or "")
        if not tid:
            continue
        layouts[tid] = _taxonomy_layout_record(model, table)
    return layouts


def _diff_maps(
    map_a: dict[str, dict[str, str]],
    map_b: dict[str, dict[str, str]],
    *,
    track_changed: bool = True,
) -> tuple[list[TaxonomyItemDiff], list[TaxonomyItemDiff], list[TaxonomyItemDiff]]:
    added: list[TaxonomyItemDiff] = []
    removed: list[TaxonomyItemDiff] = []
    changed: list[TaxonomyItemDiff] = []
    keys_a, keys_b = set(map_a), set(map_b)
    for key in sorted(keys_b - keys_a):
        added.append(
            TaxonomyItemDiff(key=key, kind="Added", detail_b=_fmt_record(map_b[key]))
        )
    for key in sorted(keys_a - keys_b):
        removed.append(
            TaxonomyItemDiff(key=key, kind="Removed", detail_a=_fmt_record(map_a[key]))
        )
    if track_changed:
        for key in sorted(keys_a & keys_b):
            if map_a[key] != map_b[key]:
                changed.append(
                    TaxonomyItemDiff(
                        key=key,
                        kind="Changed",
                        detail_a=_fmt_record(map_a[key]),
                        detail_b=_fmt_record(map_b[key]),
                    )
                )
    return added, removed, changed


def load_taxonomy_inventory(
    zip_path: Path,
    *,
    online: bool = False,
) -> tuple[
    TaxonomyPackageInfo,
    TaxonomyEntryPoint,
    dict[str, dict[str, str]],
    dict[str, dict[str, str]],
    dict[str, dict[str, str]],
    dict[str, dict[str, str]],
    dict[str, dict[str, str]],
]:
    """
    Load a taxonomy package via Arelle (prefer 'all' entry point) and return
    inventories: concepts, dimensions, rules, entry points, report layouts.
    """
    zip_path = zip_path.resolve()
    info = _load_taxonomy_package_info(zip_path)
    if info is None:
        raise ValueError(f"Not a taxonomy package (missing META-INF): {zip_path}")
    ep = pick_taxonomy_all_entry_point(info)
    entry_points = inventory_entry_points(info)
    print(f"  Loading {zip_path.name} via entry point: {ep.name or ep.href}")

    options = RuntimeOptions(
        entrypointFile=ep.href,
        packages=[str(zip_path)],
        validate=False,
        internetConnectivity="online" if online else "offline",
        keepOpen=True,
        logFile="logToBuffer",
    )
    with Session() as session:
        session.run(options)
        models = session.get_models()
        if not models:
            raise RuntimeError(f"Arelle did not load taxonomy package: {zip_path.name}")
        model = models[0]

        concepts: dict[str, dict[str, str]] = {}
        dimensions: dict[str, dict[str, str]] = {}
        for qn, concept in (getattr(model, "qnameConcepts", None) or {}).items():
            if not (getattr(concept, "isItem", False) or getattr(concept, "isTuple", False)):
                continue
            key = str(qn)
            concepts[key] = _taxonomy_concept_record(concept)
            if getattr(concept, "isDimensionItem", False):
                dimensions[key] = _taxonomy_dim_record(concept)

        rules: dict[str, dict[str, str]] = {}
        try:
            from arelle.ModelFormulaObject import ModelVariableSetAssertion
        except Exception:
            ModelVariableSetAssertion = ()  # type: ignore

        for vs in getattr(model, "modelVariableSets", None) or []:
            if ModelVariableSetAssertion and not isinstance(vs, ModelVariableSetAssertion):
                continue
            rid = (
                getattr(vs, "id", None)
                or getattr(vs, "xlinkLabel", None)
                or str(getattr(vs, "qname", None) or "")
            )
            if not rid:
                continue
            rules[str(rid)] = _taxonomy_rule_record(vs)

        for ca in getattr(model, "modelConsistencyAssertions", None) or []:
            rid = (
                getattr(ca, "id", None)
                or getattr(ca, "xlinkLabel", None)
                or str(getattr(ca, "qname", None) or ca)
            )
            rules[f"ca:{rid}"] = _taxonomy_rule_record(ca)

        layouts = inventory_report_layouts(model)

    print(
        f"    concepts={len(concepts):,}  dimensions={len(dimensions):,}  "
        f"rules={len(rules):,}  entry_points={len(entry_points):,}  "
        f"layouts={len(layouts):,}"
    )
    return info, ep, concepts, dimensions, rules, entry_points, layouts


def diff_taxonomy_inventories(
    *,
    info_a: TaxonomyPackageInfo,
    info_b: TaxonomyPackageInfo,
    concepts_a: dict[str, dict[str, str]],
    concepts_b: dict[str, dict[str, str]],
    dims_a: dict[str, dict[str, str]],
    dims_b: dict[str, dict[str, str]],
    rules_a: dict[str, dict[str, str]],
    rules_b: dict[str, dict[str, str]],
    entry_points_a: dict[str, dict[str, str]],
    entry_points_b: dict[str, dict[str, str]],
    layouts_a: dict[str, dict[str, str]],
    layouts_b: dict[str, dict[str, str]],
) -> TaxonomyDiffResult:
    """Compare old (A) vs new (B) taxonomy inventories."""
    result = TaxonomyDiffResult(
        label_a=info_a.display,
        label_b=info_b.display,
        version_a=info_a.version or "",
        version_b=info_b.version or "",
        path_a=str(info_a.path),
        path_b=str(info_b.path),
        concepts_a=len(concepts_a),
        concepts_b=len(concepts_b),
        dimensions_a=len(dims_a),
        dimensions_b=len(dims_b),
        rules_a=len(rules_a),
        rules_b=len(rules_b),
        entry_points_a=len(entry_points_a),
        entry_points_b=len(entry_points_b),
        layouts_a=len(layouts_a),
        layouts_b=len(layouts_b),
    )

    result.concepts_added, result.concepts_removed, result.concepts_changed = _diff_maps(
        concepts_a, concepts_b
    )
    result.dimensions_added, result.dimensions_removed, _ = _diff_maps(
        dims_a, dims_b, track_changed=False
    )
    result.rules_added, result.rules_removed, result.rules_changed = _diff_maps(
        rules_a, rules_b
    )
    (
        result.entry_points_added,
        result.entry_points_removed,
        result.entry_points_changed,
    ) = _diff_maps(entry_points_a, entry_points_b)
    result.layouts_added, result.layouts_removed, result.layouts_changed = _diff_maps(
        layouts_a, layouts_b
    )
    return result


def taxonomy_diff_output_dir(path_a: Path, path_b: Path, output_dir: Path) -> Path:
    """Output folder: Output/TaxonomyDiff_<stemA>_vs_<stemB>/."""

    def _safe_stem(path: Path) -> str:
        stem = path.stem.strip().rstrip(". ")
        if not stem:
            stem = "taxonomy"
        for ch in '<>:"/\\|?*':
            stem = stem.replace(ch, "_")
        return stem

    return ensure_output_dir(
        output_dir / f"TaxonomyDiff_{_safe_stem(path_a)}_vs_{_safe_stem(path_b)}"
    )


def write_taxonomy_diff_report(
    diff: TaxonomyDiffResult,
    out_path: Path,
    *,
    max_rows: int = 300,
) -> Path:
    """Write TaxonomyDiff_Report.html with summary counts and detail tables."""

    def _rows(items: list[TaxonomyItemDiff], *, show_both: bool) -> str:
        if not items:
            return "<tr><td colspan='4'><em>None</em></td></tr>"
        lines: list[str] = []
        for item in items[:max_rows]:
            if show_both:
                lines.append(
                    "<tr>"
                    f"<td>{html.escape(item.kind)}</td>"
                    f"<td>{html.escape(item.key)}</td>"
                    f"<td>{html.escape(item.detail_a or '—')}</td>"
                    f"<td>{html.escape(item.detail_b or '—')}</td>"
                    "</tr>"
                )
            else:
                detail = item.detail_b if item.kind == "Added" else item.detail_a
                lines.append(
                    "<tr>"
                    f"<td>{html.escape(item.kind)}</td>"
                    f"<td>{html.escape(item.key)}</td>"
                    f"<td colspan='2'>{html.escape(detail or '—')}</td>"
                    "</tr>"
                )
        if len(items) > max_rows:
            lines.append(
                f"<tr><td colspan='4'><em>… and {len(items) - max_rows} more</em></td></tr>"
            )
        return "\n".join(lines)

    def _section(title: str, items: list[TaxonomyItemDiff], *, show_both: bool) -> str:
        return f"""
  <h2>{html.escape(title)} ({len(items):,})</h2>
  <table>
    <thead>
      <tr>
        <th>Type</th><th>Id / QName</th>
        <th>Old</th><th>New</th>
      </tr>
    </thead>
    <tbody>
{_rows(items, show_both=show_both)}
    </tbody>
  </table>
"""

    body = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>TaxonomyDiff — {html.escape(Path(diff.path_a).name)} vs {html.escape(Path(diff.path_b).name)}</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; color: #222; max-width: 1100px; }}
    h1 {{ color: #1f4e78; margin-bottom: 6px; }}
    h2 {{ color: #1f4e78; margin-top: 28px; }}
    .compare {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
    .card {{ border: 1px solid #cfd8e3; border-radius: 6px; padding: 14px 16px; background: #f7fafc; }}
    .card h3 {{ margin: 0 0 10px; color: #1f4e78; }}
    .card .metric {{ font-size: 14px; margin: 6px 0; word-break: break-all; }}
    .card .metric strong {{ display: inline-block; min-width: 7.5em; }}
    .detail {{ background: #eef7ee; border: 1px solid #b7d7b7; border-radius: 6px;
               padding: 14px 16px; margin: 18px 0; }}
    .detail h2 {{ margin-top: 0; font-size: 18px; }}
    .detail pre {{ margin: 8px 0 0; font-family: Consolas, monospace; font-size: 14px;
                   white-space: pre; line-height: 1.55; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 12.5px; }}
    th, td {{ border: 1px solid #d9e2ec; padding: 6px 8px; text-align: left; vertical-align: top; }}
    th {{ background: #1f4e78; color: #fff; }}
    tr:nth-child(even) {{ background: #f7fafc; }}
  </style>
</head>
<body>
  <h1>Taxonomy Differences</h1>
  <p>Compare existing / old taxonomy (1st) vs latest / upcoming taxonomy (2nd).</p>

  <div class="compare">
    <div class="card">
      <h3>Old taxonomy</h3>
      <div class="metric"><strong>Package:</strong> {html.escape(diff.label_a)}</div>
      <div class="metric"><strong>Version:</strong> {html.escape(diff.version_a or 'n/a')}</div>
      <div class="metric"><strong>File:</strong> {html.escape(Path(diff.path_a).name)}</div>
      <div class="metric"><strong>Concepts:</strong> {diff.concepts_a:,}</div>
      <div class="metric"><strong>Dimensions:</strong> {diff.dimensions_a:,}</div>
      <div class="metric"><strong>Rules:</strong> {diff.rules_a:,}</div>
      <div class="metric"><strong>Entry points:</strong> {diff.entry_points_a:,}</div>
      <div class="metric"><strong>Report tables:</strong> {diff.layouts_a:,}</div>
    </div>
    <div class="card">
      <h3>New taxonomy</h3>
      <div class="metric"><strong>Package:</strong> {html.escape(diff.label_b)}</div>
      <div class="metric"><strong>Version:</strong> {html.escape(diff.version_b or 'n/a')}</div>
      <div class="metric"><strong>File:</strong> {html.escape(Path(diff.path_b).name)}</div>
      <div class="metric"><strong>Concepts:</strong> {diff.concepts_b:,}</div>
      <div class="metric"><strong>Dimensions:</strong> {diff.dimensions_b:,}</div>
      <div class="metric"><strong>Rules:</strong> {diff.rules_b:,}</div>
      <div class="metric"><strong>Entry points:</strong> {diff.entry_points_b:,}</div>
      <div class="metric"><strong>Report tables:</strong> {diff.layouts_b:,}</div>
    </div>
  </div>

  <div class="detail">
    <h2>Summary</h2>
    <pre>Concepts
Added       {len(diff.concepts_added):>6}
Removed     {len(diff.concepts_removed):>6}
Changed     {len(diff.concepts_changed):>6}

Dimensions
Added       {len(diff.dimensions_added):>6}
Removed     {len(diff.dimensions_removed):>6}

Validation Rules
Added       {len(diff.rules_added):>6}
Removed     {len(diff.rules_removed):>6}
Changed     {len(diff.rules_changed):>6}

Entry Points
Added       {len(diff.entry_points_added):>6}
Removed     {len(diff.entry_points_removed):>6}
Changed     {len(diff.entry_points_changed):>6}

Report Layout
Changed     {"Yes" if diff.report_layout_changed else "No":>6}
  Added     {len(diff.layouts_added):>6}
  Removed   {len(diff.layouts_removed):>6}
  Modified  {len(diff.layouts_changed):>6}</pre>
  </div>

{_section("Concepts — Added", diff.concepts_added, show_both=False)}
{_section("Concepts — Removed", diff.concepts_removed, show_both=False)}
{_section("Concepts — Changed", diff.concepts_changed, show_both=True)}
{_section("Dimensions — Added", diff.dimensions_added, show_both=False)}
{_section("Dimensions — Removed", diff.dimensions_removed, show_both=False)}
{_section("Validation Rules — Added", diff.rules_added, show_both=False)}
{_section("Validation Rules — Removed", diff.rules_removed, show_both=False)}
{_section("Validation Rules — Changed", diff.rules_changed, show_both=True)}
{_section("Entry Points — Added", diff.entry_points_added, show_both=False)}
{_section("Entry Points — Removed", diff.entry_points_removed, show_both=False)}
{_section("Entry Points — Changed", diff.entry_points_changed, show_both=True)}
{_section("Report Layout — Added", diff.layouts_added, show_both=False)}
{_section("Report Layout — Removed", diff.layouts_removed, show_both=False)}
{_section("Report Layout — Changed", diff.layouts_changed, show_both=True)}
</body>
</html>
"""
    out_path.write_text(body, encoding="utf-8")
    return out_path


def validate_taxonomy_diff(
    path_a: Path,
    path_b: Path,
    *,
    output_dir: Path | None = None,
    online: bool = False,
) -> int:
    """
    Compare two taxonomy packages (old vs new).

    Writes Output/TaxonomyDiff_<old>_vs_<new>/TaxonomyDiff_Report.html
    """
    path_a = path_a.resolve()
    path_b = path_b.resolve()
    output_dir = ensure_output_dir(output_dir or default_output_dir())

    for path, label in ((path_a, "old"), (path_b, "new")):
        if not path.is_file():
            raise FileNotFoundError(f"TaxonomyDiff {label} package not found: {path}")
        if path.suffix.casefold() != ZIP_SUFFIX:
            raise ValueError(
                f"TaxonomyDiff {label} package must be a .zip: {path}"
            )

    print("===== TaxonomyDiff =====")
    print(f"Old taxonomy: {path_a}")
    print(f"New taxonomy: {path_b}")

    (
        info_a,
        _ep_a,
        concepts_a,
        dims_a,
        rules_a,
        entry_points_a,
        layouts_a,
    ) = load_taxonomy_inventory(path_a, online=online)
    (
        info_b,
        _ep_b,
        concepts_b,
        dims_b,
        rules_b,
        entry_points_b,
        layouts_b,
    ) = load_taxonomy_inventory(path_b, online=online)
    diff = diff_taxonomy_inventories(
        info_a=info_a,
        info_b=info_b,
        concepts_a=concepts_a,
        concepts_b=concepts_b,
        dims_a=dims_a,
        dims_b=dims_b,
        rules_a=rules_a,
        rules_b=rules_b,
        entry_points_a=entry_points_a,
        entry_points_b=entry_points_b,
        layouts_a=layouts_a,
        layouts_b=layouts_b,
    )

    out_dir = taxonomy_diff_output_dir(path_a, path_b, output_dir)
    report_path = write_taxonomy_diff_report(diff, out_dir / "TaxonomyDiff_Report.html")

    print("\n===== TaxonomyDiff Result =====")
    print(f"Old: {diff.label_a}")
    print(f"New: {diff.label_b}")
    print()
    print("Concepts")
    print(f"Added       {len(diff.concepts_added):>6}")
    print(f"Removed     {len(diff.concepts_removed):>6}")
    print(f"Changed     {len(diff.concepts_changed):>6}")
    print()
    print("Dimensions")
    print(f"Added       {len(diff.dimensions_added):>6}")
    print(f"Removed     {len(diff.dimensions_removed):>6}")
    print()
    print("Validation Rules")
    print(f"Added       {len(diff.rules_added):>6}")
    print(f"Removed     {len(diff.rules_removed):>6}")
    print(f"Changed     {len(diff.rules_changed):>6}")
    print()
    print("Entry Points")
    print(f"Added       {len(diff.entry_points_added):>6}")
    print(f"Removed     {len(diff.entry_points_removed):>6}")
    print(f"Changed     {len(diff.entry_points_changed):>6}")
    print()
    print("Report Layout")
    print(f"Changed     {'Yes' if diff.report_layout_changed else 'No':>6}")
    print(f"  Added     {len(diff.layouts_added):>6}")
    print(f"  Removed   {len(diff.layouts_removed):>6}")
    print(f"  Modified  {len(diff.layouts_changed):>6}")
    print()
    print(f"Report: {report_path}")
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Validate XBRL/iXBRL file(s) with Arelle. Taxonomy packages are selected "
            "automatically from the Taxonomy folder using each instance schemaRef "
            "entry point. When no file is given, all files in the XBRL folder "
            "are processed. Writes Validation_Log.xml/.html and HTML_Report*/ "
            "under the Output folder. Also supports --xbrl-diff and --taxonomy-diff."
        )
    )
    parser.add_argument(
        "--taxonomy-diff",
        nargs=2,
        metavar=("OLD_TAXONOMY", "NEW_TAXONOMY"),
        help=(
            "Compare two taxonomy package .zip files (old/existing vs new/upcoming). "
            "Bare filenames are taken from the Taxonomy folder "
            f"({default_taxonomy_dir()}). "
            "Example: --taxonomy-diff boe-statistics-taxonomy-v130.zip "
            "boe-statistics-taxonomy-v131.zip"
        ),
    )
    parser.add_argument(
        "--xbrl-diff",
        nargs=2,
        metavar=("FILE_A", "FILE_B"),
        help=(
            "Compare two XBRL/iXBRL/.zip files (Report A vs Report B). "
            "Bare filenames are taken from the XBRL folder "
            f"({default_xbrl_dir()}). "
            "Example: --xbrl-diff CC_PreFix.xbrl CC_PostFix.xbrl"
        ),
    )
    parser.add_argument(
        "--xbrl",
        dest="xbrl",
        help=(
            "XBRL/iXBRL/.zip file to validate, or a folder of such files. "
            "Bare filenames are taken from the XBRL folder "
            f"({default_xbrl_dir()}). "
            'Example: --xbrl "PRA101 Capital Plus.xbrl"'
        ),
    )
    parser.add_argument(
        "--instance",
        "-i",
        dest="instance",
        help=(
            "Optional XBRL/iXBRL file or .zip (same as --xbrl <file>). "
            "Bare filenames are taken from the XBRL folder."
        ),
    )
    parser.add_argument(
        "positional",
        nargs="*",
        help=(
            "Optional positional arg: instance.xbrl or instance.zip "
            "(bare names resolve under the XBRL folder)."
        ),
    )
    parser.add_argument(
        "--xbrl-dir",
        dest="xbrl_dir",
        default=str(default_xbrl_dir()),
        help=(
            "Folder of XBRL/iXBRL/.zip instances used when no file is given "
            f"(default: {default_xbrl_dir()})"
        ),
    )
    parser.add_argument(
        "--taxonomy-dir",
        dest="taxonomy_dir",
        default=str(default_taxonomy_dir()),
        help=(
            "Folder containing taxonomy package .zip files "
            f"(default: {default_taxonomy_dir()})"
        ),
    )
    parser.add_argument(
        "--output-dir",
        dest="output_dir",
        default=str(default_output_dir()),
        help=(
            "Folder for validation logs and HTML reports "
            f"(default: {default_output_dir()})"
        ),
    )
    parser.add_argument(
        "--online",
        action="store_true",
        help="Allow internet downloads (default: offline)",
    )
    parser.add_argument(
        "--no-html-report",
        action="store_true",
        help="Skip rendering HTML table report",
    )
    args = parser.parse_args(argv)

    # TaxonomyDiff mode: two taxonomy packages (old vs new).
    if args.taxonomy_diff:
        if args.xbrl_diff or args.instance or args.xbrl or args.positional:
            parser.error(
                "When using --taxonomy-diff OLD NEW, do not also pass "
                "--xbrl-diff / --xbrl / --instance / positional instance paths."
            )
        args.taxonomy_dir = Path(args.taxonomy_dir).expanduser()
        args.output_dir = Path(args.output_dir).expanduser()
        args.xbrl_dir = Path(args.xbrl_dir).expanduser()
        try:
            args.tax_diff_a = resolve_taxonomy_input(
                args.taxonomy_diff[0],
                taxonomy_dir=args.taxonomy_dir,
                label="TaxonomyDiff old package",
            )
            args.tax_diff_b = resolve_taxonomy_input(
                args.taxonomy_diff[1],
                taxonomy_dir=args.taxonomy_dir,
                label="TaxonomyDiff new package",
            )
        except FileNotFoundError as exc:
            parser.error(str(exc))
        args.diff_a = None
        args.diff_b = None
        args.instance = None
        return args

    args.tax_diff_a = None
    args.tax_diff_b = None

    # XBRLDiff mode: two files, ignore single-instance resolution.
    if args.xbrl_diff:
        if args.instance or args.xbrl or args.positional:
            parser.error(
                "When using --xbrl-diff FILE_A FILE_B, do not also pass "
                "--xbrl / --instance / positional instance paths."
            )
        args.xbrl_dir = Path(args.xbrl_dir).expanduser()
        args.taxonomy_dir = Path(args.taxonomy_dir).expanduser()
        args.output_dir = Path(args.output_dir).expanduser()
        try:
            args.diff_a = resolve_xbrl_diff_input(
                args.xbrl_diff[0], xbrl_dir=args.xbrl_dir, side="A"
            )
            args.diff_b = resolve_xbrl_diff_input(
                args.xbrl_diff[1], xbrl_dir=args.xbrl_dir, side="B"
            )
        except FileNotFoundError as exc:
            parser.error(str(exc))
        args.instance = None
        return args

    args.diff_a = None
    args.diff_b = None

    args.xbrl_dir = Path(args.xbrl_dir).expanduser()
    args.taxonomy_dir = Path(args.taxonomy_dir).expanduser()
    args.output_dir = Path(args.output_dir).expanduser()

    instance_raw = args.instance
    xbrl_raw = args.xbrl

    # --xbrl may be a single file or a folder (batch).
    if xbrl_raw is not None:
        xbrl_path = Path(xbrl_raw).expanduser()
        looks_like_file = (
            xbrl_path.suffix.casefold() in INPUT_SUFFIXES
            or (xbrl_path.parent == Path(".") and not xbrl_path.is_absolute())
        )
        looks_like_dir = (
            xbrl_path.is_dir()
            or (not xbrl_path.exists() and xbrl_path.suffix == "")
        )
        if looks_like_file and not (xbrl_path.is_dir()):
            if instance_raw:
                parser.error(
                    "Provide the instance once: use --xbrl or --instance, not both."
                )
            try:
                instance_raw = str(
                    resolve_xbrl_input(
                        xbrl_raw, xbrl_dir=args.xbrl_dir, label="XBRL file"
                    )
                )
            except FileNotFoundError as exc:
                parser.error(str(exc))
        elif looks_like_dir:
            args.xbrl_dir = xbrl_path.expanduser().resolve()
        else:
            parser.error(
                f"XBRL path not found: {xbrl_path}. "
                "Pass a .xbrl/.zip file name (looked up in the XBRL folder), "
                "a full file path, or a folder of instance files."
            )

    if not instance_raw and len(args.positional) == 1:
        instance_raw = args.positional[0]
    elif args.positional and instance_raw:
        parser.error(
            "Provide the instance once: --xbrl file.xbrl, --instance file.xbrl, "
            "or positional path."
        )
    elif len(args.positional) > 1:
        parser.error(
            "Provide at most one instance path. To process multiple files, omit "
            "the file path and place files in the XBRL folder (or pass a folder "
            "with --xbrl / --xbrl-dir)."
        )

    if instance_raw:
        # --instance / positional may also be a bare name under XBRL/.
        try:
            args.instance = resolve_xbrl_input(
                instance_raw, xbrl_dir=args.xbrl_dir, label="XBRL file"
            )
        except FileNotFoundError as exc:
            parser.error(str(exc))
    else:
        args.instance = None
    return args


def validate(
    source_path: Path,
    *,
    taxonomy_dir: Path | None = None,
    output_dir: Path | None = None,
    online: bool = False,
    render_html: bool = True,
) -> int:
    source_path = source_path.resolve()
    taxonomy_dir = (taxonomy_dir or default_taxonomy_dir()).resolve()
    output_dir = ensure_output_dir(output_dir or default_output_dir())

    if not source_path.is_file():
        raise FileNotFoundError(f"Instance file not found: {source_path}")
    if source_path.suffix.casefold() not in INPUT_SUFFIXES:
        raise ValueError(
            f"Unsupported instance type: {source_path.suffix}. "
            f"Expected: {', '.join(sorted(INPUT_SUFFIXES))}"
        )

    # Output folder is named after the submitted file (zip or xbrl stem).
    instance_dir = instance_output_dir(source_path, output_dir)
    instance_path = prepare_instance_file(source_path, instance_dir)

    print(f"Source:        {source_path}")
    print(f"Instance:      {instance_path}")
    print(f"Taxonomy dir:  {taxonomy_dir}")
    print(f"Output dir:    {output_dir}")
    print("Resolving taxonomy package from instance entry point…")

    selection = select_taxonomy_for_instance(instance_path, taxonomy_dir)
    taxonomy_path = selection.package.path
    print(f"Entry point:   {selection.schema_ref}")
    if selection.matched_entry_point:
        ep = selection.matched_entry_point
        form = " — ".join(p for p in (ep.name, ep.description) if p)
        print(f"Matched form:  {form}")
    print(f"Taxonomy:      {selection.package.display}")
    print(f"Package file:  {taxonomy_path}")
    if selection.companion_packages:
        for comp in selection.companion_packages:
            print(f"Companion:     {comp.name}")
    elif _looks_like_reporting_framework(selection.package):
        print(
            "WARNING: EBA-style reporting framework selected but no matching "
            "Dictionary package was found in Taxonomy/. Facts may not bind and "
            "HTML report cells may stay empty. Add the matching "
            "EBA_*_Dictionary_*.zip (same major.minor version family)."
        )
    # EBA_DPM_* raw trees without META-INF are not taxonomy packages.
    dpm_hint = list(taxonomy_dir.glob("EBA_DPM*.zip")) + list(
        taxonomy_dir.glob("*DPM*.zip")
    )
    for dpm in dpm_hint:
        if dpm.resolve() == taxonomy_path.resolve():
            continue
        # Only mention if it is not a valid package.
        if _load_taxonomy_package_info(dpm) is None and "dpm" in dpm.name.casefold():
            print(
                f"Note: {dpm.name} is not an XBRL taxonomy package "
                "(no META-INF/taxonomyPackage.xml). Use Reporting Frameworks + "
                "Dictionary zips for validation/reporting."
            )
            break

    log_xml_path = validation_log_xml_path(instance_dir)
    log_html_path = validation_log_html_path(instance_dir)
    report_dir = html_report_dir(instance_dir)
    outer_index_path = instance_dir / "XBRLXray_Report.html"

    print(f"Instance out:  {instance_dir}")
    print(f"XML log:       {log_xml_path}")
    print(f"Validation log: {log_html_path}")
    if render_html:
        print(f"HTML report:   {report_dir}")
        print(f"XBRLXray report: {outer_index_path}")
    print("Running Arelle Session validation…")

    # Buffer Arelle's XML log in memory, then write Validation_Log.xml ourselves.
    # (Using logFile=<path> can append a second <log> root when the handler flushes.)
    options = RuntimeOptions(
        entrypointFile=str(instance_path),
        packages=selection.arelle_packages,
        validate=True,
        internetConnectivity="online" if online else "offline",
        keepOpen=True,
        logFile="logToBuffer",
        logFormat="[%(levelname)s] [%(messageCode)s] %(message)s - %(file)s",
        logRefObjectProperties=True,
    )

    cell_lookup: dict[str, str] = {}
    with Session() as session:
        session.run(options)
        models = session.get_models()
        try:
            log_xml = session.get_logs("xml")
        except ValueError:
            log_xml = ""

        coverage: CellCoverage | None = None
        stats: ValidationStats | None = None
        report_index: Path | None = None
        if models:
            model0 = models[0]
            stats = collect_validation_stats(model0, log_xml)
            print(f"Validation stats: {stats.summary_line}")
            if render_html:
                # Layout once per table inside render; also fill coverage from
                # that same layout (avoids a second full layout walk).
                print("Rendering HTML report tables (with cell coverage)…")
                coverage = CellCoverage(
                    instance_facts=len(getattr(model0, "facts", None) or [])
                )
                report_index = render_html_report(
                    model0,
                    report_dir,
                    outer_index_path=outer_index_path,
                    validation_log_href="Validation_Log.html",
                    coverage_out=coverage,
                    cell_lookup_out=cell_lookup,
                )
                print(f"Cell coverage: {coverage.summary_line}")
                print(f"HTML report index: {report_index}")
                print(f"XBRLXray report: {outer_index_path}")
            else:
                print("Computing cell coverage…")
                coverage = compute_cell_coverage(model0)
                print(f"Cell coverage: {coverage.summary_line}")
                print("Building fact→cell map for validation log…")
                cell_lookup = build_fact_cell_lookup(model0)

    print(f"\nModels loaded: {len(models)}")
    all_error_codes: list[str] = []
    all_unsatisfied: list[str] = []
    for index, model in enumerate(models, start=1):
        doc = getattr(model, "modelDocument", None)
        uri = getattr(doc, "uri", None) if doc is not None else None
        facts = getattr(model, "facts", None) or []
        print(f"  [{index}] uri={uri}")
        print(f"      facts={len(facts)}")
        error_codes, unsatisfied = _summarize_model_errors(model)
        all_error_codes.extend(error_codes)
        all_unsatisfied.extend(unsatisfied)

    normalized_xml = normalize_log_xml(log_xml)
    annotated_xml = annotate_validation_xml_log(normalized_xml, selection)
    log_xml_path.write_text(annotated_xml, encoding="utf-8")
    print(f"Wrote XML log: {log_xml_path}")

    errors = parse_validation_errors(annotated_xml)
    enrich_errors_with_cell_detail(errors, cell_lookup)
    write_readable_validation_log(
        errors,
        log_html_path,
        instance_path=instance_path,
        taxonomy_selection=selection,
        coverage=coverage,
        stats=stats,
    )
    print(f"Wrote validation log: {log_html_path}")
    print_readable_errors(errors)

    print("\n===== Result =====")
    print(f"Taxonomy applied: {selection.package.display}")
    print(f"Taxonomy package: {taxonomy_path.name}")
    print(f"Taxonomy version: {selection.package.version or 'n/a'}")
    if stats is not None:
        print(f"Validation stats: {stats.summary_line}")
    if coverage is not None:
        print(f"Cell coverage:    {coverage.summary_line}")
    print(f"Error messages: {len(errors) or len(all_error_codes)}")
    print(f"Unsatisfied assertions: {len(all_unsatisfied)}")
    print(f"Validation XML log:  {log_xml_path}")
    print(f"Validation HTML log: {log_html_path}")
    if render_html:
        print(f"HTML report folder: {report_dir}")
        print(f"Open report: {outer_index_path}")
    if errors or all_error_codes or all_unsatisfied:
        print("FAILED — validation reported ERROR-level findings.")
        return 1

    print("PASSED — no validation errors reported.")
    return 0


def validate_batch(
    instance_paths: list[Path],
    *,
    taxonomy_dir: Path | None = None,
    output_dir: Path | None = None,
    online: bool = False,
    render_html: bool = True,
) -> int:
    """Validate multiple instances; continue on per-file failures."""
    output_dir = ensure_output_dir(output_dir or default_output_dir())
    taxonomy_dir = (taxonomy_dir or default_taxonomy_dir()).resolve()

    print(f"Batch mode: {len(instance_paths)} file(s)")
    print(f"Taxonomy dir: {taxonomy_dir}")
    print(f"Output dir:   {output_dir}")

    results: list[tuple[Path, int, str]] = []
    for index, instance_path in enumerate(instance_paths, start=1):
        print("\n" + "#" * 72)
        print(f"# [{index}/{len(instance_paths)}] {instance_path.name}")
        print("#" * 72)
        try:
            code = validate(
                instance_path,
                taxonomy_dir=taxonomy_dir,
                output_dir=output_dir,
                online=online,
                render_html=render_html,
            )
            status = "PASSED" if code == 0 else "FAILED"
        except Exception as exc:
            code = 2
            status = f"ERROR: {exc}"
            print(f"ERROR: {exc}", file=sys.stderr)
        results.append((instance_path, code, status))

    print("\n" + "=" * 72)
    print("BATCH SUMMARY")
    print("=" * 72)
    failed = 0
    errored = 0
    for path, code, status in results:
        print(f"  [{code}] {path.name}: {status}")
        if code == 1:
            failed += 1
        elif code >= 2:
            errored += 1
    print(
        f"Total: {len(results)} | "
        f"passed: {sum(1 for _, c, _ in results if c == 0)} | "
        f"failed: {failed} | errors: {errored}"
    )
    print(f"Outputs written under: {output_dir}")

    if errored:
        return 2
    if failed:
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if getattr(args, "tax_diff_a", None) is not None and getattr(
            args, "tax_diff_b", None
        ) is not None:
            return validate_taxonomy_diff(
                args.tax_diff_a,
                args.tax_diff_b,
                output_dir=args.output_dir,
                online=args.online,
            )

        if getattr(args, "diff_a", None) is not None and getattr(args, "diff_b", None) is not None:
            return validate_xbrl_diff(
                args.diff_a,
                args.diff_b,
                taxonomy_dir=args.taxonomy_dir,
                output_dir=args.output_dir,
                online=args.online,
                render_html=not args.no_html_report,
            )

        if args.instance is not None:
            return validate(
                args.instance,
                taxonomy_dir=args.taxonomy_dir,
                output_dir=args.output_dir,
                online=args.online,
                render_html=not args.no_html_report,
            )

        instances = discover_instance_files(args.xbrl_dir.resolve())
        print(
            f"No --instance given; processing {len(instances)} file(s) "
            f"from {args.xbrl_dir}"
        )
        return validate_batch(
            instances,
            taxonomy_dir=args.taxonomy_dir,
            output_dir=args.output_dir,
            online=args.online,
            render_html=not args.no_html_report,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
