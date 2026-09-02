"""Narrative layer (Step 4): the fields that need a model, behind an interface.

The deterministic layers (Steps 1-3) fill everything groundable from the file.
The remaining EUC fields are narrative judgement — *Purpose of File*, *Key
Output / Outcome*, *Key Outputs*, and each tab's *Purpose / Description* — so
they sit behind a small :class:`Assessor` interface with two implementations:

* :class:`OfflineAssessor` — default, no network. Writes an evidence-templated
  draft from the deterministic findings (basis ``drafted``). Keeps the pipeline
  and the tests hermetic.
* :class:`ClaudeAssessor` — opt-in. Sends the *structural* evidence bundle (no
  cell values — same privacy stance as the report) to the Anthropic API and
  returns a considered narrative (basis ``inferred``).

Only structural metadata, headers and normalised formula shapes leave the
machine when the Claude assessor is used — never cell values.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Protocol


@dataclass
class Narrative:
    """The narrative fields an assessor fills."""

    purpose_of_file: str | None = None
    key_output_outcome: str | None = None
    key_outputs: list | None = None
    tabs: dict[str, str] = field(default_factory=dict)  # tab name -> purpose


class Assessor(Protocol):
    """Anything that can turn an evidence bundle into a :class:`Narrative`."""

    basis: str  # "drafted" | "inferred"
    label: str

    def narrate(self, bundle: dict) -> Narrative: ...


# ------------------------------------------------------------- evidence bundle


def build_bundle(assessment, wx) -> dict:
    """Compact, value-free evidence for the narrative step.

    Carries structure, categories, headers and formula shapes — never cell
    values — so it is safe to send to an external model.
    """
    fa = assessment.file

    def val(fld):
        return fld.value

    tabs = []
    sheet_by_name = {s.name: s for s in wx.sheets}
    for ta in assessment.tabs:
        name = ta.tab_name.value
        s = sheet_by_name.get(name)
        headers = []
        if s:
            for r in s.regions:
                headers += [h for h in r.headers if h]
        kc = ta.key_calculation_transformation_logic.value
        tabs.append({
            "name": name,
            "category": ta.tab_category.value,
            "information": ta.tab_information_analysis.value,
            "headers": headers[:20],
            "top_functions": (kc or {}).get("top_functions", []) if isinstance(kc, dict) else [],
            "upstream": ta.upstream_dependencies.value,
            "downstream": ta.downstream_dependencies.value.get("in_workbook")
            if isinstance(ta.downstream_dependencies.value, dict) else None,
            "hidden": (s.state != "visible") if s else False,
        })

    return {
        "file_name": val(fa.file_name),
        "business_area_process": val(fa.business_area_process),
        "complexity": val(fa.complexity),
        "logic_type": val(fa.logic_type),
        "key_calculations": val(fa.key_calculations_logic),
        "key_inputs": val(fa.key_inputs),
        "sheet_count": len(wx.sheets),
        "tabs": tabs,
    }


# ------------------------------------------------------------- offline assessor


def _stem(name: str) -> str:
    return name.rsplit(".", 1)[0] if "." in name else name


_PURPOSE_VERB = {
    "Calculation": "perform calculations / modelling",
    "Reconciliation": "reconcile or match figures across sources",
    "Data Transformation": "transform and reshape source data",
    "Reporting": "assemble reporting / MI output",
    "Manual Input": "capture data entered by hand",
    "Other": "support a business process",
}


class OfflineAssessor:
    """Deterministic, network-free draft from the deterministic findings."""

    basis = "drafted"
    label = "offline template"

    def narrate(self, bundle: dict) -> Narrative:
        logic = bundle.get("logic_type") or "Other"
        cats = [t["category"] for t in bundle["tabs"]]
        cat_counts = ", ".join(sorted({c for c in cats})) or "no classified tabs"
        outputs = [t["name"] for t in bundle["tabs"] if t["category"] == "Output"]
        top_fns = (bundle.get("key_calculations") or {}).get("top_functions", []) \
            if isinstance(bundle.get("key_calculations"), dict) else []

        purpose = (
            f"'{_stem(bundle['file_name'])}' is a {bundle.get('complexity','').lower()}"
            f"-complexity {logic.lower()} workbook across {bundle['sheet_count']} tab(s) "
            f"({cat_counts}). It appears to {_PURPOSE_VERB.get(logic, _PURPOSE_VERB['Other'])}"
            + (f", using {', '.join(top_fns[:4])}" if top_fns else "")
            + "."
        )
        inputs = bundle.get("key_inputs")
        if inputs:
            purpose += f" Inputs: {', '.join(map(str, inputs[:3]))}."

        outcome = (
            f"Supports {(bundle.get('business_area_process') or logic).lower()}; "
            + (f"headline output on tab(s) {', '.join(outputs)}."
               if outputs else "outputs are produced within the calculation tabs.")
        )

        key_outputs = []
        for t in bundle["tabs"]:
            if t["category"] == "Output" or (not outputs and t["information"] == "output"):
                label = t["name"]
                if t["headers"]:
                    label += f" ({', '.join(t['headers'][:5])})"
                key_outputs.append(label)

        tab_purposes = {}
        for t in bundle["tabs"]:
            fns = ", ".join(t["top_functions"][:3])
            hdr = ", ".join(t["headers"][:4])
            tab_purposes[t["name"]] = (
                f"{t['category']} tab ({t['information']})."
                + (f" Columns: {hdr}." if hdr else "")
                + (f" Key logic: {fns}." if fns else "")
                + (" Hidden sheet." if t["hidden"] else "")
            )

        return Narrative(
            purpose_of_file=purpose,
            key_output_outcome=outcome,
            key_outputs=key_outputs or ["no distinct output tab identified"],
            tabs=tab_purposes,
        )


# -------------------------------------------------------------- claude assessor


_SYSTEM = (
    "You review End-User Computing (EUC) spreadsheets for a financial-controls "
    "team. You are given a value-free structural summary of one workbook "
    "(sheet layout, tab categories, column headers, normalised formula shapes — "
    "never cell values). Write concise, factual assessment prose. Do not invent "
    "figures, systems, owners or frequencies that are not implied by the "
    "structure. Reply with a single JSON object and nothing else."
)

_INSTRUCTION = (
    "Return JSON with exactly these keys:\n"
    '  "purpose_of_file": string — 1-2 sentences on what the workbook is for.\n'
    '  "key_output_outcome": string — the business outcome it supports.\n'
    '  "key_outputs": array of strings — the main outputs produced.\n'
    '  "tabs": object mapping each tab name to a one-sentence purpose.\n'
    "Evidence:\n"
)


class ClaudeAssessor:
    """Opt-in Anthropic-backed assessor. Requires the `anthropic` package and a
    credential (ANTHROPIC_API_KEY or an `ant auth login` profile)."""

    basis = "inferred"

    def __init__(self, model: str = "claude-opus-5", max_tokens: int = 2000):
        self.model = model
        self.max_tokens = max_tokens
        self.label = f"Claude ({model})"

    def narrate(self, bundle: dict) -> Narrative:
        try:
            import anthropic
        except ImportError as e:
            raise RuntimeError(
                "the Claude assessor needs the 'anthropic' package — "
                "install it with: uv add --optional llm anthropic"
            ) from e

        client = anthropic.Anthropic()
        msg = client.messages.create(
            model=self.model,
            max_tokens=self.max_tokens,
            system=_SYSTEM,
            messages=[{"role": "user",
                       "content": _INSTRUCTION + json.dumps(bundle, default=str)}],
        )
        text = "".join(b.text for b in msg.content if b.type == "text")
        data = _parse_json(text)
        return Narrative(
            purpose_of_file=data.get("purpose_of_file"),
            key_output_outcome=data.get("key_output_outcome"),
            key_outputs=data.get("key_outputs"),
            tabs=data.get("tabs") or {},
        )


def _parse_json(text: str) -> dict:
    """Tolerant JSON extraction from a model reply (handles ```json fences)."""
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```", 2)[1]
        if text.startswith("json"):
            text = text[4:]
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise RuntimeError(f"model did not return JSON: {text[:200]!r}")
    return json.loads(text[start:end + 1])
