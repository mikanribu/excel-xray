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
* :class:`OpenAIAssessor` — opt-in. Same evidence bundle, sent to either the
  public OpenAI API or an Azure OpenAI deployment (basis ``inferred``).

Only structural metadata, headers and normalised formula shapes leave the
machine when a model-backed assessor is used — never cell values.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Protocol


PROCESS_SUBPROCESS_PROMPT = """Analyse the EUC and assign an appropriate Process and Sub-process based on the business use case supported by the file.

**Process**
Identify the end-to-end business process that the EUC supports.
The Process should represent a stable business capability or outcome rather than a specific report, workbook, system, team or individual activity.

**Sub-process**
Identify the specific functional use case performed by the EUC within the broader Process.
The Sub-process should be sufficiently specific to group EUCs that perform materially similar activities, but not so granular that each workbook receives a unique classification.
When determining the Sub-process, consider:
• Primary business objective of the EUC
• Inputs and source data used
• Key calculations, transformations or mappings performed
• Reconciliations or controls performed
• Output produced
• How the output is used by business users
• Whether other EUCs perform a similar sequence of activities
• Whether the EUCs could potentially share a common future automation solution

Follow a two-stage classification for identification:
Stage 1: Identifies the underlying business use case based on purpose, inputs, calculations, controls and outputs.

Stage 2: Maps that use case against the predefined Process / Sub-process taxonomy. If nothing considered applicable, only proposing a new category where no reasonable existing classification exists."""


RATIONALIZATION_PROMPT = """Using the assigned Process and Sub-process, compare EUCs with similar business use cases to identify potential duplication and rationalisation opportunities.

Assess each EUC against the following:

1. Potential Duplication

Identify EUCs that perform materially similar functions based on:

•          Business purpose and output

•          Input data and source systems

•          Calculation, mapping and transformation logic

•          Reconciliations and controls

•          Downstream usage

Do not classify EUCs as duplicates solely because they share the same Process or Sub-process. Explain the specific functional similarities and material differences.

2. Rationalisation Opportunity

Assess the potential for:

•          Simplification: Retain the EUC but remove unnecessary steps, calculations, manual interventions or controls.

•          Consolidation: Combine multiple similar EUCs into a common process, workbook, dataset or solution.

•          Automation: Replace repeatable manual activities with automated data sourcing, calculations, controls, reconciliation or reporting.

•          Retirement: Remove the EUC where its purpose is duplicated, no longer required, or its output can be fully provided by another process or solution.

Assessment principles

•          Consider the end-to-end business use case, not just workbook structure or file name.

•          Identify opportunities only where there is sufficient evidence from the EUC analysis.

•          Highlight material differences that may prevent consolidation or automation, such as different business rules, data sources, regulatory requirements, manual judgement or entity-specific requirements."""


RECONCILIATION_PROMPT = """Review the EUC and identify all reconciliations that are actually performed in the workbook. Provide me the number of reconciliations conducted.

Only classify an activity as a reconciliation where two or more balances, datasets, reports, systems or periods are compared to confirm agreement or identify differences. Do not treat standalone calculations as reconciliations.

For each reconciliation identified, provide:

1. Overview: What is being reconciled against what.

2. Source A: First data source, report, system or worksheet.

3. Source B: Comparison data source, report, system or worksheet.

4. Matching criteria: How the two sources are matched or compared. What is the tolerance level (if any)

5. Exception logic: Any documented rules, explanations, supporting comments, or evidence within the EUC that help explain, classify, or justify identified differences or unreconciled items."""


@dataclass
class Narrative:
    """The narrative fields an assessor fills."""

    purpose_of_file: str | None = None
    key_output_outcome: str | None = None
    key_outputs: list | None = None
    business_use_case: str | None = None
    process: str | None = None
    sub_process: str | None = None
    process_evidence: list[str] = field(default_factory=list)
    sub_process_evidence: list[str] = field(default_factory=list)
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
    # Keep the filename as context, but never accept it as evidence for a
    # Process/Sub-process label. Classifications must be supported by workbook
    # content such as worksheet names and headers.
    business_evidence = []
    for ta in assessment.tabs:
        name = ta.tab_name.value
        s = sheet_by_name.get(name)
        headers = []
        if s:
            for r in s.regions:
                headers += [h for h in r.headers if h]
        business_evidence.extend([name] + headers)
        kc = ta.key_calculation_transformation_logic.value
        tabs.append({
            "name": name,
            "visibility": ta.tab_visibility.value,
            "category": ta.tab_category.value,
            "roles": ta.tab_roles.value,
            "information": ta.tab_information_analysis.value,
            "headers": list(dict.fromkeys(headers))[:100],
            "business_description": (kc or {}).get("business_description") if isinstance(kc, dict) else None,
            "upstream": ta.upstream_dependencies.value,
            "downstream": ta.downstream_dependencies.value.get("in_workbook")
            if isinstance(ta.downstream_dependencies.value, dict) else None,
            "hidden": (s.state != "visible") if s else False,
        })

    return {
        "file_name": val(fa.file_name),
        # This repository has Logic Types and tab-category taxonomies, but no
        # separate business Process/Sub-process catalogue.
        "process_subprocess_taxonomy": [],
        "business_area_process": val(fa.business_area_process),
        "process": val(fa.process),
        "sub_process": val(fa.sub_process),
        "complexity": val(fa.complexity),
        "logic_type": val(fa.logic_type),
        "logic_types": val(fa.logic_types),
        "key_calculations": val(fa.key_calculations_logic),
        "key_inputs": val(fa.key_inputs),
        "reconciliation_candidates": val(fa.reconciliation_logic),
        "sheet_count": len(wx.sheets),
        "business_evidence": list(dict.fromkeys(x for x in business_evidence if x)),
        "tabs": tabs,
    }


# ------------------------------------------------------------- offline assessor


class OfflineAssessor:
    """Deterministic, network-free draft from the deterministic findings."""

    basis = "drafted"
    label = "offline template"

    def narrate(self, bundle: dict) -> Narrative:
        logic = bundle.get("logic_type") or "Other"
        cats = [t["category"] for t in bundle["tabs"]]
        cat_counts = ", ".join(sorted({c for c in cats})) or "no classified tabs"
        purpose = "Not established — confirm the workbook's business purpose with the process owner."
        outcome = "Not established — confirm the final business outcome and recipient with the process owner."
        tab_purposes = {}
        for t in bundle["tabs"]:
            hdr = ", ".join(t["headers"][:4])
            tab_purposes[t["name"]] = (
                f"{t['category']} tab ({t['information']})."
                + (f" Columns: {hdr}." if hdr else "")
                + (" Hidden sheet." if t["hidden"] else "")
            )

        return Narrative(
            purpose_of_file=purpose,
            key_output_outcome=outcome,
            key_outputs=[],
            process="Not established — confirm with the process owner",
            sub_process="Not established — confirm with the process owner",
            tabs=tab_purposes,
        )


# -------------------------------------------------------------- claude assessor


_SYSTEM = (
    "You review End User spreadsheets for a financial-controls "
    "team. You are given a value-free structural summary of one workbook "
    "(sheet layout, tab categories, column headers, normalised formula shapes — "
    "never cell values). Write concise, factual assessment prose. Do not invent "
    "figures, systems, owners or frequencies that are not implied by the "
    "structure. Reply with a single JSON object and nothing else."
)

_INSTRUCTION = (
    PROCESS_SUBPROCESS_PROMPT
    + "\n\nThe codebase currently has no separate business Process/Sub-process catalogue; its Logic Types and tab categories are different taxonomies and must not be reused as business processes. Reuse a stable broad classification for materially similar activities and do not create a workbook-specific category. Do not use the workbook filename as primary evidence.\n"
    "Return JSON with exactly these keys:\n"
    '  "business_use_case": string — concise Stage 1 identification from purpose, inputs, logic, controls and outputs.\n'
    '  "purpose_of_file": string — 1-2 sentences on what the workbook is for.\n'
    '  "key_output_outcome": string — the business outcome it supports.\n'
    '  "key_outputs": array of strings — final business deliverables only; exclude intermediate workings.\n'
    '  "process": string — business process, or "Not established — confirm with the process owner".\n'
    '  "sub_process": string — narrower activity, or the same not-established value.\n'
    '  "process_evidence": array of exact text snippets copied from the evidence below.\n'
    '  "sub_process_evidence": array of exact text snippets copied from the evidence below.\n'
    '  "tabs": object mapping each tab name to a one-sentence purpose.\n'
    "Only classify process/sub-process when specific business terms in the workbook name, "
    "worksheet names, or headers support it; generic terms such as Input, Output, Calculation, "
    "Model, Data, and Summary are not sufficient. "
    "Return exact source snippets in process_evidence and sub_process_evidence. "
    "Use no formula counts, function names, or formula skeletons as business conclusions. "
    "List Key Outputs only when a final business deliverable is evidenced; otherwise return [].\n"
    "Evidence:\n"
)

_RECONCILIATION_RESPONSE = (
    "Return JSON with exactly these keys: reconciliation_count (integer) and "
    "reconciliations (array). Each array item must contain worksheet, overview, "
    "source_a, source_b, matching_criteria, tolerance, exception_logic and "
    "evidence. Each evidence item must be an object with worksheet and text, "
    "where text is copied exactly from a supplied worksheet name or header. "
    "Only state matching criteria, tolerance or exception logic when directly "
    "supported by the supplied evidence; otherwise say it is not established. "
    "Use an empty array and count 0 when no genuine comparison is evidenced. "
    "Do not count standalone calculations, mappings or lookups."
)

_RATIONALIZATION_RESPONSE = (
    "Return one JSON object with a results array containing one item for each "
    "current EUC. Each item must contain comparison_id (copied from current) "
    "and these fields: "
    "potential_duplication (object with verdict and matches), "
    "similar_duplicate_files (array of comparison_id, functional_similarities, "
    "material_differences and evidence), "
    "potential_simplification (object with verdict and candidates), "
    "potential_consolidation (object with verdict and candidates), "
    "potential_automation (object with verdict and candidates), and "
    "potential_retirement (object with verdict, retirement_basis, "
    "replacement_comparison_ids, replacement_coverage and evidence). "
    "Every match or recommendation must contain an evidence array. Each evidence "
    "item must be an object with comparison_id, file_id, field and text copied "
    "exactly from that EUC's supplied purpose, output, input, calculation, control "
    "or worksheet evidence. "
    "A duplication match must explain functional similarities and material "
    "differences and cite evidence from both EUCs. Recommend retirement only "
    "where evidence supports a duplicated purpose, an explicit statement that it "
    "is no longer required, or full replacement of its required output. "
    "Do not report a positive finding or verdict unless it has a retained, "
    "evidence-backed candidate. For automation, cite explicit evidence of a "
    "repeatable manual step such as manual entry, paste, override or hardcoding."
)


def _stage_message(prompt: str, response_contract: str, payload: dict) -> str:
    """Keep approved prompt wording intact, then add the response contract/data."""
    return (prompt + "\n\n" + response_contract + "\n\nEvidence JSON:\n"
            + json.dumps(payload, ensure_ascii=False, default=str))


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
            business_use_case=data.get("business_use_case"),
            process=data.get("process"),
            sub_process=data.get("sub_process"),
            process_evidence=data.get("process_evidence") or [],
            sub_process_evidence=data.get("sub_process_evidence") or [],
            tabs=data.get("tabs") or {},
        )

    def reconcile(self, bundle: dict) -> dict:
        """Run the workbook reconciliation prompt as its own analysis stage."""
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
            messages=[{"role": "user", "content": _stage_message(
                RECONCILIATION_PROMPT, _RECONCILIATION_RESPONSE, bundle)}],
        )
        return _parse_json("".join(b.text for b in msg.content if b.type == "text"))

    def rationalize(self, payload: dict) -> dict:
        """Compare the corpus shortlist and assess evidence-backed opportunities."""
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
            messages=[{"role": "user", "content": _stage_message(
                RATIONALIZATION_PROMPT, _RATIONALIZATION_RESPONSE, payload)}],
        )
        return _parse_json("".join(b.text for b in msg.content if b.type == "text"))


# --------------------------------------------------------- openai / azure assessor


def _openai_complete(model: str, max_tokens: int, system: str, user: str, *,
                      azure_endpoint: str | None, api_version: str | None) -> dict:
    """Call an OpenAI-compatible chat endpoint and return the parsed JSON reply.

    Shared by :class:`OpenAIAssessor` and estate_insight's ``OpenAIEstateAssessor``
    so the public-API/Azure client selection lives in one place. ``model`` is
    the model id for the public API, or the deployment name for an Azure
    endpoint (Azure addresses models by the name the resource deployed them
    under, not the base model id).
    """
    try:
        from openai import AzureOpenAI, OpenAI
    except ImportError as e:
        raise RuntimeError(
            "the OpenAI assessor needs the 'openai' package — "
            "install it with: uv add --optional llm openai"
        ) from e

    if azure_endpoint:
        # api key picked up from AZURE_OPENAI_API_KEY by the client itself.
        client = AzureOpenAI(azure_endpoint=azure_endpoint,
                              api_version=api_version or "2024-10-21")
    else:
        client = OpenAI()  # picks up OPENAI_API_KEY

    resp = client.chat.completions.create(
        model=model,
        max_tokens=max_tokens,
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": user}],
    )
    text = resp.choices[0].message.content or ""
    return _parse_json(text)


class OpenAIAssessor:
    """Opt-in GPT-backed assessor. Talks to the public OpenAI API by default;
    pass `azure_endpoint` (or set $AZURE_OPENAI_ENDPOINT) to talk to an Azure
    OpenAI deployment instead — `model` then means that deployment's name.
    Requires the `openai` package and a credential:
      - public API: OPENAI_API_KEY
      - Azure:      AZURE_OPENAI_API_KEY (+ the endpoint above)."""

    basis = "inferred"

    def __init__(self, model: str = "gpt-4o", max_tokens: int = 2000,
                 azure_endpoint: str | None = None, api_version: str | None = None):
        self.model = model
        self.max_tokens = max_tokens
        self.azure_endpoint = azure_endpoint or os.environ.get("AZURE_OPENAI_ENDPOINT")
        self.api_version = api_version or os.environ.get("AZURE_OPENAI_API_VERSION")
        self.label = (f"Azure OpenAI ({model})" if self.azure_endpoint
                      else f"OpenAI ({model})")

    def narrate(self, bundle: dict) -> Narrative:
        data = _openai_complete(
            self.model, self.max_tokens, _SYSTEM,
            _INSTRUCTION + json.dumps(bundle, default=str),
            azure_endpoint=self.azure_endpoint, api_version=self.api_version,
        )
        return Narrative(
            purpose_of_file=data.get("purpose_of_file"),
            key_output_outcome=data.get("key_output_outcome"),
            key_outputs=data.get("key_outputs"),
            business_use_case=data.get("business_use_case"),
            process=data.get("process"),
            sub_process=data.get("sub_process"),
            process_evidence=data.get("process_evidence") or [],
            sub_process_evidence=data.get("sub_process_evidence") or [],
            tabs=data.get("tabs") or {},
        )

    def reconcile(self, bundle: dict) -> dict:
        """Run the workbook reconciliation prompt as its own analysis stage."""
        return _openai_complete(
            self.model, self.max_tokens, _SYSTEM,
            _stage_message(RECONCILIATION_PROMPT, _RECONCILIATION_RESPONSE, bundle),
            azure_endpoint=self.azure_endpoint, api_version=self.api_version,
        )

    def rationalize(self, payload: dict) -> dict:
        """Compare the corpus shortlist and assess evidence-backed opportunities."""
        return _openai_complete(
            self.model, self.max_tokens, _SYSTEM,
            _stage_message(RATIONALIZATION_PROMPT, _RATIONALIZATION_RESPONSE, payload),
            azure_endpoint=self.azure_endpoint, api_version=self.api_version,
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
