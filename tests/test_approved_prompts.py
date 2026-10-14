"""Keep the business-approved prompt wording exact and attached to its stage."""

from __future__ import annotations

import hashlib

from excel_xray import narrative


def test_approved_prompt_text_has_not_changed():
    expected = {
        narrative.PROCESS_SUBPROCESS_PROMPT:
            "89facbab7a8270c36780643700ccbd9a44440b172ac947aaaa6445aa8345fa22",
        narrative.RATIONALIZATION_PROMPT:
            "df7c84a5db1ffd49eecb3437e6d1f750817e9e4939c7c31cca1a223266c12203",
        narrative.RECONCILIATION_PROMPT:
            "8999fe2d4452d109e7f9ce5aed89b4ede3db0de46ce80d8d383a75e0b2f0f37a",
    }
    assert all(hashlib.sha256(text.encode()).hexdigest() == digest
               for text, digest in expected.items())


def test_each_approved_prompt_starts_its_own_analysis_stage(monkeypatch):
    calls = []

    def fake_complete(model, max_tokens, system, user, **kwargs):
        calls.append(user)
        return {}

    monkeypatch.setattr(narrative, "_openai_complete", fake_complete)
    assessor = narrative.OpenAIAssessor()
    assessor.narrate({})
    assessor.reconcile({})
    assessor.rationalize({"current": {}, "candidates": []})

    assert calls[0].startswith(narrative.PROCESS_SUBPROCESS_PROMPT)
    assert calls[1].startswith(narrative.RECONCILIATION_PROMPT + "\n\n")
    assert calls[2].startswith(narrative.RATIONALIZATION_PROMPT + "\n\n")
