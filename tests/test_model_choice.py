"""The default-model rule (core/model_choice.py): which model the sidebar preselects given the
configured name and the provider's model list. Fake lists only; nothing is fetched."""
import pytest

from core.model_choice import (CONFIGURED, NEWER, NO_LIST, REPLACED, UNLISTED, choose_default_model,
                               default_model_caption, newest_in_family, parse_anthropic_model,
                               parse_openai_model)

ANTHROPIC_LIST = [
    "claude-fable-5-1", "claude-opus-5-5", "claude-opus-5", "claude-opus-4-8", "claude-opus-4-7",
    "claude-sonnet-5-5", "claude-sonnet-4-6", "claude-haiku-5-5", "claude-haiku-4-5",
    "claude-opus-4-5-20251101", "claude-3-5-sonnet-20241022", "claude-3-opus-20240229",
]
OPENAI_LIST = [
    "gpt-6-astra", "gpt-6.1-sol", "gpt-6.1-terra", "gpt-6-luna", "gpt-5.6-luna", "gpt-5.6-sol",
    "gpt-5-nano", "gpt-5-nano-2025-08-07", "gpt-5.4-nano-2026-03-17", "gpt-4.1-nano", "gpt-4o-mini",
    "gpt-5", "gpt-5-chat-latest", "gpt-5.1-codex-mini", "o3", "gpt-image-2", "text-embedding-3-small",
    "ft:gpt-4.1-nano:org::abc123",
]


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("model_id, family, version, dated", [
    ("claude-opus-5-5", "opus", (5, 5), False),
    ("claude-opus-5", "opus", (5,), False),
    ("claude-opus-4-8", "opus", (4, 8), False),
    ("claude-opus-4-5-20251101", "opus", (4, 5), True),
    ("claude-3-5-sonnet-20241022", "sonnet", (3, 5), True),
    ("claude-3-5-sonnet-latest", "sonnet", (3, 5), False),
    ("claude-3-opus-20240229", "opus", (3,), True),
    ("claude-fable-5-1", "fable", (5, 1), False),
    ("claude-haiku-4-5", "haiku", (4, 5), False),
])
def test_anthropic_ids_split_into_family_version_and_date(model_id, family, version, dated):
    parsed = parse_anthropic_model(model_id)
    assert (parsed.family, parsed.version, parsed.dated) == (family, version, dated)


@pytest.mark.parametrize("model_id", ["claude-test-model", "claude", "gpt-6-luna", "", None, "claude-opus-5x"])
def test_anthropic_ids_outside_the_grammar_do_not_parse(model_id):
    assert parse_anthropic_model(model_id) is None


@pytest.mark.parametrize("model_id, family, version, dated", [
    ("gpt-6-luna", "luna", (6,), False),
    ("gpt-6.1-sol", "sol", (6, 1), False),
    ("gpt-5.6-terra", "terra", (5, 6), False),
    ("gpt-5-nano", "nano", (5,), False),
    ("gpt-5-nano-2025-08-07", "nano", (5,), True),
    ("gpt-5.4-nano-2026-03-17", "nano", (5, 4), True),
    ("gpt-4o-mini", "mini", (4,), False),
    ("gpt-5", "", (5,), False),
    ("gpt-5-chat-latest", "chat-latest", (5,), False),
    ("gpt-5.1-codex-mini", "codex-mini", (5, 1), False),
])
def test_openai_ids_split_into_family_version_and_date(model_id, family, version, dated):
    parsed = parse_openai_model(model_id)
    assert (parsed.family, parsed.version, parsed.dated) == (family, version, dated)


@pytest.mark.parametrize("model_id", ["o3", "gpt-image-2", "text-embedding-3-small", "ft:gpt-4.1-nano:org::abc123",
                                      "gpt-", "claude-opus-5-5", None])
def test_openai_ids_outside_the_grammar_do_not_parse(model_id):
    assert parse_openai_model(model_id) is None


# ---------------------------------------------------------------------------
# Newest of a family
# ---------------------------------------------------------------------------

def test_newest_opus_is_the_highest_version_not_fable():
    assert newest_in_family("anthropic", "opus", ANTHROPIC_LIST).model_id == "claude-opus-5-5"
    assert newest_in_family("anthropic", "sonnet", ANTHROPIC_LIST).model_id == "claude-sonnet-5-5"
    assert newest_in_family("anthropic", "fable", ANTHROPIC_LIST).model_id == "claude-fable-5-1"
    assert newest_in_family("anthropic", "mythos", ANTHROPIC_LIST) is None


def test_a_single_number_version_sorts_below_its_point_release_and_above_the_previous_major():
    assert newest_in_family("anthropic", "opus", ["claude-opus-4-8", "claude-opus-5"]).model_id == "claude-opus-5"
    assert newest_in_family("anthropic", "opus", ["claude-opus-5-5", "claude-opus-5"]).model_id == "claude-opus-5-5"
    assert newest_in_family("openai", "luna", ["gpt-5.6-luna", "gpt-6-luna"]).model_id == "gpt-6-luna"
    assert newest_in_family("openai", "luna", ["gpt-6.1-luna", "gpt-6-luna"]).model_id == "gpt-6.1-luna"


def test_an_undated_alias_beats_a_dated_snapshot_of_the_same_version():
    assert newest_in_family("anthropic", "opus", ["claude-opus-4-5-20251101", "claude-opus-4-5"]).model_id == "claude-opus-4-5"
    assert newest_in_family("openai", "nano", ["gpt-5-nano-2025-08-07", "gpt-5-nano"]).model_id == "gpt-5-nano"
    # A dated snapshot still counts when it is all there is, and a newer version beats an undated older one.
    assert newest_in_family("anthropic", "opus", ["claude-opus-4-5-20251101"]).model_id == "claude-opus-4-5-20251101"
    assert newest_in_family("openai", "nano", OPENAI_LIST).model_id == "gpt-5.4-nano-2026-03-17"


def test_ties_keep_the_first_listed_and_unparseable_ids_are_skipped():
    assert newest_in_family("anthropic", "opus", ["claude-opus-5-5", "claude-test-model", "claude-opus-5-5"]).model_id == "claude-opus-5-5"
    assert newest_in_family("openai", "luna", ["o3", None, 42, "gpt-6-luna"]).model_id == "gpt-6-luna"


# ---------------------------------------------------------------------------
# The rule
# ---------------------------------------------------------------------------

def test_without_a_list_the_configured_name_stands():
    for available in (None, [], ()):
        choice = choose_default_model("anthropic", "claude-opus-5-5", available)
        assert (choice.model, choice.status, choice.family) == ("claude-opus-5-5", NO_LIST, "opus")
    assert default_model_caption(choice, "anthropic") is None


def test_the_configured_name_is_kept_when_it_is_the_newest_of_its_family():
    choice = choose_default_model("anthropic", "claude-opus-5-5", ANTHROPIC_LIST)
    assert (choice.model, choice.status) == ("claude-opus-5-5", CONFIGURED)
    assert default_model_caption(choice, "anthropic") == "Default: the newest Opus model available to this key."
    choice = choose_default_model("openai", "gpt-6-luna", OPENAI_LIST)
    assert (choice.model, choice.status) == ("gpt-6-luna", CONFIGURED)
    assert default_model_caption(choice, "openai") == "Default: the newest Luna model available to this key."


def test_a_newer_model_of_the_family_replaces_a_listed_but_older_configured_name():
    choice = choose_default_model("anthropic", "claude-opus-4-8", ANTHROPIC_LIST)
    assert (choice.model, choice.configured, choice.status) == ("claude-opus-5-5", "claude-opus-4-8", NEWER)
    assert default_model_caption(choice, "anthropic") == (
        "Default updated: claude-opus-4-8 is configured, but claude-opus-5-5 is the newest Opus model "
        "available to this key.")
    choice = choose_default_model("openai", "gpt-5.6-luna", OPENAI_LIST)
    assert (choice.model, choice.status) == ("gpt-6-luna", NEWER)


def test_a_configured_name_that_the_key_cannot_use_is_replaced_by_the_newest_of_its_family():
    listed = [m for m in ANTHROPIC_LIST if m != "claude-opus-4-8"]
    choice = choose_default_model("anthropic", "claude-opus-4-8", listed)
    assert (choice.model, choice.status) == ("claude-opus-5-5", REPLACED)
    assert default_model_caption(choice, "anthropic") == (
        "Default updated: claude-opus-4-8 is no longer in Anthropic's model list; claude-opus-5-5 is the "
        "newest Opus model available to this key.")
    # OpenAI retired gpt-5-nano; a key that no longer lists it gets the newest nano, not a luna
    # (the family is what the configured name says, so moving to luna is a config.yaml edit).
    listed = [m for m in OPENAI_LIST if not m.startswith("gpt-5-nano")]
    choice = choose_default_model("openai", "gpt-5-nano", listed)
    assert (choice.model, choice.status) == ("gpt-5.4-nano-2026-03-17", REPLACED)
    assert "OpenAI's model list" in default_model_caption(choice, "openai")


def test_a_configured_name_with_no_family_in_the_list_is_kept_and_flagged():
    choice = choose_default_model("anthropic", "claude-opus-5-5", ["claude-test-model", "claude-sonnet-5-5"])
    assert (choice.model, choice.status, choice.family) == ("claude-opus-5-5", UNLISTED, "opus")
    assert default_model_caption(choice, "anthropic") == (
        "claude-opus-5-5 is not in Anthropic's model list and no other Opus model is; choose a listed "
        "model if a run fails.")
    # An unparseable configured name has no family to search for.
    choice = choose_default_model("anthropic", "my-finetune", ANTHROPIC_LIST)
    assert (choice.model, choice.status, choice.family) == ("my-finetune", UNLISTED, None)
    assert default_model_caption(choice, "anthropic") == (
        "my-finetune is not in Anthropic's model list; choose a listed model if a run fails.")


def test_an_unparseable_configured_name_that_is_listed_is_simply_used():
    choice = choose_default_model("anthropic", "claude-test-model", ["claude-test-model", "claude-opus-5-5"])
    assert (choice.model, choice.status, choice.family) == ("claude-test-model", CONFIGURED, None)
    assert default_model_caption(choice, "anthropic") == "Default: claude-test-model."


def test_switching_family_in_config_switches_the_rule():
    choice = choose_default_model("anthropic", "claude-sonnet-4-6", ANTHROPIC_LIST)
    assert (choice.model, choice.status) == ("claude-sonnet-5-5", NEWER)
    choice = choose_default_model("openai", "gpt-5.6-sol", OPENAI_LIST)
    assert (choice.model, choice.status) == ("gpt-6.1-sol", NEWER)


def test_fable_is_not_opus_class_for_the_rule():
    choice = choose_default_model("anthropic", "claude-opus-5-5", ["claude-fable-5-1", "claude-opus-5-5"])
    assert choice.model == "claude-opus-5-5"
    choice = choose_default_model("anthropic", "claude-fable-5-1", ["claude-fable-5-1", "claude-opus-5-5"])
    assert (choice.model, choice.status) == ("claude-fable-5-1", CONFIGURED)


def test_unknown_provider_is_an_error():
    with pytest.raises(ValueError):
        choose_default_model("gemini", "gemini-3", ["gemini-3"])
