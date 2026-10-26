"""Which model the sidebar preselects, chosen from the provider's model list.

`config.yaml` names one model per provider (`app_defaults.model_name_anthropic` and
`model_name_openai`). Those names go stale: providers release new models and retire old ones faster
than the file is edited. So when the provider's model list is available, the preselected default is
chosen from it by a preference rule, and the configured name only matters when the list cannot be
fetched or holds nothing comparable.

The rule: the newest model in the same family as the configured name.

- Anthropic ids read ``claude-<family>-<major>[-<minor>][-<yyyymmdd>]`` (older ones put the version
  first: ``claude-3-5-sonnet-20241022``). The family is the tier word: ``opus``, ``sonnet``,
  ``haiku``, ``fable``. With ``claude-opus-5-5`` configured, the default is the newest Opus model the
  key can use, and a future ``claude-opus-6`` is chosen the day it appears in the list.
- OpenAI ids read ``gpt-<major>[.<minor>]-<family>[-<yyyy-mm-dd>]``. The family is the tier word
  after the version: ``astra``, ``sol``, ``terra``, ``luna`` (formerly ``mini``, ``nano``). With
  ``gpt-6.1-terra`` configured, the default is the newest Terra model.

"Newest" means the highest version number, then an undated alias over a dated snapshot of the same
version. Ids that do not fit the grammar are never chosen as a default but stay in the list for the
user to pick. Switching family (say Sonnet instead of Opus) is done by naming a model of that family
in `config.yaml`; the rule then follows that family.

This module does no I/O and imports no Streamlit; the sidebar fetches the list and shows the result.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Iterable, Optional

# How the configured name relates to the provider's list; `default_model_caption` words each one.
NO_LIST = "no_list"          # the list could not be fetched: the configured name stands
CONFIGURED = "configured"    # the configured name is listed and nothing newer of its family is
NEWER = "newer"              # the configured name is listed, but a newer model of its family is too
REPLACED = "replaced"        # the configured name is not listed; the newest of its family stands in
UNLISTED = "unlisted"        # the configured name is not listed and neither is any model of its family

PROVIDER_LABELS = {"anthropic": "Anthropic", "openai": "OpenAI"}

_ANTHROPIC_DATE = re.compile(r"^\d{8}$")
_OPENAI_VERSION = re.compile(r"^(\d+)(?:\.(\d+))?[a-z]*$")
_OPENAI_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


@dataclass(frozen=True)
class ParsedModel:
    """A model id split into the parts the rule compares."""
    model_id: str
    family: str
    version: tuple
    dated: bool

    @property
    def rank(self):
        """Sorts newest last: the version, then an undated alias over a dated snapshot of it."""
        return (self.version, 0 if self.dated else 1)


@dataclass(frozen=True)
class DefaultModel:
    """The outcome of the rule: what to preselect, what was configured, and how they relate."""
    model: str
    configured: str
    status: str
    family: Optional[str]


def parse_anthropic_model(model_id) -> Optional[ParsedModel]:
    """``claude-opus-5-5`` -> Opus (5, 5); ``claude-3-5-sonnet-20241022`` -> Sonnet (3, 5), dated.
    ``latest`` aliases count as undated. Anything else returns None."""
    if not isinstance(model_id, str) or not model_id.startswith("claude-"):
        return None
    words, numbers, dated = [], [], False
    for token in model_id[len("claude-"):].split("-"):
        if _ANTHROPIC_DATE.match(token):
            dated = True
        elif token.isdigit():
            numbers.append(int(token))
        elif token == "latest":
            continue
        elif token.isalpha():
            words.append(token.lower())
        else:
            return None
    if not words or not numbers:
        return None
    return ParsedModel(model_id, "-".join(words), tuple(numbers), dated)


def parse_openai_model(model_id) -> Optional[ParsedModel]:
    """``gpt-6.1-sol`` -> sol (6, 1); ``gpt-5-nano-2025-08-07`` -> nano (5,), dated; ``gpt-5`` -> the
    empty family. Ids that do not start with ``gpt-<version>`` (``o3``, images, fine-tunes) return None."""
    if not isinstance(model_id, str) or not model_id.startswith("gpt-"):
        return None
    tokens = model_id[len("gpt-"):].split("-")
    match = _OPENAI_VERSION.match(tokens[0]) if tokens and tokens[0] else None
    if not match:
        return None
    version = (int(match.group(1)),) + ((int(match.group(2)),) if match.group(2) else ())
    rest, dated = tokens[1:], False
    if len(rest) >= 3 and _OPENAI_DATE.match("-".join(rest[-3:])):
        rest, dated = rest[:-3], True
    if any(not token.isalpha() for token in rest):
        return None
    return ParsedModel(model_id, "-".join(token.lower() for token in rest), version, dated)


PARSERS: dict[str, Callable[[object], Optional[ParsedModel]]] = {
    "anthropic": parse_anthropic_model,
    "openai": parse_openai_model,
}


def newest_in_family(provider: str, family: str, available: Iterable) -> Optional[ParsedModel]:
    """The newest listed model of `family`, or None. Ties keep the first listed."""
    parser = PARSERS[provider]
    best = None
    for model_id in available:
        parsed = parser(model_id)
        if parsed and parsed.family == family and (best is None or parsed.rank > best.rank):
            best = parsed
    return best


def choose_default_model(provider: str, configured: str, available) -> DefaultModel:
    """Apply the rule. `available` is the provider's model list, or None/empty when it could not be
    fetched; the configured name is then the default."""
    if provider not in PARSERS:
        raise ValueError(f"Unknown provider: {provider!r}")
    listed = [model_id for model_id in (available or []) if isinstance(model_id, str)]
    wanted = PARSERS[provider](configured)
    family = wanted.family if wanted else None
    if not listed:
        return DefaultModel(configured, configured, NO_LIST, family)
    in_list = configured in listed
    best = newest_in_family(provider, wanted.family, listed) if wanted else None
    if best is None:
        return DefaultModel(configured, configured, CONFIGURED if in_list else UNLISTED, family)
    if not in_list:
        return DefaultModel(best.model_id, configured, REPLACED, family)
    if best.model_id == configured or best.rank <= wanted.rank:
        return DefaultModel(configured, configured, CONFIGURED, family)
    return DefaultModel(best.model_id, configured, NEWER, family)


def default_model_caption(choice: DefaultModel, provider: str) -> Optional[str]:
    """The sidebar's one-line note under the model picker, or None when there is nothing to say:
    the newest model of the family is simply preselected, and a list that could not be fetched is
    already warned about. A note appears only when the configured name is not in the list."""
    label = PROVIDER_LABELS.get(provider, provider)
    family = choice.family.capitalize() if choice.family else None
    if choice.status in (NO_LIST, CONFIGURED, NEWER):
        return None
    if choice.status == REPLACED:
        return f"{choice.configured} is no longer in {label}'s model list; selected {choice.model}."
    if family:
        return (f"{choice.configured} is not in {label}'s model list and no other {family} model is; "
                f"choose a listed model if a run fails.")
    return f"{choice.configured} is not in {label}'s model list; choose a listed model if a run fails."
