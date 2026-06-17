import logging
import os
from datetime import datetime
import re
from io import BytesIO
import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from streamlit_tree_select import tree_select
import yaml
import hmac
import time
from audit_reformat import handle_audit_reformat
from audit_validation import validate_audit_sentences_sheet
from audit import detect_partial_audit, _is_retryable_llm_error, CHECKPOINT_INTERVAL
import audit_worker

logger = logging.getLogger(__name__)


def _smtp_config():
    """Pull SMTP settings from Streamlit secrets (all optional)."""
    return {
        "host": st.secrets.get("SMTP_HOST"),
        "port": st.secrets.get("SMTP_PORT", 587),
        "user": st.secrets.get("SMTP_USER"),
        "password": st.secrets.get("SMTP_PASSWORD"),
        "sender": st.secrets.get("EMAIL_FROM"),
    }


def _email_configured():
    """True when the secrets needed to actually send mail are all present."""
    cfg = _smtp_config()
    return all(cfg.get(k) for k in ("host", "user", "password", "sender"))


# Email delivery now runs inside the background worker (audit_worker._send_email)
# so the workbook is sent even when no browser is attached at finish; the UI only
# surfaces the recorded email_status. See _finalize_job_into_session below.

# Main script for streamlit app that uses LLMs to conduct audits of Clarabridge topic models

# Configure Streamlit page (must be the first Streamlit call)
st.set_page_config(page_title="Automatic Audit", initial_sidebar_state="collapsed")

def _fetch_anthropic_models(api_key):
    try:
        import anthropic
    except Exception as exc:
        return None, exc
    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.models.list()
        data = getattr(response, "data", response)
        models = []
        for item in data:
            model_id = getattr(item, "id", None)
            if model_id is None and isinstance(item, dict):
                model_id = item.get("id")
            if model_id:
                models.append(model_id)
        return models, None
    except Exception as exc:
        return None, exc


def _fetch_openai_models(api_key):
    try:
        from openai import OpenAI
    except Exception as exc:
        return None, exc
    try:
        client = OpenAI(api_key=api_key)
        response = client.models.list()
        data = getattr(response, "data", [])
        models = [item.id for item in data if getattr(item, "id", None)]
        return models, None
    except Exception as exc:
        return None, exc


def _get_model_options(llm_provider, api_key, default_model):
    if not api_key:
        return [default_model], False, "No API key found."

    if llm_provider == "anthropic":
        models, err = _fetch_anthropic_models(api_key)
    else:
        models, err = _fetch_openai_models(api_key)

    if err or not models:
        err_msg = str(err) if err else "No models returned from provider."
        return [default_model], False, err_msg

    unique_models = []
    seen = set()
    for model in models:
        if model not in seen:
            seen.add(model)
            unique_models.append(model)

    if default_model not in seen:
        unique_models.insert(0, default_model)

    return unique_models, True, None


def _validate_api_key(llm_provider, api_key):
    if not api_key:
        return False, "No API key provided."
    if llm_provider == "anthropic":
        _, err = _fetch_anthropic_models(api_key)
    else:
        _, err = _fetch_openai_models(api_key)
    if err:
        return False, str(err)
    return True, None

def get_org_and_audience():
    organization = st.secrets.get("ORGANIZATION",  "the organization")
    audience = st.secrets.get("AUDIENCE", "customers and users")

    return organization, audience

def check_password():
    """Returns True if the user entered the correct password"""

    # Already authenticated this session
    if st.session_state.get("authenticated"):
        return True

    # Track failed attempts
    if "failed_attempts" not in st.session_state:
        st.session_state["failed_attempts"] = 0

    password = st.text_input("Enter password to access this app", type="password")

    if password:
        # Compare securely using constant-time comparison
        if hmac.compare_digest(password, st.secrets["APP_PASSWORD"]):
            st.session_state["authenticated"] = True
            st.session_state["failed_attempts"] = 0
            st.rerun()  # Clears the password field from view
        else:
            st.session_state["failed_attempts"] += 1

            # Exponential backoff: 2s, 4s, 8s, 16s... capped at 60s to protect form brute-force attacks
            delay = min(2 ** st.session_state["failed_attempts"], 60)
            time.sleep(delay)

            st.error(f"Incorrect password. Please wait {delay} seconds before trying again.")
    return False

if not check_password():
    st.stop()

def _load_summary_prompt(prompts_path):
    try:
        with open(prompts_path, "r") as f:
            prompts = yaml.safe_load(f)
        return prompts.get("summary_prompt", "")
    except Exception:
        return ""


def _load_audit_defaults(prompts_path):
    try:
        with open(prompts_path, "r") as f:
            prompts = yaml.safe_load(f)
        return {
            "audit_prompt": prompts.get("audit_prompt", ""),
            "model_info": prompts.get("model_info", ""),
        }
    except Exception:
        return {"audit_prompt": "", "model_info": ""}


def _load_app_defaults(config_path):
    defaults = {
        "llm_provider": "anthropic",
        "model_name_anthropic": "claude-opus-4-5",
        "model_name_openai": "gpt-5-nano",
        "max_categories": 1000,
        "max_sentences_per_category": 1001,
        "max_tokens": 10000,
    }
    try:
        with open(config_path, "r") as f:
            config = yaml.safe_load(f)
        app_defaults = config.get("app_defaults", {})
        defaults.update({k: v for k, v in app_defaults.items() if v is not None})
    except Exception:
        pass
    return defaults


def _normalize_topic(value):
    text = str(value).strip()
    return " ".join(text.split())


def _topic_key(value):
    text = _normalize_topic(value)
    parts = [part.strip() for part in re.split(r"\s*-->\s*", text) if part.strip()]
    normalized = "-->".join(parts) if parts else text
    return normalized.casefold()


def _get_audit_stats(audit_bytes):
    sentences_sheet, header_row_idx, _, _, is_output_format = validate_audit_sentences_sheet(audit_bytes)
    df = pd.read_excel(
        BytesIO(audit_bytes),
        sheet_name=sentences_sheet,
        header=header_row_idx,
    )
    if df.empty or len(df.columns) < 3:
        return {
            "category_counts": {},
            "total_categories": 0,
            "max_sentences_per_category": 0,
        }

    # Determine column layout based on format
    if is_output_format:
        # Check if new format with ID column or old format without
        first_col_name = str(df.columns[0]).strip().casefold() if len(df.columns) > 0 else ""
        if first_col_name == "id":
            # New output format: ID, Sentence, Topic, Audit, Explanation
            id_col = df.columns[0]
            sentence_col = df.columns[1]
            category_col = df.columns[2]
        else:
            # Old output format: Sentence, Topic, Audit, Explanation
            id_col = None
            sentence_col = df.columns[0]
            category_col = df.columns[1]
    else:
        # Input format: #, Sentences, Category, ...
        id_col = df.columns[0]
        sentence_col = df.columns[1]
        category_col = df.columns[2]

    category_sentences = {}
    row_idx = 0
    for _, row in df.iterrows():
        row_idx += 1
        category = row[category_col]
        sentence = row[sentence_col]
        if id_col is not None:
            sentence_id = row[id_col]
        else:
            sentence_id = row_idx  # Synthetic ID for old output format

        if pd.isna(category) or pd.isna(sentence):
            continue
        if id_col is not None and pd.isna(sentence_id):
            continue
        category_name = str(category).strip()
        if not category_name:
            continue
        category_sentences.setdefault(category_name, set()).add(sentence_id)

    category_counts = {cat: len(ids) for cat, ids in category_sentences.items()}
    top_level_categories = set()
    for category in category_counts:
        parts = [part.strip() for part in str(category).split("-->") if part.strip()]
        if parts:
            top_level_categories.add(_topic_key(parts[0]))
    return {
        "category_counts": category_counts,
        "total_categories": len(category_counts),
        "max_sentences_per_category": max(category_counts.values(), default=0),
        "top_level_categories": top_level_categories,
    }

def _format_audit_failure_message(raw, is_retryable):
    """Return (lead, suggestions) for a failed-audit error.

    `raw` is the "ExcType: message" string for copy-paste / debugging, and
    `is_retryable` says whether the error is in our retryable set (network /
    overload / timeout) vs everything else (auth, model name typos, validation).
    The classification is captured by the worker at failure time (the original
    exception object isn't available when the UI later renders the outcome).
    """
    if is_retryable:
        lead = (
            "The audit failed because the LLM API was unreachable, overloaded, "
            f"or timed out ({raw}). This is usually temporary."
        )
        suggestions = [
            "**Wait a few minutes and try again.**",
            "**Switch model or provider** in the left sidebar — smaller models "
            "like `claude-haiku-4-5` or `gpt-5-mini` are often available "
            "when larger ones aren't.",
            "**Reduce \"Max sentences per category\"** in the sidebar to send "
            "fewer sentences for evaluation. (Un-audited sentences will be "
            "ignored in accuracy calculations.)",
            "**Try disabling any VPN or corporate proxy.** Some (e.g. NordVPN, "
            "company firewalls) cut off long LLM requests.",
        ]
    else:
        lead = (
            "The audit failed with an error that retrying alone won't fix "
            f"({raw})."
        )
        suggestions = [
            "**Verify the API key** has quota left for the selected model.",
            "**Check the model name** in the sidebar for typos.",
            "**Switch model or provider** in the left sidebar.",
            "If the error mentions tokens or message length, **reduce "
            "\"Max tokens per request\" or \"Max sentences per category\"** "
            "in the sidebar.",
        ]
    return lead, suggestions


def _strip_status_suffixes(base):
    """Remove trailing _checkpoint_YYYY-MM-DD[_HHMM] and _completed_YYYY-MM-DD[_HHMM] suffixes.

    The optional `_HHMM` segment is matched so checkpoints saved on the same
    day across different runs don't collide on filename. Older date-only
    suffixes are still stripped via the `(?:_\\d{4})?` group.
    """
    return re.sub(r"(_(?:checkpoint|completed)_\d{4}-\d{2}-\d{2}(?:_\d{4})?)+$", "", base)

def _render_locked_topic_list(partial_detection):
    """Render a read-only list of the checkpoint's locked topic selection.

    Used in place of the streamlit_tree_select widget when a checkpoint is
    uploaded — the run's topic list is fixed to whatever the original run was
    configured with, so the full model tree is hidden to avoid implying it
    can be changed.
    """
    selected = partial_detection.get("selected_categories", []) or []
    completed = partial_detection.get("completed_categories", set()) or set()
    incomplete = partial_detection.get("incomplete_categories", set()) or set()
    st.caption(
        f"Topic selection is locked to the {len(selected)} categories in the "
        f"uploaded checkpoint. The full model tree is hidden while resuming a "
        f"checkpoint."
    )
    if not selected:
        return
    with st.expander(f"View locked topic list ({len(selected)} topics)", expanded=False):
        lines = []
        for topic in selected:
            if topic in completed:
                marker = "✓"
            elif topic in incomplete:
                marker = "↻"
            else:
                marker = "○"
            lines.append(f"- {marker} {topic}")
        st.markdown("\n".join(lines))
        st.caption("✓ completed   ↻ partially audited (will be re-audited)   ○ not yet audited")


def _build_completed_filename(uploaded_file):
    original_name = getattr(uploaded_file, "name", "") or "completed_audit.xlsx"
    base, ext = os.path.splitext(original_name)
    if base.endswith("_sortable"):
        base = base[: -len("_sortable")]
    base = _strip_status_suffixes(base)
    if not ext:
        ext = ".xlsx"
    # Include HH-MM so multiple runs on the same day produce distinct
    # filenames; otherwise the browser appends "(1)", "(2)" suffixes that
    # look like an ordered checkpoint sequence but aren't.
    date_suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
    return f"{base}_completed_{date_suffix}{ext}"


def _checkpoint_filename(uploaded_file):
    """The checkpoint (partial) variant of the completed-audit filename."""
    name = _build_completed_filename(uploaded_file)
    if "_completed" in name:
        return name.replace("_completed", "_checkpoint")
    if "_checkpoint" not in name:
        base, ext = os.path.splitext(name)
        return f"{base}_checkpoint{ext}"
    return name


def _next_checkpoint_topic(current, total):
    """Topic the next checkpoint will land on.

    Checkpoints fire on multiples of CHECKPOINT_INTERVAL and always on the final
    topic, so the next one is the smallest multiple >= current, capped at total.
    """
    if not total:
        return current
    rounded_up = (
        (current + CHECKPOINT_INTERVAL - 1) // CHECKPOINT_INTERVAL
    ) * CHECKPOINT_INTERVAL
    return min(rounded_up, total)


def _get_node_field(element, field_name):
    value = element.get(field_name)
    if value is None:
        value = element.findtext(field_name)
    return value

def _parse_model_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    model_element = root if root.tag == "model" else root.find("model")
    if model_element is None:
        raise ValueError("Could not find <model> in XML.")

    tree_element = model_element.find("tree")
    if tree_element is None:
        raise ValueError("Could not find <tree> under <model> in XML.")

    model_name = model_element.get("name") or model_element.findtext("name")
    nodes_by_id = {}
    top_nodes = list(tree_element.findall("node"))
    root_node = top_nodes[0] if top_nodes else None
    root_name = _get_node_field(root_node, "name") if root_node is not None else None

    def build_node(node_element, parent_path_parts):
        node_id = _get_node_field(node_element, "id")
        node_name = _get_node_field(node_element, "name")
        node_description = _get_node_field(node_element, "description")
        node_order_number = _get_node_field(node_element, "order-number")
        node_smart_other = _get_node_field(node_element, "smart-other")

        current_path_parts = parent_path_parts + ([node_name] if node_name else [])
        if root_name and current_path_parts and current_path_parts[0] == root_name:
            current_path_parts = current_path_parts[1:]
        node_path = "-->".join([part for part in current_path_parts if part])

        node_record = {
            "id": node_id,
            "name": node_name,
            "description": node_description,
            "order-number": node_order_number,
            "smart-other": node_smart_other,
            "path": node_path,
            "children": [],
        }
        if node_id:
            nodes_by_id[node_id] = node_record

        tree_node = {
            "label": node_name or node_id or "Unnamed node",
            "value": node_path or node_name or node_id or "unknown-node",
        }

        child_tree_nodes = []
        for child in node_element.findall("node"):
            child_record, child_tree_node = build_node(child, current_path_parts)
            if child_record["id"]:
                node_record["children"].append(child_record["id"])
            child_tree_nodes.append(child_tree_node)

        if child_tree_nodes:
            tree_node["children"] = child_tree_nodes

        return node_record, tree_node

    tree_nodes = []
    if top_nodes:
        if len(top_nodes) == 1:
            _, root_tree_node = build_node(top_nodes[0], [])
            tree_nodes = root_tree_node.get("children", [])
        else:
            for node in top_nodes:
                _, tree_node = build_node(node, [])
                tree_nodes.append(tree_node)

    root_title = _get_node_field(root_node, "name") if root_node is not None else None
    root_description = model_element.get("desc") or model_element.get("description") if model_element is not None else None

    model_data = {
        "model_name": model_name,
        "nodes": nodes_by_id,
        "tree_nodes": tree_nodes,
        "root_title": root_title,
        "root_description": root_description,
    }

    return model_data

def get_api_key(provider="ANTHROPIC"):
    api_var = provider + "_API_KEY"
    try:
        secrets = st.secrets
    except Exception:
        return None

    try:
        if api_var in secrets:
            return secrets[api_var]
    except Exception:
        return None

    return None

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    # st.cache_data.clear()

    organization, audience = get_org_and_audience()

    org_for_title = ""
    if organization != "the organization":
        org_for_title = organization + " "

    st.title(f"{org_for_title}Automatic Audit")
    st.write("This app uses an LLM to audit the accuracy of CX Designer models, and provides a summary of the findings by category.")
    if "audit_run_requested" not in st.session_state:
        st.session_state["audit_run_requested"] = False

    if "warnings_acknowledged" not in st.session_state:
        st.session_state["warnings_acknowledged"] = False

    # The background-worker registry lives in the process (not session_state),
    # so a websocket reconnect / rerun re-attaches to the same running job.
    registry = audit_worker.get_registry()

    def _queue_audit_run():
        st.session_state["audit_run_requested"] = True

    def _queue_audit_run_with_warning_check():
        """Request audit, but if there are pre-flight warnings, require confirmation first."""
        if audit_warnings and not st.session_state.get("warnings_acknowledged", False):
            st.session_state["warnings_confirmation_needed"] = True
        else:
            st.session_state["warnings_acknowledged"] = False
            _queue_audit_run()

    def _acknowledge_warnings():
        st.session_state["warnings_acknowledged"] = True
        st.session_state["warnings_confirmation_needed"] = False
        _queue_audit_run()

    def _cancel_warnings():
        st.session_state["warnings_confirmation_needed"] = False

    st.subheader("Upload audit file")
    # st.write("Upload an audit file generated by CX Designer.")
    uploaded_audit = st.file_uploader(
        "To start, upload an audit file generated by CX Designer, or an audit checkpoint generated by this app.",
        type=["xlsx"],
    )

    if uploaded_audit:
        if st.session_state.get("audit_source_name") != uploaded_audit.name:
            st.session_state["audit_source_name"] = uploaded_audit.name
            st.session_state.pop("reformatted_audit_bytes", None)
            st.session_state.pop("audit_output_bytes", None)
            st.session_state.pop("audit_is_partial", None)
            st.session_state.pop("partial_audit_detection", None)
            st.session_state["warnings_acknowledged"] = False
            st.session_state["warnings_confirmation_needed"] = False

        # Detect if this is a partial audit file
        if "partial_audit_detection" not in st.session_state:
            audit_bytes_to_check = uploaded_audit.getvalue()
            detection = detect_partial_audit(audit_bytes_to_check)
            st.session_state["partial_audit_detection"] = detection

        partial_detection = st.session_state.get("partial_audit_detection", {})
        if partial_detection.get("is_partial"):
            completed_count = len(partial_detection.get("completed_categories", set()))
            incomplete_count = len(partial_detection.get("incomplete_categories", set()))
            unjudged_count = len(partial_detection.get("unjudged_categories", set()))
            selected_categories = partial_detection.get("selected_categories", []) or []
            # Lock the run's topic list to the checkpoint's original selection so
            # successive resume attempts can't silently change which topics get
            # audited (which is what produced Harold's confusing 71/72/59 set).
            st.session_state["topics_to_audit"] = list(selected_categories)
            st.warning(
                f"You uploaded an audit checkpoint with {len(selected_categories)} topics "
                f"({completed_count} completed, {incomplete_count} incomplete, {unjudged_count} not yet audited). "
                f"Topic selection is locked to the categories in the checkpoint — "
                f"only not-yet-completed categories will be re-audited, and any partially-audited "
                f"category will be re-audited entirely."
            )

        # Reformat functionality
        # with st.expander("Reformat audit (optional)"):
        #     st.write("Download a sortable version of the input file.")
        #     if st.button("Reformat audit", help="Optionally reformatted audit for review prior to processing"):
        #         if uploaded_audit is None:
        #             st.error("Please upload an audit file before processing.")
        #             return
        #         with st.spinner("Reformatting audit..."):
        #             output, output_filename, warnings = handle_audit_reformat(uploaded_audit)
        #             st.success("Reformat complete.")
        #             if warnings:
        #                 st.warning("Input audit file warnings:\n" + "\n".join(warnings))
        #             st.session_state["reformatted_audit_bytes"] = output.getvalue()
        #             st.download_button(
        #                 label="Download reformatted audit",
        #                 data=output.getvalue(),
        #                 file_name=output_filename,
        #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        #             )

        # Allow user to upload an audit file that was already reformatted
        # if st.expander("Reformatted audit file"):
        #     reformatted_audit = st.file_uploader(
        #         "Upload reformatted audit file (.xlsx)",
        #         type=["xlsx"],
        #     )

    st.subheader("Add model tree")
    model_tree = st.file_uploader(
        "Upload the model tree to select particular topics to audit, and provide the LLM with category descriptions that can improve its judgment. (Optional)",
        type=["xml"],
    )

    prompts_path = os.path.join(os.path.dirname(__file__), "prompts.yaml")
    config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
    audit_defaults = _load_audit_defaults(prompts_path)
    app_defaults = _load_app_defaults(config_path)

    if model_tree:
        if st.session_state.get("model_source_name") != model_tree.name:
            try:
                model_data = _parse_model_xml(model_tree.getvalue())
            except (ET.ParseError, ValueError) as exc:
                st.error(f"Unable to parse model XML: {exc}")
                return
            st.session_state["model_data"] = model_data
            st.session_state["model_source_name"] = model_tree.name
            # Don't override topics_to_audit when a checkpoint is loaded — the
            # checkpoint's selected_categories are the canonical run list and
            # were already written to session state above.
            partial_detection = st.session_state.get("partial_audit_detection", {})
            if not partial_detection.get("is_partial"):
                st.session_state["topics_to_audit"] = [
                    node["path"] for node in model_data["nodes"].values() if node.get("path")
                ]

        model_data = st.session_state.get("model_data")
        if not model_data:
            st.error("Model data is missing. Please re-upload the XML file.")
            return

        if uploaded_audit:
            try:
                # Skip alignment check for partial audits - they may only have a subset of categories
                partial_detection = st.session_state.get("partial_audit_detection", {})
                is_partial_audit = partial_detection.get("is_partial", False)

                if not is_partial_audit:
                    audit_bytes_for_stats = st.session_state.get("reformatted_audit_bytes")
                    if not audit_bytes_for_stats:
                        audit_bytes_for_stats = uploaded_audit.getvalue()
                    stats = _get_audit_stats(audit_bytes_for_stats)
                    tree_nodes = model_data.get("tree_nodes", [])
                    model_top_levels = {
                        _topic_key(node.get("label", ""))
                        for node in tree_nodes
                        if node.get("label")
                    }
                    audit_top_levels = stats.get("top_level_categories", set())
                    if model_top_levels and audit_top_levels:
                        if model_top_levels != audit_top_levels:
                            st.warning(
                                "Category names in XML tree do not align with audit file; check that correct files were selected."
                            )
            except Exception:
                pass

        st.write("Select nodes to be audited:")
        if model_data.get("model_name"):
            st.caption(f"{model_data['model_name']}")

        tree_busy = registry.active() is not None or st.session_state.get("audit_run_requested", False)
        partial_detection = st.session_state.get("partial_audit_detection", {})
        is_partial_audit = partial_detection.get("is_partial", False)

        if is_partial_audit:
            # Skip the tree widget entirely when resuming a checkpoint. Showing
            # the full XML tree alongside a partial selection is misleading;
            # render a read-only list of just the checkpoint's topics instead.
            _render_locked_topic_list(partial_detection)
        else:
            if tree_busy:
                st.caption("Topic selection is disabled while an audit is running.")

            @st.fragment
            def _render_tree():
                tree_state = tree_select(
                    model_data["tree_nodes"],
                    checked=st.session_state.get("topics_to_audit", []),
                    key="topics_tree",
                    disabled=tree_busy,
                )

                # Only update selected topics if not busy (prevent changes during audit)
                if not tree_busy:
                    st.session_state["topics_to_audit"] = tree_state.get("checked", [])

            _render_tree()

            st.write("Note: Topics with no rules will not appear in audit output, even if selected in model tree.")
    elif uploaded_audit and st.session_state.get("partial_audit_detection", {}).get("is_partial"):
        # Checkpoint uploaded without a model XML — still surface the locked list
        # so the user can see what will be re-audited.
        _render_locked_topic_list(st.session_state["partial_audit_detection"])

    st.subheader("Add context")

    model_data = st.session_state.get("model_data")
    include_model_description = False
    has_imported_description = False
    if model_tree and model_data:
        root_desc = model_data.get("root_description")
        if root_desc:
            has_imported_description = True
            st.info(f"**Model description (imported from XML):** {root_desc}")
            include_model_description = st.checkbox(
                "Include model description in prompts",
                value=True,
                help="Send the model description alongside the model's title to the LLM with each prompt to aid its understanding of the model's purpose.",
            )
        else:
            st.info("No description was found for this model.")

    # When the XML provides a description, the extra free-text field is just
    # optional polish — tuck it behind an expander to reduce visual clutter.
    # Otherwise it's the primary description input and stays inline.
    notes_container = (
        st.expander("Add additional information about model")
        if has_imported_description
        else st.container()
    )
    with notes_container:
        if not has_imported_description:
            st.write(
                "Write a short description of the model you're auditing to aid "
                "the LLM's understanding."
            )
        model_info = st.text_area(
            "About this model: (optional)",
            max_chars=1000,
            value=audit_defaults["model_info"],
            help="Tell the LLM about anything unique to this model, or the feedback it targets, so that it can make informed decisions.",
            placeholder=f"This model categorizes feedback about the {organization} loyalty program..."
        )

    with st.expander("Audit prompt"):
        audit_prompt = st.text_area(
            label="Task instructions:",
            value=audit_defaults["audit_prompt"],
            max_chars=2500,
            placeholder="Tell the LLM what to do...",
            help="The prompt is sent to the LLM once per category. Use {category} to refer to the category name, and {description} to refer to the category's description.",
        )

    sidebar = st.sidebar
    sidebar.subheader("LLM Settings")
    llm_provider_options = ["anthropic", "openai"]
    default_provider = app_defaults["llm_provider"]
    provider_index = (
        llm_provider_options.index(default_provider)
        if default_provider in llm_provider_options
        else 0
    )

    #use_manual_api_key = sidebar.checkbox("Use my own API key", value=False, help="Override default API connection and ")
    use_manual_api_key = False

    manual_key_valid = st.session_state.get("manual_api_key_valid", False)
    manual_key_provider = st.session_state.get("manual_api_key_provider")

    llm_provider = sidebar.selectbox(
        "Provider",
        options=llm_provider_options,
        index=provider_index,
        disabled=use_manual_api_key and not manual_key_valid,
    )
    if use_manual_api_key and manual_key_provider and manual_key_provider != llm_provider:
        st.session_state["manual_api_key_valid"] = False
        st.session_state["manual_api_key_error"] = None
        st.session_state["manual_api_key_provider"] = None
        manual_key_valid = False
        manual_key_provider = None

    default_model = (
        app_defaults["model_name_anthropic"]
        if llm_provider == "anthropic"
        else app_defaults["model_name_openai"]
    )

    api_key = get_api_key(llm_provider.upper())
    error_placeholder = sidebar.empty()

    manual_key_error = st.session_state.get("manual_api_key_error")
    manual_key_value = None
    if use_manual_api_key:
        if llm_provider == "anthropic":
            manual_key_value = sidebar.text_input(
                "Anthropic API key",
                type="password",
                help="Uses ANTHROPIC_API_KEY from the environment if left blank.",
            )
        elif llm_provider == "openai":
            manual_key_value = sidebar.text_input(
                "OpenAI API key",
                type="password",
                help="Uses OPENAI_API_KEY from the environment if left blank.",
            )
        if sidebar.button("Set API key"):
            is_valid, err = _validate_api_key(llm_provider, manual_key_value)
            st.session_state["manual_api_key_valid"] = is_valid
            st.session_state["manual_api_key_error"] = err
            st.session_state["manual_api_key_provider"] = llm_provider if is_valid else None
            if is_valid:
                st.session_state["manual_api_key_value"] = manual_key_value
            manual_key_valid = is_valid
            manual_key_error = err

        if manual_key_error:
            sidebar.error(f"API key invalid: {manual_key_error}")

    if use_manual_api_key:
        key_for_models = st.session_state.get("manual_api_key_value") if manual_key_valid else None
        if not manual_key_valid:
            model_options = [default_model]
            models_enabled = False
            model_error = "Set valid API credentials."
        else:
            model_options, models_enabled, model_error = _get_model_options(llm_provider, key_for_models, default_model)
    else:
        model_options, models_enabled, model_error = _get_model_options(llm_provider, api_key, default_model)

    model_index = model_options.index(default_model) if default_model in model_options else 0
    model_name = sidebar.selectbox(
        "Model",
        options=model_options,
        index=model_index,
        disabled=(use_manual_api_key and not manual_key_valid) or not models_enabled,
    )
    if (use_manual_api_key and not manual_key_valid) or not models_enabled:
        sidebar.warning(model_error or "Set valid API credentials")

    anthropic_api_key = None
    openai_api_key = None
    if use_manual_api_key and manual_key_valid:
        if llm_provider == "anthropic":
            anthropic_api_key = st.session_state.get("manual_api_key_value")
        elif llm_provider == "openai":
            openai_api_key = st.session_state.get("manual_api_key_value")

    if use_manual_api_key:
        if not manual_key_valid:
            error_placeholder.error("Valid API key required when using your own API key.")
        else:
            error_placeholder.empty()
    else:
        if not api_key:
            error_placeholder.error(f"{llm_provider} API key not found in Streamlit secrets.")
        else:
            error_placeholder.empty()

    sidebar.subheader("API limits", help="Adjust values to override maximum LLM spend")
    max_categories = sidebar.number_input(
        "Max categories to audit",
        min_value=1,
        value=int(app_defaults["max_categories"]),
        step=1,
    )
    max_sentences = sidebar.number_input(
        "Max sentences per category",
        min_value=1,
        value=int(app_defaults["max_sentences_per_category"]),
        step=1,
    )
    max_tokens = sidebar.number_input(
        "Max tokens per request",
        min_value=1,
        value=int(app_defaults["max_tokens"]),
        step=100,
    )

    missing_reasons = []
    if not uploaded_audit:
        missing_reasons.append("Upload an audit file")
    if llm_provider == "anthropic":
        if use_manual_api_key:
            if not manual_key_valid:
                missing_reasons.append("Provide a valid Anthropic API key")
        else:
            if not api_key:
                missing_reasons.append("Provide an API key to access LLM audit functionality")
    elif llm_provider == "openai":
        if use_manual_api_key:
            if not manual_key_valid:
                missing_reasons.append("Provide a valid OpenAI API key")
        else:
            if not api_key:
                missing_reasons.append("Provide an OpenAI API key")

    can_run_audit = not missing_reasons
    run_help = (
        None
        if can_run_audit
        else "; ".join(missing_reasons)
    )

    audit_warnings = []
    if uploaded_audit:
        try:
            audit_bytes_for_stats = st.session_state.get("reformatted_audit_bytes")
            if not audit_bytes_for_stats:
                audit_bytes_for_stats = uploaded_audit.getvalue()
            stats = _get_audit_stats(audit_bytes_for_stats)
            category_counts = stats["category_counts"]
            category_names = list(category_counts.keys())
            selected_topics = st.session_state.get("topics_to_audit")
            if selected_topics:
                categories_by_key = {_topic_key(cat): cat for cat in category_names}
                filtered = []
                seen = set()
                for topic in selected_topics:
                    match = categories_by_key.get(_topic_key(topic))
                    if match and match not in seen:
                        filtered.append(match)
                        seen.add(match)
                categories_to_audit = filtered or category_names
            else:
                categories_to_audit = category_names

            total_categories = len(categories_to_audit)
            if total_categories and int(max_categories) < total_categories:
                skipped = total_categories - int(max_categories)
                audit_warnings.append(
                    f"The input has {total_categories} categories but the max is set to {int(max_categories)}. "
                    f"{skipped} categories will be skipped entirely. "
                    f"Adjust \"Max categories to audit\" in the sidebar to change this."
                )

            if categories_to_audit:
                max_sentences_in_category = max(
                    category_counts[cat] for cat in categories_to_audit
                )
                if int(max_sentences) < max_sentences_in_category:
                    affected = sum(
                        1 for cat in categories_to_audit
                        if category_counts[cat] > int(max_sentences)
                    )
                    audit_warnings.append(
                        f"{affected} of {len(categories_to_audit)} categories have more sentences than the "
                        f"max setting of {int(max_sentences)}. Only the first {int(max_sentences)} sentences per "
                        f"category will be audited; the rest will be excluded from results. "
                        f"Adjust \"Max sentences per category\" in the sidebar to change this."
                    )

                estimated_tokens_per_sentence = 30
                effective_max_sentences = min(max_sentences_in_category, int(max_sentences))
                estimated_output_tokens = effective_max_sentences * estimated_tokens_per_sentence
                if estimated_output_tokens >= int(max_tokens):
                    audit_warnings.append(
                        f"Categories with many sentences may produce LLM responses that exceed the "
                        f"max tokens per request limit ({int(max_tokens)}). Some sentences in those "
                        f"categories may be returned without audit results. Adjust \"Max tokens per "
                        f"request\" or \"Max sentences per category\" in the sidebar to change this."
                    )

        except Exception as exc:
            audit_warnings.append(f"Unable to estimate audit limits: {exc}")

    for warning in audit_warnings:
        st.warning(warning)

    summary_prompt_default = _load_summary_prompt(prompts_path)

    with st.expander("Summary settings"):
        st.text_area(
            label="Prompt:",
            value=summary_prompt_default,
            max_chars=3000,
            placeholder="Tell the LLM how to summarize audit findings...",
            help="Sentences are batched by category.",
            key="summary_prompt",
        )
        st.number_input(
            label="Accuracy threshold:",
            min_value=0.0,
            max_value=1.0,
            value=0.80,
            step=0.01,
            help="Categories with accuracy below this threshold will have summaries generated.",
            key="accuracy_threshold",
        )

    # Re-attach to a job still running from a previous script run / session.
    # The registry lives in the process, so a websocket reconnect (which can
    # spin up a fresh session with empty session_state) re-finds the live run.
    active_job = registry.active()
    if active_job is not None and not st.session_state.get("active_run_id"):
        st.session_state["active_run_id"] = active_job.run_id

    # Launch trigger: a run was requested and nothing is currently active.
    should_run_audit = (
        st.session_state.get("audit_run_requested", False)
        and registry.active() is None
        and can_run_audit
    )
    if (
        not should_run_audit
        and st.session_state.get("audit_run_requested", False)
        and not can_run_audit
    ):
        st.session_state["audit_run_requested"] = False

    # Opt-in email delivery — survives Streamlit Cloud winding the app down
    # before the user can download.  Rendered every pass so the widget values
    # persist in session_state through the run.
    email_enabled = st.checkbox(
        "Email results to me",
        value=True,
        key="email_results_enabled",
        help=(
            "When the audit finishes (or stops/errors with partial results), "
            "email the workbook as an attachment. Useful if the app is put to "
            "sleep before you can download it."
        ),
    )
    if email_enabled:
        st.text_input(
            "Send results to (email address):",
            key="email_results_address",
            placeholder="you@example.com",
        )
        if not _email_configured():
            st.caption(
                "⚠️ Email sending isn't configured yet — ask an admin to add "
                "SMTP settings to the app secrets."
            )

    # ------------------------------------------------------------------
    # Launch / poll the background audit job
    # ------------------------------------------------------------------
    def _launch_audit_job():
        """Build params on the main thread and hand the run to the worker."""
        partial_detection = st.session_state.get("partial_audit_detection", {})
        is_partial_audit = partial_detection.get("is_partial", False)

        audit_bytes = st.session_state.get("reformatted_audit_bytes")
        if not audit_bytes:
            if uploaded_audit is None:
                st.error("Please upload an audit file before running the audit.")
                st.session_state["audit_run_requested"] = False
                return
            if is_partial_audit:
                # Checkpoints are already in our format — skip reformatting.
                audit_bytes = uploaded_audit.getvalue()
                st.session_state["reformatted_audit_bytes"] = audit_bytes
            else:
                with st.spinner("Reformatting audit..."):
                    output, _, warnings = handle_audit_reformat(uploaded_audit)
                    if warnings:
                        st.warning("Input audit file warnings:\n" + "\n".join(warnings))
                    st.session_state["reformatted_audit_bytes"] = output.getvalue()
                    audit_bytes = st.session_state["reformatted_audit_bytes"]

        if not model_tree and not st.session_state.get("topics_to_audit"):
            topics_to_audit = None
        else:
            topics_to_audit = st.session_state.get("topics_to_audit")

        if use_manual_api_key:
            final_anthropic_key = anthropic_api_key if llm_provider == "anthropic" else None
            final_openai_key = openai_api_key if llm_provider == "openai" else None
        else:
            final_anthropic_key = api_key if llm_provider == "anthropic" and api_key else None
            final_openai_key = api_key if llm_provider == "openai" and api_key else None

        existing_audit_bytes = None
        completed_categories = None
        if is_partial_audit:
            existing_audit_bytes = uploaded_audit.getvalue()
            completed_categories = partial_detection.get("completed_categories", set())

        summary_prompt_val = st.session_state.get("summary_prompt", summary_prompt_default)
        accuracy_threshold_val = st.session_state.get("accuracy_threshold", 0.80)

        audit_kwargs = dict(
            audit_excel_bytes=audit_bytes,
            prompt_template=audit_prompt,
            llm_provider=llm_provider,
            model_name=model_name,
            model_info=model_info,
            organization=organization,
            audience=audience,
            max_categories=int(max_categories),
            max_sentences_per_category=int(max_sentences),
            model_tree_bytes=model_tree.getvalue() if model_tree else None,
            topics_to_audit=topics_to_audit,
            anthropic_api_key=final_anthropic_key,
            openai_api_key=final_openai_key,
            max_tokens=int(max_tokens),
            existing_audit_bytes=existing_audit_bytes,
            completed_categories=completed_categories,
            audit_file_name=uploaded_audit.name,
            model_tree_name=model_tree.name if model_tree else None,
            include_summary=True,
            summary_prompt=summary_prompt_val,
            accuracy_threshold=accuracy_threshold_val,
            run_datetime=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            audit_warnings=audit_warnings,
            include_model_description=include_model_description,
        )
        summary_kwargs = dict(
            msg_template=summary_prompt_val,
            llm_provider=llm_provider,
            model_name=model_name,
            max_tokens=int(max_tokens),
            accuracy_threshold=accuracy_threshold_val,
            model_info=model_info,
            anthropic_api_key=final_anthropic_key,
            openai_api_key=final_openai_key,
        )
        params = audit_worker.JobParams(
            audit_kwargs=audit_kwargs,
            summary_kwargs=summary_kwargs,
            completed_filename=_build_completed_filename(uploaded_audit),
            checkpoint_filename=_checkpoint_filename(uploaded_audit),
            include_summary=True,
            email=audit_worker.EmailPayload(
                enabled=bool(st.session_state.get("email_results_enabled")),
                recipient=(st.session_state.get("email_results_address") or "").strip(),
                smtp=_smtp_config(),
            ),
        )
        try:
            job = registry.start(params)
        except audit_worker.AlreadyRunning:
            job = registry.active()
        if job is not None:
            st.session_state["active_run_id"] = job.run_id
            logger.info("Audit job launched: provider=%s, model=%s", llm_provider, model_name)
        st.session_state["audit_run_requested"] = False
        st.session_state.pop("finalized_run_id", None)
        st.rerun()

    @st.fragment(run_every="2s")
    def _audit_progress_panel():
        """Poll the active job every 2s and redraw progress (non-blocking)."""
        job = registry.get(st.session_state.get("active_run_id"))
        if job is None:
            return
        snap = job.snapshot()
        if snap.status in (
            audit_worker.JobStatus.DONE,
            audit_worker.JobStatus.STOPPED,
            audit_worker.JobStatus.ERROR,
        ):
            # Leave poll mode; the next full run finalizes + renders the output.
            st.rerun(scope="app")
            return

        current, total = snap.progress_current, snap.progress_total
        if snap.status == audit_worker.JobStatus.RUNNING_SUMMARY:
            st.write(f"Summarizing category {current} of {total}: {snap.progress_category}")
        elif snap.status == audit_worker.JobStatus.RUNNING_AUDIT:
            st.write(f"Auditing category {current} of {total}: {snap.progress_category}")
        else:
            st.write("Starting audit…")
        st.progress(current / total if total else 0.0)

        if snap.status_message:
            st.info(snap.status_message)

        # Checkpoint cadence line (audit phase only — summary has no checkpoints).
        if snap.status == audit_worker.JobStatus.RUNNING_AUDIT:
            next_cp = _next_checkpoint_topic(current, total)
            if snap.checkpoint_topic:
                st.caption(
                    f"✓ Checkpoint saved through topic {snap.checkpoint_topic}; "
                    f"next at topic {next_cp}"
                )
            else:
                st.caption(
                    f"Progress is saved as a checkpoint every {CHECKPOINT_INTERVAL} "
                    f"topics. First checkpoint after topic {next_cp}."
                )

        # Live checkpoint download — stable key, only the data changes each tick.
        if snap.checkpoint_bytes and uploaded_audit:
            st.download_button(
                label="Download audit checkpoint (.xlsx)",
                data=snap.checkpoint_bytes,
                file_name=_checkpoint_filename(uploaded_audit),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="live_checkpoint_download",
            )
            if snap.checkpoint_topic:
                st.caption(f"File contains progress through topic {snap.checkpoint_topic}")

        stop_label = (
            "Stop summary and keep audit"
            if snap.status == audit_worker.JobStatus.RUNNING_SUMMARY
            else "Stop audit"
        )
        st.button(
            stop_label,
            type="secondary",
            key="stop_audit_btn",
            on_click=job.stop_event.set,
        )

    def _finalize_job_into_session(job):
        """Promote a terminal job's output into session_state + show the outcome."""
        snap = job.snapshot()
        if st.session_state.get("finalized_run_id") != job.run_id:
            st.session_state["audit_output_bytes"] = snap.output_bytes
            st.session_state["audit_output_filename"] = snap.output_filename
            st.session_state["audit_is_partial"] = snap.is_partial
            st.session_state["finalized_run_id"] = job.run_id
        st.session_state.pop("active_run_id", None)

        if snap.status == audit_worker.JobStatus.DONE:
            if snap.is_partial:
                st.warning("Summary generation was stopped; audit results are available below.")
            else:
                st.success("Audit complete.")
        elif snap.status == audit_worker.JobStatus.STOPPED:
            st.warning("Audit stopped by user request.")
            if snap.output_bytes:
                st.info(
                    "Partial audit results are available below. Re-upload the "
                    "checkpoint file to resume the audit where it left off."
                )
        elif snap.status == audit_worker.JobStatus.ERROR:
            raw = f"{snap.error_type}: {snap.error_message}"
            lead, suggestions = _format_audit_failure_message(raw, snap.error_retryable)
            bullets = "\n".join(f"- {s}" for s in suggestions)
            st.error(f"{lead}\n\n**What to try:**\n{bullets}")
            if snap.output_bytes:
                st.info(
                    "A checkpoint of the audit's progress so far is available "
                    "below. Re-upload it to resume from where it stopped — you "
                    "can change provider, model, or per-call limits before you do."
                )

        # Surface the worker's best-effort email outcome.
        email_status = snap.email_status or ""
        if email_status.startswith("sent:"):
            st.success(f"Results emailed to {email_status.split(':', 1)[1]}.")
        elif email_status.startswith("failed:"):
            st.warning(
                "The run finished but emailing the results failed: "
                f"{email_status.split(':', 1)[1]}. You can still download the file below."
            )
        elif email_status.startswith("skipped:"):
            reason = email_status.split(":", 1)[1]
            if reason.startswith("invalid") or reason.startswith("SMTP"):
                st.warning(f"Results not emailed: {reason}.")

        for warning_msg in snap.warnings:
            st.warning(warning_msg)

    if should_run_audit:
        _launch_audit_job()

    job = registry.get(st.session_state.get("active_run_id"))
    if job is not None and not job.is_terminal():
        # A run is in flight — show the live progress panel (it owns the Stop
        # button + live checkpoint download). Reruns/reconnects re-attach here.
        _audit_progress_panel()
    else:
        if job is not None:
            _finalize_job_into_session(job)
        if st.session_state.get("warnings_confirmation_needed", False):
            st.warning("Please review the warnings above before proceeding.")
            col1, col2 = st.columns([1, 1])
            with col1:
                st.button(
                    "Continue anyway",
                    type="primary",
                    on_click=_acknowledge_warnings,
                )
            with col2:
                st.button(
                    "Cancel",
                    type="secondary",
                    on_click=_cancel_warnings,
                )
        else:
            st.button(
                "Run audit",
                type="primary",
                disabled=not can_run_audit,
                help=run_help,
                on_click=_queue_audit_run_with_warning_check,
            )

    # Final download — shown once a job has finished and been finalized.
    audit_output_bytes = st.session_state.get("audit_output_bytes")
    if audit_output_bytes and registry.active() is None:
        is_partial = st.session_state.get("audit_is_partial", False)
        if is_partial:
            st.warning(
                "The audit was interrupted before it finished. "
                "Partial results are available for download below. "
                "You can re-upload the checkpoint file to resume the audit where it left off."
            )
        download_label = "Download audit checkpoint (.xlsx)" if is_partial else "Download completed audit (.xlsx)"
        completed_filename = st.session_state.get("audit_output_filename") or _build_completed_filename(uploaded_audit)
        st.download_button(
            label=download_label,
            data=audit_output_bytes,
            file_name=completed_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()
