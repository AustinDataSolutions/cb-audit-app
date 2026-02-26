import os
from datetime import datetime
import importlib.util
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
from audit import run_audit, AuditStopRequested, detect_partial_audit, _is_retryable_llm_error

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
        "max_sentences_per_category": 51,
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


def _load_summarizer_module():
    module_path = os.path.join(os.path.dirname(__file__), "audit-report-summarizer.py")
    spec = importlib.util.spec_from_file_location("audit_report_summarizer", module_path)
    if spec is None or spec.loader is None:
        raise ImportError("Unable to load audit-report-summarizer module.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


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

def _build_completed_filename(uploaded_file):
    original_name = getattr(uploaded_file, "name", "") or "completed_audit.xlsx"
    base, ext = os.path.splitext(original_name)
    if base.endswith("_sortable"):
        base = base[: -len("_sortable")]
    if not ext:
        ext = ".xlsx"
    date_suffix = datetime.now().strftime("%Y-%m-%d")
    return f"{base}_completed_{date_suffix}{ext}"

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

    model_data = {
        "model_name": model_name,
        "nodes": nodes_by_id,
        "tree_nodes": tree_nodes,
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
    # st.cache_data.clear()

    organization, audience = get_org_and_audience()

    org_for_title = ""
    if organization != "the organization":
        org_for_title = organization + " "

    st.title(f"{org_for_title}Automatic Audit")
    st.write("This app uses an LLM to audit the accuracy of CX Designer models, and provides a summary of the findings by category.")
    if "audit_in_progress" not in st.session_state:
        st.session_state["audit_in_progress"] = False
    if "audit_run_requested" not in st.session_state:
        st.session_state["audit_run_requested"] = False
    if "partial_audit_bytes" not in st.session_state:
        st.session_state["partial_audit_bytes"] = None
    if "audit_stop_requested" not in st.session_state:
        st.session_state["audit_stop_requested"] = False
    if "summary_stop_requested" not in st.session_state:
        st.session_state["summary_stop_requested"] = False

    def _queue_audit_run():
        st.session_state["audit_run_requested"] = True
        st.session_state["partial_audit_bytes"] = None
        st.session_state["audit_stop_requested"] = False

    def _request_audit_stop():
        st.session_state["audit_stop_requested"] = True

    st.subheader("Upload audit file")
    # st.write("Upload an audit file generated by CX Designer.")
    uploaded_audit = st.file_uploader(
        "To start, upload an audit file generated by CX Designer, or an audit that has been partially completed by this app.",
        type=["xlsx"],
    )

    if uploaded_audit:
        if st.session_state.get("audit_source_name") != uploaded_audit.name:
            st.session_state["audit_source_name"] = uploaded_audit.name
            st.session_state.pop("reformatted_audit_bytes", None)
            st.session_state.pop("audit_output_bytes", None)
            st.session_state.pop("audit_is_partial", None)
            st.session_state.pop("partial_audit_detection", None)

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
            st.warning(
                f"Detected partially completed audit. "
                f"({completed_count} completed, {incomplete_count} incomplete, {unjudged_count} not yet audited). "
                f"Program will only audit the categories that have not been completed yet, "
                f"and any incomplete category will be re-audited entirely."
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
        "Uploading the model tree allows you to select particular topics to audit, and provides the LLM with category descriptions that can improve its judgment. (Optional)",
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

        tree_busy = st.session_state.get("audit_in_progress", False) or st.session_state.get("audit_run_requested", False)
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

    st.subheader("Add context")
    st.write("Write a short description of the model you're auditing to aid the LLM's understanding.")
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
                estimated_output_tokens = max_sentences_in_category * estimated_tokens_per_sentence
                if estimated_output_tokens >= int(max_tokens):
                    audit_warnings.append(
                        "LLM response likely to be truncated for some categories based on "
                        f"max tokens per request limit of {int(max_tokens)} "
                        f"(response estimated at {estimated_tokens_per_sentence} tokens per sentence; "
                        f"input file contains up to {max_sentences_in_category} sentences)."
                    )

        except Exception as exc:
            audit_warnings.append(f"Unable to estimate audit limits: {exc}")

    for warning in audit_warnings:
        st.warning(warning)

    summary_prompt_default = _load_summary_prompt(prompts_path)

    generate_summary = st.checkbox(
        "Include summary of issues",
        value=True,
        key="generate_audit_summary",
        help="Generates a note summarizing the issues for each topic that failed the audit"
    )
    if not generate_summary:
        st.session_state["summary_generation_pending"] = False

    if generate_summary:
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

    # Determine if we will run the audit this pass, so we can show the correct button
    should_run_audit = (
        st.session_state.get("audit_run_requested", False)
        and not st.session_state.get("audit_in_progress", False)
        and can_run_audit
    )
    if should_run_audit:
        # Mark in-progress immediately so the button renders as "Stop audit"
        st.session_state["audit_in_progress"] = True
        st.session_state["audit_run_requested"] = False
        st.session_state["summary_generation_pending"] = False
    elif st.session_state.get("audit_run_requested", False) and not can_run_audit:
        st.session_state["audit_run_requested"] = False

    if st.session_state.get("audit_in_progress", False):
        st.button(
            "Stop audit",
            type="secondary",
            on_click=_request_audit_stop,
        )
    else:
        st.button(
            "Run audit",
            type="primary",
            disabled=not can_run_audit,
            help=run_help,
            on_click=_queue_audit_run,
        )

    if should_run_audit:
        partial_detection = st.session_state.get("partial_audit_detection", {})
        is_partial_audit = partial_detection.get("is_partial", False)

        audit_bytes = st.session_state.get("reformatted_audit_bytes")

        if not audit_bytes:
            if uploaded_audit is None:
                st.error("Please upload an audit file before running the audit.")
                st.session_state["audit_in_progress"] = False
                return
            # For partial audits, skip reformatting - the file is already in our format
            if is_partial_audit:
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

        try:
            with st.spinner("Running audit..."):
                progress_container = st.container()
                progress_text = progress_container.empty()
                progress_bar = progress_container.progress(0)
                download_container = progress_container.empty()

                def _update_progress(current, total, category_name):
                    progress_text.write(
                        f"Auditing category {current} of {total}: {category_name}"
                    )
                    progress_bar.progress(current / total if total else 0)

                @st.fragment
                def _render_download_button(partial_bytes, partial_filename, button_key):
                    st.download_button(
                        label="Download in-progress audit (.xlsx)",
                        data=partial_bytes,
                        file_name=partial_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=button_key,
                    )

                def _save_progress(partial_bytes):
                    st.session_state["partial_audit_bytes"] = partial_bytes
                    # Update the download button with the latest partial results
                    partial_filename = _build_completed_filename(uploaded_audit).replace("_completed", "_in_progress")
                    with download_container:
                        partial_download_counter = st.session_state.get("partial_download_counter", 0) + 1
                        st.session_state["partial_download_counter"] = partial_download_counter
                        _render_download_button(
                            partial_bytes,
                            partial_filename,
                            f"download_in_progress_{partial_download_counter}",
                        )

                def _check_stop():
                    return st.session_state.get("audit_stop_requested", False)

                if use_manual_api_key:
                    final_anthropic_key = anthropic_api_key if llm_provider == "anthropic" else None
                    final_openai_key = openai_api_key if llm_provider == "openai" else None
                else:
                    final_anthropic_key = api_key if llm_provider == "anthropic" and api_key else None
                    final_openai_key = api_key if llm_provider == "openai" and api_key else None

                # Check for partial audit resume
                partial_detection = st.session_state.get("partial_audit_detection", {})
                existing_audit_bytes = None
                completed_categories = None
                if partial_detection.get("is_partial"):
                    existing_audit_bytes = uploaded_audit.getvalue()
                    # Completed categories are those fully judged; incomplete ones will be re-audited
                    completed_categories = partial_detection.get("completed_categories", set())

                output_bytes = run_audit(
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
                    log_fn=st.write,
                    warn_fn=st.warning,
                    progress_fn=_update_progress,
                    save_progress_fn=_save_progress,
                    check_stop_fn=_check_stop,
                    existing_audit_bytes=existing_audit_bytes,
                    completed_categories=completed_categories,
                )
                progress_bar.progress(1.0)
                progress_text.empty()
                progress_bar.empty()
                download_container.empty()
            st.session_state["audit_output_bytes"] = output_bytes
            st.session_state["partial_audit_bytes"] = None
            st.session_state["audit_output_filename"] = _build_completed_filename(uploaded_audit)
            st.session_state["audit_is_partial"] = False
            st.session_state["summary_generation_pending"] = generate_summary
            st.success("Audit complete.")
        except AuditStopRequested:
            st.warning("Audit stopped by user request.")
            partial_bytes = st.session_state.get("partial_audit_bytes")
            if partial_bytes:
                st.session_state["audit_output_bytes"] = partial_bytes
                st.session_state["audit_output_filename"] = _build_completed_filename(uploaded_audit)
                st.session_state["audit_is_partial"] = True
                st.info("Partial audit results are available for download.")
        except Exception as exc:
            st.error(f"Audit failed: {exc}")
            partial_bytes = st.session_state.get("partial_audit_bytes")
            if partial_bytes:
                st.session_state["audit_output_bytes"] = partial_bytes
                failed_filename = _build_completed_filename(uploaded_audit)
                if "_completed" in failed_filename:
                    failed_filename = failed_filename.replace("_completed", "_partial")
                elif "_partial" not in failed_filename:
                    base, ext = os.path.splitext(failed_filename)
                    failed_filename = f"{base}_partial{ext}"
                st.session_state["audit_output_filename"] = failed_filename
                st.session_state["audit_is_partial"] = True
                st.info(
                    "Partial audit results are available for download. "
                    "You can re-upload the in-progress file to continue the audit where it left off."
                )
            if _is_retryable_llm_error(exc):
                st.warning(
                    "This error was caused by the LLM API being overloaded or rate-limited. "
                    "You can try again later, or select a different model or provider in the left sidebar."
                )
        finally:
            st.session_state["audit_in_progress"] = False

    audit_output_bytes = st.session_state.get("audit_output_bytes")
    summary_prompt = st.session_state.get("summary_prompt", summary_prompt_default)

    summary_missing_reasons = []
    if not audit_output_bytes:
        summary_missing_reasons.append("Run the audit first")
    if llm_provider == "anthropic":
        if use_manual_api_key:
            if not manual_key_valid:
                summary_missing_reasons.append("Provide a valid Anthropic API key")
        else:
            if not api_key:
                summary_missing_reasons.append("Provide an Anthropic API key")
    elif llm_provider == "openai":
        if use_manual_api_key:
            if not manual_key_valid:
                summary_missing_reasons.append("Provide a valid OpenAI API key")
        else:
            if not api_key:
                summary_missing_reasons.append("Provide an OpenAI API key")

    can_run_summary = not summary_missing_reasons
    summary_help = (
        None
        if can_run_summary
        else "; ".join(summary_missing_reasons)
    )

    if st.session_state.get("summary_generation_pending") and can_run_summary:
        progress_text = None
        progress_bar = None
        stop_button_container = None

        # Store the audit bytes before summary in case user stops
        audit_bytes_before_summary = st.session_state.get("audit_output_bytes")

        def _request_summary_stop():
            st.session_state["summary_stop_requested"] = True

        try:
            with st.spinner("Summarizing audit..."):
                summarizer_module = _load_summarizer_module()
                SummaryStopRequested = summarizer_module.SummaryStopRequested
                if use_manual_api_key:
                    final_anthropic_key = anthropic_api_key if llm_provider == "anthropic" else None
                    final_openai_key = openai_api_key if llm_provider == "openai" else None
                else:
                    final_anthropic_key = api_key if llm_provider == "anthropic" and api_key else None
                    final_openai_key = api_key if llm_provider == "openai" and api_key else None
                progress_container = st.container()
                progress_text = progress_container.empty()
                progress_bar = progress_container.progress(0)
                stop_button_container = progress_container.empty()
                stop_button_container.button(
                    "Stop summary and download audit",
                    on_click=_request_summary_stop,
                    key="stop_summary_button",
                )

                def _update_summary_progress(current, total, category_name):
                    progress_text.write(
                        f"Summarizing category {current} of {total}: {category_name}"
                    )
                    progress_bar.progress(current / total if total else 0)

                def _check_summary_stop():
                    return st.session_state.get("summary_stop_requested", False)

                accuracy_threshold = st.session_state.get("accuracy_threshold", 0.80)
                summary_bytes = summarizer_module.summarize_audit_report(
                    audit_excel_input=audit_output_bytes,
                    msg_template=summary_prompt,
                    llm_provider=llm_provider,
                    model_name=model_name,
                    max_tokens=int(max_tokens),
                    accuracy_threshold=accuracy_threshold,
                    model_info=model_info,
                    anthropic_api_key=final_anthropic_key,
                    openai_api_key=final_openai_key,
                    log_fn=st.write,
                    warn_fn=st.warning,
                    progress_fn=_update_summary_progress,
                    check_stop_fn=_check_summary_stop,
                )
                progress_bar.progress(1.0)
            st.session_state["audit_output_bytes"] = summary_bytes
            st.session_state["summary_stop_requested"] = False
            st.success("Audit summary complete.")
        except summarizer_module.SummaryStopRequested:
            st.warning("Summary generation stopped by user request.")
            # Restore the audit bytes without summary so user can download
            if audit_bytes_before_summary:
                st.session_state["audit_output_bytes"] = audit_bytes_before_summary
                st.info("Audit results (without summary) are available for download.")
            st.session_state["summary_stop_requested"] = False
        except Exception as exc:
            st.error(f"Audit summary failed: {exc}")
            # Audit output is still available from before summary started
            if audit_bytes_before_summary:
                st.info("Audit results (without summary) are available for download.")
        finally:
            st.session_state["summary_generation_pending"] = False
            if progress_text is not None:
                progress_text.empty()
            if progress_bar is not None:
                progress_bar.empty()
            if stop_button_container is not None:
                stop_button_container.empty()
    elif st.session_state.get("summary_generation_pending") and summary_help:
        st.warning(f"Audit summary pending: {summary_help}")

    audit_output_bytes = st.session_state.get("audit_output_bytes")
    summary_pending = st.session_state.get("summary_generation_pending")
    if audit_output_bytes and not summary_pending:
        is_partial = st.session_state.get("audit_is_partial", False)
        download_label = "Download in-progress audit (.xlsx)" if is_partial else "Download completed audit (.xlsx)"
        completed_filename = st.session_state.get("audit_output_filename") or _build_completed_filename(uploaded_audit)
        st.download_button(
            label=download_label,
            data=audit_output_bytes,
            file_name=completed_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()