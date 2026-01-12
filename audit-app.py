import os
import importlib.util
import re
from io import BytesIO
import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from streamlit_tree_select import tree_select
import yaml

from audit_reformat import handle_audit_reformat
from audit_validation import validate_audit_sentences_sheet
from audit import run_audit


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
    sentences_sheet, header_row_idx, _, _ = validate_audit_sentences_sheet(audit_bytes)
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

    id_col = df.columns[0]
    sentence_col = df.columns[1]
    category_col = df.columns[2]
    category_sentences = {}
    for _, row in df.iterrows():
        category = row[category_col]
        sentence_id = row[id_col]
        sentence = row[sentence_col]
        if pd.isna(category) or pd.isna(sentence_id) or pd.isna(sentence):
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
    return f"{base}_completed{ext}"


#This script is intended to be an end-to-end audit of Clarabridge topic models powered by LLMs
#It will start with uploading the audit output from Qualtrics and reformatting it for transformation,
#then will present the user an interface to allow them to select what part of the model they want audited,
#then peform the audit by sending sentences batched by category for rebiew by an LLM and return the completed audit
# along with accuracy, summary of findings, and suggestions for improvement per category

# Configure Streamlit page
st.set_page_config(page_title="Enhanced Audit")

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
    st.title("Automated Audit")
    st.write("This app uses an LLM to audit the accuracy of CX Designer models, and provides a summary of the findings.")

    st.subheader("Upload audit file")
    # st.write("Upload an audit file generated by CX Designer.")
    uploaded_audit = st.file_uploader(
        "To start, upload the audit file generated by CX Designer.",
        type=["xlsx"],
    )

    if uploaded_audit:
        if st.session_state.get("audit_source_name") != uploaded_audit.name:
            st.session_state["audit_source_name"] = uploaded_audit.name
            st.session_state.pop("reformatted_audit_bytes", None)
            st.session_state.pop("audit_output_bytes", None)
        with st.expander("Reformat audit (optional)"):
            st.write("Download a sortable version of the input file.")
            if st.button("Reformat audit", help="Optionally reformatted audit for review prior to processing"):
                if uploaded_audit is None:
                    st.error("Please upload an audit file before processing.")
                    return
                with st.spinner("Reformatting audit..."):
                    output, output_filename, warnings = handle_audit_reformat(uploaded_audit)
                    st.success("Reformat complete.")
                    if warnings:
                        st.warning("Input audit file warnings:\n" + "\n".join(warnings))
                    st.session_state["reformatted_audit_bytes"] = output.getvalue()
                    st.download_button(
                        label="Download reformatted audit",
                        data=output.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
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

        tree_state = tree_select(
            model_data["tree_nodes"],
            checked=st.session_state.get("topics_to_audit", []),
            key="topics_tree",
        )
        st.session_state["topics_to_audit"] = tree_state.get("checked", [])

    st.subheader("Add context")
    st.write("Write a short description of the model you're auditing so that the LLM understands what it's trying to capture.")
    model_info = st.text_area(
        "About this model: (optional)", 
        max_chars=1000, 
        value=audit_defaults["model_info"],
        help="Tell the LLM about anything unique to this model, or the feedback it targets, so that it can make informed decisions.",
        placeholder="This model captures feedback about AARP Rewards, a gamified loyalty platform ..."
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
    sidebar.subheader("API Settings")
    llm_provider_options = ["anthropic", "openai"]
    default_provider = app_defaults["llm_provider"]
    provider_index = (
        llm_provider_options.index(default_provider)
        if default_provider in llm_provider_options
        else 0
    )
    llm_provider = sidebar.selectbox(
        "LLM provider",
        options=llm_provider_options,
        index=provider_index,
    )

    default_model = (
        app_defaults["model_name_anthropic"]
        if llm_provider == "anthropic"
        else app_defaults["model_name_openai"]
    )
    model_name = sidebar.text_input("Model", value=default_model)
    
    api_key = get_api_key(llm_provider.upper())
    error_placeholder = sidebar.empty()
    key_expander_open = not bool(api_key)

    anthropic_api_key = None
    openai_api_key = None
    with sidebar.expander("API key", expanded=key_expander_open):
        if llm_provider == "anthropic":
            anthropic_api_key = st.text_input(
                "Anthropic API key",
                type="password",
                help="Uses ANTHROPIC_API_KEY from the environment if left blank.",
            )
        elif llm_provider == "openai":
            openai_api_key = st.text_input(
                "OpenAI API key",
                type="password",
                help="Uses OPENAI_API_KEY from the environment if left blank.",
            )
        if sidebar.button("Set API key"):
            if llm_provider == "anthropic" and anthropic_api_key:
                api_key = anthropic_api_key
            elif llm_provider == "openai" and openai_api_key:
                api_key = openai_api_key

    manual_key_present = False
    if llm_provider == "anthropic":
        manual_key_present = bool((anthropic_api_key or "").strip())
    else:
        manual_key_present = bool((openai_api_key or "").strip())
    if not (api_key or manual_key_present):
        error_placeholder.error(f"{llm_provider} API key not found; enter key below")
    else:
        error_placeholder.empty()

    sidebar.subheader("API limits")
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
        if not (api_key or anthropic_api_key):
            missing_reasons.append("Provide an API key to access LLM audit functionality")
    elif llm_provider == "openai":
        if not (api_key or openai_api_key):
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
                audit_warnings.append(
                    f"Max categories to audit is {int(max_categories)}, but the input has {total_categories} categories to audit."
                )

            if categories_to_audit:
                max_sentences_in_category = max(
                    category_counts[cat] for cat in categories_to_audit
                )
                if int(max_sentences) < max_sentences_in_category:
                    audit_warnings.append(
                        "Max sentences per category is "
                        f"{int(max_sentences)}, but the input has up to "
                        f"{max_sentences_in_category} sentences in a category."
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
        "Generate audit summary",
        value=True,
        key="generate_audit_summary",
        help="Generates a note summarizing the issues for each topic that failed the audit"
    )
    if not generate_summary:
        st.session_state["summary_generation_pending"] = False

    if generate_summary:
        with st.expander("Summary prompt"):
            st.write(
                "Generate a summary of the audit findings by having an LLM review the notes "
                "of all sentences found to be incorrectly categorized."
            )
            st.text_area(
                label="Instructions:",
                value=summary_prompt_default,
                max_chars=3000,
                placeholder="Tell the LLM how to summarize audit findings...",
                help="Sentences are batched by category.",
                key="summary_prompt",
            )

    if st.button("Run audit", type="primary", disabled=not can_run_audit, help=run_help):
        st.session_state["summary_generation_pending"] = False
        audit_bytes = st.session_state.get("reformatted_audit_bytes")

        if not audit_bytes:
            if uploaded_audit is None:
                st.error("Please upload an audit file before running the audit.")
                return
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

                def _update_progress(current, total, category_name):
                    progress_text.write(
                        f"Auditing category {current} of {total}: {category_name}"
                    )
                    progress_bar.progress(current / total if total else 0)

                # Use api_key if set, otherwise use the text input values
                final_anthropic_key = api_key if llm_provider == "anthropic" and api_key else (anthropic_api_key or None)
                final_openai_key = api_key if llm_provider == "openai" and api_key else (openai_api_key or None)
                
                output_bytes = run_audit(
                    audit_excel_bytes=audit_bytes,
                    prompt_template=audit_prompt,
                    llm_provider=llm_provider,
                    model_name=model_name,
                    model_info=model_info,
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
                )
                progress_bar.progress(1.0)
                progress_text.empty()
                progress_bar.empty()
            st.session_state["audit_output_bytes"] = output_bytes
            st.session_state["summary_generation_pending"] = generate_summary
            st.success("Audit complete.")
        except Exception as exc:
            st.error(f"Audit failed: {exc}")

    audit_output_bytes = st.session_state.get("audit_output_bytes")
    summary_prompt = st.session_state.get("summary_prompt", summary_prompt_default)

    summary_missing_reasons = []
    if not audit_output_bytes:
        summary_missing_reasons.append("Run the audit first")
    if llm_provider == "anthropic":
        if not (api_key or anthropic_api_key):
            summary_missing_reasons.append("Provide an Anthropic API key")
    elif llm_provider == "openai":
        if not (api_key or openai_api_key):
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
        try:
            with st.spinner("Summarizing audit..."):
                summarizer_module = _load_summarizer_module()
                final_anthropic_key = api_key if llm_provider == "anthropic" and api_key else (anthropic_api_key or None)
                final_openai_key = api_key if llm_provider == "openai" and api_key else (openai_api_key or None)
                progress_container = st.container()
                progress_text = progress_container.empty()
                progress_bar = progress_container.progress(0)

                def _update_summary_progress(current, total, category_name):
                    progress_text.write(
                        f"Summarizing category {current} of {total}: {category_name}"
                    )
                    progress_bar.progress(current / total if total else 0)

                summary_bytes = summarizer_module.summarize_audit_report(
                    audit_excel_input=audit_output_bytes,
                    msg_template=summary_prompt,
                    llm_provider=llm_provider,
                    model_name=model_name,
                    max_tokens=int(max_tokens),
                    model_info=model_info,
                    anthropic_api_key=final_anthropic_key,
                    openai_api_key=final_openai_key,
                    log_fn=st.write,
                    warn_fn=st.warning,
                    progress_fn=_update_summary_progress,
                )
                progress_bar.progress(1.0)
            st.session_state["audit_output_bytes"] = summary_bytes
            st.success("Audit summary complete.")
        except Exception as exc:
            st.error(f"Audit summary failed: {exc}")
        finally:
            st.session_state["summary_generation_pending"] = False
            if progress_text is not None:
                progress_text.empty()
            if progress_bar is not None:
                progress_bar.empty()
    elif st.session_state.get("summary_generation_pending") and summary_help:
        st.warning(f"Audit summary pending: {summary_help}")

    audit_output_bytes = st.session_state.get("audit_output_bytes")
    summary_pending = st.session_state.get("summary_generation_pending")
    if audit_output_bytes and not summary_pending:
        completed_filename = _build_completed_filename(uploaded_audit)
        st.download_button(
            label="Download completed audit (.xlsx)",
            data=audit_output_bytes,
            file_name=completed_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()


#TODO: pull together extra category info to send do the LLM, e.g. full category tree, category description

#TODO: send the sentences to the LLM to judge and return judgments

#TODO: Also have the LLM summarize the main issues by category and suggest rules improvements

#TODO: assemble audited sentences back into a spreadsheet

#TODO: set up the spreadsheet to report category accuracy

#TODO: output spreadsheet
