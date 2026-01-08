import os
import importlib.util
import streamlit as st
import xml.etree.ElementTree as ET
from streamlit_tree_select import tree_select
import yaml

from audit_reformat import handle_audit_reformat
from audit import run_audit


def _load_summary_prompt(prompts_path):
    try:
        with open(prompts_path, "r") as f:
            prompts = yaml.safe_load(f)
        return prompts["audit-report-summarizer"]["rewards_msg_template"]
    except Exception:
        return ""


def _load_summarizer_module():
    module_path = os.path.join(os.path.dirname(__file__), "audit-report-summarizer.py")
    spec = importlib.util.spec_from_file_location("audit_report_summarizer", module_path)
    if spec is None or spec.loader is None:
        raise ImportError("Unable to load audit-report-summarizer module.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

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
    # Looks for secret in secrets.toml
    if api_var in st.secrets:
        return st.secrets[api_var]
    return None

def main():
    # st.cache_data.clear()
    st.title("Automated Audit")
    st.write("This app uses an LLM to audit the accuracy of CX Designer models, and provides a summary of the findings.")

    st.subheader("Prepare audit file")
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
                    output, output_filename = handle_audit_reformat(uploaded_audit)
                    st.success("Reformat complete.")
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

    default_audit_prompt = """You are auditing the accuracy of a topic in a topic model that is based on deterministic search rules.
The following sentences come from AARP members and users.
{model_info}
The sentences have been tagged with the topic '{category}'.
If a description for this topic exists, it follows here: '{description}'.
For each sentence, return a binary judgment on whether the sentence belongs in the topic, and also a brief explanation of your reasoning.
Sentences can be tagged with multiple topics.
Sentences do not need to mention AARP to be considered relevant to the topic.

Sentences:
{sentences_text}

Respond in the strict format:
ID: [sentence_id] - Judgment: [YES/NO] - Reasoning: [brief explanation]"""

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
        
        st.write("Select nodes to be audited:")
        if model_data.get("model_name"):
            st.caption(f"{model_data['model_name']}")

        tree_state = tree_select(
            model_data["tree_nodes"],
            checked=st.session_state.get("topics_to_audit", []),
            key="topics_tree",
        )
        st.session_state["topics_to_audit"] = tree_state.get("checked", [])

    st.subheader("Audit settings")
    prompts_path = os.path.join(os.path.dirname(__file__), "prompts.yaml")
    model_info = ""
    model_info = st.text_area(
        "About this model: (optional)", 
        max_chars=1000, 
        help="Tell the LLM about anything unique to this model, or the feedback it targets, so that it can make informed decisions.",
        placeholder="This model captures feedback about AARP Rewards, a gamified loyalty platform that awards points for educational and entertainment activities. Users can exchange points for tangible rewards, including gift cards."
        )

    audit_prompt = st.text_area(
        label="Task instructions:",
        value=default_audit_prompt,
        max_chars=2500,
        placeholder="Tell the LLM what to do...",
        help="The prompt is sent to the LLM once per category. Use {category} to refer to the category name, and {description} to refer to the category's description.",
    )

    with st.expander("Advanced"):
        llm_provider = st.selectbox(
            "LLM provider",
            options=["anthropic", "openai"],
        )

        default_model = "claude-opus-4-5" if llm_provider == "anthropic" else "gpt-5-nano"
        model_name = st.text_input("Model name", value=default_model)

        max_categories = st.number_input(
            "Max categories to audit",
            min_value=1,
            value=1000,
            step=1,
        )
        max_sentences = st.number_input(
            "Max sentences per category",
            min_value=1,
            value=51,
            step=1,
        )
        max_tokens = st.number_input(
            "Max tokens per request",
            min_value=1,
            value=10000,
            step=100,
        )

        api_key = get_api_key(llm_provider.upper())
        if not api_key:
            st.error(f"{llm_provider} API key not found; enter key below")

        # Initialize API key variables
        anthropic_api_key = None
        openai_api_key = None

        with st.expander("API keys"):
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
            if st.button("Set API key"):
                if llm_provider == "anthropic" and anthropic_api_key:
                    api_key = anthropic_api_key
                elif llm_provider == "openai" and openai_api_key:
                    api_key = openai_api_key

    missing_reasons = []
    if not uploaded_audit:
        missing_reasons.append("Upload an audit file")
    if llm_provider == "anthropic":
        if not (api_key or anthropic_api_key):
            missing_reasons.append("Provide an Anthropic API key")
    elif llm_provider == "openai":
        if not (api_key or openai_api_key):
            missing_reasons.append("Provide an OpenAI API key")

    can_run_audit = not missing_reasons
    run_help = (
        None
        if can_run_audit
        else "; ".join(missing_reasons)
    )

    if st.button("Run audit", type="primary", disabled=not can_run_audit, help=run_help):
        audit_bytes = st.session_state.get("reformatted_audit_bytes")

        if not audit_bytes:
            if uploaded_audit is None:
                st.error("Please upload an audit file before running the audit.")
                return
            with st.spinner("Reformatting audit..."):
                output, _ = handle_audit_reformat(uploaded_audit)
                st.session_state["reformatted_audit_bytes"] = output.getvalue()
                audit_bytes = st.session_state["reformatted_audit_bytes"]

        if not model_tree and not st.session_state.get("topics_to_audit"):
            topics_to_audit = None
        else:
            topics_to_audit = st.session_state.get("topics_to_audit")

        try:
            with st.spinner("Running audit..."):
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
                )
            st.session_state["audit_output_bytes"] = output_bytes
            st.session_state.pop("audit_summary_bytes", None)
            st.success("Audit complete.")
        except Exception as exc:
            st.error(f"Audit failed: {exc}")

    audit_output_bytes = st.session_state.get("audit_output_bytes")
    if audit_output_bytes:
        completed_filename = _build_completed_filename(uploaded_audit)
        st.download_button(
            label="Download completed audit (.xlsx)",
            data=audit_output_bytes,
            file_name=completed_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.subheader("Generate Audit Summary")
    st.write("Generate a summary of the audit findings by having an LLM review the notes of all sentences found to be incorrectly categorized.")
    summary_prompt_default = _load_summary_prompt(prompts_path)
    summary_prompt = st.text_area(
        label="Instructions:",
        value=summary_prompt_default,
        max_chars=3000,
        placeholder="Tell the LLM how to summarize audit findings...",
        help="Sentences are batched by category.",
        key="summary_prompt",
    )

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

    if st.button("Run summary", type="primary", disabled=not can_run_summary, help=summary_help):
        try:
            with st.spinner("Summarizing audit..."):
                summarizer_module = _load_summarizer_module()
                final_anthropic_key = api_key if llm_provider == "anthropic" and api_key else (anthropic_api_key or None)
                final_openai_key = api_key if llm_provider == "openai" and api_key else (openai_api_key or None)
                summary_bytes = summarizer_module.summarize_audit_report(
                    audit_excel_input=audit_output_bytes,
                    msg_template=summary_prompt,
                    llm_provider=llm_provider,
                    model_name=model_name,
                    max_tokens=int(max_tokens),
                    anthropic_api_key=final_anthropic_key,
                    openai_api_key=final_openai_key,
                    log_fn=st.write,
                )
            st.session_state["audit_summary_bytes"] = summary_bytes
            st.success("Audit summary complete.")
        except Exception as exc:
            st.error(f"Audit summary failed: {exc}")

    audit_summary_bytes = st.session_state.get("audit_summary_bytes")
    if audit_summary_bytes:
        st.download_button(
            label="Download audit summary (.xlsx)",
            data=audit_summary_bytes,
            file_name="audit_summary.xlsx",
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
