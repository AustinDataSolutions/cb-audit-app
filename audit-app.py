import streamlit as st
import os
import xml.etree.ElementTree as ET
import pandas as pd
from streamlit_tree_select import tree_select

from audit_reformat import handle_audit_reformat

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

    def build_node(node_element, parent_path_parts):
        node_id = _get_node_field(node_element, "id")
        node_name = _get_node_field(node_element, "name")
        node_description = _get_node_field(node_element, "description")
        node_order_number = _get_node_field(node_element, "order-number")
        node_smart_other = _get_node_field(node_element, "smart-other")

        current_path_parts = parent_path_parts + ([node_id] if node_id else [])
        node_path = "/".join([part for part in current_path_parts if part])

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
            "value": node_id or node_path or node_name or "unknown-node",
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
    for node in tree_element.findall("node"):
        _, tree_node = build_node(node, [])
        tree_nodes.append(tree_node)

    model_data = {
        "model_name": model_name,
        "nodes": nodes_by_id,
        "tree_nodes": tree_nodes,
    }

    return model_data

def main():
    # st.cache_data.clear()
    st.title("Enhanced Audit Script")
    st.write("Run an audit using an LLM to determine accuracy.")

    uploaded_audit = st.file_uploader(
        "Upload raw audit file (.xlsx)",
        type=["xlsx"],
    )

    if st.button("Process", type="primary"):
        if uploaded_audit is None:
            st.error("Please upload an audit file before processing.")
            return
        with st.spinner("Reformatting audit..."):
            output = handle_audit_reformat(uploaded_audit)
            st.success("Reformat complete.")
            st.download_button(
                label="Download reformatted audit",
                data=output.getvalue(),
                file_name="reformatted_audit.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    model_tree = st.file_uploader(
        "Upload model tree XML to select particular topics to audit, and provide LLM with category descriptions.",
        type=["xml"],
    )

    if model_tree:
        if st.session_state.get("model_source_name") != model_tree.name:
            try:
                model_data = _parse_model_xml(model_tree.getvalue())
            except (ET.ParseError, ValueError) as exc:
                st.error(f"Unable to parse model XML: {exc}")
                return
            st.session_state["model_data"] = model_data
            st.session_state["model_source_name"] = model_tree.name
            st.session_state["topics_to_audit"] = list(model_data["nodes"].keys())

        model_data = st.session_state.get("model_data")
        if not model_data:
            st.error("Model data is missing. Please re-upload the XML file.")
            return

        st.write("Select nodes to be audited:")
        tree_state = tree_select(
            model_data["tree_nodes"],
            checked=st.session_state.get("topics_to_audit", []),
            key="topics_tree",
        )
        st.session_state["topics_to_audit"] = tree_state.get("checked", [])

        default_audit_prompt = """You are auditing the accuracy of a topic in a topic model that is based on deterministic search rules. 
The following sentences come from AARP members and users.
The sentences have been tagged with the topic '{category}'.
If a description for this topic exists, it follows here: '{description}'.
Sentences can be tagged with multiple topics.
For each sentence, return a binary judgment on whether the sentence belongs in the topic, and also a brief explanation of your reasoning.
Sentences do not need to mention AARP to be considered relevant to the topic.

Sentences:
{sentences_text}

Respond in the strict format:
ID: [sentence_id] - Judgment: [YES/NO] - Reasoning: [brief explanation]"""

        audit_prompt = st.text_area(
            label="Set audit prompt.", value=default_audit_prompt, 
            max_chars=2500,
            placeholder="Tell the LLM what to do...",
            help="The prompt is sent to the LLM once per category. Use {category} to refer to the category name, and {description} to refer to the category's optional description"
            )

if __name__ == "__main__":
    main()


#TODO: pull together extra category info to send do the LLM, e.g. full category tree, category description

#TODO: send the sentences to the LLM to judge and return judgments

#TODO: Also have the LLM summarize the main issues by category and suggest rules improvements

#TODO: assemble audited sentences back into a spreadsheet

#TODO: set up the spreadsheet to report category accuracy

#TODO: output spreadsheet
