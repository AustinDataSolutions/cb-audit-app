import anthropic
# from openai import OpenAI
from dotenv import load_dotenv
import os
import pandas as pd
import re
import yaml
from openpyxl import Workbook, load_workbook
from datetime import datetime
import xml.etree.ElementTree as ET

'''
Accepts an audit file (xlsx) formatted in the CB Toolbox audit reformat output format.
Using the prompts.yaml file, sends batches of sentences to an LLM along with assigned topic
 for the LLM to judge accuracy.
LLM returns yes/no judgment along with an explanation of its reasoning.

Written by JT in October 2025
'''

#help(anthropic)
#help(anthropic.Anthropic)

print("Starting script...")

script_dir = os.path.dirname(os.path.abspath(__file__))
prompts_path = os.path.join(script_dir, 'prompts.yaml')
inputs_dir = os.path.join(script_dir, "inputs")

import_model_tree = False

try:
    with open(prompts_path, 'r') as f:
        prompts = yaml.safe_load(f)
        audit_config = prompts['rewards_model_audit'] #Change this as needed
        max_categories = audit_config['max_categories']
        max_sentences_per_category = audit_config['max_sentences_per_category']
        msg_template = audit_config['msg_template']
        audit_file_name = audit_config['audit_file']
        model_tree_file = audit_config['model_tree']
        audit_in_progress = audit_config.get('audit_in_progress')
        if model_tree_file:
            import_model_tree = True
        llm_provider = audit_config['llm_provider']
        if llm_provider not in ["anthropic", "openai"]:
            print("Error: llm_provider not properly set in config file. Use 'anthropic' or 'openai'.")
            exit(1)
except FileNotFoundError:
    print("Error: prompts.yaml not found")
    exit(1)
except KeyError as e:
    print(f"Error: Missing key in YAML: {e}")
    exit(1)


load_dotenv() #reads the .env file and adds it to environ via something like os.environ['ANTHROPIC_API_KEY'] = 'sk-ant-123...'

anthropic_key = os.getenv('ANTHROPIC_API_KEY')
if anthropic_key is None:
    raise RuntimeError("ANTHROPIC_API_KEY environment variable not set")
openai_api_key = os.getenv("OPENAI_API_KEY")
if openai_api_key is None:
    raise RuntimeError("OPENAI_API_KEY environment variable not set")

print("Retrieved API key")


if llm_provider == 'anthropic':
    client = anthropic.Anthropic(api_key=anthropic_key)
elif llm_provider == 'openai':
    client = OpenAI(api_key=openai_api_key)

# dir(client)

print("Retrieving audit file")
#Load audit file to be processed
excel_path = os.path.join(inputs_dir, audit_file_name)
df = pd.read_excel(excel_path)

category_sentences = {}

# Assuming column A is Sentence ID, column B is 'Sentence', and column C is 'Category'
id_col = df.columns[0]        # Column A (0-based index)
sentence_col = df.columns[1]  # Column B (0-based index)
category_col = df.columns[2]  # Column C (0-based index)

# Check if the first row is completely blank and prompt user
first_row = df.iloc[0]
if all(pd.isna(val) for val in first_row):
    user_input = input(
        "WARNING: The first row in the audit file appears to be completely blank.\n"
        "Did you remember to reformat the audit file as required?\n"
        "Press Enter to continue or type 'exit' to quit: "
    )
    if user_input.strip().lower() == 'exit':
        print("Exiting script. Please reformat the audit file and try again.")
        exit(1)

for _, row in df.iterrows():
    category = row[category_col]
    sentence_id = row[id_col]
    sentence = row[sentence_col]
    if pd.isna(category) or pd.isna(sentence) or pd.isna(sentence_id):
        continue  # Skip rows with missing category, sentence, or sentence ID
    if category not in category_sentences:
        category_sentences[category] = {}
    category_sentences[category][sentence_id] = sentence

#Load model tree if provided
category_descriptions = {}
if import_model_tree:
    print("Model tree found, logging category names and descriptions")
    model_tree_path = os.path.join(inputs_dir, model_tree_file)
    model_level = 0
    try:
        tree = ET.parse(model_tree_path)
        root = tree.getroot()
        tree_elem = root.find('tree')
        root_node = tree_elem[0] if tree_elem is not None and len(tree_elem) > 0 else None
        if root_node is not None:
            pass
        else:
            print("No <tree> element found.")
    except Exception as e:
        print(f"Error reading XML tree: {e}")

    root_name = root_node.get('name')

    def get_all_child_names_and_descriptions(parent, path_so_far=None):
        if path_so_far is None:
            path_so_far = []
        cat_name = parent.get('name')
        if cat_name is None:
            return
        #Strip root name from path as it is not included in the audit file and search will therefore fail
        if len(path_so_far) > 0:
            if path_so_far[0] == root_name:
                path_so_far = path_so_far[1:]
        updated_path = path_so_far + [cat_name]
        full_cat_path = "-->".join(updated_path)
        category_desc = parent.get('description')
        if not category_desc:
            category_desc = "None" #String for inclusion in prompt
        category_descriptions[full_cat_path] = category_desc
        for child in parent:
            get_all_child_names_and_descriptions(child, updated_path)

    try:
        get_all_child_names_and_descriptions(root_node)
    except Exception as e:
        print(f"Error processing XML tree: {e}")

outputs_dir = os.path.join(os.path.dirname(__file__), "outputs")
if not os.path.exists(outputs_dir):
    os.makedirs(outputs_dir)

#Create or resume output workbook
timestamp = datetime.now().strftime("%y%m%d%H%M")
resume_mode = False
completed_categories = set()
restart_category = None

if audit_in_progress:
    in_progress_path = os.path.join(outputs_dir, audit_in_progress)
    if os.path.exists(in_progress_path):
        print(f"Resuming audit from in-progress file: {audit_in_progress}")
        output_path = in_progress_path
        resume_mode = True
    else:
        print(f"Provided audit_in_progress file '{audit_in_progress}' not found. Starting a new audit file.")

if not resume_mode:
    output_filename = f"completed_audit_{timestamp}.xlsx"
    output_path = os.path.join(outputs_dir, output_filename)

if resume_mode:
    wb = load_workbook(output_path)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    # Write header
    ws.append(["Sentence ID", "Sentence", "Category", "NLP Judgment", "Explanation"])

# --- Add category summary worksheet ---
if "categories" not in wb.sheetnames:
    ws_categories = wb.create_sheet(title="categories")
    ws_categories.append(["Category", "Description", "Precision Rate", "Finding", "Recommendation"])
else:
    ws_categories = wb["categories"]

if resume_mode:
    existing_categories = [
        row[0]
        for row in ws_categories.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True)
        if row[0]
    ]
    if existing_categories:
        restart_category = existing_categories[-1]
        completed_categories = set(existing_categories[:-1])
        print(f"Last completed category recorded as '{restart_category}'. Re-auditing it before continuing.")

    # Remove existing rows for restart_category so results can be rewritten cleanly
    if restart_category:
        for row_idx in range(ws.max_row, 1, -1):
            if ws.cell(row=row_idx, column=3).value == restart_category:
                ws.delete_rows(row_idx)
        for row_idx in range(ws_categories.max_row, 1, -1):
            if ws_categories.cell(row=row_idx, column=1).value == restart_category:
                ws_categories.delete_rows(row_idx)

# Build ordered list of categories to audit
categories_to_audit = []
if restart_category:
    categories_to_audit.append(restart_category)
for category in category_sentences:
    if resume_mode and category in completed_categories:
        continue
    if restart_category and category == restart_category:
        continue
    categories_to_audit.append(category)

#send batches to LLM
cat_count = 0
for category in categories_to_audit:
    cat_count += 1
    print(f"Auditing category {cat_count}, {category}")
    if cat_count >= max_categories:
        print("Reached max iteration")
        break

    description = category_descriptions.get(category, "None")
    if not description:
        description = "None"
    else:
        print(f"Category description found: {description}") #temp for debug

    # Prepare the sentences for this category
    sent_tuples = list(category_sentences[category].items())
    
    # Format sentences for the prompt
    sentences_text = ""
    sent_count = 0
    for sentence_id, sentence in sent_tuples:
        sent_count += 1
        if sent_count >= (max_sentences_per_category + 1):
            break
        sentences_text += f"ID: {sentence_id} - {sentence}\n"
    
    # Create the message content
    message_content = msg_template.format(
    category=category,
    description=description,
    sentences_text=sentences_text
)
    
    print(f"Sending message to LLM for category {category}...")
    #Send message to LLM:


    if llm_provider == 'anthropic':
        message = client.messages.create(
            model="claude-opus-4-5", #"claude-sonnet-4-0",
            max_tokens=10000,
            messages=[
                {"role": "user", "content": message_content}
            ]
        )

        # Parse the LLM response
        response_text = message.content[0].text

    elif llm_provider == 'openai':
        response = client.chat.completions.create(
            model="gpt-5-nano",
            messages=[
                {"role": "user", "content": message_content}
            ]
        )

        response_text = response.choices[0].message.content

    else:
        print("ERROR: provided llm_provider is neither 'openai' nor 'anthropic'")
        exit(1)

    print(f"Received response for {category}. Preview:")
    print(response_text[:50]) #Raw text #Print start of response for auditing purposes

    # Regex to extract: ID: [sentence_id] - Judgment: [YES/NO] - Reasoning: [brief explanation]
    pattern = r"ID:\s*(.+?)\s*-\s*Judgment:\s*(YES|NO)\s*-\s*Reasoning:\s*(.+)"
    matches = re.findall(pattern, response_text, re.IGNORECASE)

    # Build a mapping from sentence_id to (judgment, explanation)
    nlp_results = {}
    for match in matches:
        sent_id = str(match[0]).strip()
        judgment = match[1].strip().upper()
        explanation = match[2].strip()
        nlp_results[sent_id] = (judgment, explanation)

    # Write results to XLSX
    for sentence_id, sentence in sent_tuples:
        judgment, explanation = nlp_results.get(str(sentence_id), ("", ""))
        ws.append([sentence_id, sentence, category, judgment, explanation])

    ws_categories.append([category]) #TODO: Add values for the remaining columns ()"Description", "Precision Rate", "Finding", "Recommendation")
    wb.save(output_path)

wb.close()

# Catch response and write it to XLSX, where NLP's judgment is inserted into column d, and explanation is inserted into column E

# Save and close XLSX

print("Script concluded")
