import os
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook
import yaml
import anthropic
from dotenv import load_dotenv
import re

#This script takes an XLSX input of a completed audit by an LLM containing a reason for the audit decision
# and returns a summary of the found issues by category
# as well as a suggestion of what rules probably need to be tweaked for each category

#TODO: Consider adding count of sentences checked and sentences found wrong to output, for greater transparancy around accuracy
# - e.g. when LLM says "All 28 comments incorrectly categorized" raises questions - it's proba

print("Starting script...")

script_dir = os.path.dirname(os.path.abspath(__file__))
prompts_path = os.path.join(script_dir, 'prompts.yaml')
inputs_dir = os.path.join(script_dir, "inputs")
outputs_dir = os.path.join(script_dir, "outputs")

try:
    with open(prompts_path, 'r') as f:
        prompts = yaml.safe_load(f)
        summarizer_config = prompts['audit-report-summarizer']
        msg_template = summarizer_config['rewards_msg_template']
        audit_file_name = summarizer_config['audit_file']
except FileNotFoundError:
    print("Error: prompts.yaml not found")
    exit(1)
except KeyError as e:
    print(f"Error: Missing key in YAML: {e}")
    exit(1)

load_dotenv()

anthropic_key = os.getenv('ANTHROPIC_API_KEY')

client = anthropic.Anthropic(api_key=anthropic_key)

#Load completed audit file to be processed
excel_path = os.path.join(inputs_dir, audit_file_name)
df = pd.read_excel(excel_path)


audit_findings = {} #turning the XLSX findings into a dict

# Assuming column A is Sentence ID, column B is 'Sentence', and column C is 'Category'
#TODO: Validate column inputs
id_col = df.columns[0]        # Column A (0-based index)
sentence_col = df.columns[1]
category_col = df.columns[2]
judgment_col = df.columns[3]
explan_col = df.columns[4]

#Collect audit findings by row and add them to audit_findings dict
#TODO: Validate if this approach also captures the header row; fix if so
for _, row in df.iterrows():
    category = row[category_col]
    sentence_id = row[id_col]
    sentence = row[sentence_col]
    judgment = row[judgment_col]
    explanation = row[explan_col]
    if pd.isna(category) or pd.isna(sentence) or pd.isna(sentence_id) or pd.isna(judgment) or pd.isna(explanation):
        print("Skipped row during ingestion")
        continue  # Skip rows with missing category, sentence, or sentence ID
    if category not in audit_findings:
        audit_findings[category] = {}
    audit_findings[category][sentence_id] = (judgment, explanation)

#Create output workbook
timestamp = datetime.now().strftime("%y%m%d%H%M")
output_filename = f"audit_summary_{timestamp}.xlsx"
output_path = os.path.join(outputs_dir, output_filename)
if not os.path.exists(outputs_dir):
    os.makedirs(outputs_dir)

# Try to load the workbook if it exists, otherwise create a new one
if os.path.exists(output_path):
    wb = load_workbook(output_path)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    # Write header
    ws.append(["Category", "Accuracy", "Issues", "Recommendation"])

#send audit explanation comments to LLM for summarization
sentence_findings = {}
percent_wrong = 0
categories_checked = 1
total_categories = len(audit_findings)
categories_processed = []  
error_count = 0

for category in audit_findings:
    try:
        print(f"Reviewing audit findings for category '{category}' ({categories_checked} of {total_categories})...")   
        inaccurate_sent_explanations = "" 
        sentcount = 0
        wrong_count = 0

        #collect the explanation of sentence issues for sentence found to be inappropriately categorized
        for sentence_id, (judgment, explanation) in audit_findings[category].items():
            sentcount += 1
            if judgment == "NO":
                wrong_count += 1
                # sentence_findings[judgment] = sentence[0]
                inaccurate_sent_explanations += f"{explanation}\n"
        
        accuracy = round(((sentcount-wrong_count)/sentcount),2)
        print(f"Detected {wrong_count} explanations out of {sentcount} sentences audited ({round(accuracy*100)}% accuracy)")

        if accuracy < 0.80:
        
            # Prep message for LLM: Combine explanations with prompt (msg_template) and category
            message_content = msg_template.format(
            category=category,
            # description=description,
            inaccurate_sent_explanations=inaccurate_sent_explanations
            )
            
            print(f"Sending explanations to LLM for summarization...")
            #Send message to LLM:
            message = client.messages.create(
                model="claude-sonnet-4-5", #claude-haiku-4-5 is cheaper
                max_tokens=10000,
                messages=[
                    {"role": "user", "content": message_content}
                ]
            )

            # Parse the LLM response
            response_text = message.content[0].text

            print("=" * 50)
            print("FULL RESPONSE:")
            print(response_text)  # Print the ENTIRE response
            print("=" * 50)

            # Regex to extract: SUMMARY: [LLM summary]; RECOMMENDATION: [LLM recommendation]
            pattern = r"SUMMARY:\s*(.+?)\s*RECOMMENDATION:\s*(.+)"
            matches = re.findall(pattern, response_text, re.IGNORECASE | re.DOTALL)

            if matches:
                summary, recommendation = matches[0]
                summary = summary.strip()
                recommendation = recommendation.strip()
            else:
                print("WARNING: REGEX FAILED TO PARSE LLM RESPONSE")
                print(f"Attempting fallback parsing...")
                
                # Fallback: try to extract each section separately
                summary_match = re.search(r"SUMMARY:\s*(.+?)(?=RECOMMENDATION:|$)", response_text, re.IGNORECASE | re.DOTALL)
                rec_match = re.search(r"RECOMMENDATION:\s*(.+?)$", response_text, re.IGNORECASE | re.DOTALL)
                
                if summary_match and rec_match:
                    summary = summary_match.group(1).strip()
                    recommendation = rec_match.group(1).strip()
                    print("Fallback parsing successful!")
                else:
                    summary = "REGEX FAILED TO PARSE LLM RESPONSE"
                    recommendation = response_text  # Store full response for manual review

            ws.append([category, accuracy, summary, recommendation])

        else:
            summary = ""
            recommendation = ""
            print("Category acuracy is high enough; moving on to next category")
            ws.append([category, accuracy, summary, recommendation])

        categories_processed.append(category)

    except Exception as e:
        print(f"ERROR processing category '{category}': {e}")
        error_count += 1
        continue  # Skip this category and continue with the next
    
    categories_checked += 1

    wb.save(output_path)

wb.close()

print("Script concluded")
print(f"Workbook saved to {output_path}")
print(f"Ecountered {error_count} errors")

# Identify missing categories
missing_categories = set(audit_findings.keys()) - set(categories_processed)
print(f"\nProcessed {len(categories_processed)} out of {total_categories} categories")
if missing_categories:
    print(f"Missing categories: {missing_categories}")