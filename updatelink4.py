import os
import requests
import PyPDF2
from docx import Document
import csv

# Function to extract links from a PDF file
def extract_links_from_pdf(pdf_file):
    links = []
    with open(pdf_file, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            if "/Annots" in page:
                for annotation in page["/Annots"]:
                    annotation_obj = annotation.get_object()  # Updated to get_object
                    if "/A" in annotation_obj and "/URI" in annotation_obj["/A"]:
                        links.append(annotation_obj["/A"]["/URI"])
    return links

# Function to extract links from a Word (.docx) file
def extract_links_from_word(docx_file):
    links = []
    doc = Document(docx_file)
    for para in doc.paragraphs:
        for run in para.runs:
            if run.hyperlink:
                links.append(run.hyperlink.target)
    return links

# Function to check if a URL is valid (not a dead link)
def check_link(url):
    try:
        response = requests.head(url, allow_redirects=True, timeout=5)
        return response.status_code == 200
    except requests.exceptions.RequestException:
        return False

# Function to process all files in a folder
def check_links_in_folder(folder):
    dead_links = {}

    for root, dirs, files in os.walk(folder):
        for file in files:
            file_path = os.path.join(root, file)
            links = []

            try:
                if file.endswith('.pdf'):
                    links = extract_links_from_pdf(file_path)
                elif file.endswith('.docx'):
                    links = extract_links_from_word(file_path)
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
                continue  # Skip to the next file

            for link in links:
                if not check_link(link):
                    if file not in dead_links:
                        dead_links[file] = []
                    dead_links[file].append(link)

    return dead_links

# Folder path to check
folder_path = "/Users/contento/funCodeProjects"

# Check for dead links
dead_links_report = check_links_in_folder(folder_path)

# Save report to CSV
csv_file = "dead_links_report.csv"
with open(csv_file, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["File", "Dead Link"])

    for file, links in dead_links_report.items():
        for link in links:
            writer.writerow([file, link])

# Print a message to indicate the report was saved
if dead_links_report:
    print(f"Dead links found and saved to {csv_file}.")
else:
    print("No dead links found!")

