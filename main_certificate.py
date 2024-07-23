import os
import csv
from docx import Document
from docx2pdf import convert
from certificate import generate_qr_code, replace_participant_name, replace_event_name, replace_host_name, add_qr_code

# Create output folders if they don't exist
os.makedirs("Output/Doc", exist_ok=True)
os.makedirs("Output/PDF", exist_ok=True)
os.makedirs("Output/QR", exist_ok=True)

# Function to read participants data from CSV
def get_participants(csv_file):
    participants = []
    with open(csv_file, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            participants.append(row)
    return participants

# Function to generate certificates for participants
def generate_certificates(template_file, participants, event_name, host_name):
    for participant in participants:
        doc = Document(template_file)
        replace_participant_name(doc, participant["Full Name"])
        replace_event_name(doc, event_name)
        replace_host_name(doc, host_name)
        
        # Generate and add QR code
        qr_code_path = generate_qr_code(participant["Full Name"], event_name)
        add_qr_code(doc, qr_code_path)
        
        # Save the certificate as DOCX
        docx_filename = f'Output/Doc/{participant["Full Name"]}_certificate.docx'
        doc.save(docx_filename)
        
        # Convert DOCX to PDF
        pdf_filename = f'Output/PDF/{participant["Full Name"]}_certificate.pdf'
        print(f"Converting {docx_filename} to {pdf_filename}...")
        convert(docx_filename, pdf_filename)

# Paths and settings
certificate_template = "Templates/Certificate Template.docx"
participants_file = "Templates/Event Participants.csv"
event_name = "SQL Server Fundamentals Challenge"
host_name = "Md Nadeem Sarwar"


# Get participants data
participants = get_participants(participants_file)

# Generate certificates
generate_certificates(certificate_template, participants, event_name, host_name)