import os
import shutil

# Source and destination directories
pdf_directory = '/certificates'
processed_directory = '/sended'  # Replace with the path where you want to move processed files

# Create the processed directory if it doesn't exist
if not os.path.exists(processed_directory):
    os.makedirs(processed_directory)

# List of filenames extracted from your provided log
file_names = [
    "Placeholder@gmail.com_certificate.pdf"
    
    
]

# Move the files
for file_name in file_names:
    source_path = os.path.join(pdf_directory, file_name)
    destination_path = os.path.join(processed_directory, file_name)

    # Check if file exists before moving
    if os.path.exists(source_path):
        shutil.move(source_path, destination_path)
        print(f"Moved {file_name}")
    else:
        print(f"File not found: {file_name}")

print("All specified files have been moved.")
