import os
import comtypes.client
from pptx import Presentation
import pandas as pd

# Replace the file path with the path to your CSV file
file_path = 'your_csv_file.csv'


def replace_text_in_slide(slide, old_text, new_text):
    # Replaces old_text with new_text in a slide
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            if old_text in shape.text:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)
                            
                            

def create_pdf_from_ppt(ppt_filename, pdf_filename):
    # Convert a PowerPoint file to PDF
    
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    
    print(f"Creating {pdf_filename} from {ppt_filename}...")

    deck = powerpoint.Presentations.Open(ppt_filename)
    deck.SaveAs(pdf_filename, 32)  # 32 for PDF format
    deck.Close()
    powerpoint.Quit()
    print("Successfully created PDF file!")
    
    
    
    
def process_names(names_list):
    # Process a list of names, creating a PDF for each using the appropriate template based on name length
    
    short_template_path = "short name.pptx"  
    long_template_path = "long name.pptx"  
    output_folder = "output_pdfs" 
    name_length_threshold=11
    
    # Creates the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for name in names_list:
        
        # Chooses the appropriate template based on the length of the name
        if len(name) >= name_length_threshold:
            if len(name) < 13 and ('I' in name.upper() or 'l' not in name):
                template_path = short_template_path
            else:
                template_path = long_template_path
        else:
            template_path = short_template_path

        # Load the PowerPoint template
        prs = Presentation(template_path)
        
        parts = name.split()

        # Check if the name has more than one space (more than two parts)
        if len(parts) > 2:
            print(f"Can't create for {name}: more than one space in the name.")
            continue 

        slide = prs.slides[0]
        replace_text_in_slide(slide, "[NAME]", name)

        abs_output_folder = os.path.abspath(output_folder)
        temp_ppt_path = os.path.join(abs_output_folder, f"{name}.pptx")
        pdf_path = os.path.join(abs_output_folder, f"{name}.pdf")
        prs.save(temp_ppt_path)

        # Convert to PDF
        create_pdf_from_ppt(temp_ppt_path, pdf_path)

        # Remove the temporary PowerPoint file
        os.remove(temp_ppt_path)
        
        
### RUNNER CODE ###

# Prints the current working directory
print("Current working directory:", os.getcwd())

# Asks the user to confirm the directory where the necessary files are located
confirmed_directory = input("Is this the correct directory? Y/N\n-->")

while confirmed_directory.upper() == "N":
    path = input("Please enter the correct directory:\n-->")
    try:
        os.chdir(path)
        print(f"Changed current directory to: {path}")
        confirmed_directory = 'Y'
    except Exception as e:
        print(f"Error changing directory: {e}")

csv_file_path = os.path.join(os.getcwd(), file_path)

# Check if the file exists
if not os.path.exists(csv_file_path):
    print(f"File not found: {csv_file_path}")
    exit()

# Reads the CSV file and converts the names column to a list
data = pd.read_csv(file_path)
temp_names = data['Name'].tolist()
names = [name.strip() for name in temp_names]


process_names(names)

