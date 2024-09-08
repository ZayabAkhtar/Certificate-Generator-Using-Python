import xlrd
import os
from PIL import Image, ImageDraw, ImageFont

# Define the path for the certificates folder
certificates_folder = os.path.join(os.getcwd(), 'certificates')

# Create the certificates folder if it doesn't exist
if not os.path.exists(certificates_folder):
    os.makedirs(certificates_folder)

# Load the Excel workbook
path = "participation.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)
rows = inputWorksheet.nrows

user = []
objects = {}

# Collect user data from the Excel sheet
for i in range(rows):
    objects['email'] = inputWorksheet.cell_value(i, 2)
    objects['name'] = inputWorksheet.cell_value(i, 1)
    objects['team'] = inputWorksheet.cell_value(i, 0)
    user.append(objects)
    objects = {}

# Define the path for the template image and font
image_path = r"D:\Zayab Akhtar\Certificate-Generator-\participation.jpg"
font_path = r"D:\Zayab Akhtar\Certificate-Generator-\Certificate.ttf"  # Update with your font file name

# Set the desired font size
font_size = 50  # Increase this value for larger text

# Generate certificates
for person in user:
    try:
        team = person['team']
        name = person['name']
        
        # Open the image and load the specified font
        image = Image.open(image_path)
        font = ImageFont.truetype(font_path, size=font_size)
        draw = ImageDraw.Draw(image)

        # Calculate the width of the name and team text
        name_width = draw.textsize(name, font=font)[0]
        team_width = draw.textsize(team, font=font)[0]

        # Get image dimensions
        image_width, image_height = image.size

        # Center align the name
        name_x = (image_width - name_width) / 2
        name_y = 525  # Adjust this as necessary

        # Center align the team
        team_x = (image_width - team_width) / 2
        team_y = 740  # Adjust this as necessary

        # Draw the text on the image
        draw.text((name_x, name_y), text=name, fill=(0, 0, 0), font=font)
        draw.text((team_x, team_y), text=team, fill=(0, 0, 0), font=font)
        
        # Save the certificate as a PDF in the certificates folder
        pdf_path = os.path.join(certificates_folder, f'{team}_{name}.pdf')
        image.save(pdf_path, 'PDF', resolution=100.0)
        print(f'Generated certificate for {name} of {team} at {pdf_path}')
        
    except Exception as e:
        print(e.__str__(), f'Could not generate certificate for {name}')

print("Certificate generation completed.")