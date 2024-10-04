from PIL import Image, ImageDraw, ImageFont
import pytesseract
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import img2pdf

def extract_text_and_bboxes(image_path):
    img = Image.open(image_path)
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    return data

def get_font_color(image, bbox):
    region = image.crop(bbox)
    region = region.resize((1, 1), Image.LANCZOS)
    return region.getpixel((0, 0))

def draw_centered_text(draw,text,bbox,font,font_color):
    text_width, text_height = draw.textsize(text,font=font)
    x_centered= bbox[0] + (bbox[2] - bbox[0] - text_width) / 2
    y_centered = bbox[1] + (bbox[3] - bbox[1] - text_height) / 2
    draw.text((x_centered, y_centered), text, fill=font_color, font=font)



def replace_text_in_image(image_path, name, rank=None, position=None):
    # Extract text and bounding boxes from the image
    data = extract_text_and_bboxes(image_path)
    img = Image.open(image_path)
    draw = ImageDraw.Draw(img)
    font_path = "arial.ttf"  # Path to your font file

    # Initialize flags to check if placeholders are found
    rank_placeholder_found = False
    pos_placeholder_found = False

    # Loop through the extracted text data
    for i, text in enumerate(data['text']):
        # Check for the "#NAME" placeholder
        if "#NAME" in text:
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            bbox = (x, y, x + w, y + h)
            font_size = h
            font_color = get_font_color(img, bbox)
            font = ImageFont.truetype(font_path, font_size)
            draw.rectangle(bbox, fill="white")  # Clear the placeholder area
            draw.text((x, y), name, fill=font_color, font=font)  # Draw the name

        # Check for the "#Rank" placeholder
        if "#Rank" in text and rank:
            rank_placeholder_found = True
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            bbox = (x, y, x + w, y + h)
            font_size = h
            font_color = get_font_color(img, bbox)
            font = ImageFont.truetype(font_path, font_size)
            draw.rectangle(bbox, fill="white")  # Clear the placeholder area
            draw.text((x, y), rank, fill=font_color, font=font)  # Draw the rank

        # Check for the "#pos" placeholder
        if "#pos" in text and position:
            pos_placeholder_found = True
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            bbox = (x, y, x + w, y + h)
            font_size = h
            font_color = get_font_color(img, bbox)
            font = ImageFont.truetype(font_path, font_size)
            draw.rectangle(bbox, fill="white")  # Clear the placeholder area
            draw.text((x, y), position, fill=font_color, font=font)  # Draw the position

    # Return the modified image
    return img


    # Optionally only replace rank/position if their placeholders were found
    if not rank_placeholder_found:
        print(f"No #Rank placeholder found in {image_path}, skipping rank replacement.")
    if not pos_placeholder_found:
        print(f"No #pos placeholder found in {image_path}, skipping position replacement.")

    return img

def convert_image_to_pdf(image_path, pdf_path):
    with open(pdf_path, "wb") as f:
        f.write(img2pdf.convert(image_path))

def send_email_with_attachment(to_email, subject, body, attachment_path):
    from_email = "your_email@gmail.com"  # Replace with your email
    password = "your_app_specific_password"  # Replace with your app-specific password

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, "rb") as attachment:
        if attachment_path.endswith('.pdf'):
            part = MIMEApplication(attachment.read(), _subtype="pdf")
        else:
            part = MIMEImage(attachment.read())
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)
        server.quit()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

# File paths
image_path = r"./certificate_1.jpg"
output_dir = "./certificates/"
excel_file_path = r"./names.xlsx"

# Read the Excel file
df = pd.read_excel(excel_file_path)

# Ensure the output directory exists
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Iterate over the names and generate certificates
for index, row in df.iterrows():
    name = row['Name']
    email = row['Email']
    rank = row.get('Rank', None)  # Use None if rank is not available
    position = row.get('Position', None)  # Use None if position is not available
    
    new_img = replace_text_in_image(image_path, name, rank, position)
    
    # Save the new image with the recipient's name
    output_image_path = os.path.join(output_dir, f"{name}_certificate.png")
    new_img.save(output_image_path)
    
    # Convert image to PDF
    pdf_output_path = os.path.join(output_dir, f"{name}_certificate.pdf")
    convert_image_to_pdf(output_image_path, pdf_output_path)
    
    # Send email with PDF attachment
    subject = "Your Certificate"
    body = f"Dear {name},\n\nPlease find your certificate attached.\n\nBest regards,\nYour Company"
    send_email_with_attachment(email, subject, body, pdf_output_path)

print("All certificates have been generated and sent.")
