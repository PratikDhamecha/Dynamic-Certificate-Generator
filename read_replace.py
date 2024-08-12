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

def replace_text_in_image(image_path, replacement, placeholder="#NAME"):
    img = Image.open(image_path)
    draw = ImageDraw.Draw(img)
    font_path = "arial.ttf"
    data = extract_text_and_bboxes(image_path)
    
    for i, text in enumerate(data['text']):
        if placeholder in text:
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            bbox = (x, y, x + w, y + h)
            font_size = h
            font_color = get_font_color(img, bbox)
            font = ImageFont.truetype(font_path, font_size)
            draw.rectangle(bbox, fill="white")
            draw.text((x, y), replacement, fill=font_color, font=font)
    
    return img

def convert_image_to_pdf(image_path, pdf_path):
    with open(pdf_path, "wb") as f:
        f.write(img2pdf.convert(image_path))

def send_email_with_attachment(to_email, subject, body, attachment_path):
    from_email = "dhamechapratik5@gmail.com"  # Replace with your email
    password = "botogyryhxtailtv"  # Replace with your app-specific password

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
    new_img = replace_text_in_image(image_path, name)
    
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
