from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from flask import Flask, jsonify, render_template, request, send_file
import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pytesseract

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')  # Frontend template

def extract_text_and_bboxes(image_path):
    """Extract text and bounding boxes from the image."""
    img = Image.open(image_path)
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    return data

def get_font_color(image, bbox):
    """Get the font color in the specified bounding box."""
    region = image.crop(bbox)
    region = region.resize((1, 1), Image.LANCZOS)
    return region.getpixel((0, 0))

def draw_centered_text(draw, text, bbox, font, font_color):
    """Draw centered text in the specified bounding box."""
    x1, y1, x2, y2 = draw.textbbox((0, 0), text, font=font)
    text_width = x2 - x1
    text_height = y2 - y1
    
    x_centered = bbox[0] + (bbox[2] - bbox[0] - text_width) / 2
    y_centered = bbox[1] + (bbox[3] - bbox[1] - text_height) / 2
    
    draw.text((x_centered, y_centered), text, fill=font_color, font=font)

def replace_text_in_image(image_file: str, name: str = None, rank: str = None, position: str = None):
    """Replace placeholders in the certificate image with actual values."""
    img = Image.open(image_file)

    # Check orientation and rotate if necessary
    if img.width < img.height:  # Vertical orientation
        img = img.rotate(90, expand=True)

    data = extract_text_and_bboxes(image_file)
    draw = ImageDraw.Draw(img)
    font_path = "arial.ttf"  # Change this to the path of your font file

    for i, text in enumerate(data['text']):
        x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
        bbox = (x, y, x + w, y + h)
        font_size = h
        font_color = get_font_color(img, bbox)
        font = ImageFont.truetype(font_path, font_size)

        if "#NAME" in text and name:
            draw.rectangle(bbox, fill="white")  # Clear area before drawing new text
            draw_centered_text(draw, name, bbox, font, font_color)

        if "#Rank" in text and rank:
            draw.rectangle(bbox, fill="white")
            draw_centered_text(draw, rank, bbox, font, font_color)

        if "#pos" in text and position:
            draw.rectangle(bbox, fill="white")
            draw_centered_text(draw, position, bbox, font, font_color)

    return img

def send_email_with_attachment(to_address: str, attachment_path: str, subject: str = "Your Certificate", body: str = "Dear [Name],\n\nPlease find your certificate attached.\n\nBest regards,\nYour Company"):
    """Send an email with an attachment."""
    default_from_address = os.getenv("DEFAULT_EMAIL", "dhamechapratik5@gmail.com")
    default_password = os.getenv("DEFAULT_PASSWORD", "botogyryhxtailtv")

    msg = MIMEMultipart()
    msg['From'] = default_from_address
    msg['To'] = to_address
    msg['Subject'] = subject

    msg.attach(MIMEText(body.replace("[Name]", to_address.split('@')[0]), 'plain'))  # Replace placeholder with actual name

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
    msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(default_from_address , default_password)
        server.sendmail(default_from_address , to_address , msg.as_string())
        server.quit()
        print(f"Email sent to {to_address} successfully.")
    except Exception as e:
        print(f"Failed to send email. Error: {str(e)}")

def convert_image_to_pdf(image_path: str , output_pdf_path: str):
    """Convert an image to a PDF file."""
    img = Image.open(image_path)
    img.convert('RGB').save(output_pdf_path)

@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads and generate certificates."""
    
    image_file = request.files['image']
    excel_file = request.files['excel']
    
    # Optional subject and body inputs from user
    subject_input = request.form.get('subject', "")
    body_input = request.form.get('body', "")

    # Save the uploaded files
    image_path = os.path.join('static', image_file.filename)
    excel_file_path = os.path.join('static', excel_file.filename)

    image_file.save(image_path)
    excel_file.save(excel_file_path)

    # Read the Excel file
    df = pd.read_excel(excel_file_path)

    output_dir = "./static/certificates/"

    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    results = []

    # Iterate over each row in the Excel sheet
    for index, row in df.iterrows():
        name = row['Name']
        email = row['Email']
        rank = row.get('Rank', None)
        position = row.get('Position', None)

        new_img = replace_text_in_image(image_path,
                                         name=name if '#NAME' in extract_text_and_bboxes(image_path)['text'] else None,
                                         rank=rank if '#Rank' in extract_text_and_bboxes(image_path)['text'] else None,
                                         position=position if '#pos' in extract_text_and_bboxes(image_path)['text'] else None)

        # Save the new image as a PNG
        output_image_path = os.path.join(output_dir, f"{name}_certificate.png")
        new_img.save(output_image_path)

        # Convert the image to PDF
        pdf_output_path = os.path.join(output_dir, f"{name}_certificate.pdf")
        convert_image_to_pdf(output_image_path, pdf_output_path)

        # Send email with the PDF attached using user-defined or default subject and body
        send_email_with_attachment(email, 
                                    pdf_output_path,
                                    subject_input if subject_input else "Your Certificate", 
                                    body_input if body_input else f"Dear {name},\n\nPlease find your certificate attached.\n\nBest regards,\nYour Company")

        results.append({
            "name": name,
            "email": email,
            "certificate": pdf_output_path
        })

    return jsonify({"message": "Certificate generated and emails sent.", "results": results})

@app.route('/download/<path:filename>' , methods=['GET'])
def download_file(filename):
   """Download a generated certificate."""
   output_dir = "./static/certificates/"
   return send_file(os.path.join(output_dir , filename) , as_attachment=True)

if __name__ == '__main__':
   app.run(debug=True)