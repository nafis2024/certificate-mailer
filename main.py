import csv, os, json, smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageDraw, ImageFont

def load_config():
    with open("config.json") as f:
        return json.load(f)

def write_text_on_image(name):
    template = "certificates/template.png"
    output_path = f"output/{name}.png"
    image = Image.open(template)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("arial.ttf", 80)
    text = name
    img_width, img_height = image.size
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    x = (img_width - text_width) / 2
    y = 600
    draw.text((x, y), text, font=font, fill="black")
    image.save(output_path)
    print(f"✅ Certificate created for {name}")
    return output_path

def send_email(config, recipient_email, recipient_name, attachment_path):
    msg = MIMEMultipart()
    msg["From"] = config["email"]
    msg["To"] = recipient_email
    msg["Subject"] = f"Certificate of Achievement - {recipient_name}"
    body = f"Dear {recipient_name},\n\nPlease find your certificate attached.\n\nBest regards."
    msg.attach(MIMEText(body, "plain"))

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
        msg.attach(part)

    with smtplib.SMTP(config["smtp_server"], config["smtp_port"]) as server:
        server.starttls()
        server.login(config["email"], config["app_password"])
        server.send_message(msg)

def main():
    config = load_config()
    with open("data.csv", newline="") as file:
        reader = csv.reader(file)
        for name, email in reader:
            output_file = write_text_on_image(name)
            send_email(config, email, name, output_file)
            print(f"📩 Sent to {name} ({email})")

if __name__ == "__main__":
    main()
