import csv
import os
import json
import smtplib
from email.message import EmailMessage
from PIL import Image, ImageDraw, ImageFont


def load_config():
    with open("config.json", "r") as f:
        return json.load(f)


def load_recipients(csv_file):
    recipients = []
    with open(csv_file, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if not row.get("name") or not row.get("email") or not row.get("piORth"):
                print("⚠️ Skipping a row: missing one of 'name', 'email', 'piORth'.")
                continue
            recipients.append(row)
    return recipients


def write_text_on_image(name, template_config, output_folder):
    template_path = template_config["template_path"]
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    image = Image.open(template_path).convert("RGBA")
    draw = ImageDraw.Draw(image)

    font_path = template_config["font_path"]
    font = ImageFont.truetype(font_path, template_config["font_size"])

    text_color = tuple(template_config["font_color"])
    y = template_config["text_position"][1]

    # Center text horizontally
    bbox = draw.textbbox((0, 0), name, font=font)
    text_width = bbox[2] - bbox[0]
    x = (image.width - text_width) / 2

    draw.text((x, y), name, font=font, fill=text_color)

    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, f"{name}.png")
    image.save(output_file)
    return output_file


def send_email(recipient_name, recipient_email, output_file, template_config, email_settings):
    msg = EmailMessage()
    msg["Subject"] = template_config["email_subject"]
    msg["From"] = email_settings["sender_email"]
    msg["To"] = recipient_email

    msg.set_content(template_config["email_body"].format(name=recipient_name))

    with open(output_file, "rb") as f:
        msg.add_attachment(f.read(), maintype="image", subtype="png", filename=os.path.basename(output_file))

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login(email_settings["sender_email"], email_settings["sender_password"])
        smtp.send_message(msg)


def main():
    print("Certificate Mailer - starting...\n")

    config = load_config()
    recipients = load_recipients(config["input_csv"])

    if not recipients:
        print("No valid recipients found.")
        return

    output_folder = config["output_folder"]
    print(f"Recipients loaded: {len(recipients)}")
    print(f"Output directory: {os.path.abspath(output_folder)}\n")

    proceed = input("Proceed to generate and send certificates? (yes/no) [no]: ").strip().lower()
    if proceed != "yes":
        print("Aborted by user.")
        return

    total = len(recipients)
    success, failed = 0, 0

    for idx, recipient in enumerate(recipients, start=1):
        name = recipient["name"].strip()
        email = recipient["email"].strip()
        group = recipient["piORth"].strip()

        print(f"[{idx}/{total}] Processing: {name} <{email}>")

        if group not in config["templates"]:
            print(f"  !! Unknown group: {group}. Skipping.")
            failed += 1
            continue

        template_cfg = config["templates"][group]

        try:
            output_file = write_text_on_image(name, template_cfg, output_folder)
            print(f"  -> Saved: {output_file}")
            send_email(name, email, output_file, template_cfg, config["email_settings"])
            print(f"  ✅ Sent to {email}\n")
            success += 1
        except Exception as e:
            print(f"  !! Failed for {name} <{email}>: {e}\n")
            failed += 1

    print("\n--- Summary ---")
    print(f"Total attempted: {total}")
    print(f"Successful: {success}")
    print(f"Failed: {failed}")
    print("Finished.")


if __name__ == "__main__":
    main()
