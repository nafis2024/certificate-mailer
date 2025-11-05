#!/usr/bin/env python3
"""
main.py - Certificate mailer (hard-coded template & centered text)

Usage:
    python3 main.py

Files expected (hard-coded):
    - certificates/template.png   <-- certificate template image
    - data.csv                    <-- CSV with lines: name,email
    - config.json                 <-- email settings (not committed)
Output:
    - output/<sanitized_name>.png
"""

import csv
import json
import os
import re
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import Optional, Tuple

from PIL import Image, ImageDraw, ImageFont

# ---------------------------
# Hard-coded project paths
# ---------------------------
TEMPLATE_PATH = Path("certificates/template.png")
DATA_CSV = Path("data.csv")
CONFIG_JSON = Path("config.json")
OUTPUT_DIR = Path("output")
DEFAULT_FONT_SIZE = 80
# Vertical coordinate where name will be drawn (adjust to your template)
TEXT_Y = 580

# ---------------------------
# Utilities
# ---------------------------
def load_config(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(
            f"Missing config file: {path}\nCreate it with keys: email, app_password, smtp_server, smtp_port"
        )
    with path.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    required = ["email", "app_password", "smtp_server", "smtp_port"]
    missing = [k for k in required if k not in cfg]
    if missing:
        raise KeyError(f"Missing keys in config.json: {missing}")
    return cfg

def sanitize_filename(name: str) -> str:
    # Remove problematic characters, limit length
    safe = re.sub(r"[^\w\-_(). ]", "", name).strip()
    if not safe:
        safe = "recipient"
    return safe[:120]

def find_font(preferred_size: int = DEFAULT_FONT_SIZE) -> Tuple[ImageFont.FreeTypeFont | ImageFont.ImageFont, int]:
    """
    Try several common fonts. If none available, fallback to Pillow default.
    Returns (font_object, size_used)
    """
    # Common font paths on Linux / macOS / Windows (codespaces and many linux envs have DejaVu)
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSerif.ttf",
        "/Library/Fonts/Arial.ttf",  # macOS
        "C:\\Windows\\Fonts\\arial.ttf",  # Windows
    ]
    for p in candidates:
        try:
            if Path(p).exists():
                return ImageFont.truetype(p, preferred_size), preferred_size
        except Exception:
            continue

    # If no TTF found, fallback to default (size ignored)
    return ImageFont.load_default(), DEFAULT_FONT_SIZE

def ensure_output_dir():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------
# Image / Certificate logic
# ---------------------------
def write_text_on_image(name: str, template: Path = TEMPLATE_PATH, y_pos: int = TEXT_Y) -> Path:
    """
    Draw `name` centered horizontally on the given template and save PNG to output/.
    Returns path to generated image.
    """
    if not template.exists():
        raise FileNotFoundError(f"Certificate template not found: {template}")

    # Load image
    image = Image.open(template).convert("RGBA")
    draw = ImageDraw.Draw(image)

    # Load font (with fallback)
    font, used_size = find_font(DEFAULT_FONT_SIZE)

    # Measure text (use textbbox for more reliable measurements)
    text = name.strip()
    # Pillow's textbbox requires specifying anchor position; measure from (0,0)
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    img_width, img_height = image.size
    x = (img_width - text_width) / 2
    y = y_pos

    # Optionally, add subtle shadow for better contrast (draw shadow then text)
    shadow_offset = 2
    try:
        # Shadow
        draw.text((x + shadow_offset, y + shadow_offset), text, font=font, fill=(0, 0, 0, 150))
    except Exception:
        # If RGBA issues, ignore shadow
        pass

    # Main text (choose black or dark gray)
    draw.text((x, y), text, font=font, fill=(10, 10, 10, 255))

    # Ensure output dir exists
    ensure_output_dir()

    # Make a safe filename
    safe_name = sanitize_filename(name)
    output_path = OUTPUT_DIR / f"{safe_name}.png"

    # Save as PNG to preserve quality
    # Convert to RGB if template was not RGBA to avoid metadata issues
    if image.mode in ("RGBA", "LA"):
        # Flatten transparency onto white background for compatibility
        background = Image.new("RGB", image.size, (255, 255, 255))
        background.paste(image, mask=image.split()[-1])  # alpha channel as mask
        background.save(output_path, format="PNG")
    else:
        image.save(output_path, format="PNG")

    return output_path

# ---------------------------
# Email logic
# ---------------------------
def make_message(sender: str, recipient_email: str, recipient_name: str, body_text: str, attachment_path: Path) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient_email
    msg["Subject"] = f"Certificate of Achievement - {recipient_name}"
    msg.attach(MIMEText(body_text, "plain"))

    # Attach file
    with attachment_path.open("rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{attachment_path.name}"')
    msg.attach(part)
    return msg

def send_email(config: dict, recipient_email: str, recipient_name: str, attachment_path: Path, body_text: Optional[str] = None):
    if body_text is None:
        body_text = f"Dear {recipient_name},\\nThank you for participating in the event. Your certificate of participation has been attached to this email.\\n\\nRegards,\\nNDITC"

    msg = make_message(config["email"], recipient_email, recipient_name, body_text, attachment_path)

    # Connect to SMTP
    server = smtplib.SMTP(config["smtp_server"], int(config["smtp_port"]), timeout=30)
    try:
        server.starttls()
        server.login(config["email"], config["app_password"])
        server.send_message(msg)
    finally:
        server.quit()

# ---------------------------
# CSV reading
# ---------------------------
def read_recipients(csv_path: Path):
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    items = []
    with csv_path.open(newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            # Skip blank rows
            if not row or all(not cell.strip() for cell in row):
                continue
            # Support possible header: if first row contains "name" or "email", skip it
            if len(items) == 0 and any(h.lower() in ("name", "email") for h in row):
                # This is a header row — skip it
                continue
            # Expect at least two columns: name, email
            if len(row) < 2:
                print(f"Skipping invalid row (need name,email): {row}")
                continue
            name, email = row[0].strip(), row[1].strip()
            if not name or not email:
                print(f"Skipping incomplete row: {row}")
                continue
            items.append((name, email))
    return items

# ---------------------------
# Main flow
# ---------------------------
def main():
    print("Certificate Mailer - starting...\n")

    # Load config
    try:
        config = load_config(CONFIG_JSON)
    except Exception as e:
        print("ERROR: Failed to load config.json ->", e)
        print("\nCreate a config.json file with contents like:\n"
              '{\n  "email": "youremail@example.com",\n  "app_password": "your_app_password",\n  "smtp_server": "smtp.gmail.com",\n  "smtp_port": 587\n}\n')
        sys.exit(1)

    # Read recipients
    try:
        recipients = read_recipients(DATA_CSV)
    except Exception as e:
        print("ERROR: Failed to read data.csv ->", e)
        sys.exit(1)

    if not recipients:
        print("No recipients found in data.csv. Exiting.")
        sys.exit(0)

    # Show summary and ask for confirmation
    print(f"Template (hard-coded): {TEMPLATE_PATH}")
    print(f"Recipients found: {len(recipients)}")
    print(f"Output folder: {OUTPUT_DIR.resolve()}\n")

    proceed = input("Proceed to generate and send certificates? (yes/no) [no]: ").strip().lower()
    if proceed not in ("yes", "y"):
        print("Aborted by user.")
        sys.exit(0)

    success_count = 0
    fail_count = 0
    for idx, (name, email) in enumerate(recipients, start=1):
        print(f"\n[{idx}/{len(recipients)}] Processing: {name} <{email}>")
        try:
            out_path = write_text_on_image(name)
            print(f"  -> Saved: {out_path}")

            # Optionally craft a custom email body; simple default used here
            send_email(config, email, name, out_path)
            print(f"  -> Email sent to {email}")
            success_count += 1
        except Exception as e:
            print(f"  !! Failed for {name} <{email}>: {e}")
            fail_count += 1

    # Summary
    print("\n--- Summary ---")
    print(f"Total attempted: {len(recipients)}")
    print(f"Successful: {success_count}")
    print(f"Failed: {fail_count}")
    print("Finished.")

if __name__ == "__main__":
    main()