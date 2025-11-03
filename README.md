# Certificate Mailer 📨

A Python script that automatically writes names on a certificate and emails them.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Place your blank certificate image in:
   ```
   certificates/template.png
   ```

3. Create your `data.csv` file with:
   ```
   name,email
   ```

4. Add your email credentials in `config.json`.

5. Run the script:
   ```bash
   python3 main.py
   ```

Certificates will be saved in `/output` and sent automatically.
