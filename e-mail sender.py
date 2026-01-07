"""
Python Email Sender Mini Project
A complete email sending application in a single file
Supports Gmail, Outlook, Yahoo, and custom SMTP servers
"""

import smtplib
import ssl
import json
import csv
import os
import getpass
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Dict, Optional
import re

class EmailSender:
    """Main email sender class with all functionality"""
    
    # Common SMTP server configurations
    SMTP_CONFIGS = {
        "gmail": {"server": "smtp.gmail.com", "port": 587, "ssl": True},
        "outlook": {"server": "smtp.office365.com", "port": 587, "ssl": True},
        "yahoo": {"server": "smtp.mail.yahoo.com", "port": 587, "ssl": True},
        "aol": {"server": "smtp.aol.com", "port": 587, "ssl": True},
        "zoho": {"server": "smtp.zoho.com", "port": 587, "ssl": True}
    }
    
    def __init__(self, config_file: str = "email_config.json"):
        """Initialize the email sender"""
        self.config_file = config_file
        self.config = self._load_config()
        self.recipients = []
        
    def _load_config(self) -> Dict:
        """Load or create configuration"""
        default_config = {
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "sender_email": "",
            "sender_name": "Python Email Sender",
            "use_ssl": True,
            "save_password": False,
            "last_used": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                print(f"✓ Configuration loaded from {self.config_file}")
                return config
            except:
                print("✗ Error loading config, using defaults")
        
        self._save_config(default_config)
        print(f"✓ Created new configuration file: {self.config_file}")
        return default_config
    
    def _save_config(self, config: Dict = None) -> None:
        """Save configuration to file"""
        if config is None:
            config = self.config
        
        config["last_used"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        with open(self.config_file, 'w') as f:
            json.dump(config, f, indent=4)
    
    def _validate_email(self, email: str) -> bool:
        """Basic email validation"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))
    
    def setup_wizard(self) -> None:
        """Interactive setup wizard"""
        print("\n" + "="*60)
        print("EMAIL SENDER SETUP WIZARD")
        print("="*60)
        
        # Choose email provider
        print("\n📧 Select your email provider:")
        print("1. Gmail")
        print("2. Outlook/Hotmail")
        print("3. Yahoo")
        print("4. Other/Custom SMTP")
        
        choice = input("\nEnter choice (1-4): ").strip()
        
        if choice == "1":
            self.config["smtp_server"] = self.SMTP_CONFIGS["gmail"]["server"]
            self.config["smtp_port"] = self.SMTP_CONFIGS["gmail"]["port"]
            print("✓ Using Gmail SMTP settings")
        elif choice == "2":
            self.config["smtp_server"] = self.SMTP_CONFIGS["outlook"]["server"]
            self.config["smtp_port"] = self.SMTP_CONFIGS["outlook"]["port"]
            print("✓ Using Outlook SMTP settings")
        elif choice == "3":
            self.config["smtp_server"] = self.SMTP_CONFIGS["yahoo"]["server"]
            self.config["smtp_port"] = self.SMTP_CONFIGS["yahoo"]["port"]
            print("✓ Using Yahoo SMTP settings")
        else:
            self.config["smtp_server"] = input("Enter SMTP server: ").strip()
            port = input("Enter SMTP port (default 587): ").strip()
            self.config["smtp_port"] = int(port) if port else 587
        
        # Get sender details
        print("\n👤 Sender Information:")
        while True:
            email = input("Your email address: ").strip()
            if self._validate_email(email):
                self.config["sender_email"] = email
                break
            print("✗ Invalid email format. Please try again.")
        
        self.config["sender_name"] = input("Your name (optional): ").strip() or "Python Email Sender"
        
        # Password handling
        print("\n🔒 Password Note:")
        print("• For Gmail, use an 'App Password' (not your regular password)")
        print("• Enable 2-factor authentication first, then create app password")
        print("• Other providers may require allowing 'less secure apps'")
        
        save_pass = input("\nSave password in config? (y/n): ").lower().strip()
        self.config["save_password"] = save_pass == 'y'
        
        if self.config["save_password"]:
            password = getpass.getpass("Enter email password/app password: ")
            self.config["sender_password"] = password
        else:
            # Remove saved password if exists
            if "sender_password" in self.config:
                del self.config["sender_password"]
        
        self._save_config()
        print("\n✅ Setup complete!")
    
    def load_recipients(self, source: str = "manual") -> List[Dict]:
        """Load recipients from different sources"""
        self.recipients = []
        
        print("\n" + "="*60)
        print("LOAD RECIPIENTS")
        print("="*60)
        
        if source == "manual":
            print("\n📝 Manual Entry:")
            print("Enter one email per line. Type 'done' when finished.")
            print("Format: email or name <email>")
            
            while True:
                entry = input("> ").strip()
                if entry.lower() == 'done':
                    break
                
                if entry:
                    # Parse "Name <email>" format
                    if '<' in entry and '>' in entry:
                        name = entry.split('<')[0].strip()
                        email = entry.split('<')[1].split('>')[0].strip()
                    else:
                        name = ""
                        email = entry.strip()
                    
                    if self._validate_email(email):
                        self.recipients.append({"name": name, "email": email})
                        print(f"✓ Added: {name or email}")
                    else:
                        print(f"✗ Invalid email: {email}")
        
        elif source == "csv":
            filename = input("Enter CSV filename (default: recipients.csv): ").strip() or "recipients.csv"
            
            if not os.path.exists(filename):
                print(f"✗ File not found: {filename}")
                create_sample = input("Create sample CSV file? (y/n): ").lower()
                if create_sample == 'y':
                    self._create_sample_csv(filename)
                return []
            
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        if 'email' in row and self._validate_email(row['email']):
                            self.recipients.append({
                                "name": row.get('name', ''),
                                "email": row['email']
                            })
                print(f"✓ Loaded {len(self.recipients)} recipients from {filename}")
            except Exception as e:
                print(f"✗ Error reading CSV: {e}")
        
        elif source == "text":
            filename = input("Enter text filename: ").strip()
            if os.path.exists(filename):
                with open(filename, 'r') as f:
                    for line in f:
                        email = line.strip()
                        if email and self._validate_email(email):
                            self.recipients.append({"name": "", "email": email})
                print(f"✓ Loaded {len(self.recipients)} recipients from {filename}")
        
        return self.recipients
    
    def _create_sample_csv(self, filename: str) -> None:
        """Create a sample CSV file"""
        sample_data = [
            {"name": "John Doe", "email": "john@example.com"},
            {"name": "Jane Smith", "email": "jane@example.com"},
            {"name": "Bob Johnson", "email": "bob@example.com"}
        ]
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["name", "email"])
                writer.writeheader()
                writer.writerows(sample_data)
            print(f"✓ Created sample CSV: {filename}")
            print("Edit this file with your actual recipients")
        except Exception as e:
            print(f"✗ Error creating CSV: {e}")
    
    def compose_email(self) -> Dict:
        """Compose email content"""
        print("\n" + "="*60)
        print("COMPOSE EMAIL")
        print("="*60)
        
        email_data = {
            "subject": "",
            "body": "",
            "is_html": False,
            "attachments": []
        }
        
        # Subject
        email_data["subject"] = input("\n📋 Subject: ").strip() or "No Subject"
        
        # Body type
        print("\n📝 Email Body Type:")
        print("1. Plain Text")
        print("2. HTML")
        choice = input("Enter choice (1-2): ").strip()
        
        email_data["is_html"] = choice == "2"
        
        # Body content
        print("\n✏️ Enter/Paste your email content.")
        print("Type 'END' on a new line when finished.")
        print("-" * 40)
        
        lines = []
        while True:
            line = input()
            if line.strip().upper() == "END":
                break
            lines.append(line)
        
        email_data["body"] = "\n".join(lines)
        
        # Attachments
        add_attachments = input("\n📎 Add attachments? (y/n): ").lower().strip()
        if add_attachments == 'y':
            print("Enter attachment filenames (one per line), type 'done' when finished:")
            while True:
                filename = input("> ").strip()
                if filename.lower() == 'done':
                    break
                if os.path.exists(filename):
                    email_data["attachments"].append(filename)
                    print(f"✓ Added: {filename}")
                else:
                    print(f"✗ File not found: {filename}")
        
        # Preview
        print("\n" + "="*60)
        print("EMAIL PREVIEW")
        print("="*60)
        print(f"Subject: {email_data['subject']}")
        print(f"Body type: {'HTML' if email_data['is_html'] else 'Plain Text'}")
        print(f"Body length: {len(email_data['body'])} characters")
        print(f"Attachments: {len(email_data['attachments'])} files")
        
        return email_data
    
    def create_message(self, recipient: Dict, email_data: Dict) -> MIMEMultipart:
        """Create email message"""
        message = MIMEMultipart("alternative")
        
        # Headers
        message["Subject"] = email_data["subject"]
        message["From"] = f"{self.config.get('sender_name', '')} <{self.config.get('sender_email', '')}>"
        
        # Recipient formatting
        if recipient.get("name"):
            message["To"] = f"{recipient['name']} <{recipient['email']}>"
        else:
            message["To"] = recipient["email"]
        
        # Email body
        if email_data["is_html"]:
            # Create HTML email with basic styling
            html_template = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        line-height: 1.6;
                        color: #333;
                        max-width: 600px;
                        margin: 0 auto;
                        padding: 20px;
                    }}
                    .header {{
                        background-color: #4CAF50;
                        color: white;
                        padding: 20px;
                        text-align: center;
                        border-radius: 5px;
                    }}
                    .content {{
                        background-color: #f9f9f9;
                        padding: 20px;
                        border: 1px solid #ddd;
                        border-radius: 5px;
                        margin-top: 20px;
                    }}
                    .footer {{
                        margin-top: 30px;
                        font-size: 12px;
                        color: #777;
                        text-align: center;
                        border-top: 1px solid #eee;
                        padding-top: 10px;
                    }}
                </style>
            </head>
            <body>
                <div class="header">
                    <h2>{email_data['subject']}</h2>
                </div>
                <div class="content">
                    {email_data['body'].replace('\n', '<br>')}
                </div>
                <div class="footer">
                    <p>Sent using Python Email Sender</p>
                    <p>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                </div>
            </body>
            </html>
            """
            message.attach(MIMEText(html_template, "html"))
            # Also include plain text version
            message.attach(MIMEText(email_data["body"], "plain"))
        else:
            message.attach(MIMEText(email_data["body"], "plain"))
        
        # Attachments
        for attachment in email_data["attachments"]:
            try:
                with open(attachment, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                
                encoders.encode_base64(part)
                filename = os.path.basename(attachment)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={filename}",
                )
                message.attach(part)
            except Exception as e:
                print(f"✗ Error attaching {attachment}: {e}")
        
        return message
    
    def send_emails(self, email_data: Dict, test_mode: bool = False) -> Dict:
        """Send emails to all recipients"""
        if not self.recipients:
            print("✗ No recipients selected")
            return {"sent": 0, "failed": 0, "total": 0}
        
        results = {
            "sent": 0,
            "failed": 0,
            "total": len(self.recipients),
            "errors": []
        }
        
        # Get password
        password = None
        if self.config.get("save_password") and "sender_password" in self.config:
            password = self.config["sender_password"]
        else:
            password = getpass.getpass(f"Enter password for {self.config.get('sender_email')}: ")
        
        if test_mode:
            print("\n🔧 TEST MODE - No emails will be sent")
            print(f"Would send to {len(self.recipients)} recipients")
            return results
        
        print(f"\n📤 Sending emails to {len(self.recipients)} recipients...")
        print("-" * 60)
        
        try:
            # Connect to server
            context = ssl.create_default_context()
            
            if self.config.get("use_ssl", True):
                server = smtplib.SMTP(self.config["smtp_server"], self.config["smtp_port"])
                server.starttls(context=context)
            else:
                server = smtplib.SMTP(self.config["smtp_server"], self.config["smtp_port"])
            
            # Login
            server.login(self.config["sender_email"], password)
            print("✓ Connected and logged in successfully")
            
            # Send emails
            for i, recipient in enumerate(self.recipients, 1):
                try:
                    message = self.create_message(recipient, email_data)
                    server.sendmail(
                        self.config["sender_email"],
                        recipient["email"],
                        message.as_string()
                    )
                    
                    recipient_name = recipient.get("name", recipient["email"])
                    print(f"✓ [{i}/{len(self.recipients)}] Sent to: {recipient_name}")
                    results["sent"] += 1
                    
                except Exception as e:
                    error_msg = f"Failed to send to {recipient['email']}: {str(e)}"
                    print(f"✗ [{i}/{len(self.recipients)}] {error_msg}")
                    results["failed"] += 1
                    results["errors"].append(error_msg)
            
            server.quit()
            print("\n✅ Sending complete!")
            
        except Exception as e:
            print(f"\n✗ Connection/Login error: {e}")
            print("⚠ Common solutions:")
            print("  1. Check your password (use App Password for Gmail)")
            print("  2. Enable 'Less secure apps' for your email provider")
            print("  3. Check if SMTP is enabled for your account")
            print("  4. Verify server/port settings")
        
        return results
    
    def send_test_email(self) -> None:
        """Send a test email to yourself"""
        print("\n" + "="*60)
        print("SEND TEST EMAIL")
        print("="*60)
        
        test_recipient = [{
            "name": "Test Recipient",
            "email": self.config.get("sender_email", "")
        }]
        
        self.recipients = test_recipient
        
        email_data = {
            "subject": "Test Email from Python Email Sender",
            "body": f"""Hello!

This is a test email sent from the Python Email Sender application.

If you're receiving this email, your SMTP configuration is working correctly!

Application Info:
- Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- SMTP Server: {self.config.get('smtp_server')}
- Sender: {self.config.get('sender_email')}

Best regards,
Python Email Sender
""",
            "is_html": False,
            "attachments": []
        }
        
        print("\n📧 Test Email Details:")
        print(f"To: {test_recipient[0]['email']}")
        print(f"Subject: {email_data['subject']}")
        
        confirm = input("\nSend test email? (y/n): ").lower().strip()
        if confirm == 'y':
            results = self.send_emails(email_data, test_mode=False)
            if results["sent"] > 0:
                print("\n✅ Test email sent successfully!")
                print("Please check your inbox (and spam folder)")
    
    def view_statistics(self) -> None:
        """View sending statistics"""
        stats_file = "email_stats.json"
        
        if os.path.exists(stats_file):
            with open(stats_file, 'r') as f:
                stats = json.load(f)
            
            print("\n" + "="*60)
            print("EMAIL STATISTICS")
            print("="*60)
            
            total_sent = stats.get("total_sent", 0)
            total_failed = stats.get("total_failed", 0)
            last_sent = stats.get("last_sent", "Never")
            
            print(f"\n📊 Totals:")
            print(f"   Emails Sent: {total_sent}")
            print(f"   Failed: {total_failed}")
            print(f"   Last Sent: {last_sent}")
            
            if "campaigns" in stats:
                print(f"\n📁 Campaigns: {len(stats['campaigns'])}")
                for i, camp in enumerate(stats["campaigns"][-5:], 1):
                    print(f"   {i}. {camp.get('subject', 'No subject')} - {camp.get('date', 'Unknown')}")
        else:
            print("\nNo statistics found. Send some emails first!")
    
    def save_statistics(self, results: Dict, email_data: Dict) -> None:
        """Save sending statistics"""
        stats_file = "email_stats.json"
        
        stats = {
            "total_sent": 0,
            "total_failed": 0,
            "last_sent": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "campaigns": []
        }
        
        if os.path.exists(stats_file):
            with open(stats_file, 'r') as f:
                stats = json.load(f)
        
        stats["total_sent"] = stats.get("total_sent", 0) + results.get("sent", 0)
        stats["total_failed"] = stats.get("total_failed", 0) + results.get("failed", 0)
        stats["last_sent"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        campaign = {
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "subject": email_data.get("subject", ""),
            "recipients": results.get("total", 0),
            "sent": results.get("sent", 0),
            "failed": results.get("failed", 0)
        }
        
        if "campaigns" not in stats:
            stats["campaigns"] = []
        
        stats["campaigns"].append(campaign)
        
        # Keep only last 50 campaigns
        if len(stats["campaigns"]) > 50:
            stats["campaigns"] = stats["campaigns"][-50:]
        
        with open(stats_file, 'w') as f:
            json.dump(stats, f, indent=4)
        
        print(f"✓ Statistics saved to {stats_file}")

def display_banner():
    """Display application banner"""
    banner = """
    ╔═══════════════════════════════════════════════╗
    ║        PYTHON EMAIL SENDER MINI PROJECT       ║
    ║          Complete Email Sending Tool          ║
    ╚═══════════════════════════════════════════════╝
    
    Features:
    • Send individual or bulk emails
    • HTML and plain text support
    • File attachments
    • Multiple recipient sources
    • Email validation
    • Statistics tracking
    • Test mode available
    
    """
    print(banner)

def main():
    """Main application loop"""
    display_banner()
    
    # Initialize email sender
    sender = EmailSender()
    
    while True:
        print("\n" + "="*60)
        print("MAIN MENU")
        print("="*60)
        print("1. 📧 Send Emails")
        print("2. ⚙️  Setup/Configuration")
        print("3. 🧪 Send Test Email")
        print("4. 📊 View Statistics")
        print("5. 👥 Manage Recipients")
        print("6. ℹ️  Help/Instructions")
        print("7. 🚪 Exit")
        
        choice = input("\nEnter your choice (1-7): ").strip()
        
        if choice == "1":
            # Send emails
            print("\n📋 How would you like to load recipients?")
            print("1. Manual entry")
            print("2. From CSV file")
            print("3. From text file")
            print("4. Use existing recipients")
            
            recipient_choice = input("Enter choice (1-4): ").strip()
            
            if recipient_choice == "1":
                sender.load_recipients("manual")
            elif recipient_choice == "2":
                sender.load_recipients("csv")
            elif recipient_choice == "3":
                sender.load_recipients("text")
            elif recipient_choice != "4":
                print("Using existing recipients")
            
            if not sender.recipients:
                print("No recipients loaded. Please load recipients first.")
                continue
            
            # Compose email
            email_data = sender.compose_email()
            
            # Send options
            print("\n⚙️ Send Options:")
            print("1. Send now")
            print("2. Test mode (no emails sent)")
            
            send_choice = input("Enter choice (1-2): ").strip()
            test_mode = send_choice == "2"
            
            if send_choice in ["1", "2"]:
                confirm = input(f"\nSend email to {len(sender.recipients)} recipients? (y/n): ").lower()
                if confirm == 'y':
                    results = sender.send_emails(email_data, test_mode)
                    
                    if not test_mode:
                        # Save statistics
                        sender.save_statistics(results, email_data)
                        
                        # Show summary
                        print("\n" + "="*60)
                        print("SENDING SUMMARY")
                        print("="*60)
                        print(f"Total: {results['total']}")
                        print(f"Successfully sent: {results['sent']}")
                        print(f"Failed: {results['failed']}")
                        
                        if results['failed'] > 0 and results['errors']:
                            print("\nErrors:")
                            for error in results['errors'][:5]:
                                print(f"  • {error}")
        
        elif choice == "2":
            # Setup
            sender.setup_wizard()
        
        elif choice == "3":
            # Test email
            sender.send_test_email()
        
        elif choice == "4":
            # View statistics
            sender.view_statistics()
        
        elif choice == "5":
            # Manage recipients
            print(f"\nCurrent recipients: {len(sender.recipients)}")
            if sender.recipients:
                print("\nFirst 5 recipients:")
                for i, recipient in enumerate(sender.recipients[:5], 1):
                    name = recipient.get('name', 'No name')
                    print(f"  {i}. {name} - {recipient['email']}")
            
            action = input("\nLoad new recipients? (y/n): ").lower()
            if action == 'y':
                print("\nLoad from:")
                print("1. Manual entry")
                print("2. CSV file")
                print("3. Text file")
                
                source_choice = input("Enter choice (1-3): ").strip()
                sources = {1: "manual", 2: "csv", 3: "text"}
                source = sources.get(int(choice), "manual")
                sender.load_recipients(source)
        
        elif choice == "6":
            # Help
            print("\n" + "="*60)
            print("HELP & INSTRUCTIONS")
            print("="*60)
            print("\n🔧 Setup Tips:")
            print("• Gmail users: Create an 'App Password' (not your regular password)")
            print("• Enable 2-factor authentication first in Google Account")
            print("• Outlook/Yahoo may require allowing 'less secure apps'")
            
            print("\n📝 Recipient Formats:")
            print("• Single email: john@example.com")
            print("• With name: John Doe <john@example.com>")
            print("• CSV file should have 'name' and 'email' columns")
            
            print("\n⚠️  Important Notes:")
            print("• Always test with yourself first")
            print("• Don't spam or violate email policies")
            print("• Check spam folder if emails aren't received")
            
            print("\n📁 Files Created:")
            print("• email_config.json - Your email settings")
            print("• email_stats.json - Sending statistics")
            print("• recipients.csv - Sample recipient list")
        
        elif choice == "7":
            print("\nThank you for using Python Email Sender!")
            print("Goodbye! 👋")
            break
        
        else:
            print("Invalid choice. Please enter 1-7.")
        
        input("\nPress Enter to continue...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProgram interrupted. Goodbye!")
    except Exception as e:
        print(f"\nAn error occurred: {e}")
        print("Please check your configuration and try again.")