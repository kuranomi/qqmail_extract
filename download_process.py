import imaplib
import email
from email.header import decode_header
import datetime
import os
from bs4 import BeautifulSoup
import re
import json

from config.settings import EMAIL_USER, EMAIL_PASS, IMAP_SERVER, IMAP_PORT, TARGET_SENDER, EXTRACT_DIR


class QQEmail:
    def __init__(self, box="INBOX"):
        self.mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        self.mail.login(EMAIL_USER, EMAIL_PASS)
        self.mail.select(box)

    def fetch_mails(self, tag="UNSEEN"):
        # status, messages = mail.search(None, 'UNSEEN')
        search_criteria = f'{tag} FROM "{TARGET_SENDER}"'
        status, messages = self.mail.search(None, search_criteria, 'CHARSET UTF-8')
        if status != 'OK':
            print("fail")
            return []

        email_ids = messages[0].split()
        print(f"{len(email_ids)} emails been found")
        return email_ids

    def tag_email(self, email_id, tag="\\Seen"):
        self.mail.store(email_id, '+FLAGS',tag)
        return True

    def get_content(self, email_id):
        content = {}
        # not tag to seen
        status, msg_data = self.mail.fetch(email_id, '(BODY.PEEK[])')

        if status != 'OK':
            print(f"fail on {email_id}")
            return False, {}

        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        # get subject
        subject, subject_encoding = decode_header(msg.get('Subject', ''))[0]
        if isinstance(subject, bytes):
            subject = subject.decode(subject_encoding or 'utf-8', errors='ignore')
        else:
            subject = subject or "[无标题]"

        content["subject"] = subject
        print(f"handling email: {subject}")
        extract_dir = f"{EXTRACT_DIR}/{str(email_id)}_{subject[5:20]}"
        os.makedirs(extract_dir, exist_ok=True)
        content["save_dir"] = extract_dir
        
        email_content = self.parse_email(msg, subject_encoding, extract_dir)
        content.update(email_content)

        file_path = f'{extract_dir}/email_content.json'
        print(f"email contend saved: {file_path}")
        if not os.path.isfile(file_path):
            with open(file_path, 'w', encoding='utf-8')as fp:
                fp.write(json.dumps(content, indent=4, ensure_ascii=False))

        print("finish parsing email...")
        return content

    def parse_email(self, msg, subject_encode, save_dir):
        if not msg.is_multipart():
            content_type = msg.get_content_type()
            body = msg.get_payload(decode=True).decode(subject_encode or'utf-8', 'ignore')
            parsed_details = self.parse_email_content(content_type, body, subject_encode)
            return parsed_details

        all_details = {}
        attachments = []
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if "attachment" not in content_disposition:
                part_detail = self.parse_email_content(content_type, part, subject_encode)
                all_details.update(part_detail)
            else:
                filename = part.get_filename()
                if filename:
                    filename = decode_header(filename)[0][0]
                    if isinstance(filename, bytes):
                        filename = filename.decode(subject_encode or 'utf-8', 'ignore')
                else:
                    filename = "attachemnt_file"

                filepath = os.path.join(save_dir, filename)
                if not os.path.isfile(filepath):
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                attachments.append(filepath)
        all_details["attachemnt_path"] = attachments
        return all_details

    def parse_email_content(self, content_type, body, subject_encode):
        res = {}
        if content_type == "text/plain":
            body = body.get_payload(decode=True).decode(subject_encode or 'utf-8', 'ignore')
            body_detail = self.get_detail_from_txt(body)
            return body_detail
        elif content_type == "text/html":
            body = body.get_payload(decode=True).decode(subject_encode or 'utf-8', 'ignore')
            image_links = self.get_detail_from_html(body)
            res["image_links"] = image_links
        return res

    def get_detail_from_html(self, body):
        soup = BeautifulSoup(body, 'html.parser')
        # get image links
        links = []
        for tag in soup.find_all(['a', 'link', 'script', 'img']):
            if tag.name == 'a' and tag.has_attr('href'):
                links.append(tag['href'])
            elif tag.name == 'link' and tag.has_attr('href'):
                links.append(tag['href'])
            elif tag.name == 'script' and tag.has_attr('src'):
                links.append(tag['src'])
            elif tag.name == 'img' and tag.has_attr('src'):
                links.append(tag['src'])
        return links

    def get_detail_from_txt(self, body):
        if not isinstance(body, str):
            return None
        message_detail = {
            "email_content": body
        }
        body_sp = body.split("\n")
        first_line = None
        for line in body_sp:
            if not line:
                continue
            first_line = line
            break
        try:
            percentages = re.findall(r'(\d+)[%％]', first_line) 
            if percentages:
                message_detail["discount"] = percentages[-1]
        except Exception:
            pass
        try:
            order_deadline = re.search(r'[(締切日|締め切り)].*?(\d+月\d+日)', body).group(1)
            if order_deadline:
                message_detail["cutoff_date"] = order_deadline
            release_date = re.search(r'(?:(?:発売).*?)?((?:\d+年)?\d+月)(?:.*?(?:発売))?', body).group(1)
            if release_date:
                message_detail["release_date"] = release_date
        except Exception:
            pass
        return message_detail

