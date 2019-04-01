from __future__ import print_function
from typing import List
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import base64
import email
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from config import *
import subprocess
import os

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
]

class File:
    name: str
    input_path: str
    output_path: str

    def __init__(self, filename):
        self.name = filename
        self.input_path = INPUT_PATH + filename
        self.output_path = OUTPUT_PATH + filename


class Message:
    id: str
    subject: str
    files: List[File]

    def __init__(self, id):
        self.id = id
        self.subject = None
        self.files = []

    def has_existing(self, file: File):
        for item in self.files:
            if item.name == file.name:
                return True


class InvoiceProcessor:
    def __init__(self):
        creds = file.Storage("token.json").get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets("credentials.json", SCOPES)
            creds = tools.run_flow(flow, store)
        self.service = build("gmail", "v1", http=creds.authorize(Http()))

    def run(self):
        results = (
            self.service.users()
            .messages()
            .list(userId="me", labelIds=[NEW_LABEL_ID])
            .execute()
        )
        for result in results.get("messages", []):
            message = Message(result.get("id"))
            self.process_invoice(message)
            if len(message.files):  # make sure we're not sending emails without attachments (eg in same thread)
                self.forward_invoice(message)
                self.cleanup(message)
            self.archive(message)

    def process_invoice(self, message: Message):
        result = self.service.users().messages().get(userId="me", id=message.id).execute()
        message.subject = [
            header.get("value")
            for header in result["payload"]["headers"]
            if header.get("name") == "Subject" or header.get("name") == "subject"
        ][0]
        for part in result.get("payload", {}).get("parts", {}):
            file = File(part["filename"])
            if "pdf" in file.name.lower() and not message.has_existing(file):
                if "data" in part["body"]:
                    file_data = base64.urlsafe_b64decode(
                        part["body"]["data"].encode("UTF-8")
                    )

                elif "attachmentId" in part["body"]:
                    attachment = (
                        self.service.users()
                        .messages()
                        .attachments()
                        .get(
                            userId="me",
                            messageId=result["id"],
                            id=part["body"]["attachmentId"],
                        )
                        .execute()
                    )
                    file_data = base64.urlsafe_b64decode(
                        attachment["data"].encode("UTF-8")
                    )

                f = open(file.input_path, "wb")
                f.write(file_data)
                f.close()

                subprocess.run(
                    'gs -sOutputFile="'
                    + file.output_path
                    + '" -sDEVICE=pdfwrite -dPDFSETTINGS=/printer -dProcessColorModel=/DeviceGray -sColorConversionStrategy=Gray -dCompatibilityLevel=1.4 -dNOPAUSE -dBATCH "'
                    + file.input_path
                    + '"',
                    shell=True,
                )

                message.files.append(file)

    def forward_invoice(self, message: Message):
        send_msg = self.create_message_with_attachment(message)
        result = (
            self.service.users().messages().send(userId="me", body=send_msg).execute()
        )

    def cleanup(self, message: Message):
        for file in message.files:
            os.remove(file.input_path)
            os.remove(file.output_path)

    def archive(self, message: Message):
        body = {"removeLabelIds": [NEW_LABEL_ID], "addLabelIds": [ARCHIVE_LABEL_ID]}
        message = (
            self.service.users()
            .messages()
            .modify(userId="me", id=message.id, body=body)
            .execute()
        )

    def create_message_with_attachment(self, message: Message):
        new_msg = MIMEMultipart()
        new_msg["to"] = TO
        new_msg["from"] = FROM
        new_msg["subject"] = message.subject

        textpart = MIMEText(MAILTEXT)
        new_msg.attach(textpart)

        for file in message.files:
            fp = open(file.output_path, "rb")
            binpart = MIMEApplication(fp.read(), _sub_type="pdf")
            fp.close()

            binpart.add_header("Content-Disposition", "attachment", filename=file.name)
            new_msg.attach(binpart)

        return {"raw": base64.urlsafe_b64encode(new_msg.as_bytes()).decode()}


if __name__ == "__main__":
    ip = InvoiceProcessor()
    ip.run()
