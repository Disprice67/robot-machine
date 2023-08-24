from base64 import b64decode
from email import message_from_bytes
from email.message import EmailMessage
from imaplib import IMAP4_SSL
from math import ceil
from os import path, remove, getcwd
from quopri import decodestring
from re import match, search
from smtplib import SMTP_SSL

from excel import File, Writer
from config import ROOT_DIR


class UnicodeReader:
    """encoded UTF-8."""

    def encoded(self, words: str):
        """Encode."""

        try:
            word_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='

            charset, encoding, encoded_text = match(word_regex, 
                                                    words).groups()
            if encoding == 'B':
                byte_string = b64decode(encoded_text)
            elif encoding == 'Q':
                byte_string = decodestring(encoded_text)
            return byte_string.decode(charset)
        except:
            return words
        

class InfoMessage:


    MESSAGE_FAILED = [("""\nПисьмо с темой: {SUBJECT}.
Не подходит для обработки,
скорее всего в письме нет файла!:("""),
                      ("""\nПисьмо с темой: {SUBJECT}.
В письме для обновления шаблона больше одного файла, 
пришлите один файл в формате '.xlsx'."""),
                      ("""\nПисьмо с темой: {SUBJECT}.
Файл не подходит для обработки.""")]

    MESSAGE_CORRECT = [("""\nПисьмо с темой: {SUBJECT}.
Шаблон успешно обновлен:)"""),
                       ("""\nОбработанный файл во вложении!:)
Название файла: {NAME}. Название обработанной 
страницы в исходном файле: {SHEETNAME}.\n""")]

    MESSAGE_INFO = [("""\nПозиций очень мало, поэтому повышающему 
коэффициенту верить не рекомендуется.
Требуется перепроверка доли ненайденных позиций вручную.\n"""),
                    ("""\nБольшой повышающий коэффициент. 
Требуется перепроверка доли ненайденных позиций вручную.\n""")]


    def __init__(self, subject: str = None,
                 filename = None,
                 sheetname = None):
        """the_object_is_responsible_for_creating_messages."""

        self.subject = subject
        self.filename = filename
        self.sheetname = sheetname
        

    def get_message(self, number, flag: bool = None, 
                    write: Writer = None, count: int = None):
        """Get_message."""

        if number == 0:
            return self.MESSAGE_FAILED[0].format(SUBJECT=self.subject)
        if number == 1:
            return self.MESSAGE_FAILED[1].format(SUBJECT=self.subject)
        if number == 2:
            return self.MESSAGE_CORRECT[0].format(SUBJECT=self.subject)
        if number == 3:
            return self.MESSAGE_FAILED[2].format(SUBJECT=self.subject, NAME=self.filename)
        if number == 4:
            return self.MESSAGE_CORRECT[1].format(NAME=self.filename, SHEETNAME=self.sheetname)
        if number == 5:
            return self.MESSAGE_INFO[0]
        if number == 6:
            return self.MESSAGE_INFO[1]
        
        if number == 8:
            return self.MESSAGE_INFO[3].format(time=ceil((6 * count) / 60))

    def finalbody(self, writer: Writer):
        """Create_final_message."""
        #const_rule
        STATIC_COUNT = 20
        body = self.get_message(4)  
        if writer.COUNT < STATIC_COUNT:
            body += self.get_message(5)
        return body


class Email:

    SUBJECT: str = None
    REPLY_EMAIL: str = None
    MSG_UID = None

    INPUT_FILE = None

    def __init__(self, mail_server: str, 
                 username: str, 
                 password: str
                 ):

        self.username = username
        self.password = password
        self.mail_server = mail_server

        self.connection = IMAP4_SSL(mail_server)
        self.connection.login(username, password)
        self.connection.select('INBOX')


    def delete_message(self):
        """Delete_messages."""

        self.connection.uid('STORE', self.MSG_UID, '+FLAGS', '(\Deleted)')


    def parse_uid(self, data: str):
        """Parse UID"""

        match = search(r'\d+ \(UID (?P<uid>\d+)\)', data)
        return match.group('uid')


    def send_email(self, 
                   body: str,
                   file = None,
                   filename = None):

        """Send email."""

        msg = EmailMessage()
        msg['From'] = self.username
        msg['Subject'] = self.SUBJECT
        msg['To'] = self.REPLY_EMAIL
        msg.set_content(body)

        if file is not None:

            with open(file, 'rb') as f:
                file_data = f.read()

            msg.add_attachment(file_data, 
                                maintype="application", 
                                subtype="xlsx",
                                filename=filename
                                )

        with SMTP_SSL(self.mail_server, 465) as smtp:
            smtp.login(self.username, self.password) 
            smtp.send_message(msg)


    def check_folder(self, folder: str):
        """Check_message_in_folder."""

        typ, items = self.connection.search(None, folder)
        items = items[0].split()

        try:
            emailid = items[0]
            return emailid
        except:
            return False
    

    def check_message(self, emailid):
        """Check_message."""

        encode = UnicodeReader()
        list_name: list = []
        resp, data = self.connection.fetch(emailid, "(RFC822)")
        u_resp, u_data = self.connection.fetch(emailid, "(UID)")

        self.MSG_UID = self.parse_uid(data=str(u_data[0]))

        #getting the mail content
        email_body = data[0][1]
        mail = message_from_bytes(email_body)

        #reply item
        self.REPLY_EMAIL = mail['From'].split()[-1]
        self.SUBJECT= encode.encoded(words=mail['Subject']).upper()
        msg = InfoMessage(self.SUBJECT)
        
        # is this part an attachment?
        for part in mail.walk():
            
            if part.get_filename() is None:
                continue
            
            filename = encode.encoded(words=part.get_filename())

            filename += '.xlsx'
            list_name.append((filename, part))

        if len(list_name) == 0:

            self.send_email(body=msg.get_message(0))
            self.delete_message()
            return False
        return list_name


    def save_attachments(self, folder: str, 
                         root_dir: str, filename,
                         part):
        """Save_attachments."""

        att_path = path.join(root_dir, filename)
        self.connection.uid('COPY', self.MSG_UID, folder)

        # finally write the stuff
        fp = open(att_path, 'wb')

        try:
            fp.write(part.get_payload(decode=True))
        except:
            pass

        fp.close()
        self.delete_message()


    def get_attachments(self, list_mail, 
                        file_main: File):
        """Get_attachment_in_messages"""

        msg = InfoMessage(self.SUBJECT)
        input_dir = ROOT_DIR + '\Входящий'
        for file in list_mail:
            
            filename, part = file
            if  self.SUBJECT == 'Обновление шаблона': #подумать как можно сделать обновление всех файлов через ключи:значение
                
                if len(list_mail) > 1:

                    #mail attribute
                    self.send_email(body=msg.get_message(1))
                    self.delete_message()

                    return False

                dashboard_file = ROOT_DIR + '\Шаблон'
                root = file_main.find_file(dashboard_file)
                remove(root)

                self.save_attachments('check_book', dashboard_file, filename, part)
                self.send_email(body=msg.get_message(2))

                return False

            else:
                self.save_attachments('requests', input_dir, filename, part)

        self.INPUT = file_main.find_file(input_dir)
        return self.INPUT


    def close_connection(self):
        """Close connection"""

        self.connection.close()
