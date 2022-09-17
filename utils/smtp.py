import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from setup import get_email_info, get_logger
from utils.excel import PATH_TO_FILE

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
logger = get_logger(__file__)


def get_recipient_string(recipient: list) -> str:
    """
    CC email string assembly
    :param recipient:
    :type recipient:
    :return:
    :rtype:
    """
    return ';'.join(set(recipient))


def _get_template_msg(path_to_msg: Path) -> str:
    """
    Get body email from file
    :param path_to_msg: Path to email template file
    :type path_to_msg: Path
    :return: Body email
    :rtype: str
    """
    with open(path_to_msg, encoding='utf-8') as f:
        return f.read()


def _prepare_html_body_msg(path_to_msg: Path, change_tuple_list: list) -> str:
    """
    Assembly of the body of the email
    :param path_to_msg: Path to email
    :type path_to_msg: Path
    :param change_tuple_list: List of tuples with substitutions
    :type change_tuple_list: list[tuple]
    :return: HTML body email
    :rtype: str
    """
    body_msg = _get_template_msg(path_to_msg)
    for val in change_tuple_list:
        body_msg.replace(*val)
    return body_msg


def get_template_msg(path_to_msg):
    """
    Get template message
    :param path_to_msg:
    :return:
    """
    with open(path_to_msg, encoding='utf-8') as f:
        return f.read()


def _get_html_body_email(path_to_msg: Path, change_tuple_list: list = []):
    """
    Get html body email with change from tuple list
    :param path_to_msg: Path to mail template
    :param change_tuple_list: List of tuples with substitutions
    :return: Html body email
    """
    new_html = get_template_msg(path_to_msg)
    for change_tuple in change_tuple_list:
        new_html = new_html.replace(*change_tuple)
    return new_html


def _word_declension(cnt_file_lines: int) -> str:
    """
    Word declination function depending on the number
    :param cnt_file_lines: Number of lines in file
    :type cnt_file_lines: int
    :return: The number of lines in the file in the desired case
    :rtype: str
    """
    string_declension_dict = {
        'one': 'строка',
        'two': 'строки',
        'three': 'строк'
    }
    cnt = cnt_file_lines % 100

    if cnt % 10 == 1 and cnt != 11:
        return f'{cnt_file_lines} {string_declension_dict["one"]}'
    elif 2 <= cnt % 10 <= 4 and (cnt > 20 or cnt < 5):
        return f'{cnt_file_lines} {string_declension_dict["two"]}'
    return f'{cnt_file_lines} {string_declension_dict["three"]}'


def _prepare_message(path_template: Path, sender_email: str, recipient: str, subject: str, cnt_file_lines: int,
                     copy_email: list) -> MIMEMultipart:
    """
    Preparing a MIMEMultipart Object
    :param path_template: Path to email template
    :type path_template: Path
    :param sender_email: Sender email
    :type sender_email: str
    :param recipient: Recipient email
    :type recipient: str
    :param subject: Subject email
    :type subject: str
    :param cnt_file_lines: Number of lines in file
    :type cnt_file_lines: int
    :param copy_email: CC email
    :type copy_email: list
    :return: MIMEMultipart object
    :rtype: MIMEMultipart
    """
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient
    msg["Subject"] = subject
    msg["CC"] = get_recipient_string(copy_email)
    change_tuple_list = [
        ('xxxResultxxx', _word_declension(cnt_file_lines))
    ]
    msg.attach(MIMEText(_get_html_body_email(path_template, change_tuple_list), 'html', 'utf-8'))
    msg = _add_attachment(msg)
    return msg


def send_success_message(path_to_email: Path, sender_email: str, recipient: str, subject_email: str,
                         cnt_file_lines: int, copy_emails: list = []):
    try:
        logger.info('Start send email')
        msg = _prepare_message(path_to_email, sender_email, recipient, subject_email, cnt_file_lines, copy_emails)
        _send(msg)
    except Exception:
        logger.info('An error occurred while sending the email', exc_info=True)
        logger.error('Error in sending email', exc_info=True)
    logger.info('End send email')


def _send(msg: MIMEMultipart, server: str = SMTP_SERVER, port: int = SMTP_PORT):
    """
    Function send email
    :param msg: MIMEMultipart object
    :type msg: MIMEMultipart
    :param server: Host server
    :type server: str
    :param port: Port server
    :type port: str
    :return:
    :rtype:
    """
    server = smtplib.SMTP(server, port)
    server.starttls()
    email_info = get_email_info()
    server.login(email_info['email_login'], email_info['email_password'])
    server.send_message(msg)
    server.quit()


def _add_attachment(msg: MIMEMultipart, path_to_attachment_file: Path = PATH_TO_FILE) -> MIMEMultipart:
    """
    Adding an attachment to an email
    :param msg: MIMEMultipart object
    :type msg: MIMEMultipart
    :param path_to_attachment_file: Path to attachment file
    :type path_to_attachment_file: Path
    :return: Update MIMEMultipart object
    :rtype: MIMEMultipart
    """
    if path_to_attachment_file.exists():
        with open(PATH_TO_FILE, "rb") as f:
            file = MIMEBase('application', 'octet-stream')
            file.set_payload(f.read())
            encoders.encode_base64(file)
        file.add_header('content-disposition', 'attachment', filename=path_to_attachment_file.name)
        msg.attach(file)
    return msg


if __name__ == "__main__":
    path_email_template = Path().absolute().joinpath('../template/email_template.html')
    # send_success_message(path_email_template, 'test@gmail.com', 'test@gmail.com', 'test', 3, [])
    print('В файле ' + str(_word_declension(100)))
