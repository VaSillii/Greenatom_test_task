from pathlib import Path

from setup import get_logger, get_email_info
from utils.excel import generate_report
from utils.operations import handler_open_browser
from utils.smtp import send_success_message

logger = get_logger(__file__)


def run_robot():
    """
    Launch robot
    :return:
    :rtype:
    """
    data = handler_open_browser()
    email_info = get_email_info()
    generate_report(data)
    path_email_template = Path().absolute().joinpath('template/email_template.html')
    send_success_message(
        path_email_template,
        email_info['email_login'],
        email_info['email_login'],
        email_info['email_subject'],
        min(len(data['USD_RUB']), len(data['JPY_RUB'])),
        # []
        email_info['email_copy']
    )


if __name__ == '__main__':
    logger.info('Start robot')
    run_robot()
    logger.info('End robot')
