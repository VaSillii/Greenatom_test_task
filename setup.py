import logging
import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path().absolute()
PATH_ENV = BASE_DIR.joinpath('.env')
PATH_LOG_FILE = BASE_DIR.joinpath('log.txt')

if PATH_ENV.exists():
    load_dotenv(PATH_ENV)


def get_chromedriver_path() -> Path:
    """
    Get path to chromedriver.
    :return: Path object to chromedriver. By default, the path to the chrome driver in the project root is returned.
    :rtype: Path
    """
    path_chromedriver_env = os.environ.get('CHROMEDRIVER_PATH', None)
    if path_chromedriver_env:
        return Path(path_chromedriver_env)
    return Path().absolute().parent.joinpath('chromedriver.exe')


def get_email_info() -> dict:
    """
    Get dictionary with information about email gmail
    :return: Email information
    :rtype: dict
    """
    return {
        'email_login': os.environ.get('EMAIL_LOGIN', 'test@gmail.com'),
        'email_password': os.environ.get('EMAIL_PASSWORD', 'top_secret_password'),
        'email_copy': os.environ.get('EMAIL_COPY', '').split(),
        'email_subject': os.environ.get('EMAIL_SUBJECT', 'Отчет о работе робота')
    }


def get_logger(path: Path, path_log_file: Path = PATH_LOG_FILE) -> logging.Logger:
    """
    Getting the logger object
    :return: Logger object
    :rtype: logging.Logger
    """
    # create logger with 'spam_application'
    logger = logging.getLogger(str(path))
    logger.setLevel(logging.DEBUG)
    # create file handler which logs even debug messages
    fh = logging.FileHandler(path_log_file)
    # fh = logging.FileHandler(Path(path_dir, 'run_robot.log'))
    fh.setLevel(logging.DEBUG)
    # create console handler with a higher log level
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    # create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    # add the handlers to the logger
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


if __name__ == '__main__':
    print(get_email_info())
