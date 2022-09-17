class ChromDriverException(Exception):
    def __init__(self, message: str = 'Incorrect path to chrome driver, check location in constants.CHROMEDRIVER_PATH'):
        super().__init__(message)
