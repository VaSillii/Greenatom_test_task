from datetime import date, time

from pydantic import BaseModel


class Report(BaseModel):
    """
    Report model
    :param date: Date of indicative exchange rate
    :param exchange_rate: Intraday clearing session
    :param time: Time of indicative exchange rate
    """
    date: date
    exchange_rate: float
    time: time
