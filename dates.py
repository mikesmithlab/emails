import datetime
import dateutil.parser
from dateutil.relativedelta import relativedelta


def parse_date(datestring : str) -> datetime.datetime:
    """Takes date strings in many formats and returns a datetime object

    Args:
        datestring (str): Will handle most things, assumes English rather than American with day before month.

    Returns:
        datetime.datetime: date
    """
    return dateutil.parser.parse(datestring, dayfirst=True)

def now() -> datetime.datetime:
    """creates datetime object for current moment

    Returns:
        datetime.datetime: current date and time
    """
    return datetime.datetime.now()

def relative_datetime(date : datetime.datetime, delta_year=0, delta_month=0, delta_week=0, delta_day=0, delta_hour=0, delta_minute=0, delta_second=0) -> datetime.datetime:
    """Change datetime object by specific amount in years, months, weeks, days etc.
       Positive numbers move forward in time

    Args:
        date (datetime.datetime): original date
        delta_year (int, optional): change in years. Defaults to 0.
        delta_month (int, optional): change in months. Defaults to 0.
        delta_week (int, optional): change in weeks. Defaults to 0.
        delta_day (int, optional): change in days. Defaults to 0.
        delta_hour (int, optional): change in hours. Defaults to 0.
        delta_minute (int, optional): change in minutes. Defaults to 0.
        delta_second (int, optional): change in seconds. Defaults to 0.

    Returns:
        datetime.datetime: new date / time
    """
    new_date = date + relativedelta(years=delta_year, months=delta_month, weeks=delta_week, days=delta_day, hours=delta_hour, minutes=delta_minute, seconds=delta_second)
    return new_date


if __name__ =='__main__':
    print(str(parse_date('2nd march 2025')))
    print(now())
    print(relative_datetime(now(),delta_year=1))