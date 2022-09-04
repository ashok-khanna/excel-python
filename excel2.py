import numpy as np
import pandas as pd
import datetime as dt


def abs_(number):
    """Returns the absolute value of the supplied number. """
    return abs(number)


def concat(*args, joiner=""):
    concatenate(*args, joiner=joiner)


def concatenate(*args, joiner=""):
    """Joins the supplied arguments into a single string. """
    return joiner.join(args)


def covar(dataframe, column1, column2):
    """Returns the covariance of the two supplied columns."""
    return dataframe[column1].cov(dataframe[column2])


def date(year, month, day):
    """Returns a Python Datetime object for the supplied year, month and day."""
    return dt.datetime(year, month, day)


def datevalue(date):
    """Returns the Excel serial number for the supplied Date object."""
    # Specifying offset value i.e.,
    # the date value for the date of 1900-01-00
    offset = 693594
    return date.toordinal() - offset


def day(date):
    """Returns the day of the month of the supplied Date / Datetime object."""
    return date.day


# Reverse order of start / end arguments vs Excel API to be consistent with others
def days(start_date, end_date):
    """Returns the numbers of days between two dates."""
    delta = end_date - start_date
    return delta.days


def days360(start_date, end_date, method=False):
    """Returns the number of days between two dates using the 360 day calendar."""
    if start_date.year == end_date.year:
        month_difference = end_date.month - start_date.month
    else:
        year_difference = end_date.year - start_date.year
        month_difference = 12 * year_difference + end_date.month - start_date.month
    if method:
        return 30 * month_difference + min(end_date.day, 30) - min(start_date.day, 30) + \
               1 if end_date.day == 31 and start_date.day == 31 else 0
    else:
        return 30 * month_difference + min(end_date.day, 30) - min(start_date.day, 30)

    return (end_date - start_date).days


def edate(start_date, months):
    """Returns the date that is the specified number of months before or after the supplied date."""
    years = months // 12
    months = months % 12
    # Above is floor division, i.e. -1 // 3 = -1
    # To test the algorithm works: -1 // 12 = -1 & -1 % 12 = 11
    # Therefore date = 11 months after 1 year before start_date
    if start_date.month + months > 12:
        years += 1
        new_month = start_date.month + months - 12
    return date(start_date.year + years, new_month, start_date.day)


def eomonth(start_date, months):
    """Returns the date that is the specified number of months before or after the supplied date."""
    years = months // 12
    months = months % 12
    # Above is floor division, i.e. -1 // 3 = -1
    # To test the algorithm works: -1 // 12 = -1 & -1 % 12 = 11
    # Therefore date = 11 months after 1 year before start_date
    if start_date.month + months > 12:
        years += 1
        new_month = start_date.month + months - 12
    return date(start_date.year + years, new_month, dt.datetime.DaysInMonth(start_date.year + years, new_month))


# https://stackoverflow.com/questions/31359150/convert-date-from-excel-in-number-format-to-date-format-python
def exceldate(number):
    """Returns a Date object for the supplied serial number."""
    if isinstance(number, int):
        return dt.datetime.fromordinal(number + 693594)
    else:
        excel_date = dt.datetime.fromordinal(int(number))
        date = dt.datetime.fromordinal(dt.datetime(1900, 1, 1).toordinal() + int(excel_date) - 2)
        hour, minute, second = _floatHourToTime(excel_date % 1)
        return date.replace(hour=hour, minute=minute, second=second)


def false():
    """Returns False."""
    return False


# https://stackoverflow.com/questions/31359150/convert-date-from-excel-in-number-format-to-date-format-python
def _floatHourToTime(fh):
    hours, hourSeconds = divmod(fh, 1)
    minutes, seconds = divmod(hourSeconds * 60, 1)
    return (
        int(hours),
        int(minutes),
        int(seconds * 60),
    )


def hour(date):
    """Returns the hour of the supplied Date / Datetime object."""
    return date.hour


def if_(condition, true_value, false_value):
    """Returns the 'True Value' if 'Condition' is True and 'False Value' otherwise."""
    if condition:
        return true_value
    else:
        return false_value

def floor_(number):
    """Returns the largest integer less than or equal to the supplied number."""
    return np.floor(number)


# Need to have separate API for series & dataframes
def rank(dataframe, column, ascending=0):
    """Returns the rank of the supplied column."""
    pd_ascending = True if ascending == 1 else False
    return dataframe[column].rank(ascending=pd_ascending)


def var(dataframe, column):
    """Returns the variance of the supplied column."""
    return dataframe[column].var()


def varp(dataframe, column):
    """Returns the population variance of the supplied column."""
    return dataframe[column].var(ddof=0)

