import numpy as np
import pandas as pd
import datetime as dt


# https://docs.python.org/3/library/stdtypes.html#common-sequence-operations
# https://www.geeksforgeeks.org/operator-overloading-in-python/
# https://stackoverflow.com/questions/1957780/how-to-override-the-operator-in-python
# https://docs.python.org/3/reference/datamodel.html

# Function Counter: 8


def abs_(number):
    """Returns the absolute value of the supplied number. """
    return abs(number)


def adddays(date, days):
    """Returns the date that is the specified number of days before or after the supplied date."""
    return date + dt.timedelta(days=days)


def addyears(date, years):
    """Returns the date that is the specified number of years before or after the supplied date."""
    return date(date.year + years, date.month, date.day)


def and_(*args, index=0):
    """Returns True if all the arguments supplied are True and False otherwise."""
    if len(args) == index + 1:
        return args[index]
    else:
        head = args[index]
        index = index + 1
        if head:
            return and_(*args, index=index)
        else:
            return False


def averageifs(dataframe, average_column, *criteria):
    if len(criteria) % 2 != 0:
        raise SyntaxError("Supplied odd number of criteria arguments to SUMIFS.")
    conds = np.ones(dataframe.shape[0], dtype=bool)
    # floor division as otherwise len(criteria) / 2 returns float
    for i in range(len(criteria) // 2):
        cond = dataframe[criteria[i * 2]].eq(criteria[i * 2 + 1])
        conds &= cond
    return dataframe.loc[conds, average_column].mean()


def averageifs_(dataframe, average_column, conditions):
    return dataframe.loc[np.logical_and.reduce(conditions), average_column].mean()


def concatenate


def countifs(dataframe, count_column, *criteria):
    if len(criteria) % 2 != 0:
        raise SyntaxError("Supplied odd number of criteria arguments to COUNTIFS.")
    conds = np.ones(dataframe.shape[0], dtype=bool)
    # floor division as otherwise len(criteria) / 2 returns float
    for i in range(len(criteria) // 2):
        cond = dataframe[criteria[i * 2]].eq(criteria[i * 2 + 1])
        conds &= cond
    return dataframe.loc[conds, count_column].count()


def countifs_(dataframe, count_column, conditions):
    return dataframe.loc[np.logical_and.reduce(conditions), count_column].count()


def date(year, month, day):
    """Returns a Python Datetime object for the supplied year, month and day."""
    return dt.datetime(year, month, day)


# Partial implementation of Excel API
def datedif(start_date, end_date, unit):
    """Returns the number of units between two dates."""
    match unit:
        case "Y":
            return end_date.year - start_date.year
        case "M":
            return (end_date.year - start_date.year) * 12 + end_date.month - start_date.month
        case "D":
            return (end_date - start_date).days


def datevalue(date):
    """Returns the Excel serial number for the supplied Date object."""
    # Specifying offset value i.e.,
    # the date value for the date of 1900-01-00
    offset = 693594
    return date.toordinal() - offset








def isblank(value):
    """Returns True if the supplied value is blank."""
    return value == ""


def iseven(number):
    """Returns True if the supplied number is even."""
    return number % 2 == 0


def isnontext(value):
    """Returns True if the supplied value is not text."""
    return not isinstance(value, str)


def isnumber(value):
    """Returns True if the supplied value is a number."""
    return isinstance(value, (int, float))


def isodd(number):
    """Returns True if the supplied number is odd."""
    return number % 2 == 1


def isoweeknum(date):
    """Returns the ISO week number of the year corresponding to a date as an integer ranging from 1 to 53."""
    return date.isocalendar()[1]


def istext(value):
    """Returns True if the supplied value is text."""
    return isinstance(value, str)


def maxifs(dataframe, max_column, *criteria):
    if len(criteria) % 2 != 0:
        raise SyntaxError("Supplied odd number of criteria arguments to SUMIFS.")
    conds = np.ones(dataframe.shape[0], dtype=bool)
    # floor division as otherwise len(criteria) / 2 returns float
    for i in range(len(criteria) // 2):
        cond = dataframe[criteria[i * 2]].eq(criteria[i * 2 + 1])
        conds &= cond
    return dataframe.loc[conds, max_column].max()


def maxifs_(dataframe, max_column, conditions):
    return dataframe.loc[np.logical_and.reduce(conditions), max_column].max()


def minifs(dataframe, min_column, *criteria):
    if len(criteria) % 2 != 0:
        raise SyntaxError("Supplied odd number of criteria arguments to SUMIFS.")
    conds = np.ones(dataframe.shape[0], dtype=bool)
    # floor division as otherwise len(criteria) / 2 returns float
    for i in range(len(criteria) // 2):
        cond = dataframe[criteria[i * 2]].eq(criteria[i * 2 + 1])
        conds &= cond
    return dataframe.loc[conds, min_column].min()


def minifs_(dataframe, min_column, conditions):
    return dataframe.loc[np.logical_and.reduce(conditions), min_column].min()


def minute(date):
    """Returns the minute of the supplied Date / Datetime object."""
    return date.minute


def month(date):
    """Returns the month of the year of the supplied Date / Datetime object."""
    return date.month


def networkdays(start_date, end_date, holidays = []):
    """Returns the number of days between two dates excluding weekends and holidays."""
    return np.busday_count(start_date, end_date, holidays=holidays)


def not_(arg):
    """Reverses the logic of the argument."""
    if arg:
        return False
    else:
        return True


def now():
    """Returns a new Datetime object representing the current time."""
    return dt.datetime.now()


def or_(*args, index=0):
    """Returns True if any the arguments supplied are True and False otherwise."""
    if len(args) == index + 1:
        return args[index]
    else:
        head = args[index]
        index = index + 1
        if head:
            return True
        else:
            return or_(*args, index=index)


def second(date):
    """Returns the second of the supplied Date / Datetime object."""
    return date.second


def sum_(*args, index=0):
    """Returns the sum of the supplied arguments."""
    if len(args) == index + 1:
        return sum(args)
    else:
        head = sum_(args[index])
        return head + sum_(*args, index=index + 1)


def sumifs(dataframe, sum_column, *criteria):
    if len(criteria) % 2 != 0:
        raise SyntaxError("Supplied odd number of criteria arguments to SUMIFS.")
    conds = np.ones(dataframe.shape[0], dtype=bool)
    # floor division as otherwise len(criteria) / 2 returns float
    for i in range(len(criteria) // 2):
        cond = dataframe[criteria[i * 2]].eq(criteria[i * 2 + 1])
        conds &= cond
    return dataframe.loc[conds, sum_column].sum()


def sumifs_(dataframe, sum_column, conditions):
    return dataframe.loc[np.logical_and.reduce(conditions), sum_column].sum()


def time(hour, minute, second):
    """Returns the decimal number associated with the supplied time as a fraction of 1 day."""
    if hour > 23:
        hour = hour % 24
    if minute > 59:
        minute = minute % 60
    if second > 59:
        second = second % 60
    return hour / 24 + minute / (24 * 60) + second / (24 * 60 * 60)


def today():
    """Returns a Date object set for the start of today."""
    return dt.date.today()
    # return dt.datetime.combine(dt.date.today(), dt.time())


def true():
    """Returns True."""
    return True


def type_(value):
    """Returns the type of the supplied value."""
    return type(value)


def vlookup(dataframe, return_column, search_column, search_value):
    series = dataframe.loc[dataframe[search_column] == search_value]
    if series.empty:
        return "#N/A"
    else:
        return series[0]


# Typically returns 0 Monday - 6 Sunday
def weekday(date, return_type=1):
    """Returns the day of the week corresponding to a date as an integer ranging from 1 (Sunday) to 7 (Saturday)."""
    weekday = date.weekday()
    match return_type:
        case 1:
            # Sun = 1, Sat = 7
            return 1 if weekday == 6 else weekday + 2
        case 2:
            # Mon = 1, Sun = 7
            return weekday + 1
        case 3:
            # Mon = 0, Sun = 6
            return weekday
        case 11:
            # Mon = 1, Sun = 7
            return weekday + 1
        case 12:
            # Tues = 1, Mon = 7
            return 7 if weekday == 0 else weekday
        case 13:
            # Weds = 1, Tues = 7
            return 7 if weekday == 1 else 6 if weekday == 0 else weekday - 1
        case 14:
            # Thurs = 1, Weds = 7
            return 7 if weekday == 2 else 6 if weekday == 1 else 5 if weekday == 0 else weekday - 2
        case 15:
            # Fri = 1, Thurs = 7
            return 7 if weekday == 3 else 6 if weekday == 2 else 5 if weekday == 1 else 4 if weekday == 0 else weekday - 3
        case 16:
            # Sat = 1, Fri = 7
            return 7 if weekday == 4 else 6 if weekday == 3 else 5 if weekday == 2 else 4 if weekday == 1 else 3 if weekday == 0 else weekday - 4
        case 17:
            # Sun = 1, Sat = 7
            return 1 if weekday == 6 else weekday + 2
        case _:
            raise RuntimeError("Some logic error in the Weekday Function")


# Partial implementation
def weeknum(date, return_type=1):
    """Returns the week number of the year corresponding to a date as an integer ranging from 1 to 53."""
    return date.isocalendar()[1]


def workday(start_date, days, holidays = []):
    """Returns the date of the supplied number of working days after the supplied date."""
    return np.busday_offset(start_date, days, holidays=holidays)


def year(date):
    """Returns the year of the supplied Date object."""
    return date.year


def yearfrac(start_date, end_date, basis=0):
    delta = end_date - start_date
    match basis:
        case 0:
            return days360(start_date, end_date) / 360
        case 1:
            number_of_days_in_year = 366 if start_date.IsLeapYear() else 365
            return delta.days / number_of_days_in_year
        case 2:
            return delta.days / 360
        case 3:
            return delta.days / 365
        case 4:
            return days360(start_date, end_date, True) / 360
