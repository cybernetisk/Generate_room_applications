import openpyxl
from datetime import timedelta, datetime

# Day_of_the_week-year-week-number
WEEK_FROM = "1-2018-3"
WEEK_TO = "1-2018-23"

CLOCK_START = "18:00"
HOURS_OPEN = 8
XLS_ROW_BASE = 31


def daterange(start: datetime, stop: datetime, step=timedelta(days=7)):
    current = start
    while current < stop:
        yield current
        current += step


def main():
    date_start = datetime.strptime(WEEK_FROM, "%u-%Y-%W")
    date_stop = datetime.strptime(WEEK_TO, "%u-%Y-%W")
    for date in daterange(date_start, date_stop):
        dates = (date + timedelta(days=1), date + timedelta(days=4))
        generate_xls(dates)


def generate_xls(dates):
    wb = openpyxl.load_workbook("templates/template.xltx")
    sheet = wb['Sheet1']

    for row_offset, date in enumerate(dates):
        time_from = datetime.strptime(CLOCK_START, "%H:%M")
        date_from = datetime.combine(date.date(), time_from.time())
        date_to = date_from + timedelta(hours=HOURS_OPEN)
        date_string_from_to = f"{date_from.strftime('%Y-%m-%d')} - {date_to.strftime('%Y-%m-%d')}"
        time_string_from_to = f"{date_from.strftime('%H:%M')} - {date_to.strftime('%H:%M')}"
        print(date_string_from_to)

        row = XLS_ROW_BASE + row_offset
        target_line = sheet[f"A{row}:H{row}"][0]  # Fetch A31:H31 for example

        target_line[1].value = "OJD, IFI2"
        target_line[3].value = "Escape, 0711"
        target_line[4].value = date_string_from_to
        target_line[5].value = time_string_from_to

    wb.save("Escape-0711-week-" + dates[0].strftime('%W') + ".xltx")


if __name__ == '__main__':
    main()
