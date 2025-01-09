from datetime import datetime

GMAIL_USER = 'vinay@axonator.com'
GMAIL_PASSWORD = '88$XDaG$34kf2%L2'
SUBJECT = "Inquiry For Demo"
START_DATE = datetime(2023, 1, 1)
END_DATE = datetime(2025, 7, 31)


startDate = START_DATE.strftime('%d-%b-%Y')
endDate = END_DATE.strftime('%d-%b-%Y')
FILTERQUERY = f'(SUBJECT "{SUBJECT}" SINCE "{startDate}" BEFORE "{endDate}")'