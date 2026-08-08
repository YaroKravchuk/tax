import FreeSimpleGUI as PySimpleGUI
import os
import pandas as pd
import textwrap
from collections import namedtuple
from datetime import datetime
from utility import BookRecords, filter_loads

# How many matching projects the picker lists at once. A year sheet can hold thousands,
# so the list is capped and a line underneath says how many matches were left out
MAX_SUGGESTIONS = 8

WINDOW_TITLE = 'Prospect LLC - Driver Logs & Invoice'

# Widths in characters. Project IDs are full site addresses, so the picker is wide, and
# the lines naming the records file are wider still because a whole folder path has to fit
FIELD_WIDTH = 72
DATE_WIDTH = 34
SOURCE_WIDTH = 78

# Rows kept for the folder path. A path too long for one line is folded onto the second
# rather than being quietly cut off, since a half-shown path is what caused the trouble
SOURCE_PATH_ROWS = 2

# How often the records file is looked at again while the form sits open, in milliseconds.
# It is one look at the file's date, so it can be done often without costing anything
SOURCE_CHECK_MS = 2000

# Colours for the lines that report on what has been typed
HINT_COLOR = 'grey30'
ERROR_COLOR = 'firebrick'
CONFIRM_COLOR = 'darkgreen'

# What the form collects. records is the already-read workbook, carried along so the
# generator does not read it a second time
UserChoices = namedtuple('UserChoices', 'records sheet_name project_id start_date end_date taxable '
                                        'should_create_driver_logs should_create_invoice should_export_pdf')


class StatusWindow:
    """A small window that stays up while the workbooks are being built.

    Building the files, and especially handing them to Excel or LibreOffice to make the
    PDFs, takes long enough that an empty screen looks like a crash. This keeps a line
    of text in front of the user the whole time.
    """

    def __init__(self, message='Working...'):
        PySimpleGUI.ChangeLookAndFeel('GreenTan')
        self.window = PySimpleGUI.Window(
            WINDOW_TITLE,
            [[PySimpleGUI.Text(message, key='-STATUS-', size=(44, 2), font=('Helvetica', 12))]],
            element_padding=(14, 14), disable_close=True, finalize=True)

    def update(self, message):
        if self.window is not None:
            self.window['-STATUS-'].update(message)
            self.window.refresh()

    # Closing twice is harmless, which lets the caller close it early and still clean up
    def close(self):
        if self.window is not None:
            self.window.close()
            self.window = None


# Function to find the projects whose ID contains every word typed, in any order and in
# any case. Project IDs start with a house number, so a plain "starts with" match would
# hide the right project from anyone searching by street name
def matching_projects(projects, query):
    words = query.lower().split()
    if not words:
        return list(projects)
    return [project for project in projects
            if all(word in project.search_text for word in words)]


# Function to find the one project a typed ID names, ignoring case and stray spaces, so
# that a hand-typed address still lines up with the spelling used in the records
def find_project(projects, text):
    typed = text.strip().lower()
    if not typed:
        return None
    return next((project for project in projects if project.search_text == typed), None)


def format_date(date):
    return date.strftime('%m/%d/%Y')


# Function to say how long ago something happened in the way a person would say it.
# "11 minutes ago" is what catches a file that should have been saved and was not;
# a bare date and time on its own leaves that sum to be done in the head
def describe_age(when):
    minutes = (datetime.now() - when).total_seconds() / 60

    if minutes < 0:               # a clock difference, so do not claim to know
        return 'just now'
    if minutes < 2:
        return 'just now'
    if minutes < 60:
        return f'{round(minutes)} minutes ago'

    hours = minutes / 60
    if hours < 24:
        return 'an hour ago' if round(hours) == 1 else f'{round(hours)} hours ago'

    days = hours / 24
    if days < 30:
        return 'yesterday' if round(days) == 1 else f'{round(days)} days ago'

    months = days / 30
    return 'a month ago' if round(months) == 1 else f'{round(months)} months ago'


# Function to describe when the records file was last saved. Both readings are given: the
# plain date and time to check against, and how long ago that was to notice it is stale
def describe_source(records):
    when = records.last_saved
    hour = when.hour % 12 or 12
    stamp = when.strftime(f'%a %d %b %Y at {hour}:%M %p')
    return f'Load records last saved {describe_age(when)}   -   {stamp}'


# Function to fold the folder holding the records onto as many lines as it takes. A path
# is only worth showing if all of it shows, so nothing here is allowed to trail off
def describe_folder(records):
    folder = os.path.dirname(records.full_path)
    lines = textwrap.wrap(folder, width=SOURCE_WIDTH - 3, break_long_words=True) or ['']
    return 'in  ' + '\n    '.join(lines[:SOURCE_PATH_ROWS])


# Function to turn a typed date into a real one. Returns the date and an explanation,
# only one of which is ever filled in
def parse_date(text, field_name):
    if not text.strip():
        return pd.NaT, None
    try:
        return pd.to_datetime(text.strip()), None
    except Exception:
        return None, f'The {field_name} "{text.strip()}" is not a date. Try 08/05/2026.'


# Function to describe a project in one line, so the right one can be recognised before
# anything is generated
def describe_project(project):
    parts = []
    if pd.notna(project.customer):
        parts.append(str(project.customer))
    parts.append('1 load' if project.loads == 1 else f'{project.loads} loads')
    if pd.notna(project.last_date):
        parts.append(f'{format_date(project.first_date)} to {format_date(project.last_date)}')
    return '   |   '.join(parts)


# Function to say how the search is going: how many projects there are, how many match,
# and whether some matches are being held back
def describe_matches(project_count, match_count, query):
    if not query.strip():
        return f'{project_count} projects on this sheet - start typing to search'
    if match_count == 0:
        return 'No projects match'
    if match_count > MAX_SUGGESTIONS:
        return f'Showing {MAX_SUGGESTIONS} of {match_count} matches - keep typing to narrow it down'
    return '1 match' if match_count == 1 else f'{match_count} matches'


def build_window(records, sheet_names, default_sheet):
    PySimpleGUI.ChangeLookAndFeel('GreenTan')

    date_fields = [
        PySimpleGUI.Column([[PySimpleGUI.Text('Start Date')],
                            [PySimpleGUI.Input(key='-START-', size=(DATE_WIDTH, 1), enable_events=True)]],
                           pad=(0, 0)),
        PySimpleGUI.Column([[PySimpleGUI.Text('End Date')],
                            [PySimpleGUI.Input(key='-END-', size=(DATE_WIDTH, 1), enable_events=True)]],
                           pad=(0, 0)),
    ]

    layout = [
        [PySimpleGUI.Text('Generate Audit Daily Load Tickets!', font=('Helvetica', 22))],
        # Which file the loads came from and when it was last saved. Editing one copy of
        # the records while the program reads another is a hard mistake to spot from the
        # output, so the form says outright what it is working from
        [PySimpleGUI.Text(describe_source(records), key='-SOURCE TIME-',
                          size=(SOURCE_WIDTH, 2), text_color=HINT_COLOR)],
        [PySimpleGUI.Text(os.path.basename(records.full_path), key='-SOURCE FILE-',
                          size=(SOURCE_WIDTH, 1), text_color=HINT_COLOR)],
        [PySimpleGUI.Text(describe_folder(records), key='-SOURCE PATH-',
                          size=(SOURCE_WIDTH, SOURCE_PATH_ROWS), text_color=HINT_COLOR)],
        [PySimpleGUI.Text('_' * 80)],
        [PySimpleGUI.Text('"Dump Trucking" Year Sheet')],
        [PySimpleGUI.Combo(sheet_names, default_value=default_sheet, size=(FIELD_WIDTH, 1),
                           key='-SHEET-', enable_events=True, readonly=True)],
        [PySimpleGUI.Text('Project ID')],
        [PySimpleGUI.Input(key='-PROJECT-', size=(FIELD_WIDTH, 1), enable_events=True)],
        [PySimpleGUI.Listbox([], key='-MATCHES-', size=(FIELD_WIDTH, MAX_SUGGESTIONS),
                             enable_events=True, no_scrollbar=False)],
        [PySimpleGUI.Text('', key='-MATCH COUNT-', size=(FIELD_WIDTH, 1), text_color=HINT_COLOR)],
        [PySimpleGUI.Text('', key='-PROJECT INFO-', size=(FIELD_WIDTH, 1), text_color=CONFIRM_COLOR)],
        date_fields,
        [PySimpleGUI.Text('', key='-DATE HINT-', size=(FIELD_WIDTH, 1), text_color=HINT_COLOR)],
        [PySimpleGUI.Checkbox('Taxable', key='-TAXABLE-')],
        [PySimpleGUI.Checkbox('Include Driver Logs', key='-DRIVER LOGS-', default=True)],
        [PySimpleGUI.Checkbox('Include Invoice', key='-INVOICE-', default=True)],
        [PySimpleGUI.Checkbox('Also Save as PDF', key='-PDF-', default=True)],
        [PySimpleGUI.Text('_' * 80)],
        [PySimpleGUI.Text('', key='-ERROR-', size=(FIELD_WIDTH, 2), text_color=ERROR_COLOR)],
        [PySimpleGUI.Submit(), PySimpleGUI.Cancel()],
    ]

    window = PySimpleGUI.Window(WINDOW_TITLE, layout, finalize=True)

    # The arrow keys walk the suggestions and Return takes one, all without the caret
    # ever leaving the text box, so typing can carry straight on
    window['-PROJECT-'].bind('<Down>', '+DOWN')
    window['-PROJECT-'].bind('<Up>', '+UP')
    window['-PROJECT-'].bind('<Return>', '+ENTER')
    window['-PROJECT-'].set_focus()

    return window


# Function to redraw the suggestion list and the two lines under it for whatever has been
# typed so far. Nothing here ever appears or disappears, so the form never jumps about
def show_matches(window, projects, query, highlighted=None):
    matches = matching_projects(projects, query)
    shown = matches[:MAX_SUGGESTIONS]

    window['-MATCHES-'].update([project.id for project in shown])
    if highlighted is not None and shown:
        window['-MATCHES-'].update(set_to_index=highlighted, scroll_to_index=highlighted)
    window['-MATCH COUNT-'].update(describe_matches(len(projects), len(matches), query))

    chosen = find_project(projects, query)
    window['-PROJECT INFO-'].update(describe_project(chosen) if chosen else '')

    return shown


# Function to put a project into the box and fill the dates in with the range it ran over,
# which is both a starting point and a hint at what dates are worth asking for
def take_project(window, projects, project):
    window['-PROJECT-'].update(project.id)

    if pd.notna(project.first_date):
        window['-START-'].update(format_date(project.first_date))
        window['-END-'].update(format_date(project.last_date))
        window['-DATE HINT-'].update('Filled in with everything this project has. '
                                     'Narrow it down for part of the job.')
    else:
        window['-START-'].update('')
        window['-END-'].update('')
        window['-DATE HINT-'].update('')

    window['-ERROR-'].update('')
    return show_matches(window, projects, project.id)


# Function to work out what has been asked for, or what is stopping the run from starting.
# Everything the form can catch is caught here, in the form, where it can still be fixed
def read_choices(values, records, projects):
    project = find_project(projects, values['-PROJECT-'])
    if project is None:
        typed = values['-PROJECT-'].strip()
        if not typed:
            return None, 'Choose a project first - start typing an address and pick it from the list.'
        return None, f'No project on this sheet matches "{typed}". Pick one from the list.'

    start_date, start_error = parse_date(values['-START-'], 'start date')
    if start_error:
        return None, start_error

    end_date, end_error = parse_date(values['-END-'], 'end date')
    if end_error:
        return None, end_error

    if pd.notna(start_date) and pd.notna(end_date) and start_date > end_date:
        return None, 'The start date is after the end date.'

    if not values['-DRIVER LOGS-'] and not values['-INVOICE-']:
        return None, ('Nothing was chosen to create. '
                      'Tick "Include Driver Logs" or "Include Invoice".')

    sheet_name = values['-SHEET-']
    if filter_loads(records, sheet_name, project.id, start_date, end_date).empty:
        return None, 'That project has no loads between those dates. Widen the date range.'

    return UserChoices(records, sheet_name, project.id, start_date, end_date, values['-TAXABLE-'],
                       values['-DRIVER LOGS-'], values['-INVOICE-'], values['-PDF-']), None


def collect_UI_input():
    # Reading the records can take a moment on a full year, so something is on screen for it
    loading = StatusWindow('Reading the load records...')
    try:
        records = BookRecords()
        sheet_names = records.year_sheet_names()
        if not sheet_names:
            raise ValueError('The load records hold no "Dump Trucking" year sheet.')
        default_sheet = records.default_year_sheet()
        projects = records.projects(default_sheet)
    finally:
        loading.close()

    window = build_window(records, sheet_names, default_sheet)
    shown = show_matches(window, projects, '')
    highlighted = None

    while True:
        event, values = window.read(timeout=SOURCE_CHECK_MS)

        if event in (PySimpleGUI.WIN_CLOSED, 'Cancel'):
            window.close()
            return None

        submitting = event == 'Submit'

        if event == PySimpleGUI.TIMEOUT_KEY:
            # Saving the records after this program read them leaves the form offering
            # yesterday's data, which is exactly the confusion this line exists to prevent
            if records.has_been_saved_since_read():
                window['-SOURCE TIME-'].update('The load records have been saved again since this '
                                               'program read them.\nClose this and run it again to '
                                               'work from the new data.', text_color=ERROR_COLOR)
            else:
                window['-SOURCE TIME-'].update(describe_source(records), text_color=HINT_COLOR)

        elif event == '-SHEET-':
            window['-MATCH COUNT-'].update('Reading that year...')
            window.refresh()
            projects = records.projects(values['-SHEET-'])
            highlighted = None
            shown = show_matches(window, projects, values['-PROJECT-'])

        elif event == '-PROJECT-':
            highlighted = None
            shown = show_matches(window, projects, values['-PROJECT-'])
            window['-ERROR-'].update('')

        elif event in ('-PROJECT-+DOWN', '-PROJECT-+UP') and shown:
            step = 1 if event.endswith('DOWN') else -1
            highlighted = 0 if highlighted is None else min(max(highlighted + step, 0), len(shown) - 1)
            window['-MATCHES-'].update(set_to_index=highlighted, scroll_to_index=highlighted)

        elif event == '-PROJECT-+ENTER':
            if highlighted is not None:
                shown = take_project(window, projects, shown[highlighted])
                highlighted = None
            elif find_project(projects, values['-PROJECT-']):
                # The box already names a real project, so Return means "go"
                submitting = True
            elif shown and values['-PROJECT-'].strip():
                # Something was typed and it narrows to a list, so Return takes the best match.
                # An empty box is left alone, since there is nothing there to have meant
                shown = take_project(window, projects, shown[0])

        elif event == '-MATCHES-' and values['-MATCHES-']:
            picked = find_project(projects, values['-MATCHES-'][0])
            if picked:
                shown = take_project(window, projects, picked)
                highlighted = None

        elif event in ('-START-', '-END-'):
            field_name = 'start date' if event == '-START-' else 'end date'
            _, error = parse_date(values[event], field_name)
            window['-ERROR-'].update(error or '')

        if submitting:
            choices, error = read_choices(values, records, projects)
            if choices is not None:
                window.close()
                return choices
            window['-ERROR-'].update(error)
