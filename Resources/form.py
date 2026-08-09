import FreeSimpleGUI as PySimpleGUI
import pandas as pd
import platform
import tkinter
from collections import namedtuple
from datetime import datetime
from utility import BookRecords, filter_loads

# How many matching projects the search shows at once. A year sheet can hold thousands,
# so the list is capped and its last line says how many were left out
MAX_SUGGESTIONS = 8

WINDOW_TITLE = 'Prospect LLC - Driver Logs & Invoice'

# Widths in characters. Project IDs are full site addresses, so the search box is wide,
# while a date is ten characters and a box any wider than that just looks unfinished
FIELD_WIDTH = 62
DATE_WIDTH = 14

# Room kept for the line that explains why the run cannot start. It holds one line, so
# the messages below are written to fit rather than leaving a band of empty form
ERROR_LIMIT = FIELD_WIDTH + 6

# What stands between the parts of the line describing a project. Those parts carry
# spaces of their own now - "Customer: All Terrain" - so a wider gap on its own no longer
# reads as the join between one part and the next. A typed "|" would, at this size, read
# as a stray 1 among the numbers, so the divider is a drawn rule: a hairline the height
# of the line, in the same colour as the rules that divide the form. Both are in pixels
INFO_RULE_WIDTH = 1
INFO_RULE_GAP = 8

# How often the records file is looked at again while the form sits open, in milliseconds.
# It is one look at the file's date, so it can be done often without costing anything
SOURCE_CHECK_MS = 2000

# The old GreenTan colours - sage page, cream fields, deep green buttons - kept, but
# laid on flat. What dated the old form was the moulding and the outlines, not the green
PAGE_COLOR = '#9FB8AD'
FIELD_COLOR = '#F7F3EC'
TEXT_COLOR = '#1B2620'
MUTED_COLOR = '#3A4842'
LINE_COLOR = '#7E9589'
ACCENT_COLOR = '#475841'
ERROR_COLOR = '#7A1D14'
CONFIRM_COLOR = '#1F4A2B'

THEME_NAME = 'ProspectGreen'
THEME = {
    'BACKGROUND': PAGE_COLOR, 'TEXT': TEXT_COLOR,
    'INPUT': FIELD_COLOR, 'TEXT_INPUT': TEXT_COLOR,
    'SCROLL': LINE_COLOR, 'BUTTON': ('#FFFFFF', ACCENT_COLOR),
    'PROGRESS': (ACCENT_COLOR, LINE_COLOR),
    'BORDER': 0, 'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 0,
    'ACCENT1': ACCENT_COLOR, 'ACCENT2': LINE_COLOR, 'ACCENT3': MUTED_COLOR,
}

# What the form collects. records is the already-read workbook, carried along so the
# generator does not read it a second time
UserChoices = namedtuple('UserChoices', 'records sheet_name project_id start_date end_date taxable '
                                        'should_create_driver_logs should_create_invoice should_export_pdf')


# Function to pick the everyday interface font of whichever computer this is running on,
# since nothing dates a window faster than a font nobody else uses
def ui_font(step=0, weight=''):
    system = platform.system()
    if system == 'Windows':
        family, size = 'Segoe UI', 10
    elif system == 'Darwin':
        family, size = 'Helvetica Neue', 13
    else:
        family, size = 'DejaVu Sans', 10
    return (family, size + step, weight) if weight else (family, size + step)


def use_theme():
    if THEME_NAME not in PySimpleGUI.LOOK_AND_FEEL_TABLE:
        PySimpleGUI.theme_add_new(THEME_NAME, THEME)
    PySimpleGUI.ChangeLookAndFeel(THEME_NAME)


class StatusWindow:
    """A small window that stays up while the workbooks are being built.

    Building the files, and especially handing them to Excel or LibreOffice to make the
    PDFs, takes long enough that an empty screen looks like a crash. This keeps a line
    of text in front of the user the whole time.
    """

    def __init__(self, message='Working...'):
        use_theme()
        self.window = PySimpleGUI.Window(
            WINDOW_TITLE,
            [[PySimpleGUI.Text(message, key='-STATUS-', size=(42, 2), font=ui_font(1))]],
            element_padding=(16, 16), disable_close=True, finalize=True)

    def update(self, message):
        if self.window is not None:
            self.window['-STATUS-'].update(message)
            self.window.refresh()

    # Closing twice is harmless, which lets the caller close it early and still clean up
    def close(self):
        if self.window is not None:
            self.window.close()
            self.window = None


class SuggestionList:
    """The matching projects, floating just under the search box.

    The list is lifted out of the form's own layout and positioned by hand, so it can
    cover what is beneath it the way a search box on a website does. Left in the layout
    it would sit there empty all the time, and shove the rest of the form up and down as
    it filled and emptied.
    """

    def __init__(self, window):
        self.window = window
        self.listbox = window['-MATCHES-']
        # FreeSimpleGUI gives every row of a layout its own frame, and it is that frame,
        # not the column inside it, that holds the space in the form
        self.frame = window['-SUGGESTIONS-'].Widget.master
        self.container = self.frame.master
        self.anchor = window['-PROJECT-'].Widget
        self.projects = []
        self.highlighted = None

        self.frame.pack_forget()
        window.TKroot.geometry('')          # let the form close up the gap it left

    @property
    def showing(self):
        return bool(self.frame.winfo_ismapped())

    def show(self, projects, hidden_count):
        self.projects = list(projects)
        rows = [project.id for project in self.projects]
        if hidden_count:
            rows.append(f'   ... and {hidden_count} more, keep typing')

        self.listbox.update(rows)
        self.listbox.Widget.configure(height=len(rows))   # only as tall as it needs to be
        self.highlighted = None
        self._place()

    def hide(self):
        if self.showing:
            self.frame.place_forget()
        self.highlighted = None

    # Function to sit the list flush under the search box. Where exactly a frame lands
    # depends on padding inside it, so it is put down once, measured, and corrected
    def _place(self):
        left = self.anchor.winfo_rootx()
        top = self.anchor.winfo_rooty() + self.anchor.winfo_height()
        x = left - self.container.winfo_rootx()
        y = top - self.container.winfo_rooty()

        self.frame.place(x=x, y=y)
        self.frame.lift()
        self.window.TKroot.update_idletasks()

        drift_x = self.listbox.Widget.winfo_rootx() - left
        drift_y = self.listbox.Widget.winfo_rooty() - top
        self.frame.place(x=x - drift_x, y=y - drift_y)
        self.frame.lift()

    # Function to walk the list with the arrow keys while the caret stays in the search
    # box, so that typing can carry straight on
    def step(self, direction):
        if not self.showing or not self.projects:
            return
        if self.highlighted is None:
            self.highlighted = 0 if direction > 0 else len(self.projects) - 1
        else:
            self.highlighted = min(max(self.highlighted + direction, 0), len(self.projects) - 1)
        self.listbox.update(set_to_index=self.highlighted, scroll_to_index=self.highlighted)

    # Function to give back the project the list is pointing at, if it is pointing at one.
    # The last line can be the "and 12 more" note, which names no project
    def chosen(self, clicked=None):
        if clicked is not None:
            return next((project for project in self.projects if project.id == clicked), None)
        if self.highlighted is None:
            return None
        return self.projects[self.highlighted]


def format_date(date):
    return date.strftime('%m/%d/%Y')


# Function to say how long ago something happened in the way a person would say it.
# "11 minutes ago" is what catches a file that should have been saved and was not
def describe_age(when):
    minutes = (datetime.now() - when).total_seconds() / 60

    if minutes < 2:                  # covers a clock a little out of step, too
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


def describe_source(records):
    return f'Records last saved {describe_age(records.last_saved)}'


# Function to find the projects whose ID contains every word typed, in any order and in
# any case. Project IDs start with a house number, so a plain "starts with" match would
# hide the right project from anyone searching by street name
def matching_projects(projects, query):
    words = query.lower().split()
    if not words:
        return list(projects)          # an empty box offers the most recent work
    return [project for project in projects
            if all(word in project.search_text for word in words)]


# Function to find the one project a typed ID names, ignoring case and stray spaces, so
# that a hand-typed address still lines up with the spelling used in the records
def find_project(projects, text):
    typed = text.strip().lower()
    if not typed:
        return None
    return next((project for project in projects if project.search_text == typed), None)


# Function to describe a project, in the parts the one line is built from, so the right
# one can be recognised before anything is generated. The count says how much work there
# is and the dates say when it was done - the first load to the last, which is every load
# the records hold. What divides the parts is drawn, not written, so they are handed back
# as they are rather than joined up here
def describe_project(project):
    parts = []
    if pd.notna(project.customer):
        parts.append(f'Customer: {project.customer}')
    parts.append('1 load' if project.loads == 1 else f'{project.loads} loads')
    if pd.notna(project.last_date):
        first, last = format_date(project.first_date), format_date(project.last_date)
        # Most projects are a single day's hauling, and naming that day twice reads as
        # though something has gone wrong with the dates
        parts.append(f'Worked {first}' if first == last else f'Worked {first} to {last}')
    return parts


# Function to write the line under the search box, with a hairline standing between one
# part and the next. Emptying the line destroys the rules standing in it, so every write
# raises its own; asking each to stretch leaves it exactly as tall as the writing
def show_project_info(window, parts, color=CONFIRM_COLOR):
    line = window['-PROJECT INFO-'].Widget
    line.configure(state='normal')
    line.delete('1.0', 'end')

    for index, part in enumerate(parts):
        if index:
            rule = tkinter.Frame(line, width=INFO_RULE_WIDTH, background=LINE_COLOR)
            line.window_create('end', window=rule, padx=INFO_RULE_GAP, stretch=True)
        line.insert('end', part)

    line.configure(state='disabled', foreground=color)


# Function to turn a typed date into a real one. Returns the date and an explanation,
# only one of which is ever filled in
def parse_date(text, field_name):
    if not text.strip():
        return pd.NaT, None
    try:
        return pd.to_datetime(text.strip()), None
    except Exception:
        return None, f'The {field_name} "{shorten(text.strip(), 20)}" is not a date. Try 08/05/2026.'


# Function to keep whatever was typed from pushing a message off the end of its one line.
# The result is never longer than the limit, the ellipsis included
def shorten(text, limit):
    if len(text) <= limit:
        return text
    return text[:max(limit - 3, 0)].rstrip() + '...'


def build_window(records, sheet_names, default_sheet):
    use_theme()

    def label(text):
        return PySimpleGUI.Text(text, font=ui_font(), text_color=MUTED_COLOR, pad=((0, 0), (6, 2)))

    date_fields = [
        PySimpleGUI.Column([[label('Start date')],
                            [PySimpleGUI.Input(key='-START-', size=(DATE_WIDTH, 1),
                                               font=ui_font(), enable_events=True, pad=(0, 0))]],
                           pad=((0, 18), 0)),
        PySimpleGUI.Column([[label('End date')],
                            [PySimpleGUI.Input(key='-END-', size=(DATE_WIDTH, 1),
                                               font=ui_font(), enable_events=True, pad=(0, 0))]],
                           pad=(0, 0)),
    ]

    layout = [
        [PySimpleGUI.Text('Generate Audit Daily Load Tickets', font=ui_font(9), pad=((0, 0), (0, 2)))],
        [PySimpleGUI.Text(describe_source(records), key='-SOURCE-', size=(FIELD_WIDTH, 1),
                          font=ui_font(-1), text_color=MUTED_COLOR, pad=((0, 0), (0, 8)))],
        [PySimpleGUI.HorizontalSeparator(color=LINE_COLOR, pad=((0, 0), (0, 4)))],

        [label('Year sheet')],
        [PySimpleGUI.Combo(sheet_names, default_value=default_sheet, size=(FIELD_WIDTH, 1),
                           key='-SHEET-', font=ui_font(), enable_events=True, readonly=True, pad=(0, 0))],

        [label('Project')],
        [PySimpleGUI.Input(key='-PROJECT-', size=(FIELD_WIDTH, 1), font=ui_font(),
                           enable_events=True, pad=(0, 0))],
        # Lifted out of the layout as soon as the window exists, and put back only while
        # something is being typed. See SuggestionList. The list is always drawn tall
        # enough for every match it holds, so there is never anything to scroll, and a
        # scrollbar cannot be drawn shorter than its own two arrows - which would hold a
        # single match open to twice the height of its one row
        [PySimpleGUI.Column([[PySimpleGUI.Listbox([], key='-MATCHES-', size=(FIELD_WIDTH, 1),
                                                  font=ui_font(), enable_events=True, pad=(0, 0),
                                                  no_scrollbar=True, background_color=FIELD_COLOR,
                                                  text_color=TEXT_COLOR)]],
                            key='-SUGGESTIONS-', pad=(0, 0))],
        # Text rather than a label, which holds one colour of writing and nothing else.
        # The rules dividing this line are drawn inside it. See show_project_info
        [PySimpleGUI.Multiline('', key='-PROJECT INFO-', size=(FIELD_WIDTH + 6, 1),
                               font=ui_font(-1), text_color=CONFIRM_COLOR,
                               background_color=PAGE_COLOR, border_width=0,
                               no_scrollbar=True, disabled=True, pad=((0, 0), (4, 1)))],

        date_fields,

        [PySimpleGUI.Checkbox('Taxable', key='-TAXABLE-', font=ui_font(), pad=((0, 14), (10, 0))),
         PySimpleGUI.Checkbox('Driver logs', key='-DRIVER LOGS-', default=True, font=ui_font(),
                              pad=((0, 14), (10, 0))),
         PySimpleGUI.Checkbox('Invoice', key='-INVOICE-', default=True, font=ui_font(),
                              pad=((0, 14), (10, 0))),
         PySimpleGUI.Checkbox('Also save as PDF', key='-PDF-', default=True, font=ui_font(),
                              pad=((0, 0), (10, 0)))],

        [PySimpleGUI.HorizontalSeparator(color=LINE_COLOR, pad=((0, 0), (14, 3)))],
        [PySimpleGUI.Text('', key='-ERROR-', size=(ERROR_LIMIT, 1),
                          font=ui_font(-1), text_color=ERROR_COLOR, pad=((0, 0), (0, 3)))],
        # Plain Buttons rather than Submit/Cancel, which give no way to drop the border
        [PySimpleGUI.Button('Submit', font=ui_font(), border_width=0, pad=((0, 8), 0),
                            size=(10, 1)),
         PySimpleGUI.Button('Cancel', font=ui_font(), border_width=0, size=(10, 1),
                            button_color=(TEXT_COLOR, PAGE_COLOR))],
    ]

    window = PySimpleGUI.Window(WINDOW_TITLE, layout, margins=(22, 18), finalize=True)

    # The arrow keys walk the suggestions and Return takes one, all without the caret
    # ever leaving the search box. Escape puts the list away
    window['-PROJECT-'].bind('<Down>', '+DOWN')
    window['-PROJECT-'].bind('<Up>', '+UP')
    window['-PROJECT-'].bind('<Return>', '+ENTER')
    window['-PROJECT-'].bind('<Escape>', '+ESCAPE')
    # Clicking into the box, or tabbing into it, drops the list down straight away rather
    # than waiting for a first letter to be typed
    window['-PROJECT-'].bind('<Button-1>', '+FOCUS')
    window['-PROJECT-'].bind('<FocusIn>', '+FOCUS')

    flatten(window)
    window['-PROJECT-'].set_focus()
    return window


# Function to take the moulded, outlined look off the stock widgets. Their sunken borders
# are what make an otherwise plain form look like a much older one
def flatten(window):
    for key in ('-PROJECT-', '-START-', '-END-'):
        window[key].Widget.configure(relief='flat', highlightthickness=1,
                                     highlightbackground=LINE_COLOR, highlightcolor=ACCENT_COLOR)

    window['-MATCHES-'].Widget.configure(relief='flat', borderwidth=0, highlightthickness=1,
                                         highlightbackground=LINE_COLOR, highlightcolor=LINE_COLOR,
                                         activestyle='none', selectbackground=ACCENT_COLOR,
                                         selectforeground='#FFFFFF')
    # The list is handed a frame of its own that keeps the plain system background, which
    # on a dark desktop is black. Painted, no edge of it can show as a bar behind the rows
    window['-MATCHES-'].Widget.master.configure(background=FIELD_COLOR)

    # Stripped of the border, padding and caret a text widget is given, the project line
    # sits exactly where the plain label it replaced used to, and cannot be typed in
    window['-PROJECT INFO-'].Widget.configure(relief='flat', borderwidth=0, highlightthickness=0,
                                              padx=0, pady=0, wrap='none', cursor='arrow',
                                              takefocus=0)

    for key in ('-TAXABLE-', '-DRIVER LOGS-', '-INVOICE-', '-PDF-'):
        window[key].Widget.configure(relief='flat', highlightthickness=0, activebackground=PAGE_COLOR)


# Function to redraw the search results for whatever has been typed so far
def show_matches(window, suggestions, projects, query, searching=False):
    chosen = find_project(projects, query)
    matches = matching_projects(projects, query)

    # Once the box holds a project's whole name there is nothing left to choose between.
    # An empty box only drops the list down while the box is actually being used, so the
    # form does not open with a list hanging over it
    if chosen is not None or (not query.strip() and not searching):
        suggestions.hide()
    elif matches:
        suggestions.show(matches[:MAX_SUGGESTIONS], max(len(matches) - MAX_SUGGESTIONS, 0))
    else:
        suggestions.hide()

    if chosen is not None:
        show_project_info(window, describe_project(chosen))
    elif query.strip() and not matches:
        show_project_info(window, ['No projects match'], MUTED_COLOR)
    else:
        show_project_info(window, [])


# Function to put a project into the box and fill the dates in with the range it ran over,
# which is both a starting point and a hint at what dates are worth asking for
def take_project(window, suggestions, projects, project):
    window['-PROJECT-'].update(project.id)

    if pd.notna(project.first_date):
        window['-START-'].update(format_date(project.first_date))
        window['-END-'].update(format_date(project.last_date))
    else:
        window['-START-'].update('')
        window['-END-'].update('')

    window['-ERROR-'].update('')
    show_matches(window, suggestions, projects, project.id)


# Function to work out what has been asked for, or what is stopping the run from starting.
# Everything the form can catch is caught here, in the form, where it can still be fixed
def read_choices(values, records, projects):
    project = find_project(projects, values['-PROJECT-'])
    if project is None:
        typed = values['-PROJECT-'].strip()
        if not typed:
            return None, 'Choose a project first, then pick it from the list.'
        return None, f'No project matches "{shorten(typed, 22)}". Pick one from the list.'

    start_date, start_error = parse_date(values['-START-'], 'start date')
    if start_error:
        return None, start_error

    end_date, end_error = parse_date(values['-END-'], 'end date')
    if end_error:
        return None, end_error

    if pd.notna(start_date) and pd.notna(end_date) and start_date > end_date:
        return None, 'The start date is after the end date.'

    if not values['-DRIVER LOGS-'] and not values['-INVOICE-']:
        return None, 'Nothing to create. Tick "Driver logs" or "Invoice".'

    sheet_name = values['-SHEET-']
    if filter_loads(records, sheet_name, project.id, start_date, end_date).empty:
        return None, 'No loads for that project between those dates.'

    return UserChoices(records, sheet_name, project.id, start_date, end_date, values['-TAXABLE-'],
                       values['-DRIVER LOGS-'], values['-INVOICE-'], values['-PDF-']), None


# Events that belong to the search box. Anything else means attention has moved on, so
# the floating list is put away rather than left hanging over the rest of the form
PICKER_EVENTS = ('-PROJECT-', '-PROJECT-+DOWN', '-PROJECT-+UP', '-PROJECT-+ENTER',
                 '-PROJECT-+FOCUS', '-MATCHES-')


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
    suggestions = SuggestionList(window)

    # The search box takes the caret as the window opens, and that counts as clicking into
    # it. Those opening events are thrown away so the form appears with nothing dropped
    # down over it, and the list waits for the box to actually be used
    window.TKroot.update_idletasks()
    window.TKroot.update()
    for _ in range(5):
        if window.read(timeout=0)[0] == PySimpleGUI.TIMEOUT_KEY:
            break

    while True:
        event, values = window.read(timeout=SOURCE_CHECK_MS)

        if event in (PySimpleGUI.WIN_CLOSED, 'Cancel'):
            window.close()
            return None

        if event not in PICKER_EVENTS:
            suggestions.hide()

        submitting = event == 'Submit'

        if event == PySimpleGUI.TIMEOUT_KEY:
            # Saving the records after this program read them leaves the form offering
            # yesterday's data, which is exactly what this line is here to catch
            if records.has_been_saved_since_read():
                window['-SOURCE-'].update('Records were saved again - close and run this once more',
                                          text_color=ERROR_COLOR)
            else:
                window['-SOURCE-'].update(describe_source(records), text_color=MUTED_COLOR)

        elif event == '-SHEET-':
            projects = records.projects(values['-SHEET-'])
            show_matches(window, suggestions, projects, values['-PROJECT-'])

        elif event in ('-PROJECT-', '-PROJECT-+FOCUS'):
            # Typing in the box, or just clicking into it, both count as searching, so
            # clearing the box goes back to offering the most recent projects
            show_matches(window, suggestions, projects, values['-PROJECT-'], searching=True)
            window['-ERROR-'].update('')

        elif event == '-PROJECT-+DOWN':
            suggestions.step(1)

        elif event == '-PROJECT-+UP':
            suggestions.step(-1)

        elif event == '-PROJECT-+ESCAPE':
            suggestions.hide()

        elif event == '-PROJECT-+ENTER':
            picked = suggestions.chosen()
            if picked is not None:
                take_project(window, suggestions, projects, picked)
            elif find_project(projects, values['-PROJECT-']):
                # The box already names a real project, so Return means "go"
                submitting = True
            elif suggestions.showing and suggestions.projects:
                take_project(window, suggestions, projects, suggestions.projects[0])

        elif event == '-MATCHES-' and values['-MATCHES-']:
            picked = suggestions.chosen(clicked=values['-MATCHES-'][0])
            if picked is not None:
                take_project(window, suggestions, projects, picked)
            else:
                suggestions.hide()          # the "and 12 more" line names no project

        elif event in ('-START-', '-END-'):
            field_name = 'start date' if event == '-START-' else 'end date'
            _, error = parse_date(values[event], field_name)
            window['-ERROR-'].update(error or '')

        if submitting:
            suggestions.hide()
            choices, error = read_choices(values, records, projects)
            if choices is not None:
                window.close()
                return choices
            window['-ERROR-'].update(error)
