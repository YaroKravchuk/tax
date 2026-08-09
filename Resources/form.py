import platform
import threading
import time
import tkinter
import pandas as pd
import ttkbootstrap as tb
from collections import namedtuple
from datetime import datetime
from ttkbootstrap.style import ThemeDefinition
from utility import BookRecords, filter_loads

# How many matching projects the search shows at once. A year sheet can hold thousands,
# so the list is capped and its last line says how many were left out
MAX_SUGGESTIONS = 8

WINDOW_TITLE = 'Prospect LLC - Driver Logs & Invoice'

# The one spelling of a date this program writes. The date boxes still take anything that
# can be typed, but everything the calendar fills in, and everything a picked project
# fills in, is written this way
DATE_FORMAT = '%m/%d/%Y'

# Widths in characters. Project IDs are full site addresses, so the search box is wide,
# while a date is ten characters and a box any wider than that just looks unfinished
FIELD_WIDTH = 62
DATE_WIDTH = 12

# The narrowest the form may open, in pixels. The widths above are counted in characters,
# and a box turns those into pixels using the computer's own default interface font -
# not the font named in ui_font, which is only what the text is drawn in. That default is
# 13 point on a Mac and 9 on Windows, so the identical form comes out 636 pixels wide here
# and markedly narrower there. This is the Mac width, which makes both open the same size.
# It is a floor and not a fixed width: a computer whose font asks for more still gets it,
# and the year sheet and project boxes, the two that stretch, take up the difference
MIN_WINDOW_WIDTH = 640

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

# How long the records are allowed to take to read before anything is put on screen to
# say so, in seconds. A small year sheet is read in a few hundredths of a second, and
# saying so took a window of its own that was gone again before it could be read - which
# is what made the program look like it opened twice. Nothing is shown for a read that
# quick now. A year sheet big enough to keep somebody waiting still says what it is doing
SLOW_READ_SECONDS = 0.4

# How long to wait between looks at a read that is still going, in seconds
READ_POLL_SECONDS = 0.02

# How wide a message is allowed to run before it wraps, in pixels. A list of the files
# just written is the longest thing ever shown in one, and its lines are file names
NOTICE_WIDTH = 460

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

# The edge drawn round the boxes that can be typed in. LINE_COLOR is only 1.5:1 against
# the sage, which left those boxes defined by their cream fill alone and their edges all
# but invisible. This is the same green carried far enough down to actually read as an edge
EDGE_COLOR = '#556458'

# The palette handed to ttkbootstrap, naming the colours above in the slots the widget
# styles look them up under. Every widget is themed from here, which is what replaced the
# old business of painting each one by hand and then filing its moulding off afterwards
THEME_NAME = 'ProspectGreen'
THEME_COLORS = {
    'primary': ACCENT_COLOR, 'secondary': LINE_COLOR, 'success': CONFIRM_COLOR,
    'info': MUTED_COLOR, 'warning': '#8A6D2F', 'danger': ERROR_COLOR,
    'light': FIELD_COLOR, 'dark': TEXT_COLOR,
    'bg': PAGE_COLOR, 'fg': TEXT_COLOR,
    'selectbg': ACCENT_COLOR, 'selectfg': '#FFFFFF',
    'border': EDGE_COLOR, 'inputfg': TEXT_COLOR, 'inputbg': FIELD_COLOR,
    'active': '#8FAA9E',
}

# What the form collects. records is the already-read workbook, carried along so the
# generator does not read it a second time
UserChoices = namedtuple('UserChoices', 'records sheet_name project_id start_date end_date taxable '
                                        'should_create_driver_logs should_create_invoice should_export_pdf')

# Everything the form is holding when Generate is pressed, before any of it has been checked
FormValues = namedtuple('FormValues', 'sheet_name project_text start_text end_text taxable '
                                      'driver_logs invoice pdf')


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


# The one theme, and the one root every window in this program hangs off. Tearing a Tk
# root down and building another one part way through a run is a known way to crash, so
# the form and the status window are both windows over a root made once and never shown
_style = None


def use_theme():
    global _style
    if _style is None:
        _style = tb.Style()
        if THEME_NAME not in _style.theme_names():
            _style.register_theme(ThemeDefinition(THEME_NAME, THEME_COLORS, mode='light'))
        _style.theme_use(THEME_NAME)
        _name_the_styles(_style)
        _style.master.withdraw()
    return _style


def app_root():
    return use_theme().master


# Function to settle the handful of things a palette on its own cannot say: what everything
# is written in, and how tall a row of the search list stands. The title is the one size
# that changes, and it is smaller than it was - it had been carrying more weight than the
# line under it, which is the line that actually changes
def _name_the_styles(style):
    style.configure('.', font=ui_font())
    for widget in ('TLabel', 'TButton', 'TCheckbutton', 'TEntry', 'TCombobox'):
        style.configure(widget, font=ui_font())

    # The themed edge round a box that can be typed in is drawn by a element that rounds
    # its corners by leaving the corner pixel unpainted, so what shows through there is
    # the cream fill - four pale specks at the corners of every box. Painting that edge
    # cream hides it altogether, and Field then draws a square one in its place. The state
    # maps go too, or the edge would come back in its own colour on focus, specks and all
    for widget in ('TEntry', 'TCombobox'):
        style.configure(widget, bordercolor=FIELD_COLOR, lightcolor=FIELD_COLOR,
                        darkcolor=FIELD_COLOR)
        style.map(widget, bordercolor=[], lightcolor=[], darkcolor=[])

    # The frame a date box lives in, which would otherwise show the sage of the page in
    # the gap around the box and its calendar button
    style.configure('Field.TFrame', background=FIELD_COLOR)

    style.configure('Title.TLabel', font=ui_font(6))
    style.configure('Field.TLabel', foreground=MUTED_COLOR)
    style.configure('Muted.TLabel', font=ui_font(-1), foreground=MUTED_COLOR)
    style.configure('Confirm.TLabel', font=ui_font(-1), foreground=CONFIRM_COLOR)
    style.configure('Error.TLabel', font=ui_font(-1), foreground=ERROR_COLOR)
    style.configure('Treeview', font=ui_font(), rowheight=24, borderwidth=0)


# Function to show a window, already the right size and already in the middle of the
# screen. A window appears the moment it is made, at whatever size Tk starts it off with,
# so one built in the ordinary way is seen three times before it settles: small in the
# corner, then the size of its contents, then moved. It is hidden the moment it is made
# instead, and shown here once, when both its size and its place are already settled
def show_centered(window, min_width=0):
    window.update_idletasks()
    # A window held to a minimum width opens at that width rather than the narrower one
    # its contents asked for, so that is the width to centre on
    width = max(window.winfo_reqwidth(), min_width)
    height = window.winfo_reqheight()
    x = (window.winfo_screenwidth() - width) // 2
    y = (window.winfo_screenheight() - height) // 3
    # Where it goes, but not how big it is. Asked for a size as well it would be held to
    # that size for good: the status window could not grow when its bar appears, and the
    # form could not grow when a message under the project box runs to a second line. A
    # minimum width is the way to widen a window without also pinning its height
    window.geometry(f'+{max(x, 0)}+{max(y, 0)}')
    window.deiconify()


class Field:
    """A box that can be typed in, inside an edge of this program's own drawing.

    The themed edge is turned off in _name_the_styles, and this puts a plain one-pixel
    frame round the box in its place: square at the corners, and one colour, which can be
    changed to follow the caret or to mark the box the run is being held up by.
    """

    def __init__(self, parent, build, focus_on=None):
        self.shell = tkinter.Frame(parent, background=EDGE_COLOR)
        self.widget = build(self.shell)
        self.widget.pack(fill='both', expand=True, padx=1, pady=1)

        # A date box is a box and a calendar button together, so the one that takes the
        # caret has to be named rather than assumed
        self.entry = focus_on(self.widget) if focus_on is not None else self.widget
        self.focused = False
        self.wrong = False
        self.entry.bind('<FocusIn>', lambda _e: self._paint(focused=True), add='+')
        self.entry.bind('<FocusOut>', lambda _e: self._paint(focused=False), add='+')

    def _paint(self, focused=None):
        if focused is not None:
            self.focused = focused
        if self.wrong:
            color = ERROR_COLOR
        elif self.focused:
            color = ACCENT_COLOR
        else:
            color = EDGE_COLOR
        self.shell.configure(background=color)

    def mark(self, wrong):
        self.wrong = wrong
        self._paint()


class StatusWindow:
    """A small window that stays up while the workbooks are being built.

    Building the files, and especially handing them to Excel or LibreOffice to make the
    PDFs, takes long enough that an empty screen looks like a crash. This keeps a few
    words in front of the user the whole time, and a bar under them whenever the work
    can actually be counted.
    """

    def __init__(self, message='Working'):
        use_theme()
        self.window = tb.Toplevel(title=WINDOW_TITLE, resizable=(False, False))
        self.window.withdraw()          # nothing is seen until it is built and placed
        self.window.protocol('WM_DELETE_WINDOW', lambda: None)   # nothing to cancel into

        body = tb.Frame(self.window, padding=(24, 20))
        body.pack(fill='both', expand=True)

        self.label = tb.Label(body, text=message, width=28)
        self.label.pack(anchor='w')

        # Put up and taken down as the work goes in and out of being countable, so the
        # window never shows a bar that is not measuring anything
        self.bar = tb.Progressbar(body, bootstyle='primary', length=260, mode='determinate')
        self.showing_bar = False

        show_centered(self.window)
        self.window.update()

    # Function to say what is happening now, and how far along it is when that is known.
    # Left without a total the bar goes away, rather than sitting still and looking stuck
    def update(self, message=None, done=None, total=None):
        if self.window is None:
            return

        if message is not None:
            self.label.configure(text=message)

        if total:
            if not self.showing_bar:
                self.bar.pack(anchor='w', pady=(12, 0))
                self.showing_bar = True
            self.bar.configure(maximum=total, value=done or 0)
        elif self.showing_bar:
            self.bar.pack_forget()
            self.showing_bar = False

        self.window.update()

    # Closing twice is harmless, which lets the caller close it early and still clean up
    def close(self):
        if self.window is not None:
            self.window.destroy()
            self.window = None
            app_root().update()


class SuggestionList:
    """The matching projects, floating just under the search box.

    A window of its own rather than a panel inside the form. That way it can hang over
    whatever is beneath it - and past the bottom edge of the form itself - the way the
    suggestions under a search box on a website do, without the form having to keep a
    band of empty space to hold it, and without any of the measuring and correcting that
    placing it inside the form used to need.
    """

    def __init__(self, form, entry):
        self.form = form
        self.entry = entry
        self.popup = None
        self.tree = None
        self.projects = []
        self.highlighted = None
        self.on_pick = None

    def _build(self):
        self.popup = tkinter.Toplevel(self.form)
        self.popup.withdraw()
        self.popup.overrideredirect(True)
        self.popup.transient(self.form)
        # A themed frame cannot draw an outline of its own, so the window behind the list
        # shows through the one pixel left round it and becomes its edge
        self.popup.configure(background=EDGE_COLOR)

        # The customer and the load count get a column each. Written into one column
        # together they started at the same place but ended wherever the customer's name
        # happened to end, which left the counts scattered down the list
        self.tree = tb.Treeview(self.popup, columns=('customer', 'loads'), show='tree',
                                selectmode='browse', bootstyle='primary')
        self.tree.pack(fill='both', expand=True, padx=1, pady=1)
        self.tree.bind('<ButtonRelease-1>', self._clicked)

    @property
    def showing(self):
        return self.popup is not None and self.popup.winfo_ismapped()

    def show(self, projects, hidden_count):
        if self.popup is None:
            self._build()

        self.projects = list(projects)
        self.tree.delete(*self.tree.get_children())
        for position, project in enumerate(self.projects):
            self.tree.insert('', 'end', iid=str(position), text=f' {project.id}',
                             values=describe_briefly(project))
        if hidden_count:
            self.tree.insert('', 'end', iid='more',
                             text=f'  ... and {hidden_count} more, keep typing', values=('', ''))

        self.tree.configure(height=len(self.tree.get_children()))   # only as tall as needed
        self.highlighted = None
        self._place()

    # Asked to go away it goes away, without first checking whether it is on screen. A
    # window that has been told to appear is not yet mapped, so a check would answer "not
    # showing" for anything put away in that moment - and the list would then be mapped
    # straight afterwards, and stay up with nothing left meaning to take it down
    def hide(self):
        if self.popup is not None:
            self.popup.withdraw()
        self.highlighted = None

    # Function to sit the list flush under the search box. It is a window on the screen,
    # so it is placed in screen coordinates and nothing inside the form can shift it
    def _place(self):
        self.entry.update_idletasks()
        width = self.entry.winfo_width()
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()

        # The load count needs room for "12 loads" and never more, the customer takes a
        # share of what is left, and the address - much the longest of the three, and the
        # one being searched on - takes everything still going
        loads = 84
        customer = min(max(int(width * 0.26), 130), 210)
        self.tree.column('#0', width=max(width - customer - loads - 2, 100), stretch=False)
        self.tree.column('customer', width=customer, stretch=False, anchor='w')
        self.tree.column('loads', width=loads, stretch=False, anchor='w')

        self.popup.geometry(f'{width}x{self.tree.winfo_reqheight() + 2}+{x}+{y}')
        self.popup.deiconify()
        self.popup.lift()

    # Function to keep the list under the search box while the form itself is dragged
    def follow(self):
        if self.showing:
            self._place()

    # Function to say whether the pointer is over the list itself. Clicking a suggestion
    # takes the caret out of the search box for a moment, and this is what tells that
    # apart from the caret having moved on somewhere else for good - without it, the list
    # would be put away by the very click that was trying to pick something from it
    def under_pointer(self):
        if not self.showing:
            return False
        x, y = self.popup.winfo_pointerxy()
        left, top = self.popup.winfo_rootx(), self.popup.winfo_rooty()
        return (left <= x < left + self.popup.winfo_width()
                and top <= y < top + self.popup.winfo_height())

    # Function to walk the list with the arrow keys while the caret stays in the search
    # box, so that typing can carry straight on
    def step(self, direction):
        if not self.showing or not self.projects:
            return
        if self.highlighted is None:
            self.highlighted = 0 if direction > 0 else len(self.projects) - 1
        else:
            self.highlighted = min(max(self.highlighted + direction, 0), len(self.projects) - 1)

        row = str(self.highlighted)
        self.tree.selection_set(row)
        self.tree.see(row)

    # Function to give back the project the list is pointing at, if it is pointing at one.
    # The last line can be the "and 12 more" note, which names no project
    def chosen(self):
        if self.highlighted is None:
            return None
        return self.projects[self.highlighted]

    def _clicked(self, _event):
        selected = self.tree.selection()
        if not selected or selected[0] == 'more':
            return                            # the "and 12 more" line names no project
        self.highlighted = int(selected[0])
        if self.on_pick is not None:
            self.on_pick(self.projects[self.highlighted])


class InfoLine:
    """The line under the search box, with hairlines standing between its parts.

    Written as one label per part with a drawn rule between them, since a rule is a thing
    that has to be built rather than a character that can be typed. Every write clears out
    what the last one left, and each rule is asked to fill the height of the line, which
    leaves it exactly as tall as the writing beside it.
    """

    def __init__(self, parent):
        self.frame = tb.Frame(parent)

    def show(self, parts, style='Confirm.TLabel'):
        for old in self.frame.winfo_children():
            old.destroy()

        for index, part in enumerate(parts):
            if index:
                rule = tkinter.Frame(self.frame, width=INFO_RULE_WIDTH, background=LINE_COLOR)
                rule.pack(side='left', fill='y', padx=INFO_RULE_GAP)
            tb.Label(self.frame, text=part, style=style).pack(side='left')


def format_date(date):
    return date.strftime(DATE_FORMAT)


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


# Function to find the projects whose address or customer holds every word typed, in any
# order and in any case. Project IDs start with a house number, so a plain "starts with"
# match would hide the right project from anyone searching by street name. The words are
# looked for across the address and the customer at once, so "terrain delridge" finds the
# Delridge job done for All Terrain without either half having to be typed in full
def matching_projects(projects, query):
    words = query.lower().split()
    if not words:
        return list(projects)          # an empty box offers the most recent work
    return [project for project in projects
            if all(word in project.match_text for word in words)]


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


# Function to describe a project in the few words that fit beside it in the search list,
# where the customer and the size of the job are what tell two similar addresses apart.
# They are handed back separately because they are shown in a column each, which is what
# keeps every load count in the list starting at the same place
def describe_briefly(project):
    customer = str(project.customer) if pd.notna(project.customer) else ''
    loads = '1 load' if project.loads == 1 else f'{project.loads} loads'
    return customer, loads


# Function to turn a typed date into a real one. Returns the date and an explanation, only
# one of which is ever filled in.
#
# The date boxes can hand back a date of their own working out, but theirs quietly falls
# back to today whenever the box is empty or holds nonsense. Empty here has to go on
# meaning "no limit", and nonsense has to be worth stopping for, so what was typed is
# read as text and parsed here instead
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


# Function to work out what has been asked for, or what is stopping the run from starting.
# Everything the form can catch is caught here, in the form, where it can still be fixed.
# The box at fault is named as well, so that it can be marked rather than only described
def read_choices(values, records, projects):
    project = find_project(projects, values.project_text)
    if project is None:
        typed = values.project_text.strip()
        if not typed:
            return None, 'Choose a project first, then pick it from the list.', 'project'
        return None, f'No project matches "{shorten(typed, 22)}". Pick one from the list.', 'project'

    start_date, start_error = parse_date(values.start_text, 'start date')
    if start_error:
        return None, start_error, 'start'

    end_date, end_error = parse_date(values.end_text, 'end date')
    if end_error:
        return None, end_error, 'end'

    if pd.notna(start_date) and pd.notna(end_date) and start_date > end_date:
        return None, 'The start date is after the end date.', 'start'

    if not values.driver_logs and not values.invoice:
        return None, 'Nothing to create. Tick "Driver logs" or "Invoice".', 'options'

    if filter_loads(records, values.sheet_name, project.id, start_date, end_date).empty:
        return None, 'No loads for that project between those dates.', 'start'

    return UserChoices(records, values.sheet_name, project.id, start_date, end_date, values.taxable,
                       values.driver_logs, values.invoice, values.pdf), None, None


class ProjectForm:
    """The form itself: which records, which project, which dates, and what to make."""

    def __init__(self, records, sheet_names, default_sheet):
        self.records = records
        self.projects = records.projects(default_sheet)
        self.choices = None

        use_theme()
        self.window = tb.Toplevel(title=WINDOW_TITLE)
        self.window.withdraw()          # nothing is seen until it is built and placed
        self.window.protocol('WM_DELETE_WINDOW', self.cancel)

        self._build(sheet_names, default_sheet)

        # Anchored to the edge drawn round the search box rather than to the box inside
        # it, so the list lines up with what can actually be seen
        self.suggestions = SuggestionList(self.window, self.project_field.shell)
        self.suggestions.on_pick = self.take_project
        self.window.bind('<Configure>', lambda _event: self.suggestions.follow())

        # The list is a window of its own, and not one the desktop manages, so nothing
        # else will ever lower it or put it away: left to itself it goes on floating over
        # whatever is in front, this program's window included. Everything that means it
        # has stopped being used has to take it down by hand. A toplevel is in the bind
        # tags of everything inside it, so one binding here catches clicks anywhere in the
        # form - and the list, being its own window, is not caught by it
        self.window.bind('<Button-1>', self._clicked_elsewhere, add='+')
        self.window.bind('<Deactivate>', lambda _e: self.suggestions.hide(), add='+')
        self.window.bind('<Unmap>', lambda _e: self.suggestions.hide(), add='+')
        self.project_box.bind('<FocusOut>', lambda _e: self._project_lost_focus(), add='+')

        self.window.update_idletasks()
        self.window.minsize(max(self.window.winfo_reqwidth(), MIN_WINDOW_WIDTH),
                            self.window.winfo_reqheight())
        self.project_box.focus_set()
        show_centered(self.window, MIN_WINDOW_WIDTH)
        self.window.after(SOURCE_CHECK_MS, self._check_source)

    # ---- building ----------------------------------------------------------------

    def _build(self, sheet_names, default_sheet):
        body = tb.Frame(self.window, padding=(24, 20))
        body.pack(fill='both', expand=True)
        body.columnconfigure(0, weight=1)
        row = _Rows()

        tb.Label(body, text='Generate Audit Daily Load Tickets',
                 style='Title.TLabel').grid(row=row.next(), column=0, sticky='w')

        self.source_line = tb.Label(body, text=describe_source(self.records), style='Muted.TLabel')
        self.source_line.grid(row=row.next(), column=0, sticky='w', pady=(2, 0))

        tb.Separator(body).grid(row=row.next(), column=0, sticky='ew', pady=(14, 12))

        # ---- which records ----
        tb.Label(body, text='Year sheet', style='Field.TLabel').grid(row=row.next(), column=0,
                                                                     sticky='w')
        self.sheet_var = tb.StringVar(value=default_sheet)
        self.sheet_field = Field(body, lambda shell: tb.Combobox(
            shell, textvariable=self.sheet_var, values=sheet_names,
            state='readonly', width=FIELD_WIDTH))
        self.sheet_box = self.sheet_field.widget
        self.sheet_field.shell.grid(row=row.next(), column=0, sticky='ew', pady=(3, 0))
        self.sheet_box.bind('<<ComboboxSelected>>', self._sheet_changed)

        # ---- which project ----
        tb.Label(body, text='Project', style='Field.TLabel').grid(row=row.next(), column=0,
                                                                  sticky='w', pady=(12, 0))
        self.project_var = tb.StringVar()
        self.project_field = Field(body, lambda shell: tb.Entry(
            shell, textvariable=self.project_var, width=FIELD_WIDTH))
        self.project_box = self.project_field.widget
        self.project_field.shell.grid(row=row.next(), column=0, sticky='ew', pady=(3, 0))

        # Doubles as the line explaining a project that cannot be used, so a complaint
        # about the project appears against the project, and not at the foot of the form
        self.project_line = InfoLine(body)
        self.project_line.frame.grid(row=row.next(), column=0, sticky='w', pady=(4, 0))

        self.project_var.trace_add('write', lambda *_: self._project_typed())
        self.project_box.bind('<Down>', lambda _e: self._step(1))
        self.project_box.bind('<Up>', lambda _e: self._step(-1))
        self.project_box.bind('<Return>', lambda _e: self._project_return())
        self.project_box.bind('<Escape>', self._project_escape)
        # Clicking into the box, or tabbing into it, drops the list down straight away
        # rather than waiting for a first letter to be typed
        self.project_box.bind('<Button-1>', lambda _e: self._show_matches(searching=True))
        self.project_box.bind('<FocusIn>', lambda _e: self._show_matches(searching=True))

        # ---- which dates ----
        dates = tb.Frame(body)
        dates.grid(row=row.next(), column=0, sticky='w', pady=(12, 0))
        self.start_box = self._date_field(dates, 'Start date', 0)
        self.end_box = self._date_field(dates, 'End date', 1)

        # ---- what to make ----
        # Four boxes in one row said three different things at once: what to make, what
        # kind of invoice, and what to save it as. Split up, each heading answers one
        # question, and no box has to be read against the ones either side of it
        groups = tb.Frame(body)
        groups.grid(row=row.next(), column=0, sticky='w', pady=(16, 0))

        self.driver_logs_var = tb.BooleanVar(value=True)
        self.invoice_var = tb.BooleanVar(value=True)
        self.taxable_var = tb.BooleanVar(value=False)
        self.pdf_var = tb.BooleanVar(value=True)

        create = self._group(groups, 'Create', 0)
        self._check(create, 'Driver logs', self.driver_logs_var)
        self._check(create, 'Invoice', self.invoice_var, command=self._invoice_toggled)

        options = self._group(groups, 'Options', 1)
        self.taxable_box = self._check(options, 'Taxable', self.taxable_var)
        self._check(options, 'Also save as PDF', self.pdf_var)

        tb.Separator(body).grid(row=row.next(), column=0, sticky='ew', pady=(18, 0))

        # One line, always here. Keeping the space means a complaint about the dates, or
        # about there being nothing to create, does not shove the buttons down the moment
        # it appears. Anything to do with the project is said against the project instead
        self.error_line = tb.Label(body, text='', style='Error.TLabel')
        self.error_line.grid(row=row.next(), column=0, sticky='w', pady=(10, 0))

        buttons = tb.Frame(body)
        buttons.grid(row=row.next(), column=0, sticky='w', pady=(6, 0))
        tb.Button(buttons, text='Generate', bootstyle='primary', width=12,
                  command=self.submit).pack(side='left')
        tb.Button(buttons, text='Cancel', bootstyle='secondary-outline', width=10,
                  command=self.cancel).pack(side='left', padx=(8, 0))

        self._invoice_toggled()

    # Function to build one date box, with its label and the calendar behind it. It starts
    # empty, nothing having been asked for yet, and picking a project fills it in
    def _date_field(self, parent, text, column):
        holder = tb.Frame(parent)
        holder.grid(row=0, column=column, sticky='w', padx=(0, 20))
        tb.Label(holder, text=text, style='Field.TLabel').pack(anchor='w')

        def build(shell):
            box = tb.DateEntry(shell, date_format=DATE_FORMAT, width=DATE_WIDTH,
                               bootstyle='primary')
            # Cream behind the box and its button, so no sage shows in the gap between
            # them or round the outside of the pair
            box.configure(style='Field.TFrame')
            return box

        field = Field(holder, build, focus_on=lambda box: box.entry)
        field.shell.pack(anchor='w', pady=(3, 0))

        field.widget.entry.delete(0, 'end')
        field.widget.entry.bind('<KeyRelease>', lambda _e: self.clear_error())
        field.widget.bind('<<DateEntrySelected>>', lambda _e: self.clear_error())
        return field

    def _group(self, parent, title, column):
        holder = tb.Frame(parent)
        holder.grid(row=0, column=column, sticky='nw', padx=(0, 44))
        tb.Label(holder, text=title, style='Field.TLabel').pack(anchor='w', pady=(0, 4))
        return holder

    def _check(self, parent, text, variable, command=None):
        box = tb.Checkbutton(parent, text=text, variable=variable, bootstyle='primary',
                             command=command or self.clear_error)
        box.pack(anchor='w', pady=1)
        return box

    # ---- what the form does ------------------------------------------------------

    # Function to keep the records line honest. Saving the records after this program read
    # them leaves the form offering yesterday's data, which is what this is here to catch
    def _check_source(self):
        if self.records.has_been_saved_since_read():
            self.source_line.configure(text='Records were saved again - close and run this once more',
                                       style='Error.TLabel')
        else:
            self.source_line.configure(text=describe_source(self.records), style='Muted.TLabel')
        self.window.after(SOURCE_CHECK_MS, self._check_source)

    def _sheet_changed(self, _event=None):
        self.projects = self.records.projects(self.sheet_var.get())
        self.suggestions.hide()
        self._show_matches()

    def _project_typed(self):
        self._show_matches(searching=True)
        self.clear_error()

    # Clicking the search box, or the edge drawn round it, is what drops the list down.
    # Anything else clicked in the form means attention has moved on, including the parts
    # of it that take no caret and so would never have reported the move themselves
    def _clicked_elsewhere(self, event):
        # Compared by the name Tk knows them by, since an event does not always arrive
        # carrying the widget itself - sometimes it carries only that name
        clicked = str(event.widget)
        if clicked in (str(self.project_box), str(self.project_field.shell)):
            return
        self.suggestions.hide()

    # Tabbing out of the search box puts the list away too, but a click on the list itself
    # also takes the caret out of the box, and that one has to be left alone
    def _project_lost_focus(self):
        if not self.suggestions.under_pointer():
            self.suggestions.hide()

    def _step(self, direction):
        self.suggestions.step(direction)
        return 'break'                 # the caret stays put instead of walking the text

    def _project_escape(self, _event):
        if self.suggestions.showing:
            self.suggestions.hide()
            return 'break'             # only the list closes; the form stays open
        self.cancel()

    def _project_return(self):
        picked = self.suggestions.chosen()
        if picked is not None:
            self.take_project(picked)
        elif find_project(self.projects, self.project_var.get()):
            self.submit()              # the box already names a real project, so "go"
        elif self.suggestions.showing and self.suggestions.projects:
            self.take_project(self.suggestions.projects[0])
        return 'break'

    # Function to redraw the search results for whatever has been typed so far
    def _show_matches(self, searching=False):
        query = self.project_var.get()
        chosen = find_project(self.projects, query)
        matches = matching_projects(self.projects, query)

        # Once the box holds a project's whole name there is nothing left to choose
        # between. An empty box only drops the list down while the box is actually being
        # used, so the form does not open with a list hanging over it
        if chosen is not None or (not query.strip() and not searching) or not matches:
            self.suggestions.hide()
        else:
            self.suggestions.show(matches[:MAX_SUGGESTIONS],
                                  max(len(matches) - MAX_SUGGESTIONS, 0))

        if chosen is not None:
            self.project_line.show(describe_project(chosen))
        elif query.strip() and not matches:
            self.project_line.show(['No projects match'], style='Muted.TLabel')
        else:
            self.project_line.show([])

    # Function to put a project into the box and fill the dates in with the range it ran
    # over, which is both a starting point and a hint at what dates are worth asking for
    def take_project(self, project):
        self.suggestions.hide()
        self.project_var.set(project.id)
        self.project_box.icursor('end')
        self.project_box.focus_set()

        if pd.notna(project.first_date):
            self._set_date(self.start_box, format_date(project.first_date))
            self._set_date(self.end_box, format_date(project.last_date))
        else:
            self._set_date(self.start_box, '')
            self._set_date(self.end_box, '')

        self.clear_error()
        self._show_matches()

    def _set_date(self, field, text):
        field.entry.delete(0, 'end')
        if text:
            field.entry.insert(0, text)

    # Taxable describes the invoice, so with no invoice being made there is nothing for it
    # to describe. Greyed out it stays visible, and stops looking like it was simply missed
    def _invoice_toggled(self):
        self.taxable_box.configure(state='normal' if self.invoice_var.get() else 'disabled')
        self.clear_error()

    # ---- what is wrong -----------------------------------------------------------

    # Function to take off whatever mark the last complaint left behind
    def clear_error(self):
        self.error_line.configure(text='')
        for field in (self.project_field, self.start_box, self.end_box):
            field.mark(False)

    # Function to say what is wrong, against the box it is wrong in. A red edge points at
    # the box, so the words do not have to name it for the eye to find it
    def show_error(self, message, field):
        self.clear_error()

        if field == 'project':
            self.project_field.mark(True)
            self.project_line.show([message], style='Error.TLabel')
            self.project_box.focus_set()
            return

        if field == 'start':
            self.start_box.mark(True)
        elif field == 'end':
            self.end_box.mark(True)
        self.error_line.configure(text=message)

    # ---- finishing ---------------------------------------------------------------

    def values(self):
        return FormValues(self.sheet_var.get(), self.project_var.get(),
                          self.start_box.entry.get(), self.end_box.entry.get(),
                          self.taxable_var.get() and self.invoice_var.get(),
                          self.driver_logs_var.get(), self.invoice_var.get(), self.pdf_var.get())

    def submit(self):
        self.suggestions.hide()
        choices, error, field = read_choices(self.values(), self.records, self.projects)
        if choices is None:
            self.show_error(error, field)
            return
        self.choices = choices
        self._finish()

    def cancel(self):
        self.choices = None
        self._finish()

    def _finish(self):
        self.suggestions.hide()
        self.window.destroy()
        app_root().quit()

    def run(self):
        self.window.lift()
        self.window.focus_force()
        app_root().mainloop()
        return self.choices


# Counts the rows of the layout, so that adding one does not mean renumbering the rest
class _Rows:
    def __init__(self):
        self.row = -1

    def next(self):
        self.row += 1
        return self.row


class Notice:
    """Something to say, with the ways of answering it, in the middle of the screen.

    Built here rather than taken ready-made, because a ready-made one puts itself in the
    middle of the window it belongs to. By the time there is anything to say the form has
    been closed, and the only window left is the hidden one everything else hangs off -
    which sits in a corner of the screen, and was taking every message there with it.
    """

    def __init__(self, message, title, buttons):
        use_theme()
        self.answer = None
        self.window = tkinter.Toplevel(app_root())
        self.window.withdraw()          # nothing is seen until it is built and placed
        self.window.title(title)
        self.window.configure(background=PAGE_COLOR)
        self.window.resizable(False, False)
        self.window.transient(app_root())

        body = tb.Frame(self.window, padding=(24, 20))
        body.pack(fill='both', expand=True)
        tb.Label(body, text=message, justify='left',
                 wraplength=NOTICE_WIDTH).pack(anchor='w')

        row = tb.Frame(body)
        row.pack(anchor='e', pady=(20, 0))
        for position, name in enumerate(buttons):
            # The last is the one being suggested, and the one Return takes
            last = position == len(buttons) - 1
            tb.Button(row, text=name, bootstyle='primary' if last else 'secondary-outline',
                      width=12, command=lambda taken=name: self._choose(taken)
                      ).pack(side='left', padx=(8, 0))

        # Closing it, or pressing Escape, answers nothing - which is never the doing part
        self.window.protocol('WM_DELETE_WINDOW', lambda: self._choose(None))
        self.window.bind('<Escape>', lambda _e: self._choose(None))
        self.window.bind('<Return>', lambda _e: self._choose(buttons[-1]))

    def _choose(self, taken):
        self.answer = taken
        self.window.destroy()

    def ask(self):
        show_centered(self.window)
        self.window.lift()
        self.window.focus_force()
        self.window.grab_set()
        self.window.wait_window()       # a loop of its own, so no mainloop is needed here
        return self.answer


# Function to report something that stopped the run, once the form itself is gone
def report_error(message, title='Something went wrong'):
    Notice(message, title, ['Close']).ask()


# Function to say what was made and where, so the files do not have to be hunted for.
# Comes back True when the folder should be opened
def report_finished(message, title='Done'):
    return Notice(message, title, ['Close', 'Open Folder']).ask() == 'Open Folder'


def collect_UI_input():
    use_theme()

    # The records are read off to one side, so that a read taking too long can be noticed
    # while it is still going on. Read on this thread it would hold everything up, and
    # nothing could be put on screen until it had already finished - which is the whole
    # difficulty: by then there is nothing left worth saying. The thread reads a file and
    # nothing else; every window here belongs to this one
    outcome = {}

    def read_records():
        try:
            records = BookRecords()
            sheet_names = records.year_sheet_names()
            if not sheet_names:
                raise ValueError('The load records hold no "Dump Trucking" year sheet.')
            default_sheet = records.default_year_sheet()
            records.projects(default_sheet)     # the slow half of opening the form
            outcome['read'] = (records, sheet_names, default_sheet)
        except BaseException as trouble:
            outcome['trouble'] = trouble

    worker = threading.Thread(target=read_records, daemon=True)
    worker.start()

    loading = None
    started = time.monotonic()
    while worker.is_alive():
        if loading is None and time.monotonic() - started > SLOW_READ_SECONDS:
            loading = StatusWindow('Reading records')
        # The desktop is answered either way, so the program never looks hung, whether or
        # not it has put anything up to be looked at
        (loading.window if loading is not None else app_root()).update()
        time.sleep(READ_POLL_SECONDS)
    worker.join()

    if loading is not None:
        loading.close()
    if 'trouble' in outcome:
        raise outcome['trouble']

    records, sheet_names, default_sheet = outcome['read']
    return ProjectForm(records, sheet_names, default_sheet).run()
