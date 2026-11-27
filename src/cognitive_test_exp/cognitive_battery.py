"""
Cognitive Test Battery  (approx. 10 minutes)
=============================================
Six tasks, presented in a random order for every participant:
    Maze, Digit span (digit memory), Stroop, Go/No-Go, Corsi block-tapping, Flanker

Before the tasks start, the researcher types in the participant number and picks
the group (control / experiment).

DATA
Results are saved as Excel workbooks in the "data" folder:
    data/P<number>_<group>_<date_time>.xlsx  -> this participant (saved after EVERY trial)
    data/all_participants.xlsx               -> all participants (updated at the end)

Each workbook has these sheets:
    Summary     one row per participant with the main scores of every test
    Maze, Digit span, Stroop, Go-NoGo, Corsi, Flanker
                one row per trial, with descriptive column headers
Practice trials are not saved.

Install:     pip install pygame openpyxl
Run:         uv run cognitive-battery   (from the project folder)
Only some:   uv run cognitive-battery --only Stroop
Timing:      start/end clock time of each test -> data/P<number>_<group>_<date_time>_timing.csv
Stop early:  press ESC (the data collected so far is saved)

Close all_participants.xlsx in Excel before running a session, otherwise
it cannot be updated (a backup copy is saved instead).
"""

import os
import random
import time
from collections import deque
from datetime import datetime
from statistics import mean

import pygame
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# =============================================================================
# SETTINGS  (change these to adjust the length / difficulty of the battery)
# =============================================================================
FULLSCREEN = False             # set to True when testing real participants
WINDOW_SIZE = (1280, 800)      # used only when FULLSCREEN is False
# project root: <project>/data
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "data")
ALL_DATA_FILE = "all_participants.xlsx"
PRACTICE = True                # short practice round with feedback before each task

ITI_MS = (400, 700)            # random fixation cross (+) between trials, min-max in ms

# Maze: (grid size, number of mazes) per difficulty level, easy -> hard
MAZE_LEVELS = [(7, 1), (10, 1), (13, 1)]
MAZE_TIME_LIMIT_S = 30

# Digit span (forward): adaptive, 2 trials per length, stops after 2 errors at one length
DIGIT_START_LEN = 3
DIGIT_MAX_LEN = 9
DIGIT_SHOW_MS = 800
DIGIT_GAP_MS = 200

# Stroop
STROOP_TRIALS = 24             # half congruent, half incongruent
STROOP_RESPONSE_WINDOW_MS = 2000

# Go/No-Go
GONOGO_TRIALS = 40
GONOGO_GO_PROPORTION = 0.75
GONOGO_STIM_MS = 500           # how long the shape stays on screen
GONOGO_RESPONSE_WINDOW_MS = 1000

# Corsi block-tapping: adaptive like digit span
CORSI_START_LEN = 3
CORSI_MAX_LEN = 9
CORSI_SHOW_MS = 700
CORSI_GAP_MS = 300

# Flanker
FLANKER_TRIALS = 24            # half congruent, half incongruent
FLANKER_RESPONSE_WINDOW_MS = 2000

# =============================================================================
# Sheets and column headers (what each data point is)
# =============================================================================
COMMON = ["Participant", "Group", "Test position (1-6)", "Trial"]

SHEETS = {
    "Summary": [
        "Participant", "Group", "Date and time", "Test order", "Finished all tests",
        "Maze: mazes completed", "Maze: total time (s)", "Maze: total wall bumps",
        "Maze: total extra moves (completed mazes)",
        "Digit span: longest correct sequence",
        "Stroop: accuracy congruent (%)", "Stroop: accuracy incongruent (%)",
        "Stroop: mean RT congruent, correct trials (ms)",
        "Stroop: mean RT incongruent, correct trials (ms)",
        "Stroop: interference effect (ms)",
        "Go/No-Go: hit rate (%)", "Go/No-Go: false alarm rate (%)",
        "Go/No-Go: mean RT hits (ms)",
        "Corsi: longest correct sequence",
        "Flanker: accuracy congruent (%)", "Flanker: accuracy incongruent (%)",
        "Flanker: mean RT congruent, correct trials (ms)",
        "Flanker: mean RT incongruent, correct trials (ms)",
        "Flanker: interference effect (ms)",
    ],
    "Maze": COMMON + [
        "Difficulty level (1=easy)", "Maze size (cells per side)",
        "Completed (1=yes, 0=time out)", "Time (s)", "Moves made",
        "Shortest possible moves", "Extra moves", "Wall bumps",
    ],
    "Digit span": COMMON + [
        "Sequence length", "Digits shown", "Answer typed", "Correct (1/0)", "Answer time (s)",
    ],
    "Stroop": COMMON + [
        "Condition", "Word shown", "Ink colour (correct answer)", "Response",
        "Correct (1/0)", "Reaction time (ms)",
    ],
    "Go-NoGo": COMMON + [
        "Trial type (go=green circle, nogo=red square)", "Pressed SPACE (1/0)",
        "Outcome", "Correct (1/0)", "Reaction time (ms)",
    ],
    "Corsi": COMMON + [
        "Sequence length", "Blocks shown (in order)", "Blocks clicked (in order)",
        "Correct (1/0)", "Response time (s)",
    ],
    "Flanker": COMMON + [
        "Condition", "Arrows shown", "Middle arrow (correct answer)", "Response",
        "Correct (1/0)", "Reaction time (ms)",
    ],
}

# =============================================================================
# Colours and stimuli
# =============================================================================
BG = (30, 30, 30)
WHITE = (235, 235, 235)
GREY = (150, 150, 150)
YELLOW = (240, 210, 50)
GREEN = (40, 190, 70)
RED = (220, 50, 50)
ORANGE = (240, 150, 40)
BLUE = (60, 100, 220)

STROOP_COLORS = {          # word: (ink colour, response key)
    "RED": ((220, 40, 40), pygame.K_r),
    "GREEN": ((40, 190, 60), pygame.K_g),
    "BLUE": ((60, 110, 240), pygame.K_b),
    "YELLOW": ((235, 210, 40), pygame.K_y),
}

# Corsi block positions (relative coordinates, 0-1), spread irregularly
CORSI_POSITIONS = [(0.15, 0.25), (0.55, 0.15), (0.80, 0.30), (0.30, 0.45), (0.60, 0.50),
                   (0.85, 0.65), (0.12, 0.75), (0.40, 0.80), (0.70, 0.88)]

DIRS = {"N": (-1, 0), "S": (1, 0), "E": (0, 1), "W": (0, -1)}
OPPOSITE = {"N": "S", "S": "N", "E": "W", "W": "E"}


class QuitBattery(Exception):
    """Raised when ESC is pressed or the window is closed."""


# =============================================================================
# Excel helpers
# =============================================================================
def add_sheet(wb, name):
    ws = wb.create_sheet(name)
    ws.append(SHEETS[name])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for i, header in enumerate(SHEETS[name], 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = max(12, min(len(header) + 2, 45))
    ws.freeze_panes = "A2"
    return ws


def new_workbook():
    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEETS:
        add_sheet(wb, name)
    return wb


def append_rows(wb, data):
    """data = {sheet name: list of row dicts}"""
    for name, rows in data.items():
        ws = wb[name] if name in wb.sheetnames else add_sheet(wb, name)
        for row in rows:
            ws.append([row.get(h, "") for h in SHEETS[name]])


def pct(values):
    return round(100 * mean(values), 1) if values else ""


def mean_rt(values):
    return round(mean(values), 1) if values else ""


class Battery:
    def __init__(self, only=None):
        self.only = only  # run only these tests (names), None = all six
        self.timing = []  # (test, start, end) wall-clock times, to align with EEG
        pygame.init()
        if FULLSCREEN:
            self.screen = pygame.display.set_mode((0, 0), pygame.FULLSCREEN)
        else:
            self.screen = pygame.display.set_mode(WINDOW_SIZE)
        pygame.display.set_caption("Cognitive Test Battery")
        self.W, self.H = self.screen.get_size()
        self.clock = pygame.time.Clock()
        self.fonts = {}
        self.participant = self.group = None
        self.order = []
        self.current_pos = 0
        self.data = {name: [] for name in SHEETS if name != "Summary"}
        self.spans = {}
        self.finished = False
        self.p_path = None

    # ------------------------------------------------------------------ basics
    def font(self, size):
        if size not in self.fonts:
            self.fonts[size] = pygame.font.SysFont("arial", size)
        return self.fonts[size]

    def get_events(self):
        events = pygame.event.get()
        for e in events:
            if e.type == pygame.QUIT or (e.type == pygame.KEYDOWN and e.key == pygame.K_ESCAPE):
                raise QuitBattery
        return events

    def draw_text(self, text, size=32, color=WHITE, y=None):
        """Draw (multi-line) text centred horizontally; y = centre of the block."""
        font = self.font(size)
        lines = text.split("\n")
        line_h = font.get_linesize()
        top = (self.H - line_h * len(lines)) // 2 if y is None else y - line_h * len(lines) // 2
        for i, line in enumerate(lines):
            surf = font.render(line, True, color)
            self.screen.blit(surf, surf.get_rect(center=(self.W // 2, top + i * line_h + line_h // 2)))

    def wait(self, ms):
        """Wait while still allowing ESC to quit."""
        end = time.perf_counter() + ms / 1000
        while time.perf_counter() < end:
            self.get_events()
            self.clock.tick(500)

    def blank(self, ms):
        self.screen.fill(BG)
        pygame.display.flip()
        self.wait(ms)

    def fixation(self, ms):
        self.screen.fill(BG)
        self.draw_text("+", 60)
        pygame.display.flip()
        self.wait(ms)

    def iti(self):
        self.fixation(random.randint(*ITI_MS))

    def message(self, text, size=30, footer="Press SPACE to continue"):
        self.screen.fill(BG)
        self.draw_text(text, size)
        if footer:
            self.draw_text(footer, 24, GREY, y=self.H - 60)
        pygame.display.flip()
        pygame.event.clear()
        while True:
            for e in self.get_events():
                if e.type == pygame.KEYDOWN and e.key == pygame.K_SPACE:
                    return
            self.clock.tick(60)

    def choose(self, text, options):
        """Show text and wait for one of the keys in `options` (dict key -> value)."""
        self.screen.fill(BG)
        self.draw_text(text, 32)
        pygame.display.flip()
        pygame.event.clear()
        while True:
            for e in self.get_events():
                if e.type == pygame.KEYDOWN and e.key in options:
                    return options[e.key]
            self.clock.tick(60)

    def text_input(self, prompt, max_len=12):
        """Let the user type digits; ENTER confirms, BACKSPACE deletes."""
        text = ""
        pygame.event.clear()
        while True:
            self.screen.fill(BG)
            self.draw_text(prompt, 32, y=self.H // 2 - 80)
            self.draw_text(text + "_", 52, YELLOW, y=self.H // 2 + 20)
            self.draw_text("Press ENTER to confirm", 24, GREY, y=self.H - 60)
            pygame.display.flip()
            for e in self.get_events():
                if e.type != pygame.KEYDOWN:
                    continue
                if e.key in (pygame.K_RETURN, pygame.K_KP_ENTER) and text:
                    return text
                if e.key == pygame.K_BACKSPACE:
                    text = text[:-1]
                elif getattr(e, "unicode", "") and e.unicode.isdigit() and len(text) < max_len:
                    text += e.unicode
            self.clock.tick(60)

    def wait_response(self, keys, timeout_ms, stim_ms=None):
        """Call directly after the stimulus is shown (flip).
        Returns (key, rt_ms), or (None, None) if no response within timeout_ms.
        If stim_ms is given, the screen is cleared after that time."""
        pygame.event.clear(pygame.KEYDOWN)       # ignore key presses made before onset
        onset = time.perf_counter()
        cleared = False
        while True:
            for e in self.get_events():
                if e.type == pygame.KEYDOWN and e.key in keys:
                    return e.key, round((time.perf_counter() - onset) * 1000, 1)
            elapsed = (time.perf_counter() - onset) * 1000
            if stim_ms and not cleared and elapsed >= stim_ms:
                self.screen.fill(BG)
                pygame.display.flip()
                cleared = True
            if elapsed >= timeout_ms:
                return None, None
            self.clock.tick(1000)

    def feedback(self, correct, too_slow=False):
        if too_slow:
            text, col = "Too slow!", ORANGE
        elif correct:
            text, col = "Correct", GREEN
        else:
            text, col = "Wrong", RED
        self.screen.fill(BG)
        self.draw_text(text, 44, col)
        pygame.display.flip()
        self.wait(600)

    def practice_then_start(self, practice_fn):
        if not PRACTICE:
            return
        self.message("Let's practise first.\nYou will get feedback after each answer.")
        practice_fn()
        self.message("Practice finished.\n\nNow the real task begins (no feedback).\n"
                     "Respond as quickly and accurately as possible.")

    # ------------------------------------------------------------------ data
    def record(self, sheet, trial, **values):
        """Store one trial and save the participant workbook straight away."""
        row = {"Participant": self.participant, "Group": self.group,
               "Test position (1-6)": self.current_pos, "Trial": trial}
        row.update(values)
        unknown = set(row) - set(SHEETS[sheet])
        if unknown:
            raise ValueError(f"Unknown column(s) for {sheet}: {unknown}")
        self.data[sheet].append(row)
        self.save_participant_file()

    def summary_row(self):
        d = self.data
        s = {"Participant": self.participant, "Group": self.group,
             "Date and time": self.session_start, "Test order": ", ".join(self.order),
             "Finished all tests": "yes" if self.finished else "no"}
        if d["Maze"]:
            s["Maze: mazes completed"] = sum(r["Completed (1=yes, 0=time out)"] for r in d["Maze"])
            s["Maze: total time (s)"] = round(sum(r["Time (s)"] for r in d["Maze"]), 2)
            s["Maze: total wall bumps"] = sum(r["Wall bumps"] for r in d["Maze"])
            extra = [r["Extra moves"] for r in d["Maze"] if r["Extra moves"] != ""]
            s["Maze: total extra moves (completed mazes)"] = sum(extra) if extra else ""
        s["Digit span: longest correct sequence"] = self.spans.get("Digit span", "")
        s["Corsi: longest correct sequence"] = self.spans.get("Corsi", "")
        for task, label in (("Stroop", "Stroop"), ("Flanker", "Flanker")):
            rts = {}
            for cond in ("congruent", "incongruent"):
                rows = [r for r in d[task] if r["Condition"] == cond]
                s[f"{label}: accuracy {cond} (%)"] = pct([r["Correct (1/0)"] for r in rows])
                rts[cond] = mean_rt([r["Reaction time (ms)"] for r in rows if r["Correct (1/0)"]])
                s[f"{label}: mean RT {cond}, correct trials (ms)"] = rts[cond]
            if rts["congruent"] != "" and rts["incongruent"] != "":
                s[f"{label}: interference effect (ms)"] = round(rts["incongruent"] - rts["congruent"], 1)
        tt = "Trial type (go=green circle, nogo=red square)"
        go = [r for r in d["Go-NoGo"] if r[tt] == "go"]
        nogo = [r for r in d["Go-NoGo"] if r[tt] == "nogo"]
        s["Go/No-Go: hit rate (%)"] = pct([r["Pressed SPACE (1/0)"] for r in go])
        s["Go/No-Go: false alarm rate (%)"] = pct([r["Pressed SPACE (1/0)"] for r in nogo])
        s["Go/No-Go: mean RT hits (ms)"] = mean_rt([r["Reaction time (ms)"] for r in go
                                                     if r["Pressed SPACE (1/0)"]])
        return s

    def save_participant_file(self):
        wb = new_workbook()
        append_rows(wb, {"Summary": [self.summary_row()], **self.data})
        wb.save(self.p_path)

    def save_timing(self):
        """Start/end clock time of each test -> <participant file>_timing.csv (to align with EEG)."""
        with open(self.p_path.replace(".xlsx", "_timing.csv"), "w", encoding="utf-8") as f:
            f.write("test,start,end,duration_s\n")
            for name, start, end in self.timing:
                f.write(f"{name},{start:%H:%M:%S},{end:%H:%M:%S},{(end - start).total_seconds():.0f}\n")

    def save_all_participants_file(self):
        path = os.path.join(DATA_DIR, ALL_DATA_FILE)
        wb = load_workbook(path) if os.path.exists(path) else new_workbook()
        append_rows(wb, {"Summary": [self.summary_row()], **self.data})
        try:
            wb.save(path)
            return path
        except PermissionError:          # file is open in Excel
            backup = os.path.join(DATA_DIR, f"all_participants_BACKUP_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
            wb.save(backup)
            print(f"WARNING: {path} was open in another program. Saved to {backup} instead.")
            return backup

    def existing_participants(self):
        path = os.path.join(DATA_DIR, ALL_DATA_FILE)
        if not os.path.exists(path):
            return set()
        ws = load_workbook(path, read_only=True)["Summary"]
        return {str(row[0]) for row in ws.iter_rows(min_row=2, values_only=True) if row[0] is not None}

    # ------------------------------------------------------------------ start
    def get_participant_info(self):
        while True:
            pid = self.text_input("Enter participant number:")
            group = self.choose("Select group:\n\n1 = Control\n2 = Experiment",
                                {pygame.K_1: "control", pygame.K_KP1: "control",
                                 pygame.K_2: "experiment", pygame.K_KP2: "experiment"})
            warning = ""
            if pid in self.existing_participants():
                warning = "\n\nWARNING: this participant number already has data!"
            ok = self.choose(f"Participant: {pid}\nGroup: {group}{warning}\n\n"
                             "SPACE = start        R = re-enter",
                             {pygame.K_SPACE: True, pygame.K_r: False})
            if ok:
                self.participant, self.group = int(pid), group
                return

    def run(self):
        try:
            pygame.mouse.set_visible(False)
            self.get_participant_info()
            tests = [("Maze", self.test_maze), ("Digit span", self.test_digit_span),
                     ("Stroop", self.test_stroop), ("Go-NoGo", self.test_gonogo),
                     ("Corsi", self.test_corsi), ("Flanker", self.test_flanker)]
            if self.only:
                tests = [t for t in tests if t[0] in self.only]
            random.shuffle(tests)
            self.order = [name for name, _ in tests]

            os.makedirs(DATA_DIR, exist_ok=True)
            now = datetime.now()
            self.session_start = now.strftime("%Y-%m-%d %H:%M")
            self.p_path = os.path.join(DATA_DIR, f"P{self.participant}_{self.group}_{now:%Y%m%d_%H%M%S}.xlsx")
            self.save_participant_file()

            self.message(f"Welcome!\n\nYou will complete {len(tests)} short task{'s' if len(tests) > 1 else ''}.\n"
                         "Each task starts with instructions and a short practice.\n\n"
                         "Please keep your hands near the keyboard.")
            for i, (name, test_fn) in enumerate(tests, 1):
                self.current_pos = i
                start = datetime.now()
                print(f"{name}: started {start:%H:%M:%S}")
                test_fn()
                end = datetime.now()
                print(f"{name}: finished {end:%H:%M:%S}")
                self.timing.append((name, start, end))
                self.save_timing()
                if i < len(tests):
                    self.message(f"Task {i} of {len(tests)} finished.\n\nTake a short breath.")
            self.finished = True
            self.message("All tasks are finished.\n\nThank you for participating!",
                         footer="Press SPACE to exit")
        except QuitBattery:
            print("Stopped early with ESC.")
        finally:
            pygame.quit()
            if self.p_path:
                self.save_participant_file()
                all_path = self.save_all_participants_file()
                print(f"Data saved to: {self.p_path}")
                print(f"           and {all_path}")

    # ================================================================ TASKS
    # ------------------------------------------------------------------ Maze
    @staticmethod
    def generate_maze(n):
        """Random maze (depth-first 'recursive backtracker'). Returns {cell: set of open sides}."""
        maze = {(r, c): set() for r in range(n) for c in range(n)}
        stack, visited = [(0, 0)], {(0, 0)}
        while stack:
            r, c = stack[-1]
            options = [(d, (r + dr, c + dc)) for d, (dr, dc) in DIRS.items()
                       if (r + dr, c + dc) in maze and (r + dr, c + dc) not in visited]
            if options:
                d, nxt = random.choice(options)
                maze[(r, c)].add(d)
                maze[nxt].add(OPPOSITE[d])
                visited.add(nxt)
                stack.append(nxt)
            else:
                stack.pop()
        return maze

    @staticmethod
    def shortest_path(maze, n):
        """Minimum number of moves from start to goal (breadth-first search)."""
        start, goal = (0, 0), (n - 1, n - 1)
        dist, queue = {start: 0}, deque([start])
        while queue:
            cell = queue.popleft()
            if cell == goal:
                return dist[cell]
            for d in maze[cell]:
                nxt = (cell[0] + DIRS[d][0], cell[1] + DIRS[d][1])
                if nxt not in dist:
                    dist[nxt] = dist[cell] + 1
                    queue.append(nxt)
        return None

    def draw_maze(self, maze, n, pos, time_left):
        self.screen.fill(BG)
        cs = min((self.W - 100) // n, (self.H - 140) // n)
        ox, oy = (self.W - cs * n) // 2, (self.H - cs * n) // 2 + 25
        goal = pygame.Rect(ox + (n - 1) * cs, oy + (n - 1) * cs, cs, cs).inflate(-cs // 4, -cs // 4)
        pygame.draw.rect(self.screen, GREEN, goal)
        pygame.draw.circle(self.screen, YELLOW, (ox + pos[1] * cs + cs // 2, oy + pos[0] * cs + cs // 2),
                           max(4, cs // 3))
        for (r, c), opens in maze.items():
            x, y = ox + c * cs, oy + r * cs
            if "N" not in opens:
                pygame.draw.line(self.screen, WHITE, (x, y), (x + cs, y), 3)
            if "S" not in opens:
                pygame.draw.line(self.screen, WHITE, (x, y + cs), (x + cs, y + cs), 3)
            if "W" not in opens:
                pygame.draw.line(self.screen, WHITE, (x, y), (x, y + cs), 3)
            if "E" not in opens:
                pygame.draw.line(self.screen, WHITE, (x + cs, y), (x + cs, y + cs), 3)
        self.draw_text(f"Time left: {max(0, time_left):.0f} s", 26, GREY, y=28)
        pygame.display.flip()

    def run_maze(self, n, level, trial, practice=False):
        maze = self.generate_maze(n)
        optimal = self.shortest_path(maze, n)
        keymove = {pygame.K_UP: "N", pygame.K_DOWN: "S", pygame.K_LEFT: "W", pygame.K_RIGHT: "E"}
        pos, moves, bumps, completed = (0, 0), 0, 0, 0
        self.fixation(600)
        pygame.event.clear()
        onset = time.perf_counter()
        while True:
            elapsed = time.perf_counter() - onset
            if elapsed >= MAZE_TIME_LIMIT_S:
                break
            for e in self.get_events():
                if e.type == pygame.KEYDOWN and e.key in keymove:
                    d = keymove[e.key]
                    if d in maze[pos]:
                        pos = (pos[0] + DIRS[d][0], pos[1] + DIRS[d][1])
                        moves += 1
                    else:
                        bumps += 1
            if pos == (n - 1, n - 1):
                completed = 1
                break
            self.draw_maze(maze, n, pos, MAZE_TIME_LIMIT_S - elapsed)
            self.clock.tick(120)
        duration = round(time.perf_counter() - onset, 2)
        if practice:
            self.feedback(completed, too_slow=not completed)
            return
        self.record("Maze", trial, **{
            "Difficulty level (1=easy)": level, "Maze size (cells per side)": n,
            "Completed (1=yes, 0=time out)": completed, "Time (s)": duration,
            "Moves made": moves, "Shortest possible moves": optimal,
            "Extra moves": max(0, moves - optimal) if completed else "",
            "Wall bumps": bumps})

    def test_maze(self):
        self.message("MAZE TASK\n\nUse the ARROW KEYS to move the yellow dot\n"
                     "to the green square as quickly as possible.\n\n"
                     "The mazes get bigger as you go.\n"
                     f"You have {MAZE_TIME_LIMIT_S} seconds per maze.")
        self.practice_then_start(lambda: self.run_maze(5, 0, 0, practice=True))
        trial = 0
        for level, (size, count) in enumerate(MAZE_LEVELS, 1):
            for _ in range(count):
                trial += 1
                self.run_maze(size, level, trial)

    # ------------------------------------------------------------ Digit span
    def digit_trial(self, length, trial, practice=False):
        seq = [random.randint(0, 9)]
        while len(seq) < length:
            d = random.randint(0, 9)
            if d != seq[-1]:                        # no immediate repeats
                seq.append(d)
        self.fixation(800)
        for d in seq:
            self.screen.fill(BG)
            self.draw_text(str(d), 140)
            pygame.display.flip()
            self.wait(DIGIT_SHOW_MS)
            self.blank(DIGIT_GAP_MS)
        onset = time.perf_counter()
        answer = self.text_input("Type the digits in the same order:", max_len=DIGIT_MAX_LEN + 2)
        duration = round(time.perf_counter() - onset, 2)
        target = "".join(map(str, seq))
        correct = int(answer == target)
        if practice:
            self.feedback(correct)
            return correct
        self.record("Digit span", trial, **{
            "Sequence length": length, "Digits shown": target, "Answer typed": answer,
            "Correct (1/0)": correct, "Answer time (s)": duration})
        return correct

    def test_digit_span(self):
        self.message("DIGIT MEMORY TASK\n\nDigits will appear one at a time.\n"
                     "Remember them, then type them in the SAME order\nand press ENTER.\n\n"
                     "The sequences get longer as you go.")
        self.practice_then_start(lambda: self.digit_trial(3, 0, practice=True))
        length, trial, best = DIGIT_START_LEN, 0, 0
        while length <= DIGIT_MAX_LEN:
            results = []
            for _ in range(2):
                trial += 1
                results.append(self.digit_trial(length, trial))
            if not any(results):
                break
            best = length
            length += 1
        self.spans["Digit span"] = best

    # ---------------------------------------------------------------- Stroop
    def test_stroop(self):
        names = list(STROOP_COLORS)
        keymap = {key: name for name, (_, key) in STROOP_COLORS.items()}
        legend = "R = red      G = green      B = blue      Y = yellow"
        self.message("STROOP TASK\n\nA colour word appears in coloured ink.\n"
                     "Respond to the INK COLOUR, not to the word.\n\n" + legend)

        def make_trials(n):
            trials = []
            for i in range(n):
                word = random.choice(names)
                if i < n // 2:
                    trials.append((word, word, "congruent"))
                else:
                    trials.append((word, random.choice([c for c in names if c != word]), "incongruent"))
            random.shuffle(trials)
            return trials

        def run(trials, practice):
            for i, (word, ink, cond) in enumerate(trials, 1):
                self.iti()
                self.screen.fill(BG)
                self.draw_text(word, 90, STROOP_COLORS[ink][0])
                self.draw_text(legend, 22, GREY, y=self.H - 50)
                pygame.display.flip()
                key, rt = self.wait_response(list(keymap), STROOP_RESPONSE_WINDOW_MS)
                response = keymap.get(key, "no response")
                correct = int(response == ink)
                if practice:
                    self.feedback(correct, too_slow=key is None)
                    continue
                self.record("Stroop", i, **{
                    "Condition": cond, "Word shown": word, "Ink colour (correct answer)": ink,
                    "Response": response, "Correct (1/0)": correct,
                    "Reaction time (ms)": rt if rt is not None else ""})

        self.practice_then_start(lambda: run(make_trials(4), True))
        run(make_trials(STROOP_TRIALS), False)

    # -------------------------------------------------------------- Go/No-Go
    def test_gonogo(self):
        self.message("GO / NO-GO TASK\n\nGREEN CIRCLE  ->  press SPACE as fast as possible\n"
                     "RED SQUARE  ->  do NOT press anything")

        def make_trials(n):
            n_go = round(n * GONOGO_GO_PROPORTION)
            trials = ["go"] * n_go + ["nogo"] * (n - n_go)
            random.shuffle(trials)
            return trials

        def run(trials, practice):
            for i, cond in enumerate(trials, 1):
                self.iti()
                self.screen.fill(BG)
                if cond == "go":
                    pygame.draw.circle(self.screen, GREEN, (self.W // 2, self.H // 2), 90)
                else:
                    pygame.draw.rect(self.screen, RED, pygame.Rect(self.W // 2 - 85, self.H // 2 - 85, 170, 170))
                pygame.display.flip()
                key, rt = self.wait_response([pygame.K_SPACE], GONOGO_RESPONSE_WINDOW_MS,
                                             stim_ms=GONOGO_STIM_MS)
                pressed = int(key is not None)
                correct = int(pressed == (cond == "go"))
                if practice:
                    self.feedback(correct, too_slow=(cond == "go" and not pressed))
                    continue
                if cond == "go":
                    outcome = "hit" if pressed else "miss"
                else:
                    outcome = "false alarm" if pressed else "correct rejection"
                self.record("Go-NoGo", i, **{
                    "Trial type (go=green circle, nogo=red square)": cond,
                    "Pressed SPACE (1/0)": pressed, "Outcome": outcome, "Correct (1/0)": correct,
                    "Reaction time (ms)": rt if rt is not None else ""})

        self.practice_then_start(lambda: run(make_trials(6), True))
        run(make_trials(GONOGO_TRIALS), False)

    # ----------------------------------------------------------------- Corsi
    def corsi_rects(self):
        size = int(min(self.W, self.H) * 0.1)
        rects = []
        for nx, ny in CORSI_POSITIONS:
            r = pygame.Rect(0, 0, size, size)
            r.center = (int(self.W * 0.15 + nx * self.W * 0.7), int(self.H * 0.12 + ny * self.H * 0.7))
            rects.append(r)
        return rects

    def draw_corsi(self, rects, lit=None, text=""):
        self.screen.fill(BG)
        for i, r in enumerate(rects):
            pygame.draw.rect(self.screen, YELLOW if i == lit else BLUE, r, border_radius=8)
        if text:
            self.draw_text(text, 26, GREY, y=30)
        pygame.display.flip()

    def corsi_trial(self, rects, length, trial, practice=False):
        seq = random.sample(range(len(rects)), length)
        self.draw_corsi(rects, text="Watch")
        self.wait(1000)
        for b in seq:
            self.draw_corsi(rects, lit=b, text="Watch")
            self.wait(CORSI_SHOW_MS)
            self.draw_corsi(rects, text="Watch")
            self.wait(CORSI_GAP_MS)

        prompt = "Click the blocks in the same order"
        pygame.event.clear()
        onset = time.perf_counter()
        clicks, lit, lit_until, duration = [], None, 0.0, None
        while len(clicks) < length or time.perf_counter() < lit_until:
            for e in self.get_events():
                if e.type == pygame.MOUSEBUTTONDOWN and e.button == 1 and len(clicks) < length:
                    for i, r in enumerate(rects):
                        if r.collidepoint(e.pos):
                            clicks.append(i)
                            lit, lit_until = i, time.perf_counter() + 0.2
                            if len(clicks) == length:
                                duration = round(time.perf_counter() - onset, 2)
                            break
            self.draw_corsi(rects, lit if time.perf_counter() < lit_until else None, prompt)
            self.clock.tick(120)

        correct = int(clicks == seq)
        if practice:
            self.feedback(correct)
            return correct
        self.record("Corsi", trial, **{
            "Sequence length": length,
            "Blocks shown (in order)": "-".join(str(b + 1) for b in seq),
            "Blocks clicked (in order)": "-".join(str(b + 1) for b in clicks),
            "Correct (1/0)": correct, "Response time (s)": duration})
        return correct

    def test_corsi(self):
        self.message("BLOCK SEQUENCE TASK\n\nSome blocks will light up one after another.\n"
                     "Afterwards, CLICK the blocks with the mouse\nin the SAME order.\n\n"
                     "The sequences get longer as you go.")
        pygame.mouse.set_visible(True)
        rects = self.corsi_rects()
        self.practice_then_start(lambda: self.corsi_trial(rects, 2, 0, practice=True))
        length, trial, best = CORSI_START_LEN, 0, 0
        while length <= min(CORSI_MAX_LEN, len(rects)):
            results = []
            for _ in range(2):
                trial += 1
                results.append(self.corsi_trial(rects, length, trial))
            if not any(results):
                break
            best = length
            length += 1
        self.spans["Corsi"] = best
        pygame.mouse.set_visible(False)

    # --------------------------------------------------------------- Flanker
    def test_flanker(self):
        self.message("FLANKER TASK\n\nYou will see a row of arrows.\n"
                     "Respond to the direction of the MIDDLE arrow only.\n\n"
                     "LEFT ARROW key = middle arrow points left\n"
                     "RIGHT ARROW key = middle arrow points right")
        keymap = {pygame.K_LEFT: "left", pygame.K_RIGHT: "right"}
        symbol = {"left": "<", "right": ">"}

        def make_trials(n):
            trials = []
            for i in range(n):
                target = random.choice(["left", "right"])
                cond = "congruent" if i < n // 2 else "incongruent"
                flank = target if cond == "congruent" else ("right" if target == "left" else "left")
                trials.append((symbol[flank] * 2 + symbol[target] + symbol[flank] * 2, target, cond))
            random.shuffle(trials)
            return trials

        def run(trials, practice):
            for i, (stim, target, cond) in enumerate(trials, 1):
                self.iti()
                self.screen.fill(BG)
                self.draw_text(stim, 100)
                pygame.display.flip()
                key, rt = self.wait_response(list(keymap), FLANKER_RESPONSE_WINDOW_MS)
                response = keymap.get(key, "no response")
                correct = int(response == target)
                if practice:
                    self.feedback(correct, too_slow=key is None)
                    continue
                self.record("Flanker", i, **{
                    "Condition": cond, "Arrows shown": stim, "Middle arrow (correct answer)": target,
                    "Response": response, "Correct (1/0)": correct,
                    "Reaction time (ms)": rt if rt is not None else ""})

        self.practice_then_start(lambda: run(make_trials(4), True))
        run(make_trials(FLANKER_TRIALS), False)


TEST_NAMES = ["Maze", "Digit span", "Stroop", "Go-NoGo", "Corsi", "Flanker"]


def main() -> None:
    import argparse

    ap = argparse.ArgumentParser(description="Cognitive test battery")
    ap.add_argument("--only", nargs="+", choices=TEST_NAMES, metavar="TEST",
                    help="run only these tests / запустить только эти тесты: " + ", ".join(TEST_NAMES))
    Battery(only=ap.parse_args().only).run()


if __name__ == "__main__":
    main()
