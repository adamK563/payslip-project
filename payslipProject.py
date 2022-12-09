import datetime
from openpyxl import load_workbook

class Shift:
  def __init__(self, start_time, end_time, worker):
    self.start_time = start_time
    self.end_time = end_time
    self.worker = worker

class ShiftManagementApp:
  def __init__(self):
    self.shifts = []

  def create_shift(self, start_time, end_time, worker):
    shift = Shift(start_time, end_time, worker)
    self.shifts.append(shift)

  def get_shifts_for_worker(self, worker):
    worker_shifts = []
    for shift in self.shifts:
      if shift.worker == worker:
        worker_shifts.append(shift)
    return worker_shifts

  def get_shift_hours(self, shift):
    return (shift.end_time - shift.start_time).total_seconds() / 3600

app = ShiftManagementApp()

# Create some shifts
app.create_shift(
  datetime.datetime(2022, 12, 9, 9),
  datetime.datetime(2022, 12, 9, 17),
  "John Doe"
)
app.create_shift(
  datetime.datetime(2022, 12, 10, 13),
  datetime.datetime(2022, 12, 10, 21),
  "John Doe"
)
app.create_shift(
  datetime.datetime(2022, 12, 9, 9),
  datetime.datetime(2022, 12, 9, 17),
  "Jane Doe"
)

# Get all shifts for a worker
worker_shifts = app.get_shifts_for_worker("John Doe")
print(worker_shifts)

# Get the total hours worked for a shift
shift = worker_shifts[0]
hours = app.get_shift_hours(shift)
print(hours)


####
#reading a colored cell in excel
####


# Load the Excel workbook
workbook = load_workbook("file.xlsx")

# Get the active sheet
sheet = workbook.active

# Scan the cells in the sheet
for row in sheet.iter_rows():
    for cell in row:
        if cell.fill.start_color.index is not None:
            # The cell is colored
            print(f"Cell {cell.coordinate} is colored {cell.fill.start_color.index}")