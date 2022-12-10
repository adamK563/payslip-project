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

#reading a colored cell in excel

# Load the Excel workbook
workbook = load_workbook("./November.xlsx")

# Get the active sheet
sheet = workbook.active

worker_name = "אדם".encode('utf-8')
morning_str = "בוקר".encode('utf-8')
evening_str = "ערב".encode('utf-8')
night_str = "לילה".encode('utf-8')
friday_str = "שישי".encode('utf-8')

for row in sheet.iter_rows():
  for cell in row:
    if cell.value == worker_name.decode('utf-8'):
      workers_row = row

# Scan the cells in the sheet
for row in sheet.iter_rows():
  if row == workers_row:
    for cell in row:
      if len(cell.coordinate) == 3:
        date =  ord(cell.coordinate[0]) - 65
      if len(cell.coordinate) == 4:
        date = ord("z") + ord(cell.coordinate[1]) - 25 

      if cell.value == morning_str.decode('utf-8'):
        
        print(f"Cell {cell.coordinate} is a morning shift")
        app.create_shift(
        datetime.datetime(2022, 11, date, 7),
        datetime.datetime(2022, 11, date, 16),
        worker_name
        )

      if cell.value == evening_str.decode('utf-8'):
        print(f"Cell {cell.coordinate} is a evening shift")
        app.create_shift(
        datetime.datetime(2022, 11, date, 16),
        datetime.datetime(2022, 11, date, 23),
        worker_name
        )
    
      if cell.value == night_str.decode('utf-8'):
        print(f"Cell {cell.coordinate} is a night shift")
        app.create_shift(
        datetime.datetime(2022, 11, date, 23),
        datetime.datetime(2022, 11, date + 1, 7),
        worker_name
        )
    
      if cell.value == friday_str.decode('utf-8'):
        print(f"Cell {cell.coordinate} is a friday shift")
        app.create_shift(
        datetime.datetime(2022, 11, date, 7),
        datetime.datetime(2022, 11, date, 16),
        worker_name
        )

# Get all shifts for a worker
worker_shifts = app.get_shifts_for_worker(worker_name)
print(worker_shifts)

# Get the total hours worked for a shift
shift = worker_shifts[0]
hours = app.get_shift_hours(shift)
print(hours)
