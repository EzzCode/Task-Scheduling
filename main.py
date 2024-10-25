# main.py
# This is the entry point of the application. It orchestrates the overall flow of the program.

from excel_io import read_tasks_from_excel, read_resources_from_excel, generate_excel_output
from scheduler import Scheduler
from utils import topological_sort

def main(tasks_file, output_file):
    # Read tasks and resources from Excel files
    tasks = read_tasks_from_excel(tasks_file)
    resources = read_resources_from_excel(tasks_file)
    
    # Combine QC Creation and QC execution tasks
    for task in tasks:
        if "QC Creation" in task.durations or "QC execution" in task.durations:
            task.durations = {"QC": task.durations.get("QC Creation", 0) + task.durations.get("QC execution", 0)}
    
    # Sort tasks topologically and by duration
    sorted_tasks = topological_sort(tasks)
    
    # Create a schedule and assign tasks
    schedule = Scheduler(resources)  
    schedule.schedule_tasks(sorted_tasks)
    # schedule.print_schedule()
    
    # Generate and save the Excel output
    generate_excel_output(schedule, output_file)

if __name__ == "__main__":
    main("input_tasks.xlsx", "output_schedule.xlsx")