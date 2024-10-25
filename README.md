# Task Scheduling Application

This Task Scheduling Application is a Python-based tool designed to automate scheduling tasks across departments and resources. It prioritizes tasks, manages dependencies, and outputs a clear, organized Excel schedule. The application uses a topological sorting algorithm to order tasks based on their dependencies and priorities.

## Features

- **Dependency Management**: Automatically considers task dependencies to ensure the correct execution order.
- **Resource Allocation**: Assigns tasks based on resource availability and departmental affiliations.
- **Priority-Based Scheduling**: Tasks are scheduled according to their priority level, where 1 is the highest priority.
- **Excel Integration**: Reads task and resource data from Excel and outputs the finalized schedule to a formatted Excel file.
- **Graphical User Interface (GUI)**: A user-friendly interface built with Tkinter for easier interaction.

## Key Components

### 1. Data Input
The application reads two sheets from an Excel file:
- **Tasks Sheet (Sheet1)**: Includes task names, priorities, durations, and associated departments.
- **Resources Sheet (Sheet2)**: Lists resources along with their respective departments.

### 2. Scheduling Algorithm
The scheduling algorithm follows these steps:
- **Topological Sorting**: Orders tasks based on dependencies.
- **Resource Assignment**: Finds resources with the earliest availability, balancing workload.
- **QC Tasks Handling**: Ensures that a resource who handles a task's QC creation also handles its QC execution.

### 3. Output Generation
The output is an organized Excel file with:
- Auto-generated colors for different tasks.
- Merged cells for department headers.
- Automatically adjusted column widths for readability.

## GUI Overview

The application’s GUI provides:
- **File Selection**: Choose input and output Excel files.
- **Dependency Rules Editor**: Allows modification of task dependencies.
- **Theme Support**: Toggle between light and dark modes.
- **Instructions**: Embedded HTML-based guide for user support.

## Installation

1. Clone this repository:
    ```bash
    git clone https://github.com/yourusername/TaskSchedulingApp.git
    ```
2. Install required packages:
    ```bash
    pip install -r requirements.txt
    ```
3. Run the application:
    ```bash
    python main.py
    ```

## Usage

1. **Load Excel Data**: Select your input Excel file containing tasks and resources.
2. **Set Dependencies**: Modify dependencies through the Dependency Rules Editor if needed.
3. **Run Scheduler**: Click the "Run Scheduler" button to generate the schedule.
4. **View Output**: The generated Excel file will open automatically, displaying the scheduled tasks.

## Dependencies

- `pandas`: For reading/writing Excel files.
- `tkinter`: GUI components.
- `openpyxl`: Excel file manipulation.
- `json`: For reading dependency rules.


## Acknowledgments

Special thanks to Orange Egypt's Digital Demand and Solutions team for the mentorship and guidance in Agile workflows and project development.

---

> **Note**: This application was developed as part of an internship project, applying real-world Agile methodologies and SDLC principles.
