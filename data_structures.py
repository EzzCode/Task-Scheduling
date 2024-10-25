# data_structures.py
# This file contains the core data structures used in the scheduling application.

from collections import defaultdict

class Task:
    def __init__(self, name, durations, priority=0, dependencies=None):
        self.name = name
        self.durations = durations
        self.priority = priority
        self.dependencies = dependencies or []
        self.start_day = None
        self.end_day = None

    def __repr__(self):
        return f"Task(name={self.name}, durations={self.durations}, priority={self.priority}, dependencies={self._dep_names()})"

    def _dep_names(self):
        return [dep.name for dep in self.dependencies]

class Resource:
    def __init__(self, name, department):
        self.name = name
        self.department = department
        self.availability = defaultdict(lambda: 1)

    def __repr__(self):
        return f"Resource(name={self.name}, department={self.department}, availability={dict(self.availability)})"

class Schedule:
    def __init__(self, resources):
        self.resources = resources
        self.assignments = defaultdict(list)
        self.resource_workloads = defaultdict(list)
        self.task_resource_map = {}
        self.qc_task_resource_map = {}

    def find_available_days(self, resource, duration, start_day):
        available_days = []
        consecutive_count = 0
        current_day = start_day
        while consecutive_count < duration:
            if resource.availability[current_day] >= 0.5:
                consecutive_count += resource.availability[current_day]
                available_days.append(current_day)
            else:
                consecutive_count = 0
                available_days = []
            current_day += 1
        return available_days

    def assign_task_to_resource(self, task, resource, days):
        duration = task.durations[resource.department]
        remaining_duration = duration
        for day in days:
            if remaining_duration > 0:
                work_time = min(resource.availability[day], remaining_duration)
                self.assignments[resource.name].append((task, day, work_time))
                resource.availability[day] -= work_time
                remaining_duration -= work_time
            if remaining_duration <= 0:
                break
        task.start_day = days[0]
        task.end_day = days[-1]

    def print_schedule(self):
        print("Schedule:")
        task_total_durations = defaultdict(float)
        for assignments in self.assignments.values():
            for task, _, work_time in assignments:
                task_total_durations[task.name] += work_time
        for resource_name, assignments in self.assignments.items():
            print(f"\nResource: {resource_name}")
            for task, day, work_time in assignments:
                task_duration = task_total_durations.get(task.name, 'N/A')
                print(f"  Task: {task.name}, Day: {day}, Work Time: {work_time}, Duration: {task_duration:.1f}")