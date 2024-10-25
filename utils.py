# utils.py
# This file contains utility functions used in the scheduling process.

from collections import deque
import sys
import os

def calculate_total_duration(task, task_duration_map, graph, memo):
    """
    Recursively calculate the total duration of a task including its dependencies.
    Uses memoization to optimize performance for tasks with shared dependencies.
    """
    if task.name in memo:
        return memo[task.name]
    
    total_duration = task.durations.get(task.name, 0)
    for dep in task.dependencies:
        dep_task = next((t for t in graph if t.name == dep), None)
        if dep_task:
            dep_duration = calculate_total_duration(dep_task, task_duration_map, graph, memo)
            total_duration = max(total_duration, dep_duration)
    
    memo[task.name] = total_duration
    return total_duration

def topological_sort(tasks):
    """
    Perform a topological sort of tasks and then sort them by priority.
    In case of equal priorities, sort by total duration.
    """
    graph = {task: set() for task in tasks}
    in_degree = {task: 0 for task in tasks}
    task_durations = {task.name: task.durations for task in tasks}
    
    # Build the graph and calculate in-degrees
    for task in tasks:
        for dep in task.dependencies:
            dep_task = next((t for t in tasks if t.name == dep), None)
            if dep_task:
                graph[dep_task].add(task)
                in_degree[task] += 1
    
    # Perform topological sort
    queue = deque([task for task in tasks if in_degree[task] == 0])
    result = []
    memo = {}
    
    while queue:
        task = queue.popleft()
        result.append(task)
        for dependent in graph[task]:
            in_degree[dependent] -= 1
            if in_degree[dependent] == 0:
                queue.append(dependent)
    
    if len(result) != len(tasks):
        raise ValueError("Circular dependency detected")
    
    # Calculate total durations
    task_duration_map = {}
    for task in result:
        task_duration_map[task] = calculate_total_duration(task, task_durations, graph, memo)
    
    # Sort tasks by priority (ascending), and by total duration in case of equal priorities
    result.sort(key=lambda t: (t.priority, -task_duration_map[t]))
    
    return result


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)