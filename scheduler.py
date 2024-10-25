# scheduler.py
# This file contains the Scheduler class which handles the task scheduling logic.

from data_structures import Schedule

class Scheduler(Schedule):
    def __init__(self, resources):
        super().__init__(resources)

    def schedule_tasks(self, sorted_tasks):
        unassigned_tasks = sorted_tasks[:]
        
        while unassigned_tasks:
            progress_made = False
            for task in unassigned_tasks[:]:
                if all(dep.end_day is not None for dep in task.dependencies):
                    start_day = max((dep.end_day for dep in task.dependencies), default=0) + 1
                    if self._assign_task(task, start_day):
                        unassigned_tasks.remove(task)
                        progress_made = True
            
            if not progress_made:
                raise Exception("Unable to schedule all tasks. Resource constraints or circular dependencies may exist.")

    def _assign_task(self, task, start_day):
        sorted_resources = self._sort_resources(task, start_day)
        
        for resource in sorted_resources:
            if resource.department in task.durations:
                if self._handle_qc_task(task, resource):
                    duration = task.durations[resource.department]
                    available_days = self.find_available_days(resource, duration, start_day)
                    
                    if available_days:
                        self.assign_task_to_resource(task, resource, available_days)
                        self.resource_workloads[resource.name].append(duration)
                        self.task_resource_map[task.name] = resource.name
                        return True
        return False

    def _sort_resources(self, task, start_day):
        def resource_sort_key(resource):
            available_days = self.find_available_days(resource, 1, start_day)
            if available_days:
                return available_days[0], sum(self.resource_workloads[resource.name])
            else:
                return float('inf'), sum(self.resource_workloads[resource.name])
        
        return sorted(self.resources, key=resource_sort_key)

    def _handle_qc_task(self, task, resource):
        if "QC" in task.durations:
            base_task_name = task.name.split(": ")[1].split(" (")[0]
            if base_task_name in self.qc_task_resource_map:
                return self.qc_task_resource_map[base_task_name] == resource.name
            else:
                self.qc_task_resource_map[base_task_name] = resource.name
        return True