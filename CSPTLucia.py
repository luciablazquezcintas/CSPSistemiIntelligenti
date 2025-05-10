from constraint import Problem
import heapq
from collections import defaultdict
import itertools
import pandas as pd


# PARSING AND CHECKING FUNCTIONS

def parse_time_slot(slot): 
    day, hours = slot.split()
    start, end = map(int, hours.split("-"))
    return day, start, end

def no_overlap(slot1, slot2):
    day1, start1, end1 = parse_time_slot(slot1)
    day2, start2, end2 = parse_time_slot(slot2)
    if day1 != day2:
        return True
    return end1 <= start2 or end2 <= start1

def shortest_path_time(start, end, graph):
    if start == end:
        return 0
    visited = set()
    heap = [(0, start)]
    while heap:
        time, current = heapq.heappop(heap)
        if current == end:
            return time
        if current in visited:
            continue
        visited.add(current)
        for neighbor, cost in graph.get(current, []):
            if neighbor not in visited:
                heapq.heappush(heap, (time + cost, neighbor))
    return float('inf')

def no_travel_conflict(slot1, slot2, location_by_slot, graph):
    day1, start1, end1 = parse_time_slot(slot1)
    day2, start2, end2 = parse_time_slot(slot2)
    if day1 != day2:
        return True
    if end1 == start2 or end2 == start1:
        loc1 = location_by_slot.get(slot1)
        loc2 = location_by_slot.get(slot2)
        if shortest_path_time(loc1, loc2, graph) > 10:
            return False
    return True


# CLASSROOM LOCATIONS AND TIME SLOTS

location_by_slot = {
    "Mon 8-10": "A", "Mon 10-12": "A", "Mon 14-16": "B", "Mon 16-18": "C",
    "Tue 8-10": "B", "Tue 10-12": "C", "Tue 14-16": "D", "Tue 16-18": "A",
    "Wed 8-10": "C", "Wed 10-12": "D", "Wed 14-16": "B", "Wed 16-18": "A",
    "Thu 8-10": "A", "Thu 10-12": "B", "Thu 14-16": "D", "Thu 16-18": "C",
    "Fri 8-10": "B", "Fri 10-12": "C", "Fri 14-16": "D", "Fri 16-18": "A"
}

campus_graph = {
    "A": [("B", 7), ("C", 15), ("D", 12)],
    "B": [("A", 7), ("C", 5), ("D", 10)],
    "C": [("A", 15), ("B", 5), ("D", 6)],
    "D": [("A", 12), ("B", 10), ("C", 6)]
}

subjects = {
    "Mobile App Dev": ["Mon 8-10", "Fri 14-16"],
    "Sistemi Intelligenti": ["Tue 10-12", "Thu 14-16"],
    "Sistemi Elettronici": ["Fri 8-10", "Thu 16-18"],
    "Elettronica Biomedica": ["Wed 10-12", "Wed 16-18"],
    "Applicazioni di Fisica": ["Mon 14-16", "Tue 14-16"],
    "Circuiti": ["Wed 8-10", "Thu 10-12"],
    "Data Science": ["Tue 8-10", "Thu 8-10"],
    "AI Fundamentals": ["Mon 10-12", "Wed 14-16"],
    "Embedded Systems": ["Tue 16-18", "Fri 10-12"],
    "Computer Networks": ["Mon 16-18", "Thu 16-18"],
    "Biomedical Signals": ["Tue 10-12", "Wed 16-18"],
    "Digital Control": ["Wed 14-16", "Fri 8-10"],
    "Physics Lab": ["Thu 8-10", "Fri 16-18"],
    "Digital Design": ["Tue 14-16", "Thu 10-12"],
    "Robotics": ["Wed 10-12", "Fri 14-16"]
}


# CSP

problem = Problem()
for subject, slots in subjects.items():
    options = []
    if len(slots) > 1:
        options += [combo for combo in itertools.combinations(slots, 2) if no_overlap(*combo)]
    options += [(slot,) for slot in slots]
    problem.addVariable(subject, options)

subject_list = list(subjects.keys())
for i in range(len(subject_list)):
    for j in range(i + 1, len(subject_list)):
        s1, s2 = subject_list[i], subject_list[j]
        def constraint(a, b, s1=s1, s2=s2):
            for slot1 in a:
                for slot2 in b:
                    if not no_overlap(slot1, slot2) or not no_travel_conflict(slot1, slot2, location_by_slot, campus_graph):
                        return False
            return True
        problem.addConstraint(constraint, (s1, s2))

solutions = problem.getSolutions()


# SCHEDULE EVALUATOR

def evaluate_schedule(solution):
    flat_solution = {subj: slot for subj, slots in solution.items() for slot in (slots if isinstance(slots, tuple) else [slots])}
    daily_slots = defaultdict(list)
    for subj, slot in flat_solution.items():
        day, start, end = parse_time_slot(slot)
        daily_slots[day].append((start, end, location_by_slot[slot]))

    total_dead_time = 0
    total_moves = 0
    total_classes_attended = sum(len(slots) if isinstance(slots, tuple) else 1 for slots in solution.values())
    for day, blocks in daily_slots.items():
        blocks.sort()
        for i in range(len(blocks) - 1):
            _, end1, loc1 = blocks[i]
            start2, _, loc2 = blocks[i + 1]
            dead_time = start2 - end1
            total_dead_time += max(0, dead_time)
            if loc1 != loc2:
                total_moves += 1

    days_with_class = len(daily_slots)
    has_friday_free = "Fri" not in daily_slots

    return {
        "solution": flat_solution,
        "dead_time": total_dead_time,
        "moves": total_moves,
        "days": days_with_class,
        "bonus_friday_off": has_friday_free,
        "classes_attended": total_classes_attended
    }

# Rank the solutions
ranked = sorted(
    [evaluate_schedule(sol) for sol in solutions],
    key=lambda x: (-x['classes_attended'], x['dead_time'], x['moves'], x['days'], not x['bonus_friday_off'])
)

# Remove duplicate solutions 
unique_solutions = []
used_combinations = set()

for result in ranked:
    solution_combination = tuple(sorted(result["solution"].items()))
    if solution_combination not in used_combinations:
        used_combinations.add(solution_combination)
        unique_solutions.append(result)
    if len(unique_solutions) == 3:
        break

# TOP 3 TO EXCEL

output_path = r"C:\Users\Lucia\Downloads\threebests0.xlsx" # hay que cambiarla cada vez que se ejecute 

days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
hours = ["8-10", "10-12", "14-16", "16-18"]

# Function to print the summary to console
def print_schedule_summary(timetable):
    print(f"\nSchedule Summary:")
    for day in days:
        print(f"\n{day}:")
        for time in hours:
            subject = timetable.at[time, day]
            if subject:
                print(f"  {time} -> {subject}")


def generate_schedule_output(result, idx=None):
    timetable = pd.DataFrame("", index=hours, columns=days)

    for subject, slot in result["solution"].items():
        day, start, end = parse_time_slot(slot)
        time_range = f"{start}-{end}"
        timetable.at[time_range, day] = subject

    # Console
    print(f"\n\nSchedule #{idx}:")
    print_schedule_summary(timetable)

    
    stats = pd.DataFrame({
        "Metric": ["Dead Time", "Building Moves", "Days with Classes", "Friday Off", "Total Assigned"],
        "Value": [
            f"{result['dead_time']}h",
            result['moves'],
            result['days'],
            "Yes" if result["bonus_friday_off"] else "No",
            f"{result['classes_attended']} of {sum(len(v) for v in subjects.values())}"
        ]
    })

    return timetable, stats


# Print summary for each schedule and save to Excel
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    for idx, result in enumerate(unique_solutions, 1):
        timetable, stats = generate_schedule_output(result, idx)

        # Export both tables to Excel
        timetable.to_excel(writer, sheet_name=f"Schedule #{idx}", startrow=0, startcol=0)
        stats.to_excel(writer, sheet_name=f"Schedule #{idx}", startrow=len(hours) + 3, startcol=0, index=False)

# PRINTING TIMETABLE SUMMARY TO CONSOLE
print("\nFile saved to:", output_path)

if not unique_solutions:
    print("No valid solutions found")
