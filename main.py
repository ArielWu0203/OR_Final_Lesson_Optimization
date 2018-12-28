import openpyxl
import pulp
import random
import re
import utils
import uuid

# define constant
grade = 4
semester = 1
lesson_number = [27, 36, 164] # required, elective, general
min_credit = 16
max_credit = 25
ratio = [0.1, 0.7, 0.2] # time, love, easy

# import all electives
wb = openpyxl.load_workbook("OR_sheets.xlsx")
time_prefer_sheet = wb.get_sheet_by_name("時間喜好")
required_sheet = wb.get_sheet_by_name("系上必修")
elective_sheet = wb.get_sheet_by_name("系上選修")
general_sheet = wb.get_sheet_by_name("通識")

# sort out all lesson information
def config_required_lesson(grade, semester, time_table, current_credit):
	# config current semester required lesson in time table
	for row in range(2, lesson_number[0]):
		lesson = required_sheet[row]
		if grade != int(lesson[0].value[0]) or not ((semester == 1 and lesson[0].value[1] == '+') or (semester == 2 and lesson[0].value[1] == '-')):
			continue
		# update current credit
		current_credit += int(lesson[2].value)
		# config required lesson to time table
		result = utils.parse_time(lesson[3].value)
		for time in result:
			for i in time["period"]:
				# None represent that you can not choose lesson at the time
				time_table[time["day"]][i] = None

def get_available_lesson(grade, semester, time_table, lesson_type):
	# get specified lesson
	lessons = {}
	# sort out all lesson
	for row in range(2, lesson_number[lesson_type]):
		lesson = elective_sheet[row]
		# comfirm that there is no overlapping class
		parse_result = utils.parse_time(lesson[3].value)
		flag = False
		for time in parse_result:
			for i in time["period"]:
				if time_table[time["day"]][i] == None:
					flag = True
		if lesson_type != 2 and grade != int(lesson[1].value[0]) or not ((semester == 1 and lesson[1].value[1] == '+') or (semester == 2 and lesson[1].value[1] == '-')):
			continue
		if flag:
			continue
		target_lesson = {
			"type": lesson_type,
			"grade": grade,
			"semester": "+" if semester == 1 else "-",
			"name": lesson[0].value,
			"time": utils.parse_time(lesson[3].value),
			"love": lesson[4].value,
			"easy": lesson[5].value,
			"credit": lesson[2].value
			}
		lessons[str(uuid.uuid1())] = target_lesson

	return lessons

def get_time_prefer_table(grade, semester, time_table):
	time = time_prefer_sheet[grade * semester + 1]
	for i in range(1, 11):
		if i % 2 == 0:
			continue
		for j in range(0, 5):
			time_table[str(int(i / 2) + 1)][j] = time[i].value
		for j in range(5, 15):
			time_table[str(int(i / 2) + 1)][j] = time[i + 1].value
	return time_table

def main():
	# define lesson table
	time_lesson_table = {
		"1+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"1-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},
		"2+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"2-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},
		"3+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"3-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},
		"4+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"4-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},

	}
	# define lesson table
	time_prefer_table = {
		"1+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"1-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},
		"2+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"2-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},
		"3+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"3-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},
		"4+": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
		},
		"4-": {
			"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
			"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
			"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
			"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
			"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday	
		},

	}
	# define available lesson list
	available_lessons = {}

	# define current credit
	current_credit = 0
	for grade in range(1, 5):
		for semester in range(1, 3):
			config_required_lesson(grade, semester, time_lesson_table["{}{}".format(grade, '+' if semester == 1 else '-')], current_credit)
			print("{} - {}".format(grade, semester))
			get_time_prefer_table(grade, semester, time_prefer_table["{}{}".format(grade, '+' if semester == 1 else '-')])
			available_lessons.update(get_available_lesson(grade, semester, time_lesson_table["{}{}".format(grade, '+' if semester == 1 else '-')], 1))
    #print("Time Lesson Table")
	#print(time_lesson_table)
	#print("Time Prefer Table")
	#print(time_prefer_table)
	#print("All Avaliable Elective Lesson")
	#print(available_lessons)
    
    # TODO: use pulp to solve optimize solution for elevtive lesson
	problem = pulp.LpProblem("Choosing Lesson", pulp.LpMaximize)
	# define all available lessons to optimization variable
	names = [key for key in available_lessons]
	lesson_variables = pulp.LpVariable.dict('x', names, lowBound=0, cat="Binary")

	# config object expression
	problem += pulp.lpSum(
			[time_prefer_table["{}{}".format(available_lessons[key]["grade"], available_lessons[key]["semester"])][str(time["day"])][time["period"][0]] / available_lessons[key]["credit"] * ratio[0] * lesson_variables[key] for time in available_lessons[key]["time"]] + 
			available_lessons[key]["love"] * ratio[1] * lesson_variables[key] + 
			available_lessons[key]["easy"] * ratio[2] * lesson_variables[key]
			for key in lesson_variables
			)
	# config basic restricted expression
	problem += pulp.LpAffineExpression(
			(lesson_variables[key], available_lessons[key]["credit"])
			for key in lesson_variables
			) <= 37
	problem += pulp.LpAffineExpression(
			(lesson_variables[key], available_lessons[key]["love"])
			for key in lesson_variables
			) >= 60
	problem += pulp.LpAffineExpression(
			(lesson_variables[key], available_lessons[key]["easy"])
			for key in lesson_variables
			) >= 60
	# config same time lesson restricted expression
	overlaps = utils.find_same_time_lesson(available_lessons)
	for target in overlaps:
		problem += pulp.lpSum(
				lesson_variables[i] for i in target
				) == 1
	print("Problem")
	print(problem)
	# start solve problem
	problem.solve()
	for v in problem.variables():
		print(available_lessons[v.name[2:].replace("_", "-")]["name"], '=', v.varValue)

	print('obj = ', pulp.value(problem.objective))
	
	# TODO: use pulp to solve optimize solution for general lesson
	# export result time table
main()
