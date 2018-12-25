import openpyxl
import pulp
import random
import re
import utils

# define constant
grade = 4
semester = 1
lesson_number = [27, 36, 164] # required, elective, general
min_credit = 16
max_credit = 25

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
	lessons = []
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
			"type": 1,
			"name": lesson[0].value,
			"time": utils.parse_time(lesson[3].value),
			"love": random.randint(0, 10),
			"easy": random.randint(0, 10)
			}
		lessons.append(target_lesson)

	return lessons

def main():
	for grade in range(1, 5):
		for semester in range(1, 3):
			# define current credit
			current_credit = 0
			# define time table
			time_table = {
				"1": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Monday
				"2": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Tuesday
				"3": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Wednesday
				"4": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], # Thursday
				"5": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # Friday
				}
			config_required_lesson(grade, semester, time_table, current_credit)
			print("{} - {}".format(grade, semester))
			print(time_table)
			available_lessons = get_available_lesson(grade, semester, time_table, 1)
			print(available_lessons)
			# TODO: use pulp to solve optimize solution for elevtive lesson
			# TODO: use pulp to solve optimize solution for general lesson
			# export result time table

main()
