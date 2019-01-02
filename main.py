import openpyxl
import plotly
import pulp
import random
import re
import utils
import uuid

# define constant
grade = 4
semester = 1
lesson_number = [28, 36, 165] # required, elective, general
min_credit = 16
max_credit = 25
#ratio = [0.1, 0.7, 0.2] # time, love, easy

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
                # lesson name represent that you can not choose lesson at the time
                time_table[time["day"]][i] = lesson[1].value

def get_available_lesson(grade, semester, time_table, lesson_type):
    # get specified lesson
    lessons = {}
    # sort out all lesson
    for row in range(2, lesson_number[lesson_type]):
        if lesson_type == 1:
            lesson = elective_sheet[row]
        else:
            lesson = general_sheet[row]
        # comfirm that there is no overlapping class
        parse_result = utils.parse_time(lesson[3].value)
        flag = False
        for time in parse_result:
            for i in time["period"]:
                if time_table[time["day"]][i] != 0:
                    flag = True
        if lesson_type != 2 and (grade != int(lesson[1].value[0]) or not ((semester == 1 and lesson[1].value[1] == '+') or (semester == 2 and lesson[1].value[1] == '-'))):
            continue
        if flag:
            continue
        if lesson_type == 2 and lesson[4].value == 0:
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

def main(time,love,easy):

    #define ratio
    ratio = [time,love,easy]
    print("ratio = ",ratio)
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

    #define objective score
    score = 0.0
    score1 = 0.0

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
    score1 = score1 + pulp.value(problem.objective)
    print("score1 = ",score1)

    # add elective to time_lesson_table
    for v in problem.variables():
        if v.varValue == 0:
            continue
        # grade+semester
        grade = str(available_lessons[v.name[2:].replace("_", "-")]["grade"])
        semester = available_lessons[v.name[2:].replace("_", "-")]["semester"]

        # class time array
        time = available_lessons[v.name[2:].replace("_", "-")]["time"]
        for item in time:
            day = item['day']
            period = item['period']
            for i in period:
                    time_lesson_table[grade+semester][day][i] = available_lessons[v.name[2:].replace("_", "-")]["name"]


    # define general lesson that chosen
    choose_general_lesson = {
            "1+": [],
            "1-": [],
            "2+": [],
            "2-": [],
            "3+": [],
            "3-": [],
            "4+": [],
            "4-": []
            }
    # config start to pulp
    for grade in range(1, 2):
        for semester in range(1,2):
            available_lessons = {}
            available_lessons.update(get_available_lesson(grade, semester, time_lesson_table["{}{}".format(grade, '+' if semester == 1 else '-')], 2))
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
                ) <= 50
            problem += pulp.LpAffineExpression(
                (lesson_variables[key], available_lessons[key]["love"])
                for key in lesson_variables
                ) >= 20
            problem += pulp.LpAffineExpression(
                (lesson_variables[key], available_lessons[key]["easy"])
                for key in lesson_variables
                ) >= 20
            # config same time lesson restricted expression
            overlaps = utils.find_same_time_lesson(available_lessons)
            for target in overlaps:
                problem += pulp.lpSum(
                    lesson_variables[i] for i in target
                    ) == 1
            print("General lessons "+str(grade)+"-"+str(semester))
            # start solve problem
            problem.solve()
            choose_lesson = []
            for v in problem.variables():
                if v.varValue > 0.0:
                    print(available_lessons[v.name[2:].replace("_", "-")]["name"], '=', v.varValue)
                    choose_lesson.append(available_lessons[v.name[2:].replace("_", "-")])
            # compute the compare value
            for target in choose_lesson:
                value = 0
                for time in target["time"]:
                    for i in time["period"]:
                        value += ratio[0] * time_prefer_table["{}{}".format(target["grade"], target["semester"])][time["day"]][i]
                    value = value / target["credit"]
                value += ratio[1] * target["love"]
                value += ratio[2] * target["easy"]
                target["value"] = value
            choose_lesson.sort(key= lambda x : x["value"], reverse=True)
            # update choose general lesson
            choose_general_lesson["{}{}".format(grade, "+" if semester == 1 else "-")] = choose_lesson
            for index, target in enumerate(choose_lesson):
                print("{} - {} - value: {}".format(index + 1, target["name"], target["value"]))

            print('obj = ', pulp.value(problem.objective))
            score =  score + pulp.value(problem.objective)
    
    print("score = ",score)
    print("total",score+score1)
    return (score+score1)

"""
    # export table
    for grade in range(1, 5):
        for semester in ['+', '-']:
            table = plotly.graph_objs.Table(
                    columnwidth = [80, 200],
                    header=dict(values=["", "Monday", "Tuesday", "Wednesday", "Thurseday", "Friday"],
                        line = dict(color="#CCCCCC"),
                        fill = dict(color="#119DEF"),
                        font = dict(color="white", size=20),
                        align = ["center"],
                        height = 60
                        ),
                    cells=dict(values=([["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E"]] + [time_lesson_table["{}{}".format(grade, semester)][key] for key in time_lesson_table["{}{}".format(grade, semester)]]),
                        line = dict(color="#CCCCCC"),
                        fill = dict(color="white"),
                        font = dict(color="#506784", size=12),
                        align = ["center"],
                        height = 60
                        )
                )
            layout = dict(width=1080, height=960)
            data = [table]
            fig = dict(data=data, layout=layout)
            filename = "{}-{} time table".format(grade, 1 if semester == '+' else 2)
            plotly.offline.plot(data, auto_open=False, filename=(filename + ".html"), image_filename=filename, image_width=1100, image_height=1200, image="svg")
            print("{} plot success!!".format(filename))
            
            table = plotly.graph_objs.Table(
                    columnwidth = [30, 240],
                    header=dict(values=["", "Name"],
                        line = dict(color="#CCCCCC"),
                        fill = dict(color="#119DEF"),
                        font = dict(color="white", size=16),
                        align = ["center"],
                        height = 40
                        ),
                    cells=dict(values=([[ index + 1 for index, key in enumerate(choose_general_lesson[str(grade) + semester])]] + [[target["name"] for target in choose_general_lesson[str(grade) + semester]]]),
                        line = dict(color="#CCCCCC"),
                        fill = dict(color="white"),
                        font = dict(color="#506784", size=12),
                        align = ["center"],
                        height = 30
                        )
                    )
            layout = dict(width=400, height=600)
            data = [table]
            fig = dict(data=data, layout=layout)
            filename = "{}-{} general choose order".format(grade, 1 if semester == '+' else 2)
            plotly.offline.plot(data, auto_open=False, filename=(filename + ".html"), image_filename=filename, image_width=400, image_height=600, image="svg")
            print("{} plot success!!".format(filename))
"""

# search for Max & min obj
def test() :
    Max = 0.0
    Min = 10000000.0
    max_ratio = [0.0,0.0,0.0]
    min_ratio = [0.0,0.0,0.0]

    for time in range(0,11,1):
        for love in range(0,11-time,1) :
            easy = 10-time-love
            ratio = [time/10,love/10,easy/10]
            score = 0.0
            score = main(ratio[0],ratio[1],ratio[2])
            if score > Max:
                Max = score
                max_ratio = ratio
            if score < Min:
                Min = score
                min_ratio = ratio
    print("ratio = ",max_ratio," Max obj = ",Max)
    print("ratio = ",min_ratio," Min obj = ",Min)

#test()

#search conditions
def test2():
    ratio = [[1.0,0.0,0.0],[0.0,1.0,0.0],[0.0,0.0,1.0],[0.5,0.5,0.0],[0.0,0.5,0.5],[0.5,0.0,0.5]]
    score = [0.0,0.0,0.0,0.0,0.0,0.0]
    for i in range(0,6) :
        score[i] = main(ratio[i][0],ratio[i][1],ratio[i][2])
    for i in range(0,6) :
        print("ratio = ",ratio[i]," obj  = ",score[i])

test2()

#print("total obj = ",main(0.1,0.7,0.2))
