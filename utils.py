import re

def parse_time(time_str):
	result = []
	matches = re.findall(r"(\[\d\]\d+~*\d*[A-Z]*)+", time_str)
	for time in matches:
		target = {
			"day": time[1],
			"period": []
			}
		if len(time) > 4:
			lower = int(time[3]) if re.match(r"\d", time[3]) else  ord(time[3]) - ord("A") + 10
			upper = int(time[5]) if re.match(r"\d", time[5]) else  ord(time[5]) - ord("A") + 10
			for i in range(lower, upper + 1):
				target["period"].append(i)
		else:
			number = int(time[3]) if re.match(r"\d", time[3]) else  ord(time[3]) - ord("A") + 10
			target["period"].append(number)
		result.append(target)
	return result

def filter_string(string):
	return re.sub(r"[\s+\.\!\/_,$%^*(+\"\']+|[+——！，。？、~@#￥%……&*（）]+", "", string)

def find_same_time_lesson(lessons):
	time_table = {
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

	buf = {}
	for key in lessons:
		times = lessons[key]["time"]
		for time in times:
			for i in time["period"]:
				field = time_table["{}{}".format(lessons[key]["grade"], '+' if lessons[key]["semester"] == 1 else '-')][time["day"]][i]
				if field != 0:
					buf[field].append(key)
					break
				else:
					time_table["{}{}".format(lessons[key]["grade"], '+' if lessons[key]["semester"] == 1 else '-')][time["day"]][i] = key
					buf[key] = []
	result = []
	for key in buf:
		target = []
		target.append(key)
		target = target + buf[key]
		if len(target) > 1:
			result.append(target)
	return result
