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
			for i in range(lower - 1, upper):
				target["period"].append(i)
		else:
			number = int(time[3]) if re.match(r"\d", time[3]) else  ord(time[3]) - ord("A") + 10
			target["period"].append(number)
		result.append(target)
	return result


