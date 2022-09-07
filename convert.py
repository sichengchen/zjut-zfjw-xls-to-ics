import xlrd
from ics import Calendar, Event
import re
import time


# Configurations
workbook_name = 'test.xls' # Name of the excel file download from Zhengfang Jiaowu system
the_first_day_of_the_semester = '2022-09-05' # YYYY-MM-DD
timezone = 8 # Timezone (8 for China/Shanghai UTC+8) 
ics_file_name = 'timetable.ics' # Name of the ics file to be generated


# Definitions
class Course:
    def __init__(self, name, time_str, location, professor):
        self._name = name
        self._time_str = time_str
        self._location = location
        self._professor = professor
        time_data = re.findall(r'\d+', self._time_str)
        day_data = re.findall(r'[一二三四五六日]', self._time_str)
        start_time_l = int(time_data[0])
        end_time_l = int(time_data[1])
        chinese_to_weekday = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '日': 0}
        time_table_start = [[8,0],[8,55],[9,55],[10,50],[11,45],[13,30],[14,25],[15,25],[16,20],[18,30],[19,25],[20,20]]
        time_table_end = [[8,45],[9,40],[10,40],[11,35],[12,30],[14,15],[15,10],[16,10],[17,5],[19,15],[20,10],[21,5]]
        self._start_time = time_table_start[start_time_l-1]
        self._end_time = time_table_end[end_time_l-1]
        self._start_week = int(time_data[2])
        self._end_week = int(time_data[3])
        self._day = chinese_to_weekday[day_data[0]]

    def __str__(self):
        return self._name + '|' + str(self._start_time) + '|' + str(self._end_time) + '|' + str(self._start_week) + '|' + str(self._end_week) + '|' + self._location + '|' + self._professor + '|' + str(self._day)
    def get_name(self):
        return self._name

    def get_time_str(self):
        return self._time_str
    
    def get_location(self):
        return self._location
    
    def get_professor(self):
        return self._professor

    def get_start_time(self):
        return self._start_time

    def get_end_time(self):
        return self._end_time

    def get_start_week(self):
        return self._start_week

    def get_end_week(self):
        return self._end_week

    def get_day(self):
        return self._day
        

# Open XLS file
book = xlrd.open_workbook(workbook_name)
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("Sheet {0} selected. nrows:{1} ncols:{2}".format(sh.name, sh.nrows, sh.ncols))


# Parse XLS file to courses_lst
keys_lst = []
courses_lst = []

for rownum in range(sh.nrows):
    row_values = sh.row_values(rownum)
    if rownum == 0:
        keys_lst = row_values
    else:
        print("Course {0}:".format(rownum))
        course = {}
        for colnum in range(sh.ncols):
            print("  {0}: {1}".format(keys_lst[colnum], row_values[colnum]))
            if(keys_lst[colnum] == "上课时间"):
                course[keys_lst[colnum]] = row_values[colnum].split(";")
            elif(keys_lst[colnum] == "上课地点"):
                course[keys_lst[colnum]] = row_values[colnum].split(";")
            else:
                course[keys_lst[colnum]] = row_values[colnum]
            
        courses_lst.append(course)

split_courses_lst = []

for course in courses_lst:
    if(len(course["上课时间"]) > 1):
        for i in range(len(course["上课时间"])):
            new_course = course.copy()
            new_course["上课时间"] = course["上课时间"][i]
            new_course["上课地点"] = course["上课地点"][i]
            split_courses_lst.append(new_course)
    else:
        new_course = course.copy()
        new_course["上课时间"] = course["上课时间"][0]
        new_course["上课地点"] = course["上课地点"][0]
        split_courses_lst.append(new_course)

class_courses_lst = []

for course in split_courses_lst:
    class_courses_lst.append(Course(course["课程名称"], course["上课时间"], course["上课地点"], course["任课教师"]))

print("\nParsed courses lst:")
for course in class_courses_lst:
    print(course)


# Write to ICS file
print("\nWriting to ICS file...")

first_day_of_semeseter = time.strptime(the_first_day_of_the_semester, "%Y-%m-%d")

c = Calendar()

for course in class_courses_lst:
    for week in range(course.get_start_week(), course.get_end_week()+1):
        e = Event()
        e.name = course.get_name()
        e.begin = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.mktime(first_day_of_semeseter) + (week-1)*7*24*60*60 + course.get_day()*24*60*60 + course.get_start_time()[0]*60*60 + course.get_start_time()[1]*60 + timezone*60*60 - 2*24*60*60 + 8*60*60))
        e.end = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.mktime(first_day_of_semeseter) + (week-1)*7*24*60*60 + course.get_day()*24*60*60 + course.get_end_time()[0]*60*60 + course.get_end_time()[1]*60 + timezone*60*60 - 2*24*60*60 + 8*60*60))
        e.location = course.get_location()
        e.description = course.get_professor()
        c.events.add(e)
        
with open(ics_file_name, 'w') as my_file:
    my_file.writelines(c.serialize_iter())

print("Done!")