# Automatic Distribute Stuff Workout System
import prettytable as pt
import random
from xlwt import Workbook, easyxf
from xlrd import open_workbook
import xlrd

# Define status code for job control
# distribute in average: -1
# distribute in specific times when drop to 0
# distribute maxiumn: -2
EVNELY_DISTRIBUTE = -1
OUT_OF_DISTRIBUTE = 0
MAXIUMN_DISTRIBUTE = -2

# Register job table
data_from = 'job_dist.xlsx'
TOTAL_WEEKS = 16


"""
This function return a tupe which contians all the
rigister teacher distribution and every teacher's
job control settings
"""


def get_teacher_expect_distribute(data_from):
    sh = open_workbook(data_from).sheet_by_name('distributes')
    nrows, ncols = sh.nrows, sh.ncols
    datas = {}
    job_controls = {}
    teacher_id_name = {}

    for x in range(1, nrows):
        data = sh.row_values(x)
        if data[2] not in datas.keys():
            datas[data[2]] = []
            job_controls[data[2]] = {}

        datas[data[2]].append(data)
        job_controls[data[2]][data[0]] = int(data[3])
        teacher_id_name[data[0]] = str(data[1])
    return (datas, job_controls, teacher_id_name)


"""
Using job controls dictionary to calculate next piority class

:Parameters
job_control     dict
"""


def get_class_piority(job_controls):
    job_piority = {}
    job_piority_temps = {}
    for key, jobs in job_controls.items():
        total = 0
        for job in jobs.values():
            total += job
        job_piority_temps[key] = total / 3

    job_piority_temps = sorted(
        job_piority_temps.items(), key=lambda d: d[1], reverse=True)

    for job in job_piority_temps:
        job_piority[job[0]] = job[1]
    return job_piority


# Hold all the teacher_datas_and_job and job_controls information from the spreadsheet
datas, job_controls, teacher_id_name = get_teacher_expect_distribute(data_from)
CLASS_COUNT = len(datas)


# This function get a week of job distributions


def get_a_week_of_distribute():
    week_of_distribute = {}  # This hold a week of  distribute
    expect_list_in_key = []  # Using this to hold the expect keys
    job_controls_bak = job_controls.copy()

    # Loop through all the item in the data set to find a
    # valide list
    class_piority = get_class_piority(job_controls_bak)

    for key in class_piority.keys():
        # print(key, value)
        job_control = job_controls_bak[key]
        times = 100
        while(True):
            times = times-1
            if times < 0:
                break

            teacher_id, name, class_no, control = random.choice(datas[key])
            # print(teacher_id, name, class_no, control)
            if teacher_id not in expect_list_in_key:
                status = job_control.get(teacher_id, -3)
                if status > 0:
                    job_control.update(
                        {teacher_id: job_control.get(teacher_id) - 1})
                    expect_list_in_key.append(teacher_id)
                    week_of_distribute[key] = str(name)
                    break
                elif status == OUT_OF_DISTRIBUTE:
                    continue
                elif status == EVNELY_DISTRIBUTE:
                    expect_list_in_key.append(teacher_id)
                    week_of_distribute[key] = str(name)
                    break
    # print(len(week_of_distribute))
    if len(week_of_distribute) != CLASS_COUNT:
        return []

    week_of_distribute = sorted(
        week_of_distribute.items(), key=lambda d: d[0])

    week_of_distribute_temp = []
    for item in week_of_distribute:
        week_of_distribute_temp.append(item[1])

    job_controls.update(job_controls_bak)
    return week_of_distribute_temp


def prettytable_output():
    tb = pt.PrettyTable()
    tb.field_names = ["WEEK"] + [str(x+1)+" class" for x in range(CLASS_COUNT)]
    for i in range(TOTAL_WEEKS):
        ready_to_add = get_a_week_of_distribute()
        if ready_to_add:
            tb.add_row(["week "+str(i+1)] + ready_to_add)
    print(tb)

# print(datas)


# print(get_a_week_of_distribute())
# print(get_class_piority(job_controls))
prettytable_output()
