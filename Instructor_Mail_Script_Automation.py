import pandas as pd
import openpyxl
from datetime import datetime, timedelta

# input, output, source 파일 세팅
SCHEDULE_FILE = '2019일정계획표(2019.08.02).xlsx'
INPUT_PARA_FILE = 'input_parameters.xlsx'
OUTPUT_SCRIPT_FILE = 'mail_script.xlsx'

EDU_SECTION_CELL_NUM = 'B4'
EDU_SECTION_COLUMN = 'A'
EDU_COLUMN = 'B'
EDU_ROOM_COLUMN = 'C'
CHANGED_EDU_ROOM_COLUMN = 'D'
TIME_COLUMN = 'J'
MONDAY_COLUMN = 'K'

GREETING_START_CELL_NUM = 'B5'
GREETING_END_CELL_NUM = 'B6'

# 파일로 받아오기
DETAILED_EDU_TIME = {
    "7": ["09:30 ~ 17:30"],
    "7-7": ["09:30 ~ 17:30", "09:30 ~ 17:30"],
    "7-7-4": ["09:30 ~ 17:30", "09:30 ~ 17:30", "09:00 ~ 13:00"],
    "4-7-7": ["14:00 ~ 18:00", "09:30 ~ 17:30", "09:30 ~ 17:30"],
    "6-6-6": ["10:00 ~ 17:00", "10:00 ~ 17:00", "10:00 ~ 17:00"],
    "7-7-7": ["09:30 ~ 17:30", "09:30 ~ 17:30", "09:30 ~ 17:30"],
    "7-7-7-7": ["09:30 ~ 17:30", "09:30 ~ 17:30", "09:30 ~ 17:30", "09:30 ~ 17:30"],
    "8-8-4": ["09:00 ~ 18:00", "09:00 ~ 18:00", "09:00 ~ 13:00"],
    "4-8-8": ["14:00 ~ 18:00", "09:00 ~ 18:00", "09:00 ~ 18:00"],
    "8-8-8-8-8": ["09:00 ~ 18:00", "09:00 ~ 18:00", "09:00 ~ 18:00", "09:00 ~ 18:00", "09:00 ~ 18:00"]
    }



def open_from_excel(fileName):
    workbook = openpyxl.load_workbook(fileName)
    return workbook

# schedule_workbook = open_from_excel(SCHEDULE_FILE)
input_para_workbook = open_from_excel(INPUT_PARA_FILE)
input_para_workbook = input_para_workbook['Sheet1']

edu_section = input_para_workbook[EDU_SECTION_CELL_NUM].value
greeting_start = input_para_workbook[GREETING_START_CELL_NUM].value
greeting_end = input_para_workbook[GREETING_END_CELL_NUM].value


###### output file 저장하는 부분 ####### 수정 필요!
df = pd.DataFrame(data={'col1': greeting_start, 'col2': greeting_end}, index=[1])
df.to_excel(OUTPUT_SCRIPT_FILE)


schedule_workbook = open_from_excel(SCHEDULE_FILE)
# 현재 월에 해당하는 sheet 불러오기
# current_month = str(datetime.now().month) + '월' ##########HERE IS TO CHANGE!!
current_month = '8월'
wb = schedule_workbook[current_month]


# 오늘 날짜에 해당하는 셀 번호 구하기
def get_today_cell_num():
    for row in wb:
        for cell in row:
            # 현재 날짜와 cell 날짜가 같다면 해당 셀 번호 저장
            try:
                if datetime.now().date() == cell.value.date():
                    current_date_cell = cell
                    break
            except:
                pass
    return current_date_cell # ex: <Cell '8월'.N50>



# 이번주 과정의 행 시작과 끝 구하기
# current_date_cell = get_today_cell_num()
current_date_cell = wb['L50']

week_start_row = current_date_cell.row

END_ROW_VALUE = '강사별 투입일수("공우식", "공우식/", "/공우식", "공우식?", "공우식?/", "/공우식?" 까지만 인식)'

def get_week_end_row(current_date_cell):
    '''현재 row에 날짜가 있거나, 현재 row, K열에 END_ROW_VALUE가 있다면 바로 그 위 행이 week_end_row'''
    week_end_row = current_date_cell.row
    while True:
        week_end_row = week_end_row + 1
        if type(wb[current_date_cell.column][week_end_row].value) == datetime or wb['K'][week_end_row].value == END_ROW_VALUE:
            week_end_row -= 1
            break
    return week_end_row


day_of_week_columns = ['K', 'L', 'M', 'N', 'O', 'P', 'Q']

# 오른쪽으로는 시작할 과정이 있는 모든 행을 찾아 리스트에 담는다.
upcoming_edu_rows = []
day_of_week_columns_index = day_of_week_columns.index(current_date_cell.column)
day_of_week_columns_index_fixed = day_of_week_columns_index

week_end_row = get_week_end_row(current_date_cell)

def append_upcoming_edu_rows_in_edu_rows(week_start_row, week_end_row):
    global day_of_week_columns_index
    for row in range(week_start_row, week_end_row):
        if wb[day_of_week_columns[day_of_week_columns_index]][row].value == None:
            while day_of_week_columns_index <= 3: # 월~목인 동안
                day_of_week_columns_index += 1
                if wb[day_of_week_columns[day_of_week_columns_index]][row].value != None:
                    upcoming_edu_rows.append(row)
                    break
            day_of_week_columns_index = day_of_week_columns_index_fixed

def append_next_week_upcoming_edu_rows_in_edu_rows(wb, next_week_row):
    next_week_monday_cell = wb[MONDAY_COLUMN][next_week_row]
    get_week_end_row(next_week_monday_cell)
    
    upcoming_edu_rows.append(row)
    

# 해당 행에서 나의 부문에 해당하는 행을 찾는다.
def get_selected_edu_sections_row():
    selected_edu_sections_row = []
    for row in upcoming_edu_rows:
        if wb[EDU_SECTION_COLUMN][row].value == edu_section:
            selected_edu_sections_row.append(row)
    return selected_edu_sections_row

# 3. 과정명, 강의실, 강사, 시간 등의 조합 알고리즘을 짠다

# 행
def append_edu_row_in_df_scripts():
    temp_list = []
    for row in selected_edu_sections_row:
        temp_list.append(row)
    df_scripts['edu_row'] = temp_list



# 과정명
def append_edu_name_in_df_scripts():
    temp_list = []
    for row in selected_edu_sections_row:
        temp_list.append(wb[EDU_COLUMN][row].value)
    df_scripts['edu_name'] = temp_list



# 강의실
def append_edu_room_in_df_scripts():
    temp_list = []
    for row in selected_edu_sections_row:
        if wb[CHANGED_EDU_ROOM_COLUMN][row].value == None:
            if type(wb[EDU_ROOM_COLUMN][row].value) == int:
                temp_list.append('KPC 서울 본부 '+ str(wb[EDU_ROOM_COLUMN][row].value) + '호 강의장')
            elif type(wb[EDU_ROOM_COLUMN][row].value) == str:
                temp_list.append('KPC ' + wb[EDU_ROOM_COLUMN][row].value + ' 본부')
        elif wb[CHANGED_EDU_ROOM_COLUMN][row].value != None:
            if type(wb[CHANGED_EDU_ROOM_COLUMN][row].value) == int:
                temp_list.append('KPC 서울 본부 '+ str(wb[CHANGED_EDU_ROOM_COLUMN][row].value) + '호 강의장')
            elif type(wb[CHANGED_EDU_ROOM_COLUMN][row].value) == str:
                temp_list.append('KPC ' + wb[CHANGED_EDU_ROOM_COLUMN][row].value + ' 본부')
    df_scripts['edu_room'] = temp_list



# 일정 도출을 위한 시작날짜, 종료날짜 구하기
def append_start_end_date_in_df_scripts():
    temp_start_list = []
    temp_end_list = []
    weekday_columns = ['K', 'L', 'M', 'N', 'O']
    for row_num in df_scripts['edu_row']:
        find_edu_start_date = False
        find_edu_end_date = False
        for column in weekday_columns:
            # 교육 시작일 찾기
            if wb[column][row_num].value != None and find_edu_start_date == False:
                edu_start_date = wb[column][week_start_row-1].value
                find_edu_start_date = True
                temp_start_list.append(edu_start_date)
            # 교육 종료일 찾기
            elif wb[column][row_num].value == None and find_edu_end_date == False and find_edu_start_date == True:
                # 빈 셀의 바로 왼쪽이 교육 종료일
                edu_end_column = weekday_columns[weekday_columns.index(column)-1]
                edu_end_date = wb[edu_end_column][week_start_row-1].value
                find_edu_end_date = True
                temp_end_list.append(edu_end_date)
            # 종료일이 금요일인 경우
            elif wb[column][row_num].value != None and find_edu_end_date == False and find_edu_start_date == True and column == 'O':
                edu_end_date = wb[column][week_start_row-1].value
                find_edu_end_date = True
                temp_end_list.append(edu_end_date)
    df_scripts['edu_start_date'] = temp_start_list
    df_scripts['edu_end_date'] = temp_end_list



# 강사
def append_instructors_in_df_scripts():
    temp_list = []
    weekday_columns = ['K', 'L', 'M', 'N', 'O']
    for row_num in df_scripts['edu_row']:
        instructors_in_one_edu = []
        for column in weekday_columns:
            if wb[column][row_num].value != None:
                instructor = wb[column][row_num].value
                # '/' 삭제
                instructor = instructor.replace("/", "")
                instructors_in_one_edu.append(instructor)
        temp_list.append(instructors_in_one_edu)
    df_scripts['instructors'] = temp_list

def append_full_edu_date_in_df_scripts():
    temp_list = []
    weekday_dict = {0: '월', 1: '화', 2: '수', 3: '목', 4: '금', 5: '토', 6: '일'}
    for index, row in df_scripts.iterrows():
        full_dates = []
        date_difference = (row['edu_end_date']-row['edu_start_date']).days + 1
        edu_date = row['edu_start_date']
        # 시작일부터 종료일까지 full_date 형식으로 저장
        for i in range(date_difference):
            start_month = '{:02d}'.format(edu_date.month)
            start_day = '{:02d}'.format(edu_date.day)
            start_weekday = weekday_dict[edu_date.weekday()]
            full_date = str(start_month) + '.' +str(start_day) + '(' + start_weekday +')'
            full_dates.append(full_date)
            # 다음날
            edu_date = edu_date + timedelta(days=1)
        temp_list.append(full_dates)
    df_scripts['full_dates'] = temp_list
        


# 강의 시간 구하기
def append_edu_time_in_df_scripts():
    ''' 1. 첫 숫자를 가져온다.
        2. 그 다음 문자가 '-'이면 다음 숫자까지 가져온다.
        3. 그 다음 문자가 '-'가 아닐 때 까지 반복한다.
    '''
    edu_time = []
    temp_list = []
    for row_num in df_scripts['edu_row']:
        time_strings = wb[TIME_COLUMN][row_num].value
        string_index = 0
        for index in range(len(time_strings)):
            
            # time으로만 이루어진 string의 마지막 index일 때
            if index+1 == len(time_strings) and time_strings[index-1] == '-':
                string_index = index
                break
            # time 이외의 문구가 있는 string의 마지막 index일 때
            elif index+1 == len(time_strings) and time_strings[index-1] != '-':
                continue
            elif time_strings[index+1] == '-':
                continue
            elif time_strings[index] == '-':
                continue
            elif time_strings[index+1] != '-':
                string_index = index
                break
        edu_time.append(time_strings[:string_index+1])
    # 또 다른 함수로 빼기 #### RE
    for time in edu_time:
        detailed_edu_time = DETAILED_EDU_TIME[time]
        temp_list.append(detailed_edu_time)
    df_scripts['edu_time'] = temp_list  
 


# print(df_scripts)

def has_only_one_instructor(instructor_list):
    instructor_set = set(instructor_list)
    if len(instructor_set) == 1:
        return True
    else:
        return False

def append_detailed_date_time_instructor_in_df_scripts():
    temp_list = []
    for index, row in df_scripts.iterrows():
        detailed_date_time_instructor = []
        full_date = row['full_dates']
        edu_time = row['edu_time']
        instructors = row['instructors']
        for index in range(len(full_date)):
            if has_only_one_instructor(instructors):
                detailed_date_time_instructor_scripts = full_date[index] + " " + edu_time[index]
            else:
                detailed_date_time_instructor_scripts = full_date[index] + " " + edu_time[index] + " : " + instructors[index] +" 지도위원님"
            detailed_date_time_instructor.append(detailed_date_time_instructor_scripts)
        temp_list.append(detailed_date_time_instructor)
    df_scripts['detailed_date_time_instructor'] = temp_list



# print(df_scripts)

def append_full_mail_script_in_df_scripts():
    temp_list = []
    for index, row in df_scripts.iterrows():
        temp_string = ''
        temp_string += '1. 과정명 : ' + row['edu_name'] + '\n'
        temp_string += '2. 일정 : ' + row['full_dates'][0] + ' ~ ' + row['full_dates'][1] + '\n'
        temp_string += '3. 강의시간: \n'
        for detailed_date_time in row['detailed_date_time_instructor']:
            temp_string += '  ' + detailed_date_time + '\n'
        temp_string += '4. 강의장소 : ' + row['edu_room'] + '\n'
        temp_string += '5. 준비사항 : 교안 파일 및 참고 자료는 USB에 담아서 준비헤주시기 바랍니다.\n'
        temp_string += '6. 수강생명단 : 첨부파일 참조\n'
        temp_list.append(temp_string)
    df_scripts['full_mail_scripts'] = temp_list



if __name__ == "__main__":
    append_upcoming_edu_rows_in_edu_rows(week_start_row, week_end_row)
    selected_edu_sections_row = get_selected_edu_sections_row()

    df_scripts = pd.DataFrame()

    append_edu_row_in_df_scripts()
    append_edu_name_in_df_scripts()
    append_edu_room_in_df_scripts()
    append_start_end_date_in_df_scripts()
    append_instructors_in_df_scripts()
    append_full_edu_date_in_df_scripts()
    append_edu_time_in_df_scripts()
    append_detailed_date_time_instructor_in_df_scripts()
    append_full_mail_script_in_df_scripts()

    for index, row in df_scripts.iterrows():
        print(row['full_mail_scripts'])

