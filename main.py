import os
import natsort
import pandas as pd
import openpyxl

# 시간 및 파일 경로 정보 입력
file_path = input("파일 경로를 입력해주세요. : ")
col_number = int(input("피부온 센서 갯수를 입력해주세요. : "))
time_interval = int(input("피부온 데이터 측정 간격을 입력해주세요. : "))
start_time = input("시작 시간을 입력해주세요. : ") + ":"
finish_time = input("끝난 시간을 입력해주세요. : ") + ":"
set_zero_time = input("기준점(0초)을 입력해주세요. : ") + ":"

# 전체 파일 리스트 불러오기
file_list = natsort.natsorted(os.listdir(file_path))

# Column 만들기
new_col = ['시간', '상대시간'] + [i+1 for i in range(col_number)]

# 새로운 빈 DataFrame 정의
new_tmp_df = pd.DataFrame()

# 현재 상태 정의
curr_state = 0

for tmp_file in file_list:
    # 각각의 파일 경로 및 파일 DataFrame으로 불러오기
    one_file_path = file_path + "\\" + tmp_file
    tmp_df = pd.read_csv(one_file_path, encoding="cp949", skiprows=19)

    # 시작시간 및 끝시간 Boolean Data 추가
    start_time_bool = []
    finish_time_bool = []
    for item in tmp_df['Date/Time']:
        start_bool = start_time in item
        finish_bool = finish_time in item
        start_time_bool.append(start_bool)
        finish_time_bool.append(finish_bool)
    
    # 시작시간 및 끝시간 Index 설정
    start_idx = start_time_bool.index(True)
    finish_idx = finish_time_bool.index(True) + int(60 / time_interval) - 1

    # 첫번째 파일에 대하여
    # 시간, 상대시간, 새로운 데이터 Column 추가
    if tmp_file == file_list[0]:
        new_time_data = list(tmp_df.loc[start_idx:finish_idx, "Date/Time"])
        for idx, time_item in enumerate(new_time_data):
            if set_zero_time in time_item:
                set_zero_idx = idx
                break

        new_rel_time_data = [i for i in range(0, -10*len(new_time_data[0:set_zero_idx + 1]), -10)] + [i for i in range(10, 10*len(new_time_data[set_zero_idx:]), 10)]
        new_rel_time_data.sort()
        new_tmp_data = list(tmp_df.loc[start_idx:finish_idx, "Value"])
        while curr_state < 3:
            if curr_state == 0:
                new_tmp_df[new_col[curr_state]] = new_time_data
            elif curr_state == 1:
                new_tmp_df[new_col[curr_state]] = new_rel_time_data
            else:
                new_tmp_df[new_col[curr_state]] = new_tmp_data
            curr_state += 1
    # 나머지 파일에 대하여 새로운 데이터 Column 추가
    else:
        new_tmp_data = list(tmp_df.loc[start_idx:finish_idx, "Value"])
        new_tmp_df[new_col[curr_state]] = new_tmp_data
        curr_state += 1

# Excel 파일로 변환 후 저장
excel_title = file_path + "\\" + input("저정할 새로운 파일 이름을 입력해주세요. : ") + ".xlsx"
new_tmp_df.to_excel(excel_title, index=False)