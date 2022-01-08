"""
Author : Jeonghoon Lee
Last modification: 2022.1.1.
whitefang79@naver.com
https://github.com/Benjijeffdad/xlsx_sorter
"""

import sys
import os
import pyexcel as px
import shutil
import time


print("Process Start.")
start_time = time.time()

# 예시로 사용할 템플릿 엑셀 파일 이름
template = sys.argv[1]

# 분석 대상이 되는 엑셀 파일이 들어 있는 폴더
directory = sys.argv[2]

# 프로그램의 작동 모드 (delete, report, separate)
mode = sys.argv[3].lower()
# 받은 변수 소문자로 만들기
# mode = sys.argv[3].lower()

# 분석 대상이 되는 엑셀 파일의 목록
file_list = os.listdir(directory)

# 템플릿 파일을 읽어와 헤더를 분리합니다.
HEADER = px.get_array(file_name=template)[0]

# 보고 모드일 때에는 보고서 파일을 생성합니다.
if mode == "report":
    # report.txt 파일을 새로이 생성합니다.
    report = open("report.txt", "w")

# 분리 모드일 때는 분리된 파일을 격리저장할 폴더를 만듭니다.
elif mode == "separate":
    # wrong_files라는 폴더를 새로이 생성합니다.
    os.mkdir("wrong_files")

elif mode != "delete":
    print("Wrong mode! (delete/report/separate) ")
    exit(1)

# for문을 퐐용해 파일을 하나씩 불러 옵니다.
for filename in file_list:
    #엑셀 파일을 읽어옵니다.
    file = px.get_array(file_name = directory + "/" + filename)
    #헤더를 분리합니다.
    header = file[0]
    # 헤더가 템플릿과 일치하는지 분석합니다.
    if header == HEADER:
        # 헤더가 템플릿과 일치하는 올바른 파일이라면
        # 아무것도 하지 않고 넘어갑니다.
        continue

    # 삭제 모드인 경우
    if mode == "delete":
        # 파일을 삭제합니다.
        os.remove(directory + "/" + filename)
    # 보고 모드인 경우
    elif mode == "report":
        #report 파일에 이름을 작성합니다.
        report.write(filename + "\n")
    # 분리 모드인 경우
    elif mode == "separate":
        # 파일을 이동시킵니다.
        shutil.move(directory + "/" + filename, "wrong_files/" + filename)

end_time = time.time()
print("Process Done")
print("The Job took " + str(end_time-start_time) + " seconds.")

"""
if mode == "delete":
    for filename in file_list:
        file = px.get_array(file_name=directory + "/" + filename)
        header = file[0]
        if header == HEADER:
            continue
        else:
            os.remove(directory + "/" + filename)
elif mode == "report":
    report = open("report.txt", "w")
    for filename in file_list:
        file = px.get_array(file_name=directory + "/" + filename)
        header = file[0]
        if header == HEADER:
            continue
        else:
            report.write(filename + "\n")
    report.close()

elif mode == "separate":
    os.mkdir("wrong_files")
    for filename in file_list:
        file = px.get_array(file_name=directory + "/" + filename)
        header = file[0]
        if header == HEADER:
            continue
        else:
            shutil.move(directory + "/" +filename, "wrong_files/" + filename)
    
else:
    print("Wrong mode! (delete/report/separate) ")
"""
