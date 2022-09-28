
# import os
# # import subprocess
#
# # # 상대경로를 통해서 배치파일 실행시키기
# # current_path = os.getcwd()
# # parent_path = os.path.abspath(os.path.join(current_path, os.pardir))
# # file_name = os.path.join(parent_path, r"MailControler\!!!Run_MDChecker.bat")
# # subprocess.call([file_name])
#
# # 상대경로를 통해서 배치파일 실행시키기
# # current_path = os.getcwd()
# # parent_path = os.path.abspath(os.path.join(current_path, os.pardir))
# # print(parent_path)
# # file_name = os.path.join(current_path, r"MailControler\!!!Run_MDChecker_ADMIN.bat")
# # 배치 파일 만들고, 해당 배치파일 실행시 작동하는 기능 구현 (수신자 최다희)
# # print(file_name)
#
# working_dir = r"D:\Project_Python\webMDChecker\MDChecker\MailControler"
# execute_file = r"!!!Run_MDChecker_ADMIN.bat"
#
# # com = r"D:\Project_Python\webMDChecker\venv\Scripts\python.exe D:\Project_Python\webMDChecker\MDChecker\MailControler\MDChecker_Controler.py ADMIN "
# com = r"D:\Project_Python\webMDChecker\venv\Scripts\python.exe D:\Project_Python\webMDChecker\MDChecker\MailControler\MDChecker_Controler.py ADMIN "
#
# # os.chdir(working_dir)
# # os.system(execute_file)
# os.system(com)


# import asyncio
#
#
# async def add(a, b):
#     print('add: {0} + {1}'.format(a, b))
#     await asyncio.sleep(1.0)  # 1초 대기. asyncio.sleep도 네이티브 코루틴
#     return a + b  # 두 수를 더한 결과 반환
#
#
# async def print_add(a, b):
#     result = await add(a, b)  # await로 다른 네이티브 코루틴 실행하고 반환값을 변수에 저장
#     print('print_add: {0} + {1} = {2}'.format(a, b, result))
#
#
# loop = asyncio.get_event_loop()  # 이벤트 루프를 얻음
# loop.run_until_complete(print_add(1, 2))  # print_add가 끝날 때까지 이벤트 루프를 실행
# loop.close()  # 이벤트 루프를 닫음


# def execute_MDChecker():
#     command = r"D:\Project_Python\webMDChecker\venv\Scripts\python.exe D:\Project_Python\webMDChecker\MDChecker\MailControler\MDChecker_Controler.py ADMIN "
#     import subprocess
#     subprocess.call(command, shell=True)
#     # os.system(command)