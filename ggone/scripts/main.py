from pyhwpx import Hwp
from KDScode import KDScode
from KCS_OCScode import KCS_OCScode

import os

def read_files_in_data_directory():
    # data 디렉토리 경로 설정
    data_dir = "../data"

    # 디렉토리가 존재하는지 확인
    if not os.path.exists(data_dir):
        print(f"디렉토리 {data_dir}가 존재하지 않습니다.")
        return

    # data 디렉토리의 파일 목록 가져오기
    files = os.listdir(data_dir)

    # 파일 읽기
    for file_name in files:
        file_path = os.path.join(data_dir, file_name)
        if 'KCS' in file_path or 'OCS' in file_path:
            KCS_OCScode(file_path)
        elif 'KDS' in file_path:
            KDScode(file_path)





#
# hwp = Hwp()
# hwp.open(file_name)
#
# #markpen(hwp, '콘크리트')
# result = strfinder(hwp)
# logging.info(f"strfinder result: {result}")


if __name__ == '__main__':
    read_files_in_data_directory()
'''
pset = hwp.get_file_info(file_name)
print(pset.Item("Format"))
print(pset.Item("VersionStr"))
print(hex(pset.Item("VersionNum")))
print(pset.Item("Encrypted"))
print(hwp.get_font_list())
''' #나중에 예외처리용

