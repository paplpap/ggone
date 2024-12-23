from pyhwpx import Hwp
from src.strfinder import strfinder
from src.KCS_3p_13 import K313
from src.K_4p_15_1 import K415_1
from src.K_4p_15_2 import K415_2 , K415_3

import re

def KCS_OCScode(file_path):
    '''
    KCS/OCS파일을 받아서 그에 맞는 처리를 진행하는 코드입니다.
    제13조(KCS, OCS코드 구성방법)
        - 표준시방서(KCS)와 전문시방서(OCS) 코드의 내용은 “1. 일반사항”, “2. 자재”, “3. 시공”의 차례로 작성한다.
        - 일반사항
            - 1.1 적용 범위
            - 1.2 참고 기준
            - 1.3 용어의 정의
    반드시 기술하며, 해당 항목이 없을 시 “내용 없음”으로 작성한다.


    :param file_path:
    :return hwp:
    '''


    hwp = Hwp()
    hwp.open(file_path)

    K313(hwp)
    K415_1(hwp)
    K415_2(hwp)
    K415_3(hwp)


if __name__ == '__main__':
    KCS_OCScode(r'../data/KCS수밀 콘크리트 24.05.hwp')




