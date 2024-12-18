from pyhwpx import Hwp

hwp = Hwp()
import re


def K416_1(hwp):
    '''
    KDS에서만 사용하는 함수이다.
    코드 작성 시 인용되는 모든 법규, 기준 및 표준은 개별 코드 “1. 일반사항”의 “1.3 참고 기준”에 기술한다.(기존 본문에 나왔던 것만 쓰기,
    본문에 등장하지 않았는데 나오는거 안됨)

    해당하는 문제에 대하여 루프를 2번 진행한다.
     1. 전체 문서(1.3 제외)를 돌면서 re 정규식에 걸리는 개별 코드를 리스트에 수집
     2. 이후 1.3에 접근하여 그 칸에 리스트의 코드가 다 있는지 확인하기
     3. 1.3에 메모를 남긴다. ( 본문에 있는데 1.3 에는 없는 코드를 )

    :param hwp:
    :return:
    '''
    hwp.MoveDocBegin()

    document_text = hwp.get_text_file()

    # 정규식 패턴 (KDS 코드)
    kds_pattern = r"KDS \d{2} \d{2} \d{2}"

    # 1.3.2 섹션 추출
    section_132_pattern = r"1\.3\.2[\s\S]*?(?=\n1\.\d|\Z)"  # 1.3.2 내용만 추출
    section_132 = re.search(section_132_pattern, document_text)

    if section_132:
        kds_in_132 = set(re.findall(kds_pattern, section_132.group()))  # 1.3.2에서 KDS 코드 추출
    else:
        kds_in_132 = set()

    # 1.3.2를 제외한 나머지 내용 추출
    remaining_text = re.sub(section_132_pattern, "", document_text)  # 1.3.2 제거
    kds_in_remaining = set(re.findall(kds_pattern, remaining_text))  # 나머지에서 KDS 코드 추출

    # 차집합 계산
    difference_set = kds_in_remaining - kds_in_132

    if difference_set != None:
        hwp.find(src='1.3.2 관련 기준')
        hwp.insert_memo(f"기준 코드 {difference_set}를 본문에서 사용하였지만, 1.3.2에 기입하지 않았습니다.")


def K416_2(hwp):
    hwp.MoveDocBegin()
    hwp.init_scan()

    hwp.find(src='1.3.2 관련 기준')
    hwp.MoveLineBegin()


    while True:
        state, text = hwp.get_text()
        print(text)
        if '1.3.2 관련 기준' in text:
            while True:
                state, text = hwp.get_text()
                print(text)
                if text == '':
                    pass
                elif text[0] != "∙":
                    hwp.move_pos(201)
                    hwp.insert_memo('1.3.2 관련 기준의 글머리표는 "∙" 으로 시작해야 합니다.')
                elif text[0] == 1:
                    break

        if state <= 1:
            break

    hwp.release_scan()



if __name__ == '__main__':
    hwp.open(r'../data/KDS 11 50 25 기초 내진 설계기준(21.05).hwp')
    #K416_1(hwp)
    K416_2(hwp)