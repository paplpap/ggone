from pyhwpx import Hwp

hwp = Hwp()
import re


def K418(hwp):
    '''
    4장 18조를 따라서 복수의 인용 표기가 필요한 경우, 다음 같이 쉼표를 이용하여 연이어 기술한다

    KDS aa bb cc 및 KDS dd ee ff를 인용하고자 할 경우
    ☞ KDS aa bb cc, KDS dd ee ff를 따른다

    :param hwp:
    :return:
    '''

    hwp.MoveDocBegin()


    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        def transform_kds_sentence(sentence):
            kds_codes = re.findall(r"KDS(?:\s\d+)+", sentence)

            # 만약 코드가 2개 이상이라면 변환 작업 수행
            if len(kds_codes) >= 2:
                # 변환된 문장 생성
                hwp.insert_memo(f'4장 18조를 따라서 복수의 인용 표기가 필요한 경우, 다음 같이 쉼표를 이용하여 연이어 기술한다')
                hwp.MoveSelNextParaBegin()

        transform_kds_sentence(text)

        hwp.Cancel()

def K418_1(hwp):
    '''
    4장 18조를 따라서 복수의 인용 표기가 필요한 경우, 다음 같이 쉼표를 이용하여 연이어 기술한다

    KDS aa bb cc의 3.1, 3.2를 인용하고자 할 경우
    ☞ KDS aa cc bb (3.1, 3.2)를 따른다.(코드 번호 (항 번호)) 꼭 지키기

    이거 정규식 수정해서 괄호 되는지 확인하기
    :param hwp:
    :return:
    '''

    hwp.MoveDocBegin()


    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        def transform_kds_sentence(sentence):
            kds_codes = re.findall(r"(KDS(?:\s\d+)+).*?(\d+\.\d+).*?(?:.*?(\d+\.\d+))?" , sentence)

            # 만약 코드가 2개 이상이라면 변환 작업 수행
            if len(kds_codes) >= 1:
                # 변환된 문장 생성
                hwp.insert_memo(f'4장 18조를 따라서 복수의 인용 표기가 필요한 경우, 다음 같이 괄호를 이용하여 연이어 기술한다')
                hwp.MoveSelNextParaBegin()

        transform_kds_sentence(text)

        hwp.Cancel()

def K418_132_1(hwp):
    '''
    관련 법규에 대하여 1.3.2 내부에서 확인하는 코드
    ☞ 법령명
    예) ·건설기술진흥법


    :param hwp:
    :return:
    '''

    hwp.MoveDocBegin()

    hwp.find(src='1.3.2 관련 기준')
    hwp.MoveSelNextParaBegin()
    hwp.Cancel()

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        if '1.4 용어의 정의' in text:
            break
        if text == '':
            continue
        elif text[0] != "∙":
            if text[-1] == '법':
                hwp.insert_memo(f'제8조①항에 따라 법규, 건설기준, 표준 등을 인용할 경우의 표기방법은 다음과 같습니다. ☞ 법령명 , 예) ·{text}')
                hwp.MoveSelNextParaBegin()


        hwp.Cancel()

def K418_132_2(hwp):
    '''
    관련 기준을 1.3.2에서 찾아서 확인하는 코드

    - 건설기준의 인용은 표 14.과 같이 제15조③항에 따라 다음과 같이 작성한다.
    - “1.3.2 관련기준”은 건설기준 코드의 경우 “코드+코드번호+기준명”을으로 표기하고,
    건설기준 코드가 아닌 관련 기준일 경우“기준명+(부처명)”으로 표기한다.


    :param hwp:
    :return:
    '''
    code182 = []

    hwp.MoveDocBegin()

    hwp.find(src='1.3.2 관련 기준')
    hwp.MoveSelNextParaBegin()
    hwp.Cancel()

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        if '1.4 용어의 정의' in text:
            break
        if text == '':
            continue
        elif text[0] == "∙":
            code182.append(text)

        hwp.Cancel()


    # 정규식 정의
    construction_code_pattern = r'^∙KDS \d{2} \d{2} \d{2} [^\n]+$'
    related_code_pattern = r'^∙?[^\n]+\s?\([^\n()]+\)$'

    hwp.MoveDocBegin()
    # 분류 로직
    for code in code182:
        if re.match(construction_code_pattern, code):
            continue
        elif re.match(related_code_pattern, code):
            continue
        else:
            if code[-1] != '법':
                hwp.find(src=f'{code}')
                hwp.insert_memo('“1.3.2 관련기준”에 적합하지 않은 기준입니다. 건설기준 코드의 경우 “코드+코드번호+기준명”을으로 표기하고, 건설기준 코드가 아닌 관련 기준일 경우“기준명+(부처명)”으로 표기해 주십시오.')

def K418_132_3(hwp):
    '''
    관련 표준에 대한 내용 // 나라표준인증에 대하여 ex(KS A 0001 표준의 서식과 작성방법) 이 1.3.2에 정확히 있는지에 대해서 확인

    :param hwp:
    :return:
    '''

    code182 = []

    hwp.MoveDocBegin()

    hwp.find(src='1.3.2 관련 기준')
    hwp.MoveSelNextParaBegin()
    hwp.Cancel()

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        if '1.4 용어의 정의' in text:
            break
        if text == '':
            continue

        code182.append(text)

        hwp.Cancel()

    # 정규식 정의
    korea_pattern = r"KS\s[A-Z]\s\d{4}(?:\s[\w가-힣―\-·\s]+)?"

    hwp.MoveDocBegin()
    # 분류 로직
    for code in code182:
        if re.match(korea_pattern, code):
            hwp.find(src=f'{code}')
            hwp.insert_memo('관련 표준은 {∙+표준+표준번호+표준명} 으로 표기해 주십시오.')



def K418_132_1_0(hwp):
    '''
    - 관련 법규
    ☞ 법령명+인용부
    예) 건설기술진흥법 제44조①항
    인용부(항목번호)는 써도되고 안써도됨
    인용부(항목번호) 쓰는 경우 괄호안에 항목번호를 써야함. 괄호 필수

    이 부분은 1.3.2를 검사하고 진행되는 방식으로 진행하였습니다. 이 방식이 아니라 그냥 본문(1.3.2 제외)에서 찾는 방식은 예외가 너무 많아짐

    :param hwp:
    :return:
    '''

    code182 = []

    hwp.MoveDocBegin()

    hwp.find(src='1.3.2 관련 기준')
    hwp.MoveSelNextParaBegin()
    hwp.Cancel()

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        if '1.4 용어의 정의' in text:
            break
        if text == '':
            continue
        elif text[-1] == "법":
            if (text[:3] != "∙KS") and (text[:4] != "∙KDS") and (text[:4] != "∙KCS") and (text[:2] != "KS") and (text[:3] != "KDS") and (text[:3] != "KCS"):
                code182.append(text)

        hwp.Cancel()


    for i in range(len(code182)):
        if code182[i][0] == "∙":
            code182[i] = code182[i][1:]

    code182 = list(set(code182))

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        for law in code182:
            if law in text:
                hwp.find(src=f'{text}')
                hwp.insert_memo('1.3.2가 아닌 부분의 관련 법규는 예) 건설기술진흥법 제44조①항 와 같이 인용부(항목번호)로 인용부(항목번호) 쓰는 경우 괄호안에 항목번호를 써야합니다.')
                hwp.MoveSelNextParaBegin()
                break

        hwp.Cancel()

def K418_132_2_0(hwp):
    '''
    -관련 기준
    ☞ 타 기준 인용 시: 코드+코드번호+(인용부)
    예) KDS 10 10 00 (1.1.1(1)①가.(가)㉮) 또는 KDS 10 10 00 (그림/표 1.1-1)
    인용부(항목번호)는 써도되고 안써도됨
    인용부(항목번호) 쓰는 경우 괄호안에 항목번호를 써야함. 괄호 필수


    ☞ 동일 기준 내 인용 시: 인용부
    예) 1.1.1(1)①가(가)㉮ 또는 그림/표 1.1-1
    ☞ 코드가 아닌 기준 인용 시 : 기준명+(인용부)
    예) 건축물의 에너지절약설계기준(제1조①항1호)   -> 이런건 특징점이 없어서 못한다. 1.3.2에 없다면 이걸 풀게 하려면 관련 기준을 모두 모아서 줘야 한다.

    ※혼란을 방지하기 위하여 항목수준 1 까지는 인용 시 제목을 기입할 수 있다.
    예) ∼와 관련한 사항은 이 기준의 1. 일반사항을 따른다.
    ∼와 관련한 사항은 KDS 10 10 00(1. 일반사항을)따른다.


    :param hwp:
    :return:
    '''
    code182 = []

    hwp.MoveDocBegin()

    hwp.find(src='1.3.2 관련 기준')
    hwp.MoveSelNextParaBegin()
    hwp.Cancel()

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        if '1.4 용어의 정의' in text:
            break
        if text == '':
            continue
        elif text[0] == "∙":
            code182.append(text)

        hwp.Cancel()

def K418_132_3_0(hwp):
    '''
    -￭ 관련 표준
    ☞ 표준+표준번호
    예) KS A 0001


    :param hwp:
    :return:
    '''
    code182 = []

    hwp.MoveDocBegin()

    hwp.find(src='1.3.2 관련 기준')
    hwp.MoveSelNextParaBegin()
    hwp.Cancel()

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        if '1.4 용어의 정의' in text:
            break
        if text == '':
            continue

        code182.append(text)

        hwp.Cancel()

    # 정규식 정의
    korea_pattern = r"KS\s[A-Z]\s\d{4}(?:\s[\w가-힣―\-·\s]+)?"

    code182out = []

    # 분류 로직
    for code in code182:
        if re.match(korea_pattern, code):
            code182out.append(code)

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        for i in code182out:
            if i in text:
                hwp.find(src=f'{text}')
                hwp.insert_memo('1.3.2가 아닌 부분의 관련 국가 표준은 표준+표준번호 까지만 쓴다. 예) KS A 0001')
                hwp.MoveSelNextParaBegin()
                break

        hwp.Cancel()


# if __name__ == '__main__':
#     hwp.open(r'../data/KDS 11 50 25 기초 내진 설계기준(21.05).hwp')
    #K418_1(hwp)
    #K418_2(hwp)
    #K418_3(hwp)


    #K418_132_1(hwp)
    #K418_132_2(hwp)
    #K418_132_3(hwp)
    #K418_132_1_0(hwp)
    #K418_132_2_0(hwp)
    #K418_132_3_0(hwp)