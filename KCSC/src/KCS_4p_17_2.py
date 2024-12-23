import sys
import os
from pyhwpx import Hwp
import pyperclip
import re

''' 1. 일반사항 에서 왜 하이라이팅 되는지 확인'''
one_titles_KCS = [
        "1. 일반사항",
        "1.1 적용 범위",
        "1.2 참고 기준",
        "1.3 용어의 정의"
        "2. 자재",
        "3. 시공"
]

titles_KCS = [
        "1.2 참고 기준",
        "1.2.2 관련 기준",
        "1.3 용어의 정의",
]

all_symbols = [  # 검사할 기호 리스트 ∙
        '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨',
        '가', '나', '다', '라', '마', '바', '사', '아', '자', '차', '카', '타', '파', '하',
        '(가)', '(나)', '(다)', '(라)', '(마)', '(바)', '(사)', '(아)', '(자)', '(차)', '(카)', '(타)', '(파)', '(하)',
        '㉮', '㉯', '㉰', '㉱', '㉲', '㉳', '㉴', '㉵', '㉶', '㉷', '㉸', '㉹', '㉺', '㉻',
]

def starts_with_title_pattern(word):
    '''
    title 숫자로 시작하는 문장인지 검사
    :param sentence: 문장
    :return: bool
    '''
    pattern = r'^(\d+(\.\d+){1,3})|(\d\.)$'
    return bool(re.match(pattern, word))

def starts_with_content_pattern(word, symbols):
    '''
    항목 수준 기호 시작하는 문장인지 검사 (하위 목차인지)
    :param sentence: 문장
    :return: bool
    '''
    if word in symbols:  # (1)이후로 내용
        return True
    pattern = r'^(\(+\d+\))$'
    return bool(re.match(pattern, word))

def check_dot(text):
    '''
    ∙로 시작하는 단순 나열 항목인지
    :param line: 문장
    :param symbols: 문장 앞에 붙어도 되는 symbols
    :return: bool
    '''
    if text[0]=='∙' and text[1]==' ': # KCS213000.hwp 2024년 지침 기준
        return True
    return False

def check_symbol(text, symbols):
    '''
    첫 글자가 symbol인지
    :param line: 문장
    :param symbols: 문장 앞에 붙어도 되는 symbols
    :return: Title, Content, text로 type 분류
    '''
    word1 = text.split()[0]
    word2 = text[0]
    word3 = text[0:3]

    for word in [word1, word2, word3]:
        if starts_with_title_pattern(word): # 1.1.1.1까지 제목
            return 'Title'
    for word in [word1, word2, word3]:
        if starts_with_content_pattern(word, symbols):
            return 'Content'
    return 'text'

def after_table(hwp):
    """
    목차 직후의 첫 문장으로 이동
    :param hwp: 한글 문서
    :return:
    """
    hwp.MoveDocBegin()  # 집필위원 직전 문단으로 이동
    hwp.find(src='1. 일반사항')
    while hwp.MoveSelNextParaBegin():
        pass
    hwp.MoveLineDown()
    hwp.MoveLineDown()
    hwp.MoveLineBegin()
    hwp.Cancel()

def KCS_417_2(hwp):
    '''
    KCS에 대해 제 17조 행잉패러그래프 검사

    읽다가 titles_KDS을 만난 경우 flag_dot -> 점만 써야함
    one_titles_KDS을 만난 경우 flag -> 내용 없음 빼고 기호 써야함
    그냥 숫자 제목 형식을 만나면 -> flag, flag_dot 초기화

    flag_dot 상태 -> 점으로 시작 또는 내용 없음
    flag 상태 -> 숫자로 시작 또는 내용 없음
    둘 다 아님 -> 무조건 숫자로 시작

    :param hwp: 한글 문서
    :return:
    '''
    global titles_KCS

    table_list = []
    after_table(hwp)
    while hwp.MoveSelNextParaBegin():
        text = hwp.get_selected_text().strip()
        if '집필위원' in text:
            break
        if text[0:2]=='표 ':
            table_list.append(text)
            hwp.Cancel()
            while hwp.MoveSelNextParaBegin():
                text = hwp.get_selected_text().strip()
                if text!='':
                    table_list.append(text)
                else:
                    break
                hwp.Cancel()

        elif '표 ' in text and check_symbol(text, all_symbols)=='text':
            table_list.append(text)
            while hwp.MoveSelNextParaBegin():
                text = hwp.get_selected_text().strip()
                if text!='':
                    lst = text.split('\r\n\r\n')
                    lst = [line.strip() for line in lst if line]
                    table_list.extend(lst)
                else:
                    break
                hwp.Cancel()

        hwp.Cancel()

    after_table(hwp)
    flag_dot = False
    flag = False

    while hwp.MoveSelNextParaBegin():
        text = hwp.get_selected_text().strip()

        if '집필위원' in text:
            break
        if text.replace('\r', '').replace('\n', '').replace(' ','')=='':
            hwp.Cancel()
            continue

        type = check_symbol(text, all_symbols)
        # print(f"text: {type}, {text}")
        if type=='Title':
            flag = False
            flag_dot = False
        if text in titles_KCS: # 단순 나열 dot 사용 제외
            flag_dot = True
        elif text in one_titles_KCS:  # 내용 없음은 제외
            flag = True

        if flag_dot: # 단순 나열 시
            if type=='Title' or check_dot(text) or text=='내용 없음':
                pass
            else:

                hwp.markpen_on_selection()
                hwp.MovePrevParaBegin()
                hwp.insert_memo('제 17조 이 문장에는 행잉패러그래프가 지켜지지 않았습니다.')
                hwp.MoveNextParaBegin()

        elif flag: # one_titles에 대해 검사
            if type=='Title' or type=='Content' or text=='내용 없음':
                pass
            else:
                hwp.markpen_on_selection()
                hwp.MovePrevParaBegin()
                hwp.insert_memo('제 17조 이 문장에는 행잉패러그래프가 지켜지지 않았습니다.')
                hwp.MoveNextParaBegin()

        else:
            if type=='text': # 항목 수준을 지키지 않는 문장 중에
                if text in table_list: # 표에 포함되면 제외
                    hwp.Cancel()
                    continue
                hwp.markpen_on_selection()
                hwp.MovePrevParaBegin()
                hwp.insert_memo('3. 이 문장에는 행잉패러그래프가 지켜지지 않았습니다.')
                hwp.MoveNextParaBegin()

        hwp.Cancel()

if __name__ == '__main__':
    file_path = r'../data/KCS213000.hwp'
    hwp = Hwp()
    hwp.open(file_path, format="HWP")
    KCS_417_2(hwp)