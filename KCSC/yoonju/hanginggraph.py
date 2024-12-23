import sys
import os
from pyhwpx import Hwp
import pyperclip
import re

''' 뒤에서부터 읽는 코드
복잡해서 hanginggraph2.py로 다시'''

one_titles_KCS = [
        "1. 일반사항",
        "1.1 적용 범위",
        "1.2 참고 기준",
        "1.3 용어의 정의"
        "2. 자재",
        "3. 시공"
]

titles_KDS = [
        "1.3.2 관련 기준",
        "1.4 용어의 정의",
]

all_symbols = ['∙',  # 검사할 기호 리스트 ∙
        '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨',
        '가', '나', '다', '라', '마', '바', '사', '아', '자', '차', '카', '타', '파', '하',
        '(가)', '(나)', '(다)', '(라)', '(마)', '(바)', '(사)', '(아)', '(자)', '(차)', '(카)', '(타)', '(파)', '(하)',
        '㉮', '㉯', '㉰', '㉱', '㉲', '㉳', '㉴', '㉵', '㉶', '㉷', '㉸', '㉹', '㉺', '㉻',
]

def starts_with_number_pattern(word):
    '''
    숫자로 시작하는 테이블인지 검사 (상위 목차인지)
    :param sentence: 문장
    :return: bool
    '''
    pattern = r'^(\d+(\.\d+){0,3}|\(\d+\)|\d+\.)$'
    return bool(re.match(pattern, word))

def check_symbol(text, symbols):
    '''
    첫 글자가 symbol인지
    :param line: 문장
    :param symbols: 문장 앞에 붙어도 되는 symbols
    :return: bool
    '''
    word1 = text.split()[0]
    word2 = text[0]

    if word1 in symbols or word2 in symbols or starts_with_number_pattern(word1) or starts_with_number_pattern(word2):
        return True
    else:
        return False

def move_block_first(hwp):
    cnt = 0
    while hwp.MoveSelPrevParaBegin():
        cnt += 1
        text = hwp.get_selected_text()
        hwp.Cancel()
        # print(f"함수에서 text: {text}")

        if text=='': # 1.3.2, 1.4인지 확인
            hwp.MoveSelNextParaBegin()
            hwp.Cancel()
            hwp.MoveSelNextParaBegin()
            Nexttext = hwp.get_selected_text()
            # print(f"함수에서 Nexttext: {Nexttext}")
            hwp.Cancel()

            for _ in range(cnt-1):
                hwp.MoveSelNextParaBegin()
                hwp.Cancel()
            hwp.MoveSelPrevParaBegin()
            finaltext = hwp.get_selected_text()
            # print(f"함수에서 finaltext: {finaltext}")
            if Nexttext in titles_KDS:
                return True
            else:
                return False

        hwp.Cancel()

def K417_2(hwp):
    global titles_KDS

    hwp.MoveDocBegin() # 집필위원 직전 문단으로 이동
    hwp.find(src='집필위원')
    hwp.MoveLineUp() #
    hwp.MoveSelPrevParaBegin()
    hwp.Cancel() 

    while hwp.MoveSelPrevParaBegin():
        text = hwp.get_selected_text().strip()

        if text == '1. 일반사항\r\n':
            break
        elif text.replace('\r', '').replace('\n', '').replace(' ','')=='':
            continue

        elif text[0]=='표':
            print(f"표 시작: {text}")
            while hwp.MoveSelPrevParaBegin():
                text = hwp.get_selected_text().strip()
                print(f"표 중간: {list(text)}")
                # print(f'text: {text}')
                if text=='':
                    break
                hwp.Cancel()
            print(f"표 이후: {text}")

        elif check_symbol(text, all_symbols): # 항목수준으로 시작하면
            pass
        else: # 틀린 기호/ 기호 없이 시작함
            # print(f"하이라이팅 문장: {text}")
            hwp.markpen_on_selection()
            hwp.insert_memo('이 문장에는 행잉패러그래프가 지켜지지 않았습니다.')

        hwp.Cancel()

    hwp.MoveDocBegin()  # 집필위원 직전 문단으로 이동
    hwp.find(src='1. 일반사항\r\n')
    hwp.MoveSelNextParaBegin()
    hwp.Cancel()

if __name__ == '__main__':
    hwp = Hwp()
    file_path = r'../data/KDS118025돌쌓기옹벽.hwp'
    hwp.open(file_path, format="HWP")
    K417_2(hwp)

