import sys
import os
from pyhwpx import Hwp
import pyperclip
import re

all_symbols = [  # 검사할 기호 리스트 ∙
        '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨',
        '가', '나', '다', '라', '마', '바', '사', '아', '자', '차', '카', '타', '파', '하',
        '(가)', '(나)', '(다)', '(라)', '(마)', '(바)', '(사)', '(아)', '(자)', '(차)', '(카)', '(타)', '(파)', '(하)',
        '㉮', '㉯', '㉰', '㉱', '㉲', '㉳', '㉴', '㉵', '㉶', '㉷', '㉸', '㉹', '㉺', '㉻',
]

def starts_with_title_pattern(word):
    '''
    title 숫자로 시작하는 테이블인지 검사 (상위 목차인지)
    :param sentence: 문장
    :return: bool
    '''
    pattern = r'^(\d+(\.\d+){1,3})|(\d\.)$'
    return bool(re.match(pattern, word))

def starts_with_content_pattern(word, symbols):
    '''
    title 숫자로 시작하는 테이블인지 검사 (상위 목차인지)
    :param sentence: 문장
    :return: bool
    '''
    if word in symbols:  # (1)이후로 내용
        return True
    pattern = r'^(\(+\d+\))$'
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
    word3 = text[0:3]

    for word in [word1, word2, word3]:
        if starts_with_title_pattern(word): # 1.1.1.1까지 제목
            return 'Title'
    for word in [word1, word2, word3]:
        if starts_with_content_pattern(word, symbols):
            return 'Content'
    return False

def after_table(hwp):
    hwp.MoveDocBegin()  # 집필위원 직전 문단으로 이동
    hwp.find(src='1. 일반사항')
    while hwp.MoveSelNextParaBegin():
        pass
    hwp.MoveLineDown()
    hwp.MoveLineDown()
    hwp.Cancel()

def K417_1(hwp):
    global titles_KDS
    after_table(hwp)

    while hwp.MoveSelNextParaBegin():
        flag = False
        text = hwp.get_selected_text().strip()
        print(f"text: {text}")
        if '집필위원' in text:
            break

        if text.replace('\r', '').replace('\n', '').replace(' ','')=='':
            hwp.Cancel()
            continue

        type = check_symbol(text, all_symbols)
        if type == 'Title' and text[-1]=='.':
            flag = True
        elif type == 'Content' and text[-1]!='.':
            flag = True

        if flag:
            hwp.markpen_on_selection()
            hwp.MoveSelPrevParaBegin()
            hwp.insert_memo('제 17조 기술 형식이 잘못되었습니다.')
            hwp.MoveSelNextParaBegin()
            hwp.Cancel()

        hwp.Cancel()

# if __name__ == '__main__':
#     hwp = Hwp()
#     file_path = r'../data/KCS수밀콘크리트.hwp'
#     hwp.open(file_path, format="HWP")
#     K417_1(hwp)