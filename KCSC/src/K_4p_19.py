import sys
import os
from pyhwpx import Hwp
import pyperclip
import re

def extract_words(text):
    """
    문장에서 약어를 추출함
    :param text: 문서에서 읽은 한 문장
    :return: 약어 리스트
    """
    # 1단계: 괄호 안의 내용을 추출
    pattern = r'\(([A-Za-z,\s]+)\)'
    matches = re.findall(pattern, text)

    # 2단계: ','로 나누고 조건 확인
    final_results = []
    for match in matches:
        values = [item.strip() for item in match.split(',')]  # ','로 나누고 공백 제거
        # 조건: ','로 나눈 값이 2개 이하이고, 모두 대문자인 요소가 1개인 경우
        uppercase_values = [value for value in values if value.isupper()]
        if len(values) <= 2 and len(uppercase_values) == 1:
            final_results.extend(uppercase_values)  # 해당 대문자 값 추가

    return final_results


def extract_consecutive_english_words(text):
    """
    약어의 의미상 영단어를 추출
    :param text: 문서에서 읽은 한 문장
    :return: 연속된 영단어
    """
    match = re.search(r'([A-Za-z]+(?: [A-Za-z]+)*)$', text)
    return match.group(0)

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
    hwp.Cancel()

def K419(hwp):
    """
    제 4조 19항 검사
    1) 한 문장씩 읽으면 형식에 맞는 약어 추출
    2) 문서에서 약어 선언이 잘못된 경우 오류
    3) 선언한 약어와 다른 단어를 사용한 경우 오류
    :param hwp: 한글 문서
    :return:
    """
    after_table(hwp)
    word_list = []
    mismatch_list = []
    while hwp.MoveSelNextParaBegin(): # 문서에서 사용된 약어 찾기
        text = hwp.get_selected_text().strip()
        hwp.Cancel()
        words = extract_words(text)

        for word in words:
            text = text.split(word)[0].strip()
            try:
                word2 = extract_consecutive_english_words(text[:-1])
                word_list.append(word2)
            except:
                mismatch_list.append(word)
            # 한글 정식 명칭이 있는 경우 한글 단어 추가 불가능 이 부분 삭제 x
            # if text[-1] == ',':
            #     word2 = text[:-2].split('(')[-2].split()[-1]
            #     word_list.append(word2)

    title_list = {word.title():0 for word in word_list} # 추출한 약어 리스트
    mismatch_list.extend([word for word in word_list if word not in title_list])
    mismatch_list = list(set(mismatch_list))

    for mismatch in mismatch_list: # 약어 선언이 잘못된 경우
        after_table(hwp)
        hwp.find(src=mismatch)
        hwp.markpen_on_selection()
        hwp.insert_memo('제 19조 약어 선언이 잘못되었습니다.')
        hwp.Cancel()

    after_table(hwp)
    while hwp.MoveSelNextParaBegin(): # 약어 선언을 했음에도 잘못 사용한 경우
        text = hwp.get_selected_text().strip()
        if '집필위원' in text:
            break

        for word, cnt in title_list.items():
            if word.lower() in text.lower() or word in text:
                title_list[word] += 1
                if title_list[word] == 1:
                    continue
                hwp.Cancel()
                hwp.MoveSelPrevParaBegin()
                hwp.markpen_on_selection()
                hwp.insert_memo('제 19조 위에서 선언한 약어를 사용해야 합니다.')
                text = hwp.get_selected_text().strip()
                hwp.Cancel()
                hwp.MoveSelNextParaBegin()

            hwp.Cancel()
        hwp.Cancel()
    hwp.Cancel()

# if __name__ == '__main__':
#     hwp = Hwp()
#     file_path = r'../data/KDS118025돌쌓기옹벽.hwp'
#     hwp.open(file_path, format="HWP")
#     K419(hwp)


