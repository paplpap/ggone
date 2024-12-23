import sys
import os
from pyhwpx import Hwp
import pyperclip
import re

def after_table(hwp):
    hwp.MoveDocBegin()  # 집필위원 직전 문단으로 이동
    hwp.find(src='1. 일반사항')
    while hwp.MoveSelNextParaBegin():
        pass
    hwp.MoveLineDown()
    hwp.MoveLineDown()
    hwp.Cancel()

def highlighting(hwp, blocks, reasons):
    global titles_KDS
    after_table(hwp)
    check = [False] * len(blocks)

    while hwp.MoveSelNextParaBegin():
        text = hwp.get_selected_text().strip()
        # print(f"text: {text}")
        if '집필위원' in text:
            break

        if text.replace('\r', '').replace('\n', '').replace(' ','')=='':
            hwp.Cancel()
            continue

        if text in blocks: # block을 찾아서 memo 추가하기
            index = blocks.index(text)
            check[index] = True
            reason = reasons[index]

            hwp.markpen_on_selection()
            hwp.MoveSelPrevParaBegin()
            hwp.insert_memo(reason)
            hwp.MoveSelNextParaBegin()
            hwp.Cancel()

        hwp.Cancel()

    hwp.MoveDocBegin()  # 집필위원 직전 문단으로 이동
    hwp.find(src='목  차')

    hwp.markpen_on_selection()

    for inx in range(len(blocks)):
        if not check[inx]:
            reason = reasons[inx]
            hwp.insert_memo(reason)
            hwp.Cancel()

# if __name__ == '__main__':
#     file_path = r'../data/KDS101000설계기준총칙.hwp'
#     hwp = Hwp()
#     hwp.open(file_path)
#     incorrect_titles = ['1.2 적용범위', '1.3 참고 기준', '1.5 기호의 정의']
#     explanation = ['<1.2 적용 범위>가 있어야하지만 빠졌습니다.', '<1.3 참고 기준>가 있어야하지만 빠졌습니다.', '<1.5 기호의 정의>가 있어야하지만 빠졌습 니다.']
#
#     highlighting(hwp, incorrect_titles, explanation)