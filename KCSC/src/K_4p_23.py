from pyhwpx import Hwp

hwp = Hwp()
import re


def K423_1(hwp):
    '''

    - 서술형 문장의 끝은 마침표 온점(.)을 사용하며, 느낌표(!)나 물음표(?)는 사용하지 않는다.
    => 서술형 문장은 그냥 나열과 어떤식으로 다른가? >> 컴퓨터에서 자동으로 그걸 못걸러낸다. >> !,? 만 찾아서 . 으로 바꾸는걸 먼저 만든다.
    - 기간, 거리 또는 범위 등을 나타내는 이음표는 물결표(～)를 쓰고, 줄표(―)나 붙임표(-)를 사용하지 않는다. (3~4m X, 3m~4m O)
    => 조항 번호는 - ,― 으로 나타내서 전체탐색으로는 잘못된 부분을 많이 잡는다.
    :param hwp:
    :return:
    '''

    hwp.MoveDocBegin()

    while hwp.MoveSelNextWord():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서


        if ('!' in text) or ('?' in text):
            hwp.insert_memo('서술형 문장의 끝은 마침표 온점(.)을 사용하며, 느낌표(!)나 물음표(?)는 사용하지 않는다.')
            hwp.MoveSelNextWord()

        #if ('―' in text) or ('-' in text):
        #    hwp.insert_memo('기간, 거리 또는 범위 등을 나타내는 이음표는 물결표(～)를 쓰고, 줄표(―)나 붙임표(-)를 사용하지 않는다.')
        #    hwp.MoveSelNextWord()

        hwp.Cancel()



if __name__ == '__main__':
    hwp.open(r'../data/KDS 11 50 25 기초 내진 설계기준(21.05).hwp')
    K423_1(hwp)

