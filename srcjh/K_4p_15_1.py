from pyhwpx import Hwp
hwp = Hwp()
import re

def K415_1(hwp):
    '''
    4장 15조 코드 작성 시 「, 」, “, ”, ‘, ’ 등 인용부호를 사용하지 않는다.
    해당하는 코드를 구현하였습니다.
    문서 전체에서 해당하는 기호를 탐색하게 하였습니다.

    :param hwp:
    :return:
    '''
    hwp.MoveDocBegin()

    # 인용부호 리스트
    quotes = ['「', '」', '“', '”', '‘', '’', '‚', '‛']

    while hwp.MoveSelNextWord():  # 다음 단어를 선택(다음단어가 없으면 break)

        text = hwp.get_selected_text()  # 문자열을 가져와서

        for i in quotes:
            if i in text:
                hwp.insert_memo('코드 작성 시 「, 」, “, ”, ‘, ’ 등 인용부호를 사용하지 않는다.')
                hwp.MoveSelNextWord()
                break
        hwp.Cancel()




if __name__ == '__main__':
    #hwp.open(r'../data/KCS수밀 콘크리트 24.05.hwp')
    #K415_1(hwp)
    hwp.open(r'../data/KDS_11_80_25_돌(블록)쌓기옹벽.hwp')
    K415_1(hwp)
