from pyhwpx import Hwp

hwp = Hwp()
import re


def K420(hwp):
    '''
    단위에 대한 리스트를 준비하고, 해당하는 리스트에 단위가 있고, 다음 문자가 spacebar이면 통과 아니면 메모남김
    °(도), “(분), ‘(초)는 붙여 써도 되는 글자 이므로 unit리스트에 없음

    :param hwp:
    :return:
    '''

    hwp.MoveDocBegin()

    unit = ['＄', '％', '￦', 'Ｆ', 'Å', '￠', '￡', '￥', '¤', '℉', '‰',
            '?', '㎕', '㎖', '㎗', 'ℓ', '㎘', '㏄', '㎣', '㎤', '㎥', '㎦', '㎙', '㎚', '㎛', '㎜',
            '㎝', '㎞', '㎟', '㎠', '㎡', '㎙', '㏊', '㎍', '㎎', '㎏', '㏏', '㎈', '㎉', '㏈', '㎧',
            '㎨', '㎰', '㎱', '㎲', '㎳', '㎴', '㎵', '㎶', '㎷', '㎸', '㎹', '㎀', '㎁', '㎂', '㎃',
            '㎄', '㎺', '㎻', '㎼', '㎽', '㎾', '㎿', '㎐', '㎑', '㎒', '㎓', '㎔', 'Ω', '㏀', '㏁',
            '㎊', '㎋', '㎌', '㏖', '㏅', '㎭', '㎮', '㎯', '㏛', '㎩', '㎪', '㎫', '㎬', '㏝', '㏐',
            '㏓', '㏃', '㏉', '㏜', '㏆']

    while hwp.MoveSelNextParaBegin():  # 다음 문장를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서

        for i in range(len(text)-1):
            if text[i] in unit:
                if text[i+1] != ' ':
                    hwp.insert_memo('°(도), “(분), ‘(초)를 제외한 단위는 숫자 다음 반칸(Alt+Spacebar)을 띄우고 기재한다.')
                    hwp.MoveSelNextParaBegin()
                    break

        hwp.Cancel()



if __name__ == '__main__':
    hwp.open(r'../data/KDS 11 50 25 기초 내진 설계기준(21.05).hwp')
    K420(hwp)


