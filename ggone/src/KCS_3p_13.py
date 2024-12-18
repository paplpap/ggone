from pyhwpx import Hwp
hwp = Hwp()

def K313(hwp):

    hwp.MoveDocBegin()
    src = [
        '1. 일반사항',
        '1.1 적용범위',
        '1.2 참고기준',
        '1.3 용어의정의',
        '2. 자재',
        '3. 시공'
    ]
    i_no_space = [
        '1. 일반 사항',
        '1.1 적용 범위',
        '1.2 참고 기준',
        '1.3 용어의 정의',
        '2. 자재',
        '3. 시공'
    ]

    def prob_text_in():
        for i in range(len(src)):
            if not hwp.find(src=src[i]):
                if not hwp.find(src=i_no_space[i]):
                    hwp.insert_memo(f'\n{src[i]}이 적혀있지 안습니다. 해당하는 조항은 반드시 적혀있어야 합니다.')
            hwp.markpen_on_selection() # 시연용으로 사용, 실 제출시에는 삭제 요함

    prob_text_in() # 목차까지 탐색
    prob_text_in() # 본문 탐색



if __name__ == '__main__':
    hwp.open(r'../data/KCS 10 30 05 건설공사 측량(21.12).hwp')
    K313(hwp)
