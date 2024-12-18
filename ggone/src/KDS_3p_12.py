from pyhwpx import Hwp
hwp = Hwp()

def K312(hwp):

    hwp.MoveDocBegin()
    src = [
        '1. 일반사항',
        '1.1 목적',
        '1.2 적용 범위',
        '1.3 참고 기준',
        '1.3.1 관련 법규',
        '1.3.2 관련 기준',
        '1.4 용어의 정의',
        '1.5 기호의 정의',
        '2. 조사 및 계획',
        '3. 재료',
        '4. 설계'
    ]
    i_no_space = [
        '1. 일반 사항',
        '1.1 목적',
        '1.2 적용범위',
        '1.3 참고기준',
        '1.3.1 관련법규',
        '1.3.2 관련기준',
        '1.4 용어의정의',
        '1.5 기호의정의',
        '2. 조사및계획',
        '3. 재료',
        '4. 설계'
    ]

    def prob_text_in():
        for i in range(len(src)):
            if not hwp.find(src=src[i]):
                if not hwp.find(src=i_no_space[i]):
                    hwp.insert_text(f'\r\n{src[i]} 의 내용이 없으면 "내용없음"을 적어주셔야 합니다.')

            #hwp.markpen_on_selection() # 시연용으로 사용, 실 제출시에는 삭제 요함

    def prob_text_in_two():
        hwp.MoveDocBegin()
        hwp.find(src='1. 일반사항')
        for i in range(len(src)):
            if not hwp.find(src=src[i]):
                if not hwp.find(src=i_no_space[i]):
                    hwp.insert_text(f'\r\n{src[i]} 의 내용이 없으면 "내용없음"을 적어주셔야 합니다.')
                    break
            #hwp.markpen_on_selection() # 시연용으로 사용, 실 제출시에는 삭제 요함

    prob_text_in() # 목차까지 탐색
    prob_text_in_two() # 본문 탐색

if __name__ == '__main__':
    hwp.open(r'../data/KDS 11 30 05 연약지반 설계일반(21.12).hwp')
    K312(hwp)
