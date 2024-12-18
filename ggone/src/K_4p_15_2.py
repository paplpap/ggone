from pyhwpx import Hwp

hwp = Hwp()
import re


def K415_2(hwp):
    '''
    4장 15조 외국어를 한글과 병기할 때 “한글(한자, 일본어, 영어 등)”의 순으로 기재하며, 대소문자로 구분되는 외국어의 경우 약어를 제외하고 소문자로 표기한다.

    해당하는 코드를 구현하였습니다.

    :param hwp:
    :return:
    '''
    hwp.MoveDocBegin()



    patterns = {
        "영어": r"^(?!.*@)[a-zA-Z ]+$"
    }

    def normalize_english_case(text):
        """
        텍스트에서 영어 문자를 검사하여 모든 문자가 대문자인 경우는 유지하고,
        그렇지 않으면 소문자로 변환하는 함수.
        변환된 경우에는 변환된 텍스트를 반환하고, 그렇지 않은 경우 False를 반환.

        """

        def convert(match):
            word = match.group()
            return word if word.isupper() else word.lower()

        # 영어 패턴에 대해 변환
        transformed_text = re.sub(patterns["영어"], convert, text)

        # 변환이 발생했는지 확인
        if transformed_text != text:
            return transformed_text
        return False


    while hwp.MoveSelNextWord():  # 다음 단어를 선택(다음단어가 없으면 break)

        text = hwp.get_selected_text()  # 문자열을 가져와서
        text = normalize_english_case(text)
        if text:
            hwp.insert_memo(f"대소문자로 구분되는 외국어의 경우 약어를 제외하고 소문자로 표기한다.")
            hwp.MoveSelNextWord()
        hwp.Cancel()



def K415_3(hwp):
    # 정규식 패턴 정의 (괄호 포함)
    patterns = {
        "한글": r"\((?=[가-힣])[^0-9]*[가-힣]+[^0-9]*\)",
        "한자": r"\((?=[一-龥])[^0-9]*[一-龥]+[^0-9]*\)",
        "일본어": r"\((?=[ぁ-んァ-ン])[^0-9]*[ぁ-んァ-ン]+[^0-9]*\)",
        "영어": r"\((?=[a-zA-Z])[^0-9]*[a-zA-Z]+[^0-9]*\)",
    }

    def check_order(text):
        """
        입력 텍스트에서 "한글(한자, 일본어, 영어)"의 순서를 확인하는 함수.
        영어 문자는 필요에 따라 소문자로 변환함.
        순서가 맞으면 False, 아니면 True를 반환.
        """
        # 텍스트에서 각 언어 추출
        extracted = {key: re.search(pattern, text) for key, pattern in patterns.items()}

        # 순서대로 등장하는지 확인
        sequence = []
        for key, match in extracted.items():
            if match:  # 해당 언어가 발견되면 위치 저장
                sequence.append((key, match.start()))

        # 정렬된 순서와 비교
        expected_order = ["한글", "한자", "일본어", "영어"]
        actual_order = [key for key, _ in sorted(sequence, key=lambda x: x[1])]  # 위치 기준 정렬

        if actual_order == [key for key in expected_order if key in actual_order]:
            return False  # 순서가 맞음
        else:
            return True  # 순서가 틀림


    hwp.MoveDocBegin()

    while hwp.MoveSelNextWord():  # 다음 단어를 선택(다음단어가 없으면 break)
        text = hwp.get_selected_text()  # 문자열을 가져와서
        text = check_order(text)
        if text:
            hwp.insert_memo(f"한글과 병기시 '한글(한자, 일본어, 영어)'의 순서로 입력.")
            hwp.MoveSelNextWord()
        hwp.Cancel()




if __name__ == '__main__':
    #hwp.open(r'../data/KCS수밀 콘크리트 24.05.hwp')
    #K415_2(hwp)

    hwp.open(r'../data/KDS_11_80_25_돌(블록)쌓기옹벽.hwp')
    K415_2(hwp)
   # K415_3(hwp)

