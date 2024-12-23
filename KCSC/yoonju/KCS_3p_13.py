import sys
import os
from pyhwpx import Hwp
from Highlighting import highlighting
import re

one_titles = [
        "1. 일반사항",
        "1.1 적용 범위",
        "1.2 참고 기준",
        "1.3 용어의 정의"
]

main_titles = [
        "1. 일반사항",
        "2. 자재",
        "3. 시공"
]

except_title = [ # 하위 숫자가 있는 title
]

def remove_duplicates_with_order(titles, explanations):
    '''
    하이라이팅 칠 문장의 중복 제거
    :param titles: 하이라이팅 칠 문장들
    :param explanations: 하이라이팅 친 이유
    :return: 하이라이팅 칠 문장, 그 이유
    '''
    seen = set()
    unique_titles = []
    unique_explanations = []

    for title, exp in zip(titles, explanations):
        if title not in seen:
            unique_titles.append(title)
            unique_explanations.append(exp)
            seen.add(title)

    return unique_titles, unique_explanations

def extract_top_titles(titles):
    # Sort the titles to ensure hierarchical processing
    titles = sorted(titles)
    hierarchy = {}

    # Build the hierarchy
    for title in titles:
        parts = title.split(" ")
        key = parts[0]
        if "." in key:
            parent_key = ".".join(key.split(".")[:-1])
            if parent_key:
                if parent_key in hierarchy:
                    hierarchy[parent_key].append(title)
                else:
                    hierarchy[parent_key] = [title]

    # Extract titles that have children and format them correctly
    titles_with_children = [key for key, value in hierarchy.items() if len(value) > 1]
    titles_with_children = [title+'.' if len(title) == 1 else title for title in titles_with_children]
    titles_with_children = [title+' ' for title in titles_with_children]

    # Extract full titles that match the patterns of titles with children
    top_titles = []
    for num in titles_with_children:
        for title in titles:
            if num in title:
                top_titles.append(title)
                break

    return top_titles

def check_title_order(found_titles, require_titles):
    '''
    title의 누락 및 순서 검사
    :param found_titles: 이 문서에서 찾은 title
    :param require_titles: 지침에 정해진 title
    :return: 하이라이팅 칠 문장, 그 이유
    '''

    incorrect_titles = []
    explanation = []

    mismatched = [orig for orig, sorted_elem in zip(found_titles, sorted(found_titles)) if orig != sorted_elem]
    incorrect_titles.extend(mismatched)
    explanation.extend([f"제 13조 <{title}>의 순서/위치가 잘못되었습니다." for title in mismatched])

    missing_titles = [title for title in require_titles if title not in found_titles]
    incorrect_titles.extend(missing_titles)
    explanation.extend([f"제 13조 <{title}>가 있어야하지만 빠졌습니다." for title in missing_titles])

    # 있지만 순서가 틀린 titles 여기 뭔가 이상함 적용범위가... 있어야하는데 빠진거랑 관련이 있는듯
    union = [title for title in require_titles if title in found_titles]
    require_titles = [title for title in require_titles if title in union]
    found_titles = [title for title in found_titles if title in union]
    misordered_titles = [
        title for title, expected in zip(found_titles, require_titles)
        if title != expected
    ]

    incorrect_titles.extend(misordered_titles)
    explanation.extend([f"제 13조 <{title}>의 순서/위치가 잘못되었습니다." for title in misordered_titles])

    return incorrect_titles, explanation

def check_title_content(documents, title_list, found_titles, require_titles):
    '''
    title이 비어있을 경우 내용 없음 작성
    :param documents: 문서
    :param title_list:
    :param found_titles:
    :param require_titles:
    :return:
    '''
    incorrect_titles = []
    explanation = []

    except_title =  extract_top_titles(title_list) # 내용 없음이 없어야 하는 title
    extract_number = lambda title: re.match(r'^\d+(\.\d+)*', title).group() if re.match(r'^\d+(\.\d+)*', title) else None
    filter_num = set(filter(None, map(extract_number, found_titles))) & set(filter(None, map(extract_number, require_titles))) # 문서에 있는 title의 숫자
    filtered_found_titles = [title for title in found_titles if extract_number(title) in filter_num]
    find_titles = [x for x in filtered_found_titles if x not in except_title] # 이 문서에서 '내용 없음'을 작성해야함

    lines = documents.split('\n')
    lines = [line.replace('\r', '') for line in lines]

    for title in find_titles:
        start = title_list[title_list.index(title)]
        start_idx = lines.index(start)
        end = title_list[title_list.index(title)+1]
        end_idx = lines.index(end)
        paragraph = ' '.join(lines[start_idx+1: end_idx])

        paragraph = paragraph.replace(' ', '')
        if len(paragraph)!=0:
            pass
        else:
            incorrect_titles.append(title)
            explanation.append(f"제 13조 <{title}>에 기술 또는 '내용 없음'을 작성해야 합니다.")

    for title in except_title:
        start = title_list[title_list.index(title)]
        start_idx = lines.index(start)
        end = title_list[title_list.index(title)+1]
        end_idx = lines.index(end)
        paragraph = ' '.join(lines[start_idx+1: end_idx])

        paragraph = paragraph.replace(' ', '')
        if len(paragraph)==0:
            pass
        else:
            incorrect_titles.append(title)
            explanation.append(f"제 13조 <{title}>에 '내용 없음'을 기술하지 말아야합니다.")

    return incorrect_titles, explanation

def read_hwp(hwp):
    '''
    한글 파일을 문장으로 읽기
    :param hwp: hwp 파일
    :return: 한글 파일의 str 형태
    '''
    hwp.init_scan()
    texts = hwp.get_text_file()

    return texts

def read_after_table(hwp):
    '''
    목차 이후의 한글 파일을 한 문장씩 리스트로 읽기
    :param hwp: 한글 파일의 str 형태
    :return: 목차 이후로 한글 파일의 리스트 형태 (한 문장씩)
    '''
    texts = read_hwp(hwp)
    lines = texts.split('\n')
    inx = 0

    for i, line in enumerate(lines):
        if line.strip() == "1. 일반사항":
            inx = i

    if inx == 0:
        return 0
    else:
        return "\n".join(lines[inx:]).strip()

def KCS_configure_13(hwp):
    '''
    제 13조를 검사하고 잘못된 문장을 모음
    :param hwp: 목차 이후로 한 문장씩 읽은 한글 파일
    :return: 하이라이팅 해야하는 문장들, 이유들
    '''
    documents = read_after_table(hwp)
    if documents == 0:
        return False, False

    found_titles = re.findall(r"^\d+\.\d*(?:\.\d+)* [^\n]*", documents, re.MULTILINE)
    found_titles = [title.replace('\r', '') for title in found_titles]
    found_main = [title for title in found_titles if re.match(r'^\d+\.\s', title)]
    found_one = [title for title in found_titles if re.match(r'^1\.', title)]
    # found_titles = [
    #         title for title in found_titles
    #         if re.match(r'^(\d+(\.\d+){1,3})|(\d\.)$', title.split(' ')[0])
    # ]

    incorrect_titles1, explanation1 = check_title_order(found_main, main_titles)
    incorrect_titles2, explanation2 = check_title_order(found_one, one_titles)
    incorrect_titles3, explanation3 = check_title_content(documents, found_titles, found_one, one_titles)

    incorrect_titles = incorrect_titles1 + incorrect_titles2 + incorrect_titles3
    explanation = explanation1 + explanation2 + explanation3

    return incorrect_titles, explanation

def KCS_313(hwp):
    '''
    제 13조를 검사하는 함수
    :param hwp: 한글 문서
    :return: 문서가 하이라이팅 됌
    '''
    incorrect_titles, explanation = KCS_configure_13(hwp)
    if incorrect_titles == explanation == False:
        print(f"<1. 일반사항>을 추가하셔야 오류 검사가 가능합니다.")
        return 0
    incorrect_titles, explanation = remove_duplicates_with_order(incorrect_titles, explanation)
    highlighting(hwp, incorrect_titles, explanation)
    # for inx, (title, exp) in enumerate(zip(incorrect_titles, explanation)):
    #     print(f"{title}")
    #     print(f"{exp}")


# if __name__=="__main__":
#     file_path = r'../data/KCS213000.hwp'
#     hwp = Hwp()
#     hwp.open(file_path)
#     KCS_313(hwp)
