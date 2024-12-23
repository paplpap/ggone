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

# Example usage
example_titles = [
    '1. 일반사항', '1.1 적용범위', '1.2 참고 기준', '1.2.2 관련 기준', '1.2.1 관련 법규',
    '1.3 용어의 정의', '1.4 수밀 콘크리트 일반', '1.5 제출물', '2. 자재', '2.1 구성재',
    '2.2 배합', '2.3 재료 품질관리', '3. 시공', '3.1 시공일반', '3.2 운반', '3.3 타설',
    '3.4 양생', '3.5 현장품질관리'
]

# Testing the function
new_titles = extract_top_titles(example_titles)
print(f"new_titles: {new_titles}")
