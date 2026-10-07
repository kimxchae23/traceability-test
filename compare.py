from difflib import SequenceMatcher
from copy import copy

from openpyxl import load_workbook
from docx import Document

import os
import re


# ============================================================
# 1. 추적성 Excel 파일 자동 탐색
# ============================================================

def find_excel_file():

    candidates = []

    for file_name in os.listdir("."):

        # Excel 임시파일 제외
        if file_name.startswith("~$"):
            continue

        # xlsx만 확인
        if not file_name.lower().endswith(".xlsx"):
            continue

        # 프로그램이 만든 결과 파일 제외
        if "_rev" in file_name.lower():
            continue

        # 요건단계 / 설계단계 추적성 분석표 탐색
        if (
            "요건단계 추적성 분석표" in file_name
            or
            "설계단계 추적성 분석표" in file_name
        ):

            candidates.append(file_name)

    if len(candidates) == 0:

        raise FileNotFoundError(
            "요건단계 또는 설계단계 추적성 분석표 "
            "Excel 파일을 찾을 수 없습니다."
        )

    if len(candidates) > 1:

        raise ValueError(
            f"추적성 분석표가 여러 개 있습니다: {candidates}"
        )

    return candidates[0]


# ============================================================
# 2. 설정
# ============================================================

OLD_WORD = "old_srs.docx"
NEW_WORD = "new_srs.docx"

# Excel 자동 검색
EXCEL_FILE = find_excel_file()

# 결과 파일
OUTPUT_FILE = EXCEL_FILE.replace(
    ".xlsx",
    "_Rev01.xlsx"
)

# 비교하려는 개정 문서 종류
# "SRS", "DS", "SDD"
DOCUMENT_TYPE = "SRS"


# ============================================================
# 3. Word 장절 번호 인식
#
# 예:
# 5.1 입력
# 5.1.1 입력 처리
# 7.2.3 오류 처리
# ============================================================

SECTION_PATTERN = re.compile(
    r"^(\d+(?:\.\d+)+)\s+(.*)"
)


# ============================================================
# 4. Word 문서를 장절별로 파싱
# ============================================================

def parse_sections(file_path):

    doc = Document(file_path)

    sections = {}

    current_section = None
    current_content = []

    for paragraph in doc.paragraphs:

        text = paragraph.text.strip()

        if not text:
            continue

        match = SECTION_PATTERN.match(text)

        # 새로운 장절 발견
        if match:

            # 이전 장절 저장
            if current_section is not None:

                sections[current_section] = {
                    "content": "\n".join(
                        current_content
                    ).strip()
                }

            current_section = match.group(1)

            # 장절번호 뒤 제목/텍스트 포함
            first_text = match.group(2).strip()

            current_content = [
                first_text
            ]

        else:

            if current_section is not None:

                current_content.append(
                    text
                )

    # 마지막 장절 저장
    if current_section is not None:

        sections[current_section] = {
            "content": "\n".join(
                current_content
            ).strip()
        }

    return sections


# ============================================================
# 5. 장절 번호 정렬
# ============================================================

def section_sort_key(section):

    return [
        int(number)
        for number in section.split(".")
    ]


# ============================================================
# 6. 텍스트 정규화
# ============================================================

def normalize_text(text):

    if text is None:
        return ""

    # 줄바꿈 / 여러 공백을 한 칸으로 통일
    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# 7. 텍스트 유사도
#
# 0.0 ~ 1.0
# ============================================================

def text_similarity(
    text1,
    text2
):

    text1 = normalize_text(
        text1
    )

    text2 = normalize_text(
        text2
    )

    if not text1 or not text2:
        return 0.0

    return SequenceMatcher(
        None,
        text1,
        text2
    ).ratio()


# ============================================================
# 8. 이후 절들이 다시 정렬되는지 확인
#
# 신규 절 삽입 / 기존 절 삭제 후
# 뒤의 절들이 다시 이어지는지 확인
# ============================================================

def confirm_alignment(
    old_items,
    new_items,
    old_index,
    new_index,
    check_count=2,
    threshold=0.80
):

    available = min(
        check_count,
        len(old_items) - old_index,
        len(new_items) - new_index
    )

    if available <= 0:
        return False

    matched = 0

    for offset in range(
        available
    ):

        old_content = (
            old_items[
                old_index + offset
            ][1]["content"]
        )

        new_content = (
            new_items[
                new_index + offset
            ][1]["content"]
        )

        score = text_similarity(
            old_content,
            new_content
        )

        if score >= threshold:

            matched += 1

        else:

            break

    return matched == available


# ============================================================
# 9. 기존 Word / 개정 Word 비교
#
# 상태:
# 동일
# 수정
# 추가
# 삭제
# 이동
# 이동+수정
# 변경확인필요
# ============================================================

def compare_documents(
    old_file,
    new_file
):

    old_sections = parse_sections(
        old_file
    )

    new_sections = parse_sections(
        new_file
    )

    old_items = sorted(
        old_sections.items(),
        key=lambda x: section_sort_key(
            x[0]
        )
    )

    new_items = sorted(
        new_sections.items(),
        key=lambda x: section_sort_key(
            x[0]
        )
    )

    changes = []

    i = 0
    j = 0

    # 최대 몇 절 앞까지 볼 것인지
    LOOKAHEAD = 3

    # 사실상 동일한 내용
    SAME_THRESHOLD = 0.97

    # 같은 feature/절의 수정으로 볼 최소 유사도
    MODIFY_THRESHOLD = 0.65

    # 삽입/삭제 후 다시 연결됐다고 보는 기준
    ALIGN_THRESHOLD = 0.80

    while (
        i < len(old_items)
        and
        j < len(new_items)
    ):

        old_section, old_data = (
            old_items[i]
        )

        new_section, new_data = (
            new_items[j]
        )

        old_content = (
            old_data["content"]
        )

        new_content = (
            new_data["content"]
        )

        similarity = text_similarity(
            old_content,
            new_content
        )

        # ====================================================
        # 1. 내용이 거의 동일
        # ====================================================

        if similarity >= SAME_THRESHOLD:

            if old_section == new_section:

                status = "동일"

            else:

                status = "이동"

            changes.append({

                "section":
                    old_section,

                "old_section":
                    old_section,

                "new_section":
                    new_section,

                "status":
                    status,

                "old_content":
                    f"{old_section} {old_content}",

                "new_content":
                    f"{new_section} {new_content}",

                "similarity":
                    similarity
            })

            i += 1
            j += 1

            continue

        # ====================================================
        # 2. 같은 장절인데 내용이 일부 수정
        # ====================================================

        if (
            old_section == new_section
            and
            similarity >= MODIFY_THRESHOLD
        ):

            changes.append({

                "section":
                    old_section,

                "old_section":
                    old_section,

                "new_section":
                    new_section,

                "status":
                    "수정",

                "old_content":
                    f"{old_section} {old_content}",

                "new_content":
                    f"{new_section} {new_content}",

                "similarity":
                    similarity
            })

            i += 1
            j += 1

            continue

        # ====================================================
        # 3. NEW 문서에 절이 삽입되었는지 확인
        #
        # NEW 쪽을 몇 칸 앞으로 이동했을 때
        # OLD 현재 절과 다시 연결되는지 검사
        # ====================================================

        insertion_found = False

        for offset in range(
            1,
            LOOKAHEAD + 1
        ):

            if (
                j + offset
                >=
                len(new_items)
            ):

                break

            candidate_new_section, candidate_new_data = (
                new_items[
                    j + offset
                ]
            )

            candidate_similarity = (
                text_similarity(
                    old_content,
                    candidate_new_data[
                        "content"
                    ]
                )
            )

            if (
                candidate_similarity
                <
                ALIGN_THRESHOLD
            ):

                continue

            # 뒤의 절까지 연속적으로 맞는지 확인
            if confirm_alignment(
                old_items,
                new_items,
                i,
                j + offset,
                check_count=2,
                threshold=ALIGN_THRESHOLD
            ):

                # j ~ j+offset-1은 신규 추가
                for inserted_index in range(
                    j,
                    j + offset
                ):

                    (
                        inserted_section,
                        inserted_data
                    ) = (
                        new_items[
                            inserted_index
                        ]
                    )

                    changes.append({

                        "section":
                            inserted_section,

                        "old_section":
                            None,

                        "new_section":
                            inserted_section,

                        "status":
                            "추가",

                        "old_content":
                            "",

                        "new_content":
                            (
                                f"{inserted_section} "
                                f"{inserted_data['content']}"
                            ),

                        "similarity":
                            0.0
                    })

                j += offset

                insertion_found = True

                break

        if insertion_found:

            continue

        # ====================================================
        # 4. OLD 문서의 절이 삭제되었는지 확인
        # ====================================================

        deletion_found = False

        for offset in range(
            1,
            LOOKAHEAD + 1
        ):

            if (
                i + offset
                >=
                len(old_items)
            ):

                break

            (
                candidate_old_section,
                candidate_old_data
            ) = (
                old_items[
                    i + offset
                ]
            )

            candidate_similarity = (
                text_similarity(
                    candidate_old_data[
                        "content"
                    ],
                    new_content
                )
            )

            if (
                candidate_similarity
                <
                ALIGN_THRESHOLD
            ):

                continue

            if confirm_alignment(
                old_items,
                new_items,
                i + offset,
                j,
                check_count=2,
                threshold=ALIGN_THRESHOLD
            ):

                # i ~ i+offset-1은 삭제
                for deleted_index in range(
                    i,
                    i + offset
                ):

                    (
                        deleted_section,
                        deleted_data
                    ) = (
                        old_items[
                            deleted_index
                        ]
                    )

                    changes.append({

                        "section":
                            deleted_section,

                        "old_section":
                            deleted_section,

                        "new_section":
                            None,

                        "status":
                            "삭제",

                        "old_content":
                            (
                                f"{deleted_section} "
                                f"{deleted_data['content']}"
                            ),

                        "new_content":
                            "",

                        "similarity":
                            0.0
                    })

                i += offset

                deletion_found = True

                break

        if deletion_found:

            continue

        # ====================================================
        # 5. 번호는 달라졌지만 내용은 상당히 유사
        #
        # 이동 + 일부 수정
        # ====================================================

        if similarity >= MODIFY_THRESHOLD:

            changes.append({

                "section":
                    old_section,

                "old_section":
                    old_section,

                "new_section":
                    new_section,

                "status":
                    "이동+수정",

                "old_content":
                    f"{old_section} {old_content}",

                "new_content":
                    f"{new_section} {new_content}",

                "similarity":
                    similarity
            })

            i += 1
            j += 1

            continue

        # ====================================================
        # 6. 자동 판단하기 어려움
        # ====================================================

        changes.append({

            "section":
                old_section,

            "old_section":
                old_section,

            "new_section":
                new_section,

            "status":
                "변경확인필요",

            "old_content":
                f"{old_section} {old_content}",

            "new_content":
                f"{new_section} {new_content}",

            "similarity":
                similarity
        })

        i += 1
        j += 1

    # ========================================================
    # OLD에만 남은 절 = 삭제
    # ========================================================

    while i < len(old_items):

        old_section, old_data = (
            old_items[i]
        )

        changes.append({

            "section":
                old_section,

            "old_section":
                old_section,

            "new_section":
                None,

            "status":
                "삭제",

            "old_content":
                (
                    f"{old_section} "
                    f"{old_data['content']}"
                ),

            "new_content":
                "",

            "similarity":
                0.0
        })

        i += 1

    # ========================================================
    # NEW에만 남은 절 = 추가
    # ========================================================

    while j < len(new_items):

        new_section, new_data = (
            new_items[j]
        )

        changes.append({

            "section":
                new_section,

            "old_section":
                None,

            "new_section":
                new_section,

            "status":
                "추가",

            "old_content":
                "",

            "new_content":
                (
                    f"{new_section} "
                    f"{new_data['content']}"
                ),

            "similarity":
                0.0
        })

        j += 1

    return changes


# ============================================================
# 10. 변경 결과 터미널 출력
# ============================================================

def print_changes(changes):

    print()
    print("=" * 100)

    print(
        f"{'OLD':<10} "
        f"{'NEW':<10} "
        f"{'상태':<15} "
        f"{'유사도':<8} "
        f"변경 내용"
    )

    print("=" * 100)

    for change in changes:

        old_section = (
            change.get(
                "old_section"
            )
            or "-"
        )

        new_section = (
            change.get(
                "new_section"
            )
            or "-"
        )

        status = change[
            "status"
        ]

        similarity = change.get(
            "similarity",
            0
        )

        old_content = change.get(
            "old_content",
            ""
        )

        new_content = change.get(
            "new_content",
            ""
        )

        print(
            f"{old_section:<10} "
            f"{new_section:<10} "
            f"{status:<15} "
            f"{similarity:.2f}"
        )

        if status in [
            "수정",
            "이동+수정",
            "변경확인필요"
        ]:

            print(
                f"    기존: "
                f"{old_content}"
            )

            print(
                f"    개정: "
                f"{new_content}"
            )

        elif status == "추가":

            print(
                f"    신규: "
                f"{new_content}"
            )

        elif status == "삭제":

            print(
                f"    삭제: "
                f"{old_content}"
            )

        elif status == "이동":

            print(
                "    내용 동일 / "
                "장절 번호 이동"
            )

        print("-" * 100)


# ============================================================
# 11. Excel 파일명에서 단계 판별
# ============================================================

def detect_stage(
    excel_path
):

    file_name = os.path.basename(
        excel_path
    )

    if "요건단계" in file_name:

        return "요건"

    elif "설계단계" in file_name:

        return "설계"

    else:

        raise ValueError(
            "Excel 파일명에서 "
            "'요건단계' 또는 "
            "'설계단계'를 찾을 수 없습니다."
        )


# ============================================================
# 12. Sheet 이름 / 순서로 정방향·역방향 판별
# ============================================================

def get_direction(
    ws,
    sheet_index
):

    sheet_name = ws.title

    if "정방향" in sheet_name:

        return "정방향"

    elif "역방향" in sheet_name:

        return "역방향"

    # 이름에 없으면 Sheet 순서 사용
    if sheet_index == 0:

        return "정방향"

    elif sheet_index == 1:

        return "역방향"

    return None


# ============================================================
# 13. 단계 + 방향 + 문서 종류로 Contents 열 결정
#
# B = ID
# C = Contents
#
# E = ID
# F = Contents
# ============================================================

def get_target_column(
    stage,
    direction,
    document_type
):

    # ========================================================
    # 요건 단계
    # ========================================================

    if stage == "요건":

        # 정방향
        #
        # C = DS Contents
        # F = SRS Contents

        if direction == "정방향":

            mapping = {
                "DS": 3,
                "SRS": 6
            }

        # 역방향
        #
        # C = SRS Contents
        # F = DS Contents

        elif direction == "역방향":

            mapping = {
                "SRS": 3,
                "DS": 6
            }

        else:

            raise ValueError(
                f"알 수 없는 검사 방향: "
                f"{direction}"
            )

    # ========================================================
    # 설계 단계
    # ========================================================

    elif stage == "설계":

        # 정방향
        #
        # C = SRS Contents
        # F = SDD Contents

        if direction == "정방향":

            mapping = {
                "SRS": 3,
                "SDD": 6
            }

        # 역방향
        #
        # C = SDD Contents
        # F = SRS Contents

        elif direction == "역방향":

            mapping = {
                "SDD": 3,
                "SRS": 6
            }

        else:

            raise ValueError(
                f"알 수 없는 검사 방향: "
                f"{direction}"
            )

    else:

        raise ValueError(
            f"알 수 없는 단계: "
            f"{stage}"
        )

    if document_type not in mapping:

        raise ValueError(
            f"{stage}단계 / "
            f"{direction} 검사에서 "
            f"{document_type} 문서를 "
            f"찾을 수 없습니다."
        )

    return mapping[
        document_type
    ]


# ============================================================
# 14. Excel Contents에서 장절번호 추출
#
# 예:
# 5.1.3 범위 판정
#
# -> 5.1.3
# ============================================================

def get_section_number(
    cell_value
):

    if cell_value is None:

        return None

    text = str(
        cell_value
    ).strip()

    match = re.match(
        r"^(\d+(?:\.\d+)+)",
        text
    )

    if match:

        return match.group(1)

    return None


# ============================================================
# 15. Comment 빨간 글씨 처리
# ============================================================

def set_review_comment(
    cell,
    text
):

    cell.value = text

    new_font = copy(
        cell.font
    )

    # 빨간색
    new_font.color = "FFFF0000"

    cell.font = new_font


# ============================================================
# 16. Excel 자동 업데이트
# ============================================================

def update_excel(
    excel_path,
    output_path,
    changes,
    document_type
):

    wb = load_workbook(
        excel_path
    )

    stage = detect_stage(
        excel_path
    )

    print()
    print("==============================")
    print(f"Excel 파일: {excel_path}")
    print(f"감지된 단계: {stage}단계")
    print("==============================")

    # G열 = Comment
    COMMENT_COL = 7

    # 실제 데이터 시작 행
    DATA_START_ROW = 3

    # ========================================================
    # 첫 번째 / 두 번째 Sheet 처리
    # ========================================================

    for sheet_index, ws in enumerate(
        wb.worksheets[:2]
    ):

        direction = get_direction(
            ws,
            sheet_index
        )

        target_col = get_target_column(
            stage,
            direction,
            document_type
        )

        column_name = (
            "C"
            if target_col == 3
            else "F"
        )

        print()
        print(
            f"===== {ws.title} ====="
        )

        print(
            f"검사 방향: {direction}"
        )

        print(
            f"{document_type} Contents 위치: "
            f"{column_name}열"
        )

        # ====================================================
        # ★ 중요
        #
        # Excel을 수정하기 전에
        # 원본 장절 번호 → 행 번호를 전부 저장
        #
        # 이후 Contents가 변경돼도
        # 최초 행 위치를 기준으로만 수정
        # ====================================================

        original_section_rows = {}

        for row in range(
            DATA_START_ROW,
            ws.max_row + 1
        ):

            content = ws.cell(
                row=row,
                column=target_col
            ).value

            excel_section = (
                get_section_number(
                    content
                )
            )

            if excel_section is None:

                continue

            if (
                excel_section
                not in
                original_section_rows
            ):

                original_section_rows[
                    excel_section
                ] = []

            original_section_rows[
                excel_section
            ].append(
                row
            )

        print(
            "원본 Excel 장절 매핑:",
            original_section_rows
        )

        # ====================================================
        # Word 변경사항 처리
        # ====================================================

        for change in changes:

            status = change[
                "status"
            ]

            # 동일은 아무것도 하지 않음
            if status == "동일":

                continue

            old_section = (
                change.get(
                    "old_section"
                )
            )

            new_section = (
                change.get(
                    "new_section"
                )
            )

            print()

            print(
                f"변경 감지: "
                f"{document_type} "
                f"{old_section or new_section} "
                f"/ {status}"
            )

            # =================================================
            # 신규 추가
            #
            # 기존 Excel에 매핑 정보가 없으므로
            # 자동 행 추가 금지
            # =================================================

            if status == "추가":

                print(
                    f">>> 신규 장절 "
                    f"{new_section}: "
                    f"기존 Excel에 매핑된 "
                    f"항목 없음"
                )

                continue

            # =================================================
            # 기존 Excel 검색은 반드시 OLD 장절 기준
            # =================================================

            if old_section is None:

                continue

            matched_rows = (
                original_section_rows.get(
                    old_section,
                    []
                )
            )

            if not matched_rows:

                print(
                    f">>> 기존 장절 "
                    f"{old_section}을 "
                    f"{ws.title}에서 찾지 못함"
                )

                continue

            # =================================================
            # 해당 OLD 장절이 연결되어 있던
            # 모든 Excel 행 처리
            # =================================================

            for row in matched_rows:

                print(
                    f">>> 매칭 성공: "
                    f"{old_section} "
                    f"/ {ws.title} "
                    f"/ {row}행"
                )

                # =============================================
                # 수정
                # =============================================

                if status == "수정":

                    ws.cell(
                        row=row,
                        column=target_col
                    ).value = (
                        change[
                            "new_content"
                        ]
                    )

                    comment_cell = (
                        ws.cell(
                            row=row,
                            column=COMMENT_COL
                        )
                    )

                    set_review_comment(
                        comment_cell,
                        "재검토 필요"
                    )

                # =============================================
                # 장절 이동
                # =============================================

                elif status == "이동":

                    ws.cell(
                        row=row,
                        column=target_col
                    ).value = (
                        change[
                            "new_content"
                        ]
                    )

                    comment_cell = (
                        ws.cell(
                            row=row,
                            column=COMMENT_COL
                        )
                    )

                    set_review_comment(
                        comment_cell,
                        "장절 이동 - 재검토 필요"
                    )

                # =============================================
                # 이동 + 수정
                # =============================================

                elif status == "이동+수정":

                    ws.cell(
                        row=row,
                        column=target_col
                    ).value = (
                        change[
                            "new_content"
                        ]
                    )

                    comment_cell = (
                        ws.cell(
                            row=row,
                            column=COMMENT_COL
                        )
                    )

                    set_review_comment(
                        comment_cell,
                        (
                            "장절 이동 및 수정 - "
                            "재검토 필요"
                        )
                    )

                # =============================================
                # 삭제
                #
                # 기존 Contents는 유지
                # =============================================

                elif status == "삭제":

                    comment_cell = (
                        ws.cell(
                            row=row,
                            column=COMMENT_COL
                        )
                    )

                    set_review_comment(
                        comment_cell,
                        (
                            f"{document_type} 삭제 - "
                            f"추적성 재검토 필요"
                        )
                    )

                # =============================================
                # 자동으로 판단하기 어려움
                #
                # Contents는 건드리지 않음
                # =============================================

                elif status == "변경확인필요":

                    comment_cell = (
                        ws.cell(
                            row=row,
                            column=COMMENT_COL
                        )
                    )

                    set_review_comment(
                        comment_cell,
                        "변경 관계 확인 필요"
                    )

    # ========================================================
    # 결과 저장
    # ========================================================

    wb.save(
        output_path
    )

    print()
    print("==============================")
    print("Excel 업데이트 완료")
    print(f"저장 파일: {output_path}")
    print("==============================")


# ============================================================
# 17. 실제 실행
# ============================================================

changes = compare_documents(
    OLD_WORD,
    NEW_WORD
)

print_changes(
    changes
)

update_excel(
    EXCEL_FILE,
    OUTPUT_FILE,
    changes,
    DOCUMENT_TYPE
)