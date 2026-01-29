"""
Outlook 메일 파일(.msg)을 PDF로 변환하는 Python 스크립트

요구사항:
    - Windows OS
    - Microsoft Outlook 설치
    - Microsoft Word 설치
    - pip install pywin32

사용법:
    python outlook_to_pdf.py [입력폴더] [출력폴더]
    python outlook_to_pdf.py C:\Mail C:\Mail\PDF
"""

import os
import sys
import re
from pathlib import Path

try:
    import win32com.client
    from win32com.client import constants
except ImportError:
    print("pywin32가 설치되어 있지 않습니다.")
    print("설치 명령어: pip install pywin32")
    sys.exit(1)


def clean_filename(filename: str, max_length: int = 200) -> str:
    """파일명에서 사용 불가능한 문자 제거"""
    # Windows에서 사용 불가능한 문자 제거
    cleaned = re.sub(r'[\\/:*?"<>|]', '_', filename)
    # 앞뒤 공백 및 점 제거
    cleaned = cleaned.strip('. ')
    # 길이 제한
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length]
    return cleaned


def get_unique_path(filepath: Path) -> Path:
    """중복 파일명 처리 - 번호 추가"""
    if not filepath.exists():
        return filepath

    counter = 1
    while True:
        new_path = filepath.parent / f"{filepath.stem}_{counter}{filepath.suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def convert_msg_to_pdf(msg_path: Path, output_folder: Path, outlook, word) -> bool:
    """단일 .msg 파일을 PDF로 변환"""
    try:
        namespace = outlook.GetNamespace("MAPI")

        # 메일 열기
        mail = namespace.OpenSharedItem(str(msg_path))

        # PDF 파일 경로 생성
        pdf_name = clean_filename(msg_path.stem) + ".pdf"
        pdf_path = get_unique_path(output_folder / pdf_name)

        # Word 문서 생성
        doc = word.Documents.Add()

        # 메일 정보 추가
        content = doc.Content
        content.InsertAfter(f"보낸 사람: {mail.SenderName} <{mail.SenderEmailAddress}>\n")
        content.InsertAfter(f"받는 사람: {mail.To}\n")
        if mail.CC:
            content.InsertAfter(f"참조: {mail.CC}\n")
        content.InsertAfter(f"날짜: {mail.ReceivedTime}\n")
        content.InsertAfter(f"제목: {mail.Subject}\n")
        content.InsertAfter("-" * 50 + "\n\n")

        # 메일 본문 추가
        content.InsertAfter(mail.Body)

        # PDF로 저장 (17 = wdFormatPDF)
        doc.SaveAs2(str(pdf_path), 17)
        doc.Close(False)

        # 메일 닫기
        mail.Close(0)  # 0 = olDiscard

        print(f"✓ 변환 완료: {pdf_path.name}")
        return True

    except Exception as e:
        print(f"✗ 변환 실패: {msg_path.name} - {e}")
        return False


def convert_folder(input_folder: str, output_folder: str, include_subfolders: bool = True):
    """폴더 내 모든 .msg 파일을 PDF로 변환"""

    input_path = Path(input_folder)
    output_path = Path(output_folder)

    # 입력 폴더 확인
    if not input_path.exists():
        print(f"오류: 입력 폴더가 존재하지 않습니다 - {input_folder}")
        return

    # 출력 폴더 생성
    output_path.mkdir(parents=True, exist_ok=True)

    # .msg 파일 찾기
    if include_subfolders:
        msg_files = list(input_path.rglob("*.msg"))
    else:
        msg_files = list(input_path.glob("*.msg"))

    if not msg_files:
        print("변환할 .msg 파일이 없습니다.")
        return

    print(f"총 {len(msg_files)}개의 .msg 파일을 찾았습니다.\n")

    # Outlook 및 Word 시작
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
    except Exception as e:
        print(f"오류: Outlook 또는 Word를 시작할 수 없습니다 - {e}")
        return

    # 변환 실행
    success_count = 0
    fail_count = 0

    for msg_file in msg_files:
        if convert_msg_to_pdf(msg_file, output_path, outlook, word):
            success_count += 1
        else:
            fail_count += 1

    # 정리
    word.Quit()

    # 결과 출력
    print(f"\n{'='*50}")
    print(f"변환 완료!")
    print(f"성공: {success_count}개")
    print(f"실패: {fail_count}개")
    print(f"저장 위치: {output_path}")


def main():
    if len(sys.argv) >= 3:
        # 명령줄 인자로 폴더 지정
        input_folder = sys.argv[1]
        output_folder = sys.argv[2]
    elif len(sys.argv) == 2:
        # 입력 폴더만 지정, 출력은 하위 PDF 폴더
        input_folder = sys.argv[1]
        output_folder = os.path.join(input_folder, "PDF_Output")
    else:
        # 대화형 모드
        print("Outlook 메일(.msg) → PDF 변환기\n")
        input_folder = input("메일 파일이 있는 폴더 경로: ").strip().strip('"')
        output_folder = input("PDF 저장할 폴더 경로 (Enter=자동): ").strip().strip('"')

        if not output_folder:
            output_folder = os.path.join(input_folder, "PDF_Output")

    print(f"\n입력 폴더: {input_folder}")
    print(f"출력 폴더: {output_folder}\n")

    convert_folder(input_folder, output_folder)


if __name__ == "__main__":
    main()
