"""
이메일 파서 사용 예시
"""
from email_parser import EmailParser
from excel_exporter import ExcelExporter


def example_parse_single_file():
    """예시 1: 단일 파일 파싱"""
    print("예시 1: 단일 파일 파싱")
    print("-" * 50)

    parser = EmailParser()

    # .eml 파일 파싱
    try:
        msg = parser.parse_eml_file('test.eml')

        # 파싱된 정보 출력
        print(f"제목: {msg.subject}")
        print(f"발신자: {msg.sender}")
        print(f"수신자: {msg.recipients}")
        print(f"날짜: {msg.date}")
        print(f"본문 (처음 100자): {msg.body[:100]}...")
        print()

    except FileNotFoundError:
        print("test.eml 파일을 찾을 수 없습니다.")
        print()


def example_parse_directory():
    """예시 2: 디렉토리 파싱 및 엑셀 저장"""
    print("예시 2: 디렉토리 파싱 및 엑셀 저장")
    print("-" * 50)

    parser = EmailParser()
    exporter = ExcelExporter()

    try:
        # 디렉토리 내 모든 .eml 파일 파싱
        messages = parser.parse_directory('./emails')

        print(f"총 {len(messages)}개의 이메일을 파싱했습니다.")

        # 스레드 구성
        parser.build_threads(messages)
        print(f"{len(parser.threads)}개의 스레드를 발견했습니다.")

        # 질의응답 추출
        qa_pairs = parser.get_all_qa_pairs()
        print(f"{len(qa_pairs)}개의 질의응답 쌍을 추출했습니다.")

        # 엑셀로 저장
        output_file = exporter.export_qa_pairs(qa_pairs, 'output_qa.xlsx')
        print(f"엑셀 파일이 저장되었습니다: {output_file}")
        print()

    except Exception as e:
        print(f"오류: {str(e)}")
        print()


def example_parse_multiple_files():
    """예시 3: 여러 파일 파싱"""
    print("예시 3: 여러 파일 파싱")
    print("-" * 50)

    parser = EmailParser()
    exporter = ExcelExporter()

    # 파싱할 파일 목록
    files = ['email1.eml', 'email2.eml', 'email3.eml']

    try:
        # 여러 파일 파싱
        messages = parser.parse_eml_files(files)

        print(f"총 {len(messages)}개의 이메일을 파싱했습니다.")

        # 원본 이메일 형식으로 저장
        raw_messages = [msg.to_dict() for msg in messages]
        output_file = exporter.export_raw_emails(raw_messages, 'output_raw.xlsx')
        print(f"원본 이메일 형식으로 저장되었습니다: {output_file}")
        print()

    except Exception as e:
        print(f"오류: {str(e)}")
        print()


def example_programmatic_usage():
    """예시 4: 프로그래밍 방식으로 사용"""
    print("예시 4: 프로그래밍 방식으로 사용")
    print("-" * 50)

    parser = EmailParser()

    try:
        # 파일 파싱
        msg = parser.parse_eml_file('test.eml')

        # 딕셔너리로 변환
        msg_dict = msg.to_dict()

        # JSON처럼 사용 가능
        print("파싱된 이메일 정보:")
        for key, value in msg_dict.items():
            if key == 'body':
                print(f"  {key}: {value[:50]}...")
            else:
                print(f"  {key}: {value}")
        print()

    except FileNotFoundError:
        print("test.eml 파일을 찾을 수 없습니다.")
        print()


if __name__ == '__main__':
    print("\n" + "=" * 50)
    print("    이메일 파서 사용 예시")
    print("=" * 50 + "\n")

    # 예시 실행 (실제 파일이 있을 때만 작동)
    example_parse_single_file()
    # example_parse_directory()
    # example_parse_multiple_files()
    # example_programmatic_usage()

    print("\n" + "=" * 50)
    print("예시 실행 완료!")
    print("=" * 50 + "\n")

    print("실제 사용법:")
    print("  - GUI: python gui.py")
    print("  - CLI: python cli.py -d ./emails")
    print("\n자세한 내용은 EMAIL_PARSER_README.md를 참고하세요.")
