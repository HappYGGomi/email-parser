"""
CLI Interface - 명령줄 인터페이스
"""
import argparse
import sys
from pathlib import Path
from datetime import datetime
from email_parser import EmailParser
from excel_exporter import ExcelExporter


def print_banner():
    """배너 출력"""
    banner = """
    ╔═══════════════════════════════════════════╗
    ║     이메일 파서 (Email Parser CLI)        ║
    ║   .eml/.msg 파일을 엑셀로 변환해드립니다  ║
    ╚═══════════════════════════════════════════╝
    """
    print(banner)


def parse_files(file_paths, output_file=None, export_raw=False):
    """파일 파싱 및 엑셀 저장"""
    try:
        print(f"\n[1/4] 이메일 파일 파싱 중...")
        parser = EmailParser()

        # 파일 파싱
        messages = []
        for file_path in file_paths:
            try:
                msg = parser.parse_file(file_path)  # .eml 및 .msg 모두 지원
                messages.append(msg)
                print(f"  ✓ {file_path}")
            except Exception as e:
                print(f"  ✗ {file_path}: {str(e)}")

        if not messages:
            print("\n오류: 파싱된 이메일이 없습니다.")
            return False

        print(f"\n[2/4] 총 {len(messages)}개의 이메일을 파싱했습니다.")

        # 스레드 구성
        print(f"\n[3/4] 이메일 스레드 구성 중...")
        parser.build_threads(messages)
        print(f"  ✓ {len(parser.threads)}개의 스레드를 발견했습니다.")

        # 질의응답 추출
        qa_pairs = parser.get_all_qa_pairs()
        print(f"  ✓ {len(qa_pairs)}개의 질의응답 쌍을 추출했습니다.")

        # 엑셀 저장
        print(f"\n[4/4] 엑셀 파일 생성 중...")
        exporter = ExcelExporter()

        # 출력 파일명 설정
        if not output_file:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f'email_qa_{timestamp}.xlsx'

        # 저장
        if export_raw:
            # 원본 이메일 저장
            raw_messages = [msg.to_dict() for msg in messages]
            exporter.export_raw_emails(raw_messages, output_file)
        else:
            # 질의응답 저장
            exporter.export_qa_pairs(qa_pairs, output_file)

        print(f"  ✓ 엑셀 파일이 저장되었습니다: {output_file}")
        print(f"\n완료!")

        return True

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def parse_directory(directory, output_file=None, export_raw=False, recursive=True):
    """디렉토리 파싱"""
    try:
        print(f"\n[1/4] 디렉토리 스캔 중: {directory}")
        parser = EmailParser()

        directory_path = Path(directory)
        if not directory_path.is_dir():
            print(f"오류: 디렉토리가 아닙니다: {directory}")
            return False

        # .eml 및 .msg 파일 찾기
        if recursive:
            eml_files = list(directory_path.glob('**/*.eml'))
            msg_files = list(directory_path.glob('**/*.msg'))
        else:
            eml_files = list(directory_path.glob('*.eml'))
            msg_files = list(directory_path.glob('*.msg'))

        all_files = eml_files + msg_files

        if not all_files:
            print(f"오류: 이메일 파일(.eml 또는 .msg)을 찾을 수 없습니다.")
            return False

        print(f"  ✓ {len(eml_files)}개의 .eml 파일, {len(msg_files)}개의 .msg 파일을 발견했습니다.")

        # 파일 경로 리스트로 변환
        file_paths = [str(f) for f in all_files]

        # 파일 파싱 수행
        return parse_files(file_paths, output_file, export_raw)

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """메인 함수"""
    print_banner()

    parser = argparse.ArgumentParser(
        description='이메일 파서 - .eml 및 .msg 파일을 파싱하여 엑셀로 변환',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  # 단일 파일 파싱
  python cli.py -f email1.eml
  python cli.py -f email1.msg

  # 여러 파일 파싱 (.eml과 .msg 혼합 가능)
  python cli.py -f email1.eml email2.msg email3.eml

  # 디렉토리 파싱 (하위 디렉토리 포함, .eml과 .msg 모두)
  python cli.py -d ./emails

  # 디렉토리 파싱 (하위 디렉토리 제외)
  python cli.py -d ./emails --no-recursive

  # 출력 파일명 지정
  python cli.py -d ./emails -o output.xlsx

  # 원본 이메일 내보내기
  python cli.py -d ./emails --raw
        """
    )

    parser.add_argument('-f', '--files', nargs='+', help='.eml 또는 .msg 파일 경로 (하나 이상)')
    parser.add_argument('-d', '--directory', help='이메일 파일(.eml/.msg)이 있는 디렉토리 경로')
    parser.add_argument('-o', '--output', help='출력 엑셀 파일명 (기본값: email_qa_YYYYMMDD_HHMMSS.xlsx)')
    parser.add_argument('--raw', action='store_true', help='원본 이메일 형식으로 저장 (질의응답 추출 안함)')
    parser.add_argument('--no-recursive', action='store_true', help='하위 디렉토리 검색 안함')

    args = parser.parse_args()

    # 파일 또는 디렉토리 중 하나는 필수
    if not args.files and not args.directory:
        parser.print_help()
        print("\n오류: -f 또는 -d 옵션 중 하나를 지정해야 합니다.")
        sys.exit(1)

    # 파일 파싱
    if args.files:
        success = parse_files(args.files, args.output, args.raw)
    # 디렉토리 파싱
    elif args.directory:
        recursive = not args.no_recursive
        success = parse_directory(args.directory, args.output, args.raw, recursive)

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
