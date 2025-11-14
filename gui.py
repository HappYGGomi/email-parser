"""
GUI Interface - 그래픽 사용자 인터페이스 (tkinter)
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
import threading
import sys
from email_parser import EmailParser
from excel_exporter import ExcelExporter


class EmailParserGUI:
    """이메일 파서 GUI 클래스"""

    def __init__(self, root):
        self.root = root
        self.root.title("이메일 파서 (Email Parser)")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # 변수 초기화
        self.selected_files = []
        self.selected_directory = None
        self.parser = EmailParser()
        self.exporter = ExcelExporter()

        # UI 구성
        self.setup_ui()

    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

        # 제목
        title_label = ttk.Label(
            main_frame,
            text="📧 이메일 파서",
            font=("맑은 고딕", 18, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20))

        # 입력 방법 선택
        input_frame = ttk.LabelFrame(main_frame, text="1. 입력 방법 선택", padding="10")
        input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)

        # 파일 선택 버튼
        self.btn_select_files = ttk.Button(
            input_frame,
            text="파일 선택",
            command=self.select_files
        )
        self.btn_select_files.grid(row=0, column=0, padx=(0, 10), sticky=tk.W)

        self.label_files = ttk.Label(input_frame, text="선택된 파일 없음")
        self.label_files.grid(row=0, column=1, sticky=(tk.W, tk.E))

        # 디렉토리 선택 버튼
        self.btn_select_dir = ttk.Button(
            input_frame,
            text="디렉토리 선택",
            command=self.select_directory
        )
        self.btn_select_dir.grid(row=1, column=0, padx=(0, 10), pady=(10, 0), sticky=tk.W)

        self.label_dir = ttk.Label(input_frame, text="선택된 디렉토리 없음")
        self.label_dir.grid(row=1, column=1, pady=(10, 0), sticky=(tk.W, tk.E))

        # 옵션 설정
        options_frame = ttk.LabelFrame(main_frame, text="2. 옵션 설정", padding="10")
        options_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # 출력 형식
        self.var_export_mode = tk.StringVar(value="qa")
        ttk.Radiobutton(
            options_frame,
            text="질의응답 형식으로 저장",
            variable=self.var_export_mode,
            value="qa"
        ).grid(row=0, column=0, sticky=tk.W)

        ttk.Radiobutton(
            options_frame,
            text="원본 이메일 형식으로 저장",
            variable=self.var_export_mode,
            value="raw"
        ).grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        # 하위 디렉토리 검색
        self.var_recursive = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="하위 디렉토리 포함 검색",
            variable=self.var_recursive
        ).grid(row=2, column=0, sticky=tk.W, pady=(10, 0))

        # 실행 버튼
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, pady=(0, 10))

        self.btn_parse = ttk.Button(
            action_frame,
            text="파싱 시작",
            command=self.start_parsing,
            width=20
        )
        self.btn_parse.grid(row=0, column=0, padx=5)

        self.btn_clear = ttk.Button(
            action_frame,
            text="초기화",
            command=self.clear_all,
            width=20
        )
        self.btn_clear.grid(row=0, column=1, padx=5)

        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="3. 실행 로그", padding="10")
        log_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            width=80,
            height=20,
            font=("Consolas", 9)
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 프로그레스 바
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

    def select_files(self):
        """파일 선택 다이얼로그"""
        files = filedialog.askopenfilenames(
            title="이메일 파일 선택",
            filetypes=[
                ("Email files", "*.eml *.msg"),
                ("EML files", "*.eml"),
                ("MSG files", "*.msg"),
                ("All files", "*.*")
            ]
        )

        if files:
            self.selected_files = list(files)
            self.selected_directory = None
            count = len(self.selected_files)
            self.label_files.config(text=f"{count}개 파일 선택됨")
            self.label_dir.config(text="선택된 디렉토리 없음")
            self.log(f"✓ {count}개의 파일이 선택되었습니다.")

    def select_directory(self):
        """디렉토리 선택 다이얼로그"""
        directory = filedialog.askdirectory(title="이메일 디렉토리 선택")

        if directory:
            self.selected_directory = directory
            self.selected_files = []
            self.label_dir.config(text=directory)
            self.label_files.config(text="선택된 파일 없음")
            self.log(f"✓ 디렉토리가 선택되었습니다: {directory}")

    def log(self, message):
        """로그 메시지 출력"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        """로그 지우기"""
        self.log_text.delete(1.0, tk.END)

    def clear_all(self):
        """모든 선택 초기화"""
        self.selected_files = []
        self.selected_directory = None
        self.label_files.config(text="선택된 파일 없음")
        self.label_dir.config(text="선택된 디렉토리 없음")
        self.clear_log()
        self.log("초기화되었습니다.")

    def start_parsing(self):
        """파싱 시작"""
        # 입력 검증
        if not self.selected_files and not self.selected_directory:
            messagebox.showwarning("경고", "파일 또는 디렉토리를 선택해주세요.")
            return

        # 출력 파일 선택
        output_file = filedialog.asksaveasfilename(
            title="저장할 엑셀 파일 이름",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"email_qa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not output_file:
            return

        # 버튼 비활성화
        self.btn_parse.config(state='disabled')
        self.btn_select_files.config(state='disabled')
        self.btn_select_dir.config(state='disabled')
        self.clear_log()

        # 프로그레스 바 시작
        self.progress.start()

        # 별도 스레드에서 파싱 실행
        thread = threading.Thread(
            target=self.parse_thread,
            args=(output_file,),
            daemon=True
        )
        thread.start()

    def parse_thread(self, output_file):
        """파싱 스레드"""
        try:
            export_mode = self.var_export_mode.get()
            recursive = self.var_recursive.get()

            # 파일 수집
            if self.selected_files:
                files_to_parse = self.selected_files
                self.log(f"\n[1/4] {len(files_to_parse)}개의 파일을 파싱합니다...")
            else:
                directory = Path(self.selected_directory)
                if recursive:
                    eml_files = list(directory.glob('**/*.eml'))
                    msg_files = list(directory.glob('**/*.msg'))
                else:
                    eml_files = list(directory.glob('*.eml'))
                    msg_files = list(directory.glob('*.msg'))

                files_to_parse = eml_files + msg_files

                self.log(f"\n[1/4] 디렉토리 스캔 완료: {len(eml_files)}개의 .eml 파일, {len(msg_files)}개의 .msg 파일 발견")

                if not files_to_parse:
                    self.root.after(0, lambda: messagebox.showwarning("경고", "이메일 파일(.eml 또는 .msg)을 찾을 수 없습니다."))
                    self.parse_complete()
                    return

                files_to_parse = [str(f) for f in files_to_parse]

            # 파싱
            messages = []
            for file_path in files_to_parse:
                try:
                    msg = self.parser.parse_file(file_path)  # .eml 및 .msg 모두 지원
                    messages.append(msg)
                    self.log(f"  ✓ {Path(file_path).name}")
                except Exception as e:
                    self.log(f"  ✗ {Path(file_path).name}: {str(e)}")

            if not messages:
                self.root.after(0, lambda: messagebox.showerror("오류", "파싱된 이메일이 없습니다."))
                self.parse_complete()
                return

            self.log(f"\n[2/4] 총 {len(messages)}개의 이메일을 파싱했습니다.")

            # 스레드 구성
            self.log(f"\n[3/4] 이메일 스레드 구성 중...")
            self.parser.build_threads(messages)
            self.log(f"  ✓ {len(self.parser.threads)}개의 스레드를 발견했습니다.")

            # 엑셀 저장
            self.log(f"\n[4/4] 엑셀 파일 생성 중...")

            if export_mode == "raw":
                # 원본 이메일 저장
                raw_messages = [msg.to_dict() for msg in messages]
                self.exporter.export_raw_emails(raw_messages, output_file)
                self.log(f"  ✓ 원본 이메일 {len(messages)}개를 저장했습니다.")
            else:
                # 질의응답 저장
                qa_pairs = self.parser.get_all_qa_pairs()
                self.log(f"  ✓ {len(qa_pairs)}개의 질의응답 쌍을 추출했습니다.")
                self.exporter.export_qa_pairs(qa_pairs, output_file)

            self.log(f"\n✓ 완료! 파일이 저장되었습니다:\n  {output_file}")

            # 성공 메시지
            self.root.after(0, lambda: messagebox.showinfo("완료", f"파싱이 완료되었습니다!\n\n{output_file}"))

        except Exception as e:
            error_msg = f"오류 발생: {str(e)}"
            self.log(f"\n✗ {error_msg}")
            import traceback
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("오류", error_msg))

        finally:
            self.parse_complete()

    def parse_complete(self):
        """파싱 완료 후 처리"""
        # 프로그레스 바 정지
        self.progress.stop()

        # 버튼 활성화
        self.btn_parse.config(state='normal')
        self.btn_select_files.config(state='normal')
        self.btn_select_dir.config(state='normal')


def main():
    """메인 함수"""
    root = tk.Tk()
    app = EmailParserGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
