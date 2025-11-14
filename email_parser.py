"""
Email Parser Module - .eml 및 .msg 파일을 파싱하여 이메일 스레드를 분석
"""
import email
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional
import re

# .msg 파일 파싱을 위한 라이브러리
try:
    import extract_msg
    MSG_SUPPORT = True
except ImportError:
    MSG_SUPPORT = False
    print("Warning: extract-msg not installed. .msg file support disabled.")


class EmailMessage:
    """이메일 메시지 클래스"""

    def __init__(self, msg: email.message.EmailMessage):
        self.msg = msg
        self.subject = self._get_subject()
        self.sender = self._get_sender()
        self.recipients = self._get_recipients()
        self.date = self._get_date()
        self.body = self._get_body()
        self.in_reply_to = msg.get('In-Reply-To', '')
        self.message_id = msg.get('Message-ID', '')

    def _get_subject(self) -> str:
        """제목 추출"""
        subject = self.msg.get('Subject', '')
        return str(subject).replace('\n', '').replace('\r', '')

    def _get_sender(self) -> str:
        """발신자 추출"""
        return str(self.msg.get('From', ''))

    def _get_recipients(self) -> str:
        """수신자 추출"""
        return str(self.msg.get('To', ''))

    def _get_date(self) -> Optional[datetime]:
        """날짜 추출"""
        date_str = self.msg.get('Date')
        if date_str:
            try:
                return parsedate_to_datetime(date_str)
            except:
                return None
        return None

    def _get_body(self) -> str:
        """이메일 본문 추출"""
        body = ""

        if self.msg.is_multipart():
            for part in self.msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))

                # 첨부파일 제외
                if "attachment" in content_disposition:
                    continue

                # 텍스트 본문 추출
                if content_type == "text/plain":
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            body = payload.decode('utf-8', errors='ignore')
                            break
                    except:
                        pass

                # HTML이 있고 plain text가 없으면 HTML 사용
                if not body and content_type == "text/html":
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            body = payload.decode('utf-8', errors='ignore')
                            # 간단한 HTML 태그 제거
                            body = re.sub(r'<[^>]+>', '', body)
                    except:
                        pass
        else:
            try:
                payload = self.msg.get_payload(decode=True)
                if payload:
                    body = payload.decode('utf-8', errors='ignore')
            except:
                body = str(self.msg.get_payload())

        return body.strip()

    def to_dict(self) -> Dict:
        """딕셔너리로 변환"""
        return {
            'subject': self.subject,
            'sender': self.sender,
            'recipients': self.recipients,
            'date': self.date.strftime('%Y-%m-%d %H:%M:%S') if self.date else '',
            'body': self.body,
            'message_id': self.message_id,
            'in_reply_to': self.in_reply_to
        }


class EmailThread:
    """이메일 스레드 클래스"""

    def __init__(self):
        self.messages: List[EmailMessage] = []

    def add_message(self, message: EmailMessage):
        """메시지 추가"""
        self.messages.append(message)

    def sort_by_date(self):
        """날짜순으로 정렬"""
        self.messages.sort(key=lambda x: x.date if x.date else datetime.min)

    def extract_qa_pairs(self) -> List[Dict]:
        """질의응답 쌍 추출"""
        qa_pairs = []

        self.sort_by_date()

        for i, msg in enumerate(self.messages):
            # 첫 번째 메시지를 질문으로 간주
            if i == 0:
                qa_pair = {
                    'thread_subject': msg.subject,
                    'question_sender': msg.sender,
                    'question_date': msg.date.strftime('%Y-%m-%d %H:%M:%S') if msg.date else '',
                    'question': msg.body,
                    'answer_sender': '',
                    'answer_date': '',
                    'answer': ''
                }

                # 다음 메시지가 있으면 답변으로 설정
                if i + 1 < len(self.messages):
                    next_msg = self.messages[i + 1]
                    qa_pair['answer_sender'] = next_msg.sender
                    qa_pair['answer_date'] = next_msg.date.strftime('%Y-%m-%d %H:%M:%S') if next_msg.date else ''
                    qa_pair['answer'] = next_msg.body

                qa_pairs.append(qa_pair)

            # 홀수 번째 메시지들을 추가 질문으로 처리
            elif i % 2 == 0:
                qa_pair = {
                    'thread_subject': msg.subject,
                    'question_sender': msg.sender,
                    'question_date': msg.date.strftime('%Y-%m-%d %H:%M:%S') if msg.date else '',
                    'question': msg.body,
                    'answer_sender': '',
                    'answer_date': '',
                    'answer': ''
                }

                # 다음 메시지가 있으면 답변으로 설정
                if i + 1 < len(self.messages):
                    next_msg = self.messages[i + 1]
                    qa_pair['answer_sender'] = next_msg.sender
                    qa_pair['answer_date'] = next_msg.date.strftime('%Y-%m-%d %H:%M:%S') if next_msg.date else ''
                    qa_pair['answer'] = next_msg.body

                qa_pairs.append(qa_pair)

        return qa_pairs


class EmailParser:
    """이메일 파서 메인 클래스"""

    def __init__(self):
        self.threads: Dict[str, EmailThread] = {}

    def parse_eml_file(self, file_path: str) -> EmailMessage:
        """단일 .eml 파일 파싱"""
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        with open(file_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        return EmailMessage(msg)

    def parse_msg_file(self, file_path: str) -> EmailMessage:
        """단일 .msg 파일 파싱 (Microsoft Outlook)"""
        if not MSG_SUPPORT:
            raise ImportError("extract-msg 라이브러리가 설치되지 않았습니다. 'pip install extract-msg'를 실행하세요.")

        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        # .msg 파일 파싱
        msg_obj = extract_msg.Message(str(file_path))

        # extract_msg.Message를 email.message.EmailMessage로 변환
        msg = email.message.EmailMessage()

        # 헤더 설정
        if msg_obj.subject:
            msg['Subject'] = msg_obj.subject
        if msg_obj.sender:
            msg['From'] = msg_obj.sender
        if msg_obj.to:
            msg['To'] = msg_obj.to
        if msg_obj.date:
            msg['Date'] = msg_obj.date.strftime('%a, %d %b %Y %H:%M:%S %z') if hasattr(msg_obj.date, 'strftime') else str(msg_obj.date)

        # 본문 설정
        body = msg_obj.body if msg_obj.body else ""
        msg.set_content(body)

        # 리소스 정리
        msg_obj.close()

        return EmailMessage(msg)

    def parse_file(self, file_path: str) -> EmailMessage:
        """파일 확장자에 따라 자동으로 파싱 (.eml 또는 .msg)"""
        file_path = Path(file_path)
        extension = file_path.suffix.lower()

        if extension == '.eml':
            return self.parse_eml_file(str(file_path))
        elif extension == '.msg':
            return self.parse_msg_file(str(file_path))
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {extension}")

    def parse_files(self, file_paths: List[str]) -> List[EmailMessage]:
        """여러 이메일 파일 파싱 (.eml 또는 .msg)"""
        messages = []

        for file_path in file_paths:
            try:
                msg = self.parse_file(file_path)
                messages.append(msg)
            except Exception as e:
                print(f"파일 파싱 실패 {file_path}: {str(e)}")

        return messages

    def parse_eml_files(self, file_paths: List[str]) -> List[EmailMessage]:
        """여러 .eml 파일 파싱"""
        messages = []

        for file_path in file_paths:
            try:
                msg = self.parse_eml_file(file_path)
                messages.append(msg)
            except Exception as e:
                print(f"파일 파싱 실패 {file_path}: {str(e)}")

        return messages

    def parse_directory(self, directory: str, recursive: bool = True) -> List[EmailMessage]:
        """디렉토리 내 모든 이메일 파일 파싱 (.eml 및 .msg)"""
        directory = Path(directory)

        if not directory.is_dir():
            raise NotADirectoryError(f"디렉토리가 아닙니다: {directory}")

        # .eml 및 .msg 파일 모두 검색
        if recursive:
            eml_files = list(directory.glob('**/*.eml'))
            msg_files = list(directory.glob('**/*.msg')) if MSG_SUPPORT else []
        else:
            eml_files = list(directory.glob('*.eml'))
            msg_files = list(directory.glob('*.msg')) if MSG_SUPPORT else []

        all_files = eml_files + msg_files

        return self.parse_files([str(f) for f in all_files])

    def build_threads(self, messages: List[EmailMessage]):
        """메시지들을 스레드로 구성"""
        self.threads.clear()

        for msg in messages:
            # 제목에서 Re:, Fwd: 등을 제거하여 스레드 키 생성
            subject_key = self._normalize_subject(msg.subject)

            if subject_key not in self.threads:
                self.threads[subject_key] = EmailThread()

            self.threads[subject_key].add_message(msg)

        # 각 스레드를 날짜순으로 정렬
        for thread in self.threads.values():
            thread.sort_by_date()

    def _normalize_subject(self, subject: str) -> str:
        """제목 정규화 (Re:, Fwd: 등 제거)"""
        subject = re.sub(r'^(Re:|RE:|Fwd:|FW:|Fw:)\s*', '', subject, flags=re.IGNORECASE)
        return subject.strip().lower()

    def get_all_qa_pairs(self) -> List[Dict]:
        """모든 스레드에서 질의응답 쌍 추출"""
        all_qa_pairs = []

        for subject, thread in self.threads.items():
            qa_pairs = thread.extract_qa_pairs()
            all_qa_pairs.extend(qa_pairs)

        return all_qa_pairs


if __name__ == '__main__':
    # 테스트 코드
    parser = EmailParser()

    # 단일 파일 테스트
    # msg = parser.parse_eml_file('test.eml')
    # print(msg.to_dict())

    print("Email Parser Module loaded successfully!")
