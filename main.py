# main.py
import sys, os, io, re, webbrowser, tempfile, requests
from PyQt5 import QtWidgets
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from pptx import Presentation
from pptx.util import Inches

SCOPES = [
    'https://www.googleapis.com/auth/presentations.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

class SlideDownloader(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Google Slides → PPTX Downloader')
        self.resize(500, 300)
        layout = QtWidgets.QVBoxLayout()
        self.url_input = QtWidgets.QLineEdit()
        self.url_input.setPlaceholderText('Paste full Google Slides URL here')
        layout.addWidget(self.url_input)
        self.auth_btn = QtWidgets.QPushButton('Authenticate & Download')
        self.auth_btn.clicked.connect(self.handle_download)
        layout.addWidget(self.auth_btn)
        self.log = QtWidgets.QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)
        self.setLayout(layout)

    def log_msg(self, msg: str):
        self.log.append(msg)
        QtWidgets.QApplication.processEvents()

    def extract_presentation_id(self, url: str) -> str:
        m = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
        if m: return m.group(1)
        m2 = re.search(r'presentation/d/([a-zA-Z0-9_-]+)', url)
        return m2.group(1) if m2 else None

    def authenticate(self):
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        return flow.run_local_server(port=0)

    def download_slide_deck(self, pres_id: str, creds):
        slides_svc = build('slides', 'v1', credentials=creds)
        drive_svc  = build('drive', 'v3', credentials=creds)
        self.log_msg('Fetching presentation metadata...')
        pres = slides_svc.presentations().get(presentationId=pres_id).execute()
        ppt = Presentation()
        total = len(pres.get('slides', []))
        for i, slide in enumerate(pres.get('slides', []), start=1):
            self.log_msg(f'Processing slide {i}/{total}...')
            new_slide = ppt.slides.add_slide(ppt.slide_layouts[6])
            for el in slide.get('pageElements', []):
                if 'image' in el:
                    url = el['image'].get('contentUrl')
                    if url:
                        self.log_msg(f'Downloading image: {url}')
                        resp = requests.get(url); resp.raise_for_status()
                        fd, tmp = tempfile.mkstemp(suffix='.png')
                        with os.fdopen(fd, 'wb') as f: f.write(resp.content)
                        new_slide.shapes.add_picture(tmp, Inches(1), Inches(1))
                if 'video' in el:
                    src = el['video'].get('source', '')
                    if src:
                        self.log_msg(f'Found video: {src}')
                        if 'drive.google.com' in src:
                            QtWidgets.QMessageBox.information(
                                self, 'Grant Access',
                                'Please ensure this Drive video is shared publicly, then try again.'
                            )
                            webbrowser.open(src)
                        else:
                            tb = new_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(0.5))
                            tb.text_frame.text = f'Video link: {src}'
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Save as…', filter='PowerPoint Files (*.pptx)')
        if path:
            if not path.lower().endswith('.pptx'): path += '.pptx'
            ppt.save(path); self.log_msg(f'Saved to: {path}')

    def handle_download(self):
        pid = self.extract_presentation_id(self.url_input.text().strip())
        if not pid:
            return self.log_msg('Error: invalid Slides URL')
        try:
            creds = self.authenticate()
            self.download_slide_deck(pid, creds)
        except Exception as e:
            self.log_msg(f'Error: {e}')

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    SlideDownloader().show()
    sys.exit(app.exec_())

