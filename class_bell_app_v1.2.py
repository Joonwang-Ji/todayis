import sys
import getpass
import subprocess
import json
import os
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime, timedelta, timezone
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox,
                             QLabel, QTabWidget, QSystemTrayIcon, QMenu, QCheckBox, QAction,
                             QDialog, QFormLayout, QComboBox, QInputDialog, QHeaderView, QLineEdit,
                             QGridLayout, QSlider, QSpacerItem, QItemDelegate, QScrollArea)
from PyQt5.QtCore import QTimer, QTime, Qt, QRegExp, QEvent, QSharedMemory, QUrl, QThread, pyqtSignal
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon, QGuiApplication, QRegExpValidator, QPixmap
import pytz
import requests
import time

# 로그 파일 설정
app_data_dir = Path(os.getenv('APPDATA')) / '오늘은'
app_data_dir.mkdir(exist_ok=True)
log_file = app_data_dir / 'app.log'
config_file = app_data_dir / 'config.json'

handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=10, encoding="utf-8")
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.handlers.clear()
logger.addHandler(handler)

def resource_path(relative_path):
    """리소스 파일 경로 반환"""
    if hasattr(sys, '_MEIPASS'):
        path = os.path.join(sys._MEIPASS, relative_path)
    else:
        path = os.path.join(os.path.abspath("."), relative_path)
    if not os.path.exists(path):
        logger.error(f"Resource not found: {path}")
        QMessageBox.warning(None, "리소스 오류", f"필수 파일을 찾을 수 없습니다: {relative_path}")
    return path

class ConfigLoader(QThread):
    configLoaded = pyqtSignal(dict)

    def run(self):
        config = {}
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
        self.configLoaded.emit(config)

class TimeDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        reg_exp = QRegExp("[0-9:]*")
        validator = QRegExpValidator(reg_exp, editor)
        editor.setValidator(validator)
        editor.setMaxLength(5)
        editor.textChanged.connect(lambda text: self.adjust_max_length(editor, text))
        return editor

    def adjust_max_length(self, editor, text):
        if ":" in text:
            editor.setMaxLength(5)
        else:
            editor.setMaxLength(4)

class BellDialog(QDialog):
    def __init__(self, start_bell, end_bell, misc_bell, parent=None):
        super().__init__(parent)
        self.setWindowTitle("벨소리 설정")
        self.setFixedSize(400, 200)
        self.start_bell_input = QLineEdit(self)
        self.end_bell_input = QLineEdit(self)
        self.misc_bell_input = QLineEdit(self)  # 추가
        self.start_bell_input.setText(start_bell)
        self.end_bell_input.setText(end_bell)
        self.misc_bell_input.setText(misc_bell)  # 추가

        layout = QVBoxLayout()
        start_row = QHBoxLayout()
        start_label = QLabel("수업 시작 벨소리:", self)
        start_label.setFixedWidth(100)
        start_row.addWidget(start_label)
        start_row.addWidget(self.start_bell_input)
        start_button = QPushButton("변경", self)
        start_button.setFixedWidth(50)
        start_button.clicked.connect(self.choose_start_bell)
        start_row.addWidget(start_button)
        layout.addLayout(start_row)

        end_row = QHBoxLayout()
        end_label = QLabel("수업 종료 벨소리:", self)
        end_label.setFixedWidth(100)
        end_row.addWidget(end_label)
        end_row.addWidget(self.end_bell_input)
        end_button = QPushButton("변경", self)
        end_button.setFixedWidth(50)
        end_button.clicked.connect(self.choose_end_bell)
        end_row.addWidget(end_button)
        layout.addLayout(end_row)

        # "기타 시작" 설정 추가
        misc_row = QHBoxLayout()
        misc_label = QLabel("기타 시작 벨소리:", self)
        misc_label.setFixedWidth(100)
        misc_row.addWidget(misc_label)
        misc_row.addWidget(self.misc_bell_input)
        misc_button = QPushButton("변경", self)
        misc_button.setFixedWidth(50)
        misc_button.clicked.connect(self.choose_misc_bell)
        misc_row.addWidget(misc_button)
        layout.addLayout(misc_row)

        action_layout = QHBoxLayout()
        action_layout.addStretch()
        ok_button = QPushButton("확인", self)
        cancel_button = QPushButton("취소", self)
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        action_layout.addWidget(ok_button)
        action_layout.addWidget(cancel_button)
        layout.addLayout(action_layout)
        self.setLayout(layout)

    def choose_start_bell(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "시작 벨소리 선택", "", "Audio Files (*.mp3 *.wav)")
        if file_name:
            self.start_bell_input.setText(file_name)

    def choose_end_bell(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "종료 벨소리 선택", "", "Audio Files (*.mp3 *.wav)")
        if file_name:
            self.end_bell_input.setText(file_name)

    def choose_misc_bell(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "기타 시작 벨소리 선택", "", "Audio Files (*.mp3 *.wav)")
        if file_name:
            self.misc_bell_input.setText(file_name)

    def get_bells(self):
        return (self.start_bell_input.text().strip(),
                self.end_bell_input.text().strip(),
                self.misc_bell_input.text().strip())  # 반환값에 misc_bell 추가

class RepeatScheduleDialog(QDialog):
    def __init__(self, days, parent=None):
        super().__init__(parent)
        self.setWindowTitle("반복 스케줄 생성")
        self.days = days
        layout = QFormLayout()

        self.start_hour = QComboBox()
        self.start_hour.addItems([f"{h:02d}" for h in range(24)])
        self.start_hour.setCurrentText("09")
        self.start_minute = QComboBox()
        self.start_minute.addItems([f"{m:02d}" for m in range(0, 60, 5)])
        self.start_minute.setCurrentText("00")
        start_time_layout = QHBoxLayout()
        start_time_layout.addWidget(self.start_hour)
        start_time_layout.addWidget(QLabel(":"))
        start_time_layout.addWidget(self.start_minute)
        layout.addRow("수업시작 시간 (HH:MM):", start_time_layout)

        self.duration_input = QComboBox()
        self.duration_input.addItems([str(m) for m in range(0, 121, 5)])
        self.duration_input.setCurrentText("60")
        layout.addRow("수업시간 (분):", self.duration_input)

        self.break_input = QComboBox()
        self.break_input.addItems([str(m) for m in range(0, 61, 5)])
        self.break_input.setCurrentText("10")
        layout.addRow("쉬는시간 (분):", self.break_input)

        self.repeat_count_input = QComboBox()
        self.repeat_count_input.addItems([str(i) for i in range(1, 21)])
        self.repeat_count_input.setCurrentText("1")
        layout.addRow("반복 횟수:", self.repeat_count_input)

        self.day_checkboxes = {}
        day_layout = QVBoxLayout()
        for day in self.days:
            checkbox = QCheckBox(day)
            self.day_checkboxes[day] = checkbox
            day_layout.addWidget(checkbox)
        layout.addRow("요일 선택:", day_layout)

        self.apply_btn = QPushButton("적용")
        self.apply_btn.clicked.connect(self.accept)
        layout.addWidget(self.apply_btn)
        self.setLayout(layout)

    def get_schedule_data(self):
        selected_days = [day for day, checkbox in self.day_checkboxes.items() if checkbox.isChecked()]
        if not selected_days:
            selected_days = [self.days[0]]
        return {
            "start_time": f"{self.start_hour.currentText()}:{self.start_minute.currentText()}",
            "duration": int(self.duration_input.currentText()),
            "break_time": int(self.break_input.currentText()),
            "repeat_count": int(self.repeat_count_input.currentText()),
            "days": selected_days
        }

class SupportQRDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("카카오페이 후원 QR")
        self.setFixedSize(300, 300)  # QR 이미지 크기에 맞게 조정
        layout = QVBoxLayout()

        # QR 이미지 로드
        qr_path = resource_path("todayis_qr.jpeg")
        if os.path.exists(qr_path):
            qr_image = QPixmap(qr_path).scaled(250, 250, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            qr_label = QLabel()
            qr_label.setPixmap(qr_image)
            qr_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(qr_label)
        else:
            error_label = QLabel("QR 이미지를 찾을 수 없습니다.")
            error_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(error_label)

        # 닫기 버튼
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.close)
        close_btn.setFixedWidth(100)
        layout.addWidget(close_btn, alignment=Qt.AlignCenter)

        self.setLayout(layout)

class ClassBellApp(QMainWindow):
    city_map = {
        "Seoul": {"name": "서울특별시", "nx": 60, "ny": 127},
        "Busan": {"name": "부산광역시", "nx": 98, "ny": 76},
        "Daegu": {"name": "대구광역시", "nx": 89, "ny": 90},
        "Incheon": {"name": "인천광역시", "nx": 55, "ny": 124},
        "Gwangju": {"name": "광주광역시", "nx": 58, "ny": 74},
        "Daejeon": {"name": "대전광역시", "nx": 67, "ny": 100},
        "Ulsan": {"name": "울산광역시", "nx": 102, "ny": 84},
        "Sejong": {"name": "세종특별자치시", "nx": 66, "ny": 103},
        "Suwon": {"name": "수원시", "nx": 60, "ny": 121},
        "Yongin": {"name": "용인시", "nx": 61, "ny": 120},
        "Seongnam": {"name": "성남시", "nx": 61, "ny": 123},
        "Bucheon": {"name": "부천시", "nx": 57, "ny": 125},
        "Hwaseong": {"name": "화성시", "nx": 58, "ny": 121},
        "Ansan": {"name": "안산시", "nx": 58, "ny": 123},
        "Anyang": {"name": "안양시", "nx": 59, "ny": 124},
        "Pyeongtaek": {"name": "평택시", "nx": 62, "ny": 114},
        "Siheung": {"name": "시흥시", "nx": 57, "ny": 123},
        "Gimpo": {"name": "김포시", "nx": 56, "ny": 126},
        "Gwangju-si": {"name": "광주시", "nx": 61, "ny": 126},
        "Goyang": {"name": "고양시", "nx": 57, "ny": 127},
        "Namyangju": {"name": "남양주시", "nx": 62, "ny": 127},
        "Paju": {"name": "파주시", "nx": 56, "ny": 130},
        "Uijeongbu": {"name": "의정부시", "nx": 61, "ny": 129},
        "Yangju": {"name": "양주시", "nx": 61, "ny": 130},
        "Icheon": {"name": "이천시", "nx": 68, "ny": 121},
        "Osan": {"name": "오산시", "nx": 62, "ny": 118},
        "Anseong": {"name": "안성시", "nx": 65, "ny": 116},
        "Gunpo": {"name": "군포시", "nx": 59, "ny": 123},
        "Uiwang": {"name": "의왕시", "nx": 60, "ny": 123},
        "Hanam": {"name": "하남시", "nx": 63, "ny": 126},
        "Yeoju": {"name": "여주시", "nx": 69, "ny": 123},
        "Dongducheon": {"name": "동두천시", "nx": 61, "ny": 132},
        "Gwacheon": {"name": "과천시", "nx": 60, "ny": 124},
        "Guri": {"name": "구리시", "nx": 62, "ny": 127},
        "Pocheon": {"name": "포천시", "nx": 64, "ny": 134},
        "Chuncheon": {"name": "춘천시", "nx": 73, "ny": 134},
        "Wonju": {"name": "원주시", "nx": 76, "ny": 121},
        "Gangneung": {"name": "강릉시", "nx": 92, "ny": 131},
        "Donghae": {"name": "동해시", "nx": 95, "ny": 127},
        "Taebaek": {"name": "태백시", "nx": 95, "ny": 119},
        "Sokcho": {"name": "속초시", "nx": 87, "ny": 138},
        "Samcheok": {"name": "삼척시", "nx": 96, "ny": 125},
        "Cheongju": {"name": "청주시", "nx": 69, "ny": 106},
        "Chungju": {"name": "충주시", "nx": 76, "ny": 114},
        "Jecheon": {"name": "제천시", "nx": 81, "ny": 118},
        "Cheonan": {"name": "천안시", "nx": 63, "ny": 110},
        "Gongju": {"name": "공주시", "nx": 63, "ny": 102},
        "Boryeong": {"name": "보령시", "nx": 54, "ny": 100},
        "Asan": {"name": "아산시", "nx": 60, "ny": 110},
        "Seosan": {"name": "서산시", "nx": 54, "ny": 110},
        "Nonsan": {"name": "논산시", "nx": 62, "ny": 97},
        "Gyeryong": {"name": "계룡시", "nx": 63, "ny": 99},
        "Dangjin": {"name": "당진시", "nx": 54, "ny": 112},
        "Jeonju": {"name": "전주시", "nx": 63, "ny": 89},
        "Gunsan": {"name": "군산시", "nx": 56, "ny": 92},
        "Iksan": {"name": "익산시", "nx": 60, "ny": 91},
        "Jeongeup": {"name": "정읍시", "nx": 58, "ny": 83},
        "Namwon": {"name": "남원시", "nx": 68, "ny": 80},
        "Gimje": {"name": "김제시", "nx": 59, "ny": 88},
        "Mokpo": {"name": "목포시", "nx": 50, "ny": 67},
        "Yeosu": {"name": "여수시", "nx": 73, "ny": 66},
        "Suncheon": {"name": "순천시", "nx": 70, "ny": 70},
        "Naju": {"name": "나주시", "nx": 56, "ny": 71},
        "Gwangyang": {"name": "광양시", "nx": 73, "ny": 70},
        "Pohang": {"name": "포항시", "nx": 102, "ny": 94},
        "Gyeongju": {"name": "경주시", "nx": 100, "ny": 89},
        "Gimcheon": {"name": "김천시", "nx": 81, "ny": 96},
        "Andong": {"name": "안동시", "nx": 91, "ny": 106},
        "Gumi": {"name": "구미시", "nx": 84, "ny": 96},
        "Yeongju": {"name": "영주시", "nx": 89, "ny": 111},
        "Yeongcheon": {"name": "영천시", "nx": 95, "ny": 93},
        "Sangju": {"name": "상주시", "nx": 81, "ny": 103},
        "Mungyeong": {"name": "문경시", "nx": 81, "ny": 107},
        "Gyeongsan": {"name": "경산시", "nx": 91, "ny": 90},
        "Changwon": {"name": "창원시", "nx": 90, "ny": 77},
        "Jinju": {"name": "진주시", "nx": 81, "ny": 75},
        "Tongyeong": {"name": "통영시", "nx": 87, "ny": 68},
        "Sacheon": {"name": "사천시", "nx": 80, "ny": 71},
        "Gimhae": {"name": "김해시", "nx": 95, "ny": 77},
        "Miryang": {"name": "밀양시", "nx": 92, "ny": 83},
        "Geoje": {"name": "거제시", "nx": 90, "ny": 70},
        "Yangsan": {"name": "양산시", "nx": 97, "ny": 79},
        "Jeju": {"name": "제주시", "nx": 52, "ny": 38},
        "Seogwipo": {"name": "서귀포시", "nx": 52, "ny": 33}
    }

    def __init__(self):
        self.shared_mem = QSharedMemory("ClassBellAppUniqueKey")
        if not self.shared_mem.create(1):
            QMessageBox.information(None, "오늘은 v1.2", "프로그램이 이미 실행 중입니다.")
            sys.exit(0)

        super().__init__()
        self.setWindowTitle("오늘은 v1.2")
        self.setFixedSize(800, 700)
        self.setWindowIcon(QIcon(resource_path("icon.ico")))

        screen = QGuiApplication.primaryScreen().geometry()
        self.move((screen.width() - 800) // 2, (screen.height() - 700) // 2)

        self.copied_data = []  # 복사된 일정을 저장할 리스트
        self.days = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
        self.schedule_lists = {"기본 스케줄": {day: [] for day in self.days}}
        self.start_bell = resource_path("start_bell.wav")
        self.end_bell = resource_path("end_bell.wav")
        self.misc_bell = resource_path("start_bell.wav")  # "기타 시작" 기본값은 "수업 시작"과 동일
        # 수정: 현재 시간 직전으로 초기화 (동일 분 내 중복 방지)
        current_time = QTime.currentTime().toString("HH:mm")
        self.last_played_time = current_time  # None 대신 현재 시간으로 설정
        

        self.is_inserting = False  # 삽입 상태 플래그 추가
        self.is_first_run = True  # 최초 실행 플래그
        self.last_realtime_update = None  # 갱신 시간 추적
        self.status_timer = None  # 타이머 인스턴스 초기화 (상태메시지 용)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_time)
        self.timer.start(1000)

        self.daily_timer = QTimer(self)
        self.daily_timer.timeout.connect(self.update_today_schedule)
        self.daily_timer.start(24 * 60 * 60 * 1000)

        self.seoul_tz = pytz.timezone("Asia/Seoul")
        self.player = QMediaPlayer()
        self.player.stateChanged.connect(self.on_player_state_changed)

        self.tray_icon = QSystemTrayIcon(QIcon(resource_path("icon.ico")), self)
        self.tray_icon.setToolTip("오늘은 v1.2")
        tray_menu = QMenu()
        tray_menu.addAction("보이기", self.show)
        tray_menu.addAction("종료하기", self.exit_program)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        self.tray_icon.activated.connect(self.on_tray_icon_activated)

        self.weather_data = "날씨 정보를 불러오는 중..."
        self.city_name = None
        self.temp = None
        self.temp_min = None
        self.temp_max = None
        self.weather_desc = None
        self.icon = None

        # 예보 날씨 타이머 (24시간마다)
        self.weather_timer_long = QTimer(self)
        self.weather_timer_long.timeout.connect(self.update_forecast_weather)
        self.weather_timer_long.start(24 * 60 * 60 * 1000)

        # 실시간 날씨 타이머 (매 정시 갱신)
        self.weather_timer_short = QTimer(self)
        self.weather_timer_short.timeout.connect(self.update_realtime_weather)
        now = datetime.now(self.seoul_tz)
        next_hour = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
        initial_delay = int((next_hour - now).total_seconds() * 1000)


        self.load_config()
        self.setup_ui()
        self.apply_style()
        self.edit_table.installEventFilter(self)

        # 최초 날씨 갱신 및 타이머 설정
        self.update_realtime_weather()  # 최초 실행 시 즉시 갱신
        QTimer.singleShot(initial_delay, lambda: [
            self.weather_timer_short.start(60 * 60 * 1000),  # 매시간 갱신 시작
            self.update_realtime_weather(),  # 첫 정각에 즉시 갱신
            logger.info(f"Realtime weather timer started at {datetime.now(self.seoul_tz).strftime('%H:%M')}")
        ])
        self.update_forecast_weather()


        # QAction으로 단축키 정의
        self.action_start = QAction("수업시작 설정", self)
        self.action_start.setShortcut(Qt.Key_S)
        self.action_start.triggered.connect(lambda: self.set_schedule_type(self.edit_table.selectedItems(),
                                                                           "수업시작") if self.edit_table.selectedItems() and any(
            item.text().strip() for item in self.edit_table.selectedItems()) else self.set_status_message(
            "빈 셀에서는 동작하지 않습니다."))
        self.addAction(self.action_start)

        self.action_end = QAction("종료 설정", self)
        self.action_end.setShortcut(Qt.Key_D)
        self.action_end.triggered.connect(lambda: self.set_schedule_type(self.edit_table.selectedItems(),
                                                                         "종료") if self.edit_table.selectedItems() and any(
            item.text().strip() for item in self.edit_table.selectedItems()) else self.set_status_message(
            "빈 셀에서는 동작하지 않습니다."))
        self.addAction(self.action_end)

        self.action_misc = QAction("기타시작 설정", self)
        self.action_misc.setShortcut(Qt.Key_E)
        self.action_misc.triggered.connect(lambda: self.set_schedule_type(self.edit_table.selectedItems(),
                                                                          "기타시작") if self.edit_table.selectedItems() and any(
            item.text().strip() for item in self.edit_table.selectedItems()) else self.set_status_message(
            "빈 셀에서는 동작하지 않습니다."))
        self.addAction(self.action_misc)

        self.action_insert = QAction("스케줄 삽입", self)
        self.action_insert.setShortcut(Qt.Key_I)
        self.action_insert.triggered.connect(self.add_below_schedule)  # 삽입은 빈 셀에서도 동작 허용
        self.addAction(self.action_insert)

        self.update_edit_table()
        self.update_today_schedule()  # 추가: 초기화 시 오늘의 시간표 즉시 업데이트
        self.show()

    def fetch_forecast_weather(self):
        cache_file = app_data_dir / "forecast_cache.json"
        cache_duration = timedelta(hours=24)
        now = datetime.now(self.seoul_tz)

        # 캐시 파일이 존재하는 경우
        if os.path.exists(cache_file):
            try:
                with open(cache_file, "r", encoding="utf-8") as f:
                    cache_data = json.load(f)
                last_updated = datetime.fromisoformat(cache_data["timestamp"])
                if last_updated.tzinfo is None:
                    last_updated = self.seoul_tz.localize(last_updated)
                cache_date = last_updated.date()
                current_date = now.date()

                # 날짜가 다르거나 24시간 초과 시 새 데이터 가져오기
                if (now - last_updated < cache_duration) and (cache_date == current_date):
                    self.temp_min = cache_data["temp_min"]
                    self.temp_max = cache_data["temp_max"]
                    self.city_name = cache_data.get("city_name", "서울특별시")
                    self.update_weather_data()
                    return
                else:
                    logger.info("Cache expired or date changed, fetching new forecast data")
            except Exception as e:
                logger.error(f"Forecast cache read error: {str(e)}")

        # API 호출 및 새 데이터 가져오기
        api_key = "UVBfVbhomHosY6RfYywnTw3LYQ3IoKWeDgIpcEM%2Fs3zqYABIXGRSMEggQ37qCVDaWgRwasS5GSpDzYkV17zTtQ%3D%3D"
        retries = 3
        base_time_options = ["2300", "2000", "1700", "1400", "1100", "0800", "0500", "0200"]

        for attempt in range(retries):
            try:
                ip_response = requests.get("https://ipinfo.io/json", timeout=5)
                ip_data = ip_response.json()
                city_from_ip = ip_data.get("city", "Seoul")
                city_info = self.city_map.get(city_from_ip, {"name": "서울특별시", "nx": 60, "ny": 127})
                self.city_name, nx, ny = city_info["name"], city_info["nx"], city_info["ny"]

                base_date = now.strftime("%Y%m%d")
                base_time_str = "0500"
                current_hour = now.hour
                for base_time in base_time_options:
                    if int(base_time[:2]) <= current_hour or (current_hour < 2 and base_time == "2300"):
                        base_time_str = base_time
                        break
                if int(base_time_str[:2]) > current_hour:
                    base_date = (now - timedelta(days=1)).strftime("%Y%m%d")

                url_fcst = f"http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getVilageFcst?serviceKey={api_key}&numOfRows=1000&pageNo=1&base_date={base_date}&base_time={base_time_str}&nx={nx}&ny={ny}&dataType=JSON"
                response_fcst = requests.get(url_fcst, timeout=15)
                response_fcst.raise_for_status()
                data_fcst = response_fcst.json()
                if data_fcst["response"]["header"]["resultCode"] != "00":
                    raise ValueError(f"Forecast API Error: {data_fcst['response']['header']['resultMsg']}")

                all_temps = []
                for item in data_fcst["response"]["body"]["items"]["item"]:
                    if item["fcstDate"] == base_date and item["category"] == "TMP":
                        all_temps.append(float(item["fcstValue"]))
                if not all_temps:
                    raise ValueError("No TMP data in forecast response")
                self.temp_min, self.temp_max = min(all_temps), max(all_temps)

                cache_data = {
                    "timestamp": now.isoformat(),
                    "temp_min": self.temp_min,
                    "temp_max": self.temp_max,
                    "city_name": self.city_name
                }
                with open(cache_file, "w", encoding="utf-8") as f:
                    json.dump(cache_data, f, ensure_ascii=False)
                self.update_weather_data()
                break
            except (requests.RequestException, ValueError) as e:
                logger.error(f"Forecast fetch attempt {attempt + 1} failed: {str(e)}")
                if attempt < retries - 1:
                    time.sleep(5)
                elif attempt == retries - 1:
                    self.weather_data = "⚠️ 예보 날씨 정보를 가져올 수 없습니다."


    def fetch_realtime_weather(self):
        cache_file = app_data_dir / "realtime_cache.json"
        cache_duration = timedelta(hours=1)
        now = datetime.now(self.seoul_tz)

        if os.path.exists(cache_file):
            try:
                with open(cache_file, "r", encoding="utf-8") as f:
                    cache_data = json.load(f)
                last_updated_naive = datetime.fromisoformat(cache_data["timestamp"])
                if last_updated_naive.tzinfo is None:
                    last_updated = self.seoul_tz.localize(last_updated_naive)
                else:
                    last_updated = last_updated_naive
                time_diff = now - last_updated
                if time_diff < cache_duration:
                    self.temp = cache_data["temp"]
                    self.weather_desc = cache_data["weather_desc"]
                    self.icon = cache_data["icon"]
                    self.city_name = cache_data.get("city_name", "서울특별시")
                    self.update_weather_data()
                    logger.debug(f"Using cached realtime weather, age: {time_diff.total_seconds() / 60:.1f} minutes")
                    return
                else:
                    logger.info("Realtime cache expired, fetching new data")
            except Exception as e:
                logger.error(f"Realtime cache read error: {str(e)}")

        api_key = "UVBfVbhomHosY6RfYywnTw3LYQ3IoKWeDgIpcEM%2Fs3zqYABIXGRSMEggQ37qCVDaWgRwasS5GSpDzYkV17zTtQ%3D%3D"
        retries = 3
        for attempt in range(retries):
            try:
                if not self.city_name or self.city_name not in self.city_map:
                    ip_response = requests.get("https://ipinfo.io/json", timeout=5)
                    ip_data = ip_response.json()
                    city_from_ip = ip_data.get("city", "Seoul")
                    city_info = self.city_map.get(city_from_ip, {"name": "서울특별시", "nx": 60, "ny": 127})
                    self.city_name, nx, ny = city_info["name"], city_info["nx"], city_info["ny"]
                else:
                    nx, ny = self.city_map[self.city_name]["nx"], self.city_map[self.city_name]["ny"]

                now = datetime.now()
                real_time = now
                real_date = real_time.strftime("%Y%m%d")
                real_hour = real_time.strftime("%H00")
                if int(real_time.strftime("%M")) < 45:
                    real_time = real_time - timedelta(hours=1)
                    real_date = real_time.strftime("%Y%m%d")
                    real_hour = real_time.strftime("%H00")

                url_real = f"http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getUltraSrtNcst?serviceKey={api_key}&numOfRows=10&pageNo=1&base_date={real_date}&base_time={real_hour}&nx={nx}&ny={ny}&dataType=JSON"
                response_real = requests.get(url_real, timeout=15)
                response_real.raise_for_status()
                data_real = response_real.json()
                if data_real["response"]["header"]["resultCode"] != "00":
                    raise ValueError(f"Realtime API Error: {data_real['response']['header']['resultMsg']}")

                items_real = data_real["response"]["body"]["items"]["item"]
                temp = pty = r06 = wsd = reh = None
                for item in items_real:
                    if item["category"] == "T1H": temp = float(item["obsrValue"])
                    elif item["category"] == "PTY": pty = item["obsrValue"]
                    elif item["category"] == "RN1": r06 = float(item["obsrValue"]) if item["obsrValue"] != "강수없음" else 0
                    elif item["category"] == "WSD": wsd = float(item["obsrValue"])
                    elif item["category"] == "REH": reh = float(item["obsrValue"])

                if temp is None:
                    raise ValueError("No T1H data in realtime response")

                pty_map = {"0": "맑음", "1": "비", "2": "비/눈", "3": "눈", "4": "소나기"}
                base_desc = pty_map.get(pty, "맑음")
                weather_desc = base_desc
                if pty == "1" and r06:
                    weather_desc = "이슬비" if r06 < 1 else "약한 비" if r06 < 5 else "강한 비" if r06 > 20 else "비"
                elif pty == "4" and r06:
                    weather_desc = "약한 소나기" if r06 < 5 else "강한 소나기" if r06 > 20 else "소나기"
                if wsd:
                    if wsd > 14: weather_desc += " (강풍)"
                    elif pty == "0" and wsd > 7: weather_desc = f"바람부는 {weather_desc}"
                if reh and reh > 90 and wsd and wsd < 3 and pty == "0":
                    weather_desc = "안개" if base_desc == "맑음" else f"{base_desc} 속 안개"

                icon = "☀️" if "맑음" in weather_desc and "안개" not in weather_desc else "🌧️" if "비" in weather_desc or "소나기" in weather_desc else "❄️" if "눈" in weather_desc else "🌫️" if "안개" in weather_desc else "💨" if "바람" in weather_desc or "강풍" in weather_desc else "⚠️"

                self.temp = temp
                self.weather_desc = weather_desc
                self.icon = icon

                # 최고/최저 온도 동적 갱신 및 캐시 업데이트
                forecast_cache_file = app_data_dir / "forecast_cache.json"
                if os.path.exists(forecast_cache_file):
                    with open(forecast_cache_file, "r", encoding="utf-8") as f:
                        forecast_data = json.load(f)
                    updated = False
                    if self.temp > forecast_data["temp_max"]:
                        forecast_data["temp_max"] = self.temp
                        updated = True
                    if self.temp < forecast_data["temp_min"]:
                        forecast_data["temp_min"] = self.temp
                        updated = True
                    if updated:
                        forecast_data["timestamp"] = datetime.now().isoformat()
                        with open(forecast_cache_file, "w", encoding="utf-8") as f:
                            json.dump(forecast_data, f, ensure_ascii=False)
                        self.temp_min = forecast_data["temp_min"]
                        self.temp_max = forecast_data["temp_max"]

                # 캐시 저장
                cache_data = {
                    "timestamp": now.isoformat(),  # 항상 시간대 포함 (+09:00)
                    "temp": self.temp,
                    "weather_desc": self.weather_desc,
                    "icon": self.icon,
                    "city_name": self.city_name
                }
                with open(cache_file, "w", encoding="utf-8") as f:
                    json.dump(cache_data, f, ensure_ascii=False)
                self.update_weather_data()
                break
            except (requests.RequestException, ValueError) as e:
                logger.error(f"Realtime fetch attempt {attempt + 1} failed: {str(e)}")
                if attempt < retries - 1:
                    time.sleep(5)  # 재시도 전 5초 대기
                elif attempt == retries - 1:
                    self.weather_data = "⚠️ 실시간 날씨 정보를 가져올 수 없습니다."

    def update_weather_data(self):
        if self.temp is not None and self.temp_min is not None and self.temp_max is not None and self.weather_desc is not None and self.icon is not None and self.city_name is not None:
            if self.temp > self.temp_max:
                self.temp_max = self.temp
            if self.temp < self.temp_min:
                self.temp_min = self.temp
            now = datetime.now(self.seoul_tz)
            if self.is_first_run and self.last_realtime_update:
                # 최초 실행 시 실제 갱신 시간
                update_time = self.last_realtime_update.strftime('%H:%M')
                self.is_first_run = False  # 최초 갱신 후 False로 전환
            else:
                # 이후에는 정각 표시
                update_time = now.replace(minute=0, second=0, microsecond=0).strftime('%H:00')
            self.weather_data = f"{self.icon} {self.city_name} 날씨: {self.weather_desc}, {self.temp:.1f}°C (최저 {self.temp_min:.1f}°C / 최고 {self.temp_max:.1f}°C, {update_time} 갱신)"
        else:
            self.weather_data = "날씨 정보를 불러오는 중..."
        if hasattr(self, 'weather_label'):
            self.weather_label.setText(self.weather_data)

    def update_forecast_weather(self):
        self.fetch_forecast_weather()
        if hasattr(self, 'weather_label'):
            self.weather_label.setText(self.weather_data)

    def update_realtime_weather(self):
        self.fetch_realtime_weather()
        if hasattr(self, 'weather_label'):
            self.weather_label.setText(self.weather_data)
        now = datetime.now(self.seoul_tz)
        self.last_realtime_update = now
        # is_first_run은 여기서 변경하지 않음 (최초 UI 갱신 후만 False로)
        logger.info(f"Realtime weather updated at {now.strftime('%H:%M')}")

    def set_status_message(self, message, timeout=5000):
        """상태 메시지를 설정하고 지정된 시간 후 지움"""
        # 기존 타이머가 있으면 중지
        if self.status_timer is not None and self.status_timer.isActive():
            self.status_timer.stop()

        self.status_label.setText(message)

        # 빈 메시지가 아닌 경우에만 타이머 설정
        if message:
            self.status_timer = QTimer(self)
            self.status_timer.setSingleShot(True)
            self.status_timer.timeout.connect(lambda: self.status_label.setText(""))
            self.status_timer.start(timeout)
        else:
            self.status_label.setText("")  # 즉시 빈 문자열 설정

    def on_player_state_changed(self, state):
        if state == QMediaPlayer.StoppedState:
            pass

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)

        self.tab_widget = QTabWidget()
        self.today_widget = QWidget()
        self.edit_widget = QWidget()
        self.help_widget = QWidget()
        self.tab_widget.addTab(self.today_widget, "오늘 시간표")
        self.tab_widget.addTab(self.edit_widget, "시간표 관리")
        self.tab_widget.addTab(self.help_widget, "사용설명서")
        main_layout.addWidget(self.tab_widget)

        # 오늘 시간표 탭
        today_layout = QHBoxLayout(self.today_widget)
        today_layout.setContentsMargins(10, 10, 10, 10)
        today_layout.setSpacing(0)

        self.today_table = QTableWidget(0, 2)
        self.today_table.setHorizontalHeaderLabels(["시간", "구분"])
        self.today_table.setColumnWidth(0, 60)
        self.today_table.setColumnWidth(1, 87)
        self.today_table.setFixedWidth(200)
        self.today_table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.today_table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.today_table.setFont(QFont("맑은 고딕", 16))
        self.today_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.today_table.verticalHeader().setDefaultAlignment(Qt.AlignRight)
        self.today_table.horizontalHeader().setStyleSheet("""
            QHeaderView::section {
                background-color: #F0F0F0;
                font-size: 13px;
                padding: 4px;
            }
        """)
        self.today_table.setStyleSheet("QTableWidget::item { padding:8px; }")
        self.today_table.setEditTriggers(QTableWidget.NoEditTriggers)  # 편집 불가 설정 추가
        today_layout.addWidget(self.today_table)

        right_container = QWidget()
        right_container.setFixedWidth(600)
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(14, 0, 0, 0)

        self.today_header = QLabel(self.get_today_header())
        self.today_header.setAlignment(Qt.AlignCenter)
        self.today_header.setWordWrap(False)
        self.today_header.setMaximumWidth(540)
        self.today_header.setStyleSheet("""
            background-color: #7FB3D5;
            border-radius: 10px;
            padding: 10px;
            color: #FFFFFF;
            font-size: 96px;
        """)
        right_layout.addWidget(self.today_header)

        self.weather_label = QLabel(self.weather_data)
        self.weather_label.setAlignment(Qt.AlignCenter)
        self.weather_label.setFont(QFont("맑은 고딕", 10))
        self.weather_label.setMaximumWidth(540)
        self.weather_label.setStyleSheet("""
            background-color: #F7F9FA;
            border: 1px solid #E0E4E8;
            border-radius: 8px;
            padding: 8px;
            color: #4A5A6A;
        """)
        right_layout.addWidget(self.weather_label)
        right_layout.addSpacerItem(QSpacerItem(200, 10))

        volume_layout = QHBoxLayout()
        volume_label = QLabel("<b>[ 볼륨 조절 ]</b>")
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setValue(self.volume_value)
        self.volume_slider.setMaximumWidth(400)
        self.volume_slider.valueChanged.connect(self.set_volume)
        self.volume_slider.setToolTip("슬라이더를 움직여 볼륨 조절")
        self.volume_slider.setStyleSheet("""
            QSlider::groove:horizontal {
                height: 8px;
                background: #A9CCE3;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                background: #FFFFFF;
                border: 1px solid #85C1E9;
                width: 16px;
                height: 16px;
                margin: -4px 0;
                border-radius: 8px;
            }
        """)
        self.volume_value = QLabel("100%")
        volume_layout.addWidget(volume_label)
        volume_layout.addWidget(self.volume_slider)
        volume_layout.addWidget(self.volume_value)
        right_layout.addLayout(volume_layout)

        right_layout.addSpacerItem(QSpacerItem(20, 10))

        btn_layout = QGridLayout()
        btn_layout.setContentsMargins(0, 0, 50, 0)
        self.play_bell_btn = QPushButton("🎵 벨소리 재생")
        self.play_bell_btn.clicked.connect(self.play_start_bell)
        self.play_bell_btn.setToolTip("수업시작 벨소리를 재생합니다.")
        self.play_bell_btn.setFixedWidth(380)
        btn_layout.addWidget(self.play_bell_btn, 0, 0)

        self.bell_btn = QPushButton("🔔 벨소리 변경")
        self.bell_btn.clicked.connect(self.choose_bells)
        self.bell_btn.setToolTip("시작 및 종료 벨소리를 설정합니다.")
        self.bell_btn.setFixedWidth(150)
        btn_layout.addWidget(self.bell_btn, 0, 1)

        checkbox_layout = QHBoxLayout()
        self.autostart_checkbox = QCheckBox("윈도우 부팅시 자동 시작")
        self.autostart_checkbox.setChecked(self.autostart_state)
        self.autostart_checkbox.stateChanged.connect(self.toggle_autostart)
        self.autostart_checkbox.setToolTip("윈도우 시작 시 프로그램을 자동 실행합니다.")
        checkbox_layout.addStretch()
        checkbox_layout.addWidget(self.autostart_checkbox)
        btn_layout.addLayout(checkbox_layout, 1, 0)

        self.exit_btn = QPushButton("🚪 종료")
        self.exit_btn.clicked.connect(self.exit_program)
        self.exit_btn.setToolTip("프로그램을 완전히 종료합니다.")
        self.exit_btn.setFixedWidth(150)
        btn_layout.addWidget(self.exit_btn, 1, 1)
        right_layout.addLayout(btn_layout)

        right_layout.addSpacerItem(QSpacerItem(20, 10))  # 기존 스페이서 (예시)

        # info_label 위에 빈 공간을 추가
        #right_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        info_text = (
            "<b>[노엠스쿨]</b> <a href='https://noemschool.org'>https://noemschool.org</a> ☎ 0507-1301-6215 (신입생 상시 모집)<br><br>"
            "이 프로그램은 대전 소재의 기독 대안학교 '노엠스쿨'(여자 중고등학교)에서 학교의 필요에 의해 제작한 프로그램입니다.<br>"
            "좋은 곳에 널리 쓰이길 바라며 무료 배포합니다.<br>"
            "- 2025년 3월의 어느 오늘<br><br>"
            
            "<b>* 프로그램 문의:</b> <a href='mailto:todayis@j2w.kr'>todayis@j2w.kr</a><br>"
            "<b>* 프로그램 후원:</b> 카카오페이 후원으로 함께 만들어 갑니다. <a href='showqr://support'>QR 보기</a>"  # 링크 추가
        )

        self.info_label = QLabel(info_text, self.today_widget)
        self.info_label.setContentsMargins(0,5,0,0)
        self.info_label.setFont(QFont("맑은 고딕", 11))
        self.info_label.setAlignment(Qt.AlignLeft)
        self.info_label.setWordWrap(True)
        self.info_label.setMaximumWidth(540)
        self.info_label.setOpenExternalLinks(True)
        self.info_label.setOpenExternalLinks(False)  # 외부 링크 자동 열기 비활성화
        self.info_label.linkActivated.connect(self.handle_link_activated)  # 링크 클릭 이벤트 연결

        self.info_label.setStyleSheet("""
            background-color: #E8ECEF;
            border: 2px solid #D3D3D3;
            border-radius: 8px;
            padding: 15px;
        """)
        right_layout.addWidget(self.info_label)
        right_layout.addStretch()
        today_layout.addWidget(right_container)

        # 시간표 관리 탭
        edit_layout = QVBoxLayout(self.edit_widget)
        edit_layout.setContentsMargins(10, 10, 10, 10)

        list_layout = QHBoxLayout()
        self.schedule_label = QLabel("시간표 목록:")
        self.schedule_label.setFont(QFont("맑은 고딕", 12))
        list_layout.addWidget(self.schedule_label)
        list_layout.setSpacing(6)
        self.schedule_combo = QComboBox()
        self.schedule_combo.setFixedWidth(140)
        self.schedule_combo.addItems(self.schedule_lists.keys())
        self.schedule_combo.setCurrentText(self.current_list)
        list_layout.addWidget(self.schedule_combo)
        list_layout.addStretch()
        self.create_list_btn = QPushButton("📝 목록생성")
        self.create_list_btn.clicked.connect(self.create_list)
        self.create_list_btn.setToolTip("새로운 빈 시간표를 생성합니다.")
        self.create_list_btn.setFixedSize(120, 40)
        list_layout.addWidget(self.create_list_btn)
        self.rename_list_btn = QPushButton("🔧 이름변경")
        self.rename_list_btn.clicked.connect(self.rename_list)
        self.rename_list_btn.setToolTip("선택된 시간표 목록의 이름을 변경합니다.")
        self.rename_list_btn.setFixedSize(120, 40)
        list_layout.addWidget(self.rename_list_btn)
        self.clone_list_btn = QPushButton("📋 복제")
        self.clone_list_btn.clicked.connect(self.clone_list)
        self.clone_list_btn.setToolTip("선택된 시간표 목록을 복제합니다.")
        self.clone_list_btn.setFixedSize(120, 40)
        list_layout.addWidget(self.clone_list_btn)
        self.delete_list_btn = QPushButton("🗑️ 삭제")
        self.delete_list_btn.clicked.connect(self.delete_list)
        self.delete_list_btn.setToolTip("선택된 시간표 목록을 삭제합니다.")
        self.delete_list_btn.setFixedSize(120, 40)
        list_layout.addWidget(self.delete_list_btn)
        edit_layout.addLayout(list_layout)

        self.edit_table = QTableWidget(40, 7)
        self.edit_table.setHorizontalHeaderLabels(self.days)
        self.edit_table.setSelectionBehavior(QTableWidget.SelectItems)
        self.edit_table.setSelectionMode(QTableWidget.ExtendedSelection)

        self.edit_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)  # 원래 설정 복구
        for i in range(40):
            self.edit_table.setRowHeight(i, 30)  # 모든 행 높이를 30px로 고정 (값 조정 가능)
        self.edit_table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)  # 동적 크기 조정 방지

        self.edit_table.itemChanged.connect(self.on_item_changed)
        self.edit_table.setItemDelegate(TimeDelegate(self))
        self.edit_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.edit_table.customContextMenuRequested.connect(self.show_context_menu)
        self.edit_table.doubleClicked.connect(self.on_double_click)
        self.edit_table.setFont(QFont("맑은 고딕", 10))

        self.edit_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.edit_table.verticalHeader().setDefaultAlignment(Qt.AlignRight)
        self.edit_table.setMaximumWidth(800)
        for i in range(7):
            self.edit_table.setColumnWidth(i, 100) #101
        self.edit_table.setMinimumHeight(500)
        self.edit_table.horizontalHeader().setFixedHeight(40)
        self.edit_table.verticalHeader().setDefaultSectionSize(25)
        self.edit_table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.edit_table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.edit_table.setAttribute(Qt.WA_InputMethodEnabled, True)  # IME 활성화
        edit_layout.addWidget(self.edit_table)

        self.schedule_combo.currentTextChanged.connect(self.load_list)

        edit_control_layout = QHBoxLayout()
        self.reset_btn = QPushButton("♻️ 스케줄 초기화")
        self.reset_btn.clicked.connect(self.reset_schedule)
        self.reset_btn.setToolTip("현재 시간표를 초기화합니다.")
        self.reset_btn.setFixedWidth(150)
        edit_control_layout.addWidget(self.reset_btn)
        self.repeat_schedule_btn = QPushButton("🔄 반복 스케줄 생성")
        self.repeat_schedule_btn.clicked.connect(self.create_repeat_schedule)
        self.repeat_schedule_btn.setToolTip("반복 스케줄을 생성합니다.")
        self.repeat_schedule_btn.setFixedWidth(150)
        edit_control_layout.addWidget(self.repeat_schedule_btn)
        self.save_file_btn = QPushButton("💾 파일 저장")
        self.save_file_btn.clicked.connect(self.save_schedule)
        self.save_file_btn.setToolTip("현재 시간표를 파일로 저장합니다.")
        self.save_file_btn.setFixedWidth(150)
        edit_control_layout.addWidget(self.save_file_btn)
        self.load_file_btn = QPushButton("📂 파일 불러오기")
        self.load_file_btn.clicked.connect(self.load_schedule)
        self.load_file_btn.setToolTip("저장된 시간표 파일을 불러옵니다.")
        self.load_file_btn.setFixedWidth(150)
        edit_control_layout.addWidget(self.load_file_btn)
        edit_control_layout.addStretch()
        edit_layout.addLayout(edit_control_layout)

        self.status_label = QLabel("")
        self.status_label.setFont(QFont("맑은 고딕", 10))
        edit_layout.addWidget(self.status_label)

        # 사용설명서 탭
        help_layout = QVBoxLayout(self.help_widget)
        help_layout.setContentsMargins(10, 10, 10, 10)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
                QScrollArea {
                    border: 1px solid #D3D3D3;
                    border-radius: 6px;
                    background-color: #F9F9F9;
                }
        """)

        help_content = QWidget()
        help_content_layout = QVBoxLayout(help_content)

        help_text = """
        <h2>🌟 오늘은 v1.2 사용설명서</h2>
        <p>안녕하세요, 선생님! "오늘은"은 수업 시간에 맞춰 벨을 울리고 시간표를 관리하는 간단한 프로그램이에요. 처음 사용하셔도 쉽게 익힐 수 있도록 설명드릴게요!</p>

        <h3>1. 프로그램 시작하기</h3>
        <ul>
            <li><b>실행</b>: 바탕화면 아이콘을 더블클릭하거나, 컴퓨터를 켜면 자동으로 시작돼요.</li>
            <li><b>창</b>: 프로그램이 열리면 세 개의 탭이 보여요. "오늘 시간표", "시간표 관리", "사용설명서"</li>
        </ul>

        <h3>2. 주요 기능</h3>
        <p>탭마다 할 수 있는 일을 알려드릴게요!</p>

        <h4>🎯 첫 번째 탭: 오늘 시간표</h4>
        <ul>
            <li><b>시간/날씨 확인</b>: 오늘 날짜와 시간과 날씨가 표시돼요.</li>
            <li><b>오늘 일정</b>: 왼쪽에 오늘 알림 일정이 보여요.</li>
            <li><b>🎵 벨소리 재생</b>: 버튼을 누르면 시작 벨을 바로 울릴 수 있어요.</li>
            <li><b>🔔 벨소리 변경</b>: 시작과 종료 벨을 원하는 소리로 바꿀 수 있어요. (mp3, wav 형식)</li>
            <li><b>🔊 볼륨조절</b>: 슬라이더를 움직여 벨소리 크기를 조절해요. (0~100%)</li>
            <li><b>🚀 자동 시작</b>: 체크하면 컴퓨터 켤 때마다 프로그램이 자동 실행돼요.</li>
            <li><b>🚪 종료</b>: 버튼을 누르면 프로그램이 완전히 꺼져요.</li>
        </ul>

        <h4>📅 두 번째 탭: 시간표 관리</h4>
        <ul>
            <li><b>시간표 목록</b>: 시간표를 새로 만들거나 복제, 이름변경, 삭제할 수 있어요.</li>
            <ul>
                <li>📝 <b>목록생성</b>: 새 시간표를 만들어요.</li>
                <li>🔧 <b>이름변경</b>: 현재 시간표 이름을 바꿔요.</li>
                <li>📋 <b>복제</b>: 현재 시간표를 똑같이 복사해요.</li>
                <li>🗑️ <b>삭제</b>: 필요 없는 시간표를 지워요. ("기본 스케줄"은 못 지움)</li>
            </ul>
            <li><b>시간 입력</b>: 테이블에서 요일별 스케줄을 입력해요.</li>
            <ul>
                <li>더블클릭하거나 Enter 키로 셀을 편집해요.</li>
                <li>형식: "HH:MM" (예: 09:00) 또는 "HHMM". (예: 0900)</li>
                <li>입력 후 Enter → 다음 줄로 이동하며 계속 추가 가능해요.</li>
            </ul>
            <li><b>단축키</b>: 단축키나 우클릭으로 셀 설정을 쉽게 할 수 있어요.</li>
            <ul>
                <li><b>S 키</b>: "수업시작"으로 설정. (🟡 노란색)</li>
                <li><b>D 키</b>: "종료"로 설정. (⚪ 회색)</li>
                <li><b>E 키</b>: "기타시작"으로 설정. (🟢 연두색)</li>
                <li><b>I 키</b>: 스케줄 삽입.</li>
                <li><b>Backspace 키</b>: 선택한 시간/빈 칸 삭제.</li>
                <li><b>Ctrl+C,V 키</b>: 선택한 시간 복사/붙여넣기.</li>
                <li><b>ESC 키</b>: 편집종료. 삽입 취소.</li>
                <li>우클릭 메뉴: "복사","붙여넣기,","삭제", "삽입", "빈 칸 삭제", "수업시작/종료/기타시작" 선택.</li>
                <li><b>주의</b>: 단축키 중 'S, D, E, I'는 한글 입력 모드에서는 동작하지 않을 수 있습니다. 영어 모드에서 사용해 주세요.</li>
            </ul>
            <li><b>🔄 반복 스케줄 생성</b>: 똑같은 패턴으로 여러 시간을 추가해요.</li>
            <ul>
                <li>시작 시간, 수업 시간, 쉬는 시간, 반복 횟수, 요일을 선택.</li>
                <li>예) 09:00 시작, 60분 수업, 10분 쉬고, 5번 반복 → 5교시 자동 생성!</li>
            </ul>
            <li><b>♻️ 스케줄 초기화</b>: 현재 시간표를 모두 지워요.</li>
            <li><b>💾 파일 저장 / 📂 파일 불러오기</b>: 현재 시간표를 파일로 저장하거나 불러와요.</li>
        </ul>

        <h4>📖 세 번째 탭: 사용설명서</h4>
        <ul>
            <li>지금 보고 계신 이 설명서예요! 스크롤해서 읽어보세요.</li>
        </ul>

        <h3>3. 유용한 팁</h3>
        <ul>
            <li>🕛 <b>시간</b>: 밤 12시는 "00:00"으로 입력하세요.</li>
            <li>🎶 <b>벨소리</b>: mp3나 wav 파일을 준비해서 "벨소리 변경"으로 설정해 보세요.</li>
            <li>💻 <b>컴퓨터</b>: 노트북이라면 전원을 연결해 두세요.</li>
            <li>🚀 <b>자동 시작</b>: 체크는 관리자 모드로 실행 시 설정 가능해요.</li>
            <li>❓ <b>문제 시</b>: "종료" 후 다시 켜보세요. 안 되면 <a href="mailto:todayis@j2w.kr">todayis@j2w.kr</a>로 문의!</li>
        </ul>

        <p>이 프로그램은 노엠스쿨에서 선생님들과 학생들을 위해 만든 무료 프로그램이에요. <br>편리하게 사용해 주세요! 🌈<br>
        <b>* 노엠스쿨 홈페이지:</b> <a href="https://noemschool.org">https://noemschool.org</a></p>
        """
        help_label = QLabel(help_text)
        help_label.setFont(QFont("맑은 고딕", 12))
        help_label.setWordWrap(True)
        help_label.setOpenExternalLinks(True)
        help_label.setStyleSheet("padding: 10px;")

        help_content_layout.addWidget(help_label)
        help_content_layout.addStretch()
        scroll_area.setWidget(help_content)
        help_layout.addWidget(scroll_area)

        # edit_table 스타일은 별도로 적용
        self.edit_table.setStyleSheet("""
                QTableWidget::item {
                    text-align: center;
                    padding: 0px;
                    height: 30px;
                }
        """)

    def apply_style(self):
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#F5F5F5"))
        palette.setColor(QPalette.Button, QColor("#A9CCE3"))
        palette.setColor(QPalette.ButtonText, QColor(0, 0, 0))
        self.setPalette(palette)

        self.setStyleSheet("""
            QPushButton {
                background-color: #A9CCE3;
                border: none;
                padding: 10px;
                border-radius: 6px;
                font-family: "맑은 고딕";
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #85C1E9;
            }
            QPushButton:disabled {
                background-color: #D5DBDB;
            }
            QTableWidget {
                border: 1px solid #D3D3D3;
                border-radius: 6px;
                font-family: "맑은 고딕";
                font-size: 12px;
            }
            QTableWidget::item {
                text-align: center;
            }
            QTableWidget QHeaderView::section {
                background-color: #F0F0F0;
                font-weight: bold;
                border: none;
                padding: 5px;
            }
            QTabWidget::pane {
                border: 1px solid #D3D3D3;
                border-radius: 6px;
            }
            QTabBar::tab {
                background: #E6E6E6;
                padding: 10px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background: #A9CCE3;
                color: black;
            }
            QLabel {
                font-family: "맑은 고딕";
            }
            QMainWindow { 
                background-color: #f0f0f0; 
            }
        """)

    def get_today_header(self):
        utc_time = datetime.now(timezone.utc)
        seoul_time = utc_time.astimezone(self.seoul_tz)
        date_str = seoul_time.strftime('%Y년 %m월 %d일')
        day_str = self.days[seoul_time.weekday()][0]
        time_str = QTime.currentTime().toString('HH:mm:ss')
        return f"<div style='font-size: 40px;'>{date_str} ({day_str})</div><div style='font-size: 96px;'>{time_str}</div>"

    def update_today_schedule(self):
        current_day_idx = datetime.now(self.seoul_tz).weekday()
        current_day = self.days[current_day_idx]
        current_time = QTime.currentTime().toString("HH:mm:ss")
        self.today_table.setRowCount(0)

        for item in self.schedule_lists[self.current_list][current_day]:
            schedule_time = item["time"] + ":00"
            if item["active"] == "예" and schedule_time > current_time:
                row = self.today_table.rowCount()
                self.today_table.insertRow(row)
                time_item = QTableWidgetItem(item["time"])
                time_item.setTextAlignment(Qt.AlignCenter)
                desc_item = QTableWidgetItem(item["desc"])
                desc_item.setTextAlignment(Qt.AlignCenter)

                if "수업시작" in item["desc"] or "교시 시작" in item["desc"]:
                    time_item.setBackground(QColor("#FFFF99"))
                    desc_item.setBackground(QColor("#FFFF99"))
                elif "종료" in item["desc"] or "교시 종료" in item["desc"]:
                    time_item.setBackground(QColor("#F5F5F5"))
                    desc_item.setBackground(QColor("#F5F5F5"))
                elif "기타시작" in item["desc"]:
                    time_item.setBackground(QColor("#CCFFCC"))
                    desc_item.setBackground(QColor("#CCFFCC"))

                self.today_table.setItem(row, 0, time_item)
                self.today_table.setItem(row, 1, desc_item)

    def update_edit_table(self):
        self.edit_table.blockSignals(True)
        max_rows = max(40, max([len(self.schedule_lists[self.current_list][day]) for day in self.days]))
        self.edit_table.setRowCount(max_rows)

        for col, day in enumerate(self.days):
            schedule = self.schedule_lists[self.current_list][day]
            if not self.is_inserting:  # 삽입 중이 아니면 빈 셀 제거 및 정렬
                schedule = [s for s in schedule if s["time"]]
                schedule = sorted(schedule, key=lambda x: x["time"] or "00:00")
            # 삽입 중이면 원래 순서 유지 (빈 셀 포함)
            row = 0
            for item in schedule:
                table_item = QTableWidgetItem(item["time"])
                table_item.setTextAlignment(Qt.AlignCenter)
                if "수업시작" in item["desc"] or "교시 시작" in item["desc"]:
                    table_item.setIcon(QIcon(resource_path("start_icon.png")))
                    table_item.setBackground(QColor("#FFFF99"))
                elif "종료" in item["desc"] or "교시 종료" in item["desc"]:
                    table_item.setIcon(QIcon(resource_path("end_icon.png")))
                    table_item.setBackground(QColor("#F5F5F5"))
                elif "기타시작" in item["desc"]:
                    table_item.setIcon(QIcon(resource_path("start_icon.png")))
                    table_item.setBackground(QColor("#CCFFCC"))
                self.edit_table.setItem(row, col, table_item)
                self.edit_table.setRowHeight(row, 30)
                row += 1
            while row < max_rows:
                empty_item = QTableWidgetItem("")
                empty_item.setTextAlignment(Qt.AlignCenter)
                empty_item.setData(Qt.UserRole, None)
                self.edit_table.setItem(row, col, empty_item)
                self.edit_table.setRowHeight(row, 30)
                row += 1
        self.edit_table.setStyleSheet("QTableWidget::item { padding: 0px; }")
        self.edit_table.blockSignals(False)
        self.update_today_schedule()
        logger.debug(f"Table updated, row count: {self.edit_table.rowCount()}")

    def play_start_bell(self):
        self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.start_bell))))
        self.player.play()

    def set_volume(self, value):
        self.player.setVolume(value)
        self.volume_value.setText(f"{value}%")

    def choose_bells(self):
        dialog = BellDialog(self.start_bell, self.end_bell, self.misc_bell, parent=self)  # 그대로 유지
        if dialog.exec_():
            self.start_bell, self.end_bell, self.misc_bell = dialog.get_bells()
            self.save_config()

    def show_context_menu(self, pos):
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.set_status_message("선택된 셀이 없습니다.")
            return
        menu = QMenu(self)
        copy_action = QAction("복사(Ctrl+C)", self)
        copy_action.triggered.connect(self.copy_schedule)
        menu.addAction(copy_action)
        paste_action = QAction("붙여넣기(Ctrl+V)", self)
        paste_action.triggered.connect(self.paste_schedule)
        menu.addAction(paste_action)
        delete_action = QAction("삭제(Backspace)", self)
        delete_action.triggered.connect(self.delete_schedule)
        menu.addAction(delete_action)
        add_below_action = QAction("삽입(I)", self)
        add_below_action.triggered.connect(self.add_below_schedule)
        menu.addAction(add_below_action)
        remove_empty_action = menu.addAction("빈 칸 삭제(ESC)")
        remove_empty_action.triggered.connect(self.remove_empty_rows)
        set_start_action = QAction("수업시작 설정(S)", self)
        set_start_action.triggered.connect(lambda: self.set_schedule_type(selected_items, "수업시작"))
        menu.addAction(set_start_action)
        set_end_action = QAction("종료 설정(D)", self)
        set_end_action.triggered.connect(lambda: self.set_schedule_type(selected_items, "종료"))
        menu.addAction(set_end_action)
        set_misc_action = QAction("기타시작 설정(E)", self)
        set_misc_action.triggered.connect(lambda: self.set_schedule_type(selected_items, "기타시작"))
        menu.addAction(set_misc_action)
        menu.exec_(self.edit_table.viewport().mapToGlobal(pos))

    def remove_empty_rows(self):
        logger.debug("Remove empty rows triggered via context menu")
        for day in self.days:
            self.schedule_lists[self.current_list][day] = [s for s in self.schedule_lists[self.current_list][day] if s["time"]]
        self.update_edit_table()
        self.set_status_message("빈 행이 제거되었습니다.")

    def set_schedule_type(self, items, desc_type):
        if not items:
            self.set_status_message("선택된 셀이 없습니다.")
            return
        updated = False
        for item in items:
            row = item.row()
            col = item.column()
            day = self.days[col]
            time = item.text().strip()  # 빈 문자열 제거
            # 시간이 비어 있으면 동작하지 않음
            if not time or not QTime.fromString(time, "HH:mm").isValid():
                continue  # 빈 셀이거나 유효하지 않은 시간은 건너뜀
            found = False
            for i, schedule in enumerate(self.schedule_lists[self.current_list][day]):
                if schedule["time"] == time:
                    self.schedule_lists[self.current_list][day][i]["desc"] = desc_type
                    found = True
                    updated = True
                    break
            if not found:
                self.schedule_lists[self.current_list][day].append({"time": time, "desc": desc_type, "active": "예"})
                updated = True
        if updated:
            self.update_edit_table()
            self.save_config()
            self.set_status_message(f"{desc_type} 설정 완료")
        else:
            self.set_status_message("빈 셀이거나 유효하지 않은 시간입니다. 동작이 실행되지 않았습니다.")

    def delete_schedule(self):
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.set_status_message("삭제할 셀이 선택되지 않았습니다.")
            return

        self.edit_table.blockSignals(True)
        for item in selected_items:
            row = item.row()
            col = item.column()
            day = self.days[col]
            time = item.text()
            if time:
                self.schedule_lists[self.current_list][day] = [s for s in self.schedule_lists[self.current_list][day] if s["time"] != time]
        self.update_edit_table()
        self.update_today_schedule()
        self.save_config()
        self.set_status_message("선택된 시간이 삭제되었습니다.")
        self.edit_table.blockSignals(False)

    def add_below_schedule(self):
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.set_status_message("추가할 위치가 선택되지 않았습니다.")
            return
        item = selected_items[0]
        row = item.row()
        col = item.column()
        day = self.days[col]

        try:
            self.edit_table.blockSignals(True)
            logger.debug(f"Adding new schedule below row {row}, col {col} in {day}")

            # 모든 요일에서 기존 빈 셀 제거
            for d in self.days:
                self.schedule_lists[self.current_list][d] = [
                    s for s in self.schedule_lists[self.current_list][d] if s["time"]
                ]

            # 삽입 상태 설정
            self.is_inserting = True

            # 데이터 리스트에 새 항목 삽입
            current_schedule = self.schedule_lists[self.current_list][day]
            insert_pos = row + 1
            new_schedule = {"time": "", "desc": "수업시작", "active": "예"}
            if insert_pos > len(current_schedule):
                current_schedule.append(new_schedule)
            else:
                current_schedule.insert(insert_pos, new_schedule)
            logger.debug(f"Inserted temporary empty schedule at position {insert_pos} in {day}")

            # 테이블 갱신
            self.update_edit_table()

            # 새 셀로 이동 및 편집 모드 진입
            new_row = insert_pos
            new_item = self.edit_table.item(new_row, col)
            if not new_item or new_item.text() != "":
                new_item = QTableWidgetItem("")
                new_item.setTextAlignment(Qt.AlignCenter)
                new_item.setData(Qt.UserRole, "new_empty")
                self.edit_table.setItem(new_row, col, new_item)
            self.edit_table.setCurrentCell(new_row, col)
            self.edit_table.editItem(new_item)
            self.edit_table.blockSignals(False)
            QTimer.singleShot(0, lambda: self.ensure_edit_mode(new_row, col))

            self.edit_table.setRowHeight(new_row, 30)
            self.edit_table.scrollTo(self.edit_table.model().index(new_row, col))
            logger.debug(f"UI refreshed, focused on {new_row}, {col}, row count: {self.edit_table.rowCount()}")

            self.set_status_message("아래에 새 시간이 추가되었습니다. 시간을 입력하세요. (삽입취소: ESC)")
        except Exception as e:
            logger.error(f"Error in add_below_schedule: {str(e)}")
            self.edit_table.blockSignals(False)
            self.is_inserting = False
            self.set_status_message(f"셀 추가 중 오류 발생: {str(e)}")

    def ensure_edit_mode(self, row, col):
        """편집 모드와 포커스를 강제로 유지"""
        item = self.edit_table.item(row, col)
        if item and not self.edit_table.isPersistentEditorOpen(item):
            self.edit_table.editItem(item)
        editor = self.edit_table.cellWidget(row, col)
        if editor:
            editor.setFocus()
            self.edit_table.setFocus()
        logger.debug(f"Edit mode ensured for item at {row}, {col}")

    def create_repeat_schedule(self):
        dialog = RepeatScheduleDialog(self.days, self)
        if dialog.exec_():
            data = dialog.get_schedule_data()
            start_time = QTime.fromString(data["start_time"], "HH:mm")
            duration = data["duration"]
            break_time = data["break_time"]
            repeat_count = data["repeat_count"]
            selected_days = data["days"]

            if not start_time.isValid():
                self.set_status_message("잘못된 시작 시간입니다.")
                return

            for day in selected_days:
                current_time = start_time
                for i in range(repeat_count):
                    start_str = current_time.toString("HH:mm")
                    if not any(s["time"] == start_str for s in self.schedule_lists[self.current_list][day]):
                        self.schedule_lists[self.current_list][day].append({"time": start_str, "desc": "수업시작", "active": "예"})
                    end_time = current_time.addSecs(duration * 60)
                    end_str = end_time.toString("HH:mm")
                    if not any(s["time"] == end_str for s in self.schedule_lists[self.current_list][day]):
                        self.schedule_lists[self.current_list][day].append({"time": end_str, "desc": "종료", "active": "예"})
                    current_time = end_time.addSecs(break_time * 60)
            self.update_edit_table()
            self.update_today_schedule()
            self.save_config()
            self.set_status_message(f"선택된 요일에 반복 스케줄이 추가되었습니다.")

    def on_item_changed(self, item):
        row = item.row()
        col = item.column()
        new_time = item.text().strip()
        day = self.days[col]

        if not new_time:
            # 빈 입력 시 아무 동작 안 함 (행 유지)
            if item.data(Qt.UserRole) == "new_empty":
                self.set_status_message("시간을 입력하세요.")
            return

        formatted_time = None
        if len(new_time) == 4 and new_time.isdigit():
            formatted_time = f"{new_time[:2]}:{new_time[2:]}"
            if not QTime.fromString(formatted_time, "HH:mm").isValid():
                item.setText("")
                self.set_status_message(f"유효하지 않은 시간: {new_time}")
                return
            item.setText(formatted_time)
        elif new_time and QTime.fromString(new_time, "HH:mm").isValid():
            formatted_time = new_time
        else:
            item.setText("")
            self.set_status_message(f"시간 형식이 잘못됨: {new_time}")
            return

        if formatted_time:
            self.schedule_lists[self.current_list][day] = [s for s in self.schedule_lists[self.current_list][day] if
                                                           s["time"]]
            if row < len(self.schedule_lists[self.current_list][day]):
                old_time = self.schedule_lists[self.current_list][day][row]["time"]
                old_desc = self.schedule_lists[self.current_list][day][row]["desc"]
                if old_time != formatted_time:
                    self.schedule_lists[self.current_list][day] = [s for s in
                                                                   self.schedule_lists[self.current_list][day] if
                                                                   s["time"] != old_time]
                    if not any(s["time"] == formatted_time for s in self.schedule_lists[self.current_list][day]):
                        self.schedule_lists[self.current_list][day].append(
                            {"time": formatted_time, "desc": old_desc, "active": "예"})
                        self.schedule_lists[self.current_list][day].sort(key=lambda x: x["time"])
            else:
                if not any(s["time"] == formatted_time for s in self.schedule_lists[self.current_list][day]):
                    desc = "수업시작" if row % 2 == 0 else "종료"
                    if row > 0 and self.edit_table.item(row - 1, col):
                        prev_item = self.edit_table.item(row - 1, col)
                        if prev_item and prev_item.text():
                            prev_desc = next((s["desc"] for s in self.schedule_lists[self.current_list][day] if
                                              s["time"] == prev_item.text()), None)
                            if prev_desc:
                                desc = "종료" if "시작" in prev_desc else "수업시작"
                    self.schedule_lists[self.current_list][day].append(
                        {"time": formatted_time, "desc": desc, "active": "예"})
                    self.schedule_lists[self.current_list][day].sort(key=lambda x: x["time"])
            item.setData(Qt.UserRole, None)  # 새 셀 플래그 제거
            self.save_config()
            self.set_status_message(f"시간 수정: {formatted_time}")
            self.update_edit_table()

    def on_double_click(self):
        selected_items = self.edit_table.selectedItems()
        if selected_items:
            self.edit_table.editItem(selected_items[0])

    def eventFilter(self, source, event):
        if source == self.edit_table and event.type() == QEvent.KeyPress:
            key = event.key()
            current_items = self.edit_table.selectedItems()

            if key in (Qt.Key_Return, Qt.Key_Enter):
                current_row = self.edit_table.currentRow()
                current_col = self.edit_table.currentColumn()
                if current_row >= 0 and current_col >= 0:
                    current_item = self.edit_table.item(current_row, current_col)
                    if not current_item:
                        current_item = QTableWidgetItem("")
                        self.edit_table.setItem(current_row, current_col, current_item)
                    editor = self.edit_table.cellWidget(current_row, current_col)
                    if not editor:
                        self.edit_table.editItem(current_item)
                        editor = self.edit_table.cellWidget(current_row, current_col)
                    if editor and current_item:
                        new_text = editor.text().strip()
                        self.edit_table.closePersistentEditor(current_item)
                        if new_text:
                            formatted_time = None
                            if len(new_text) == 4 and new_text.isdigit():
                                formatted_time = f"{new_text[:2]}:{new_text[2:]}"
                                if not QTime.fromString(formatted_time, "HH:mm").isValid():
                                    current_item.setText("")
                                    self.set_status_message(f"유효하지 않은 시간: {new_text}")
                                    return True
                            elif new_text and QTime.fromString(new_text, "HH:mm").isValid():
                                formatted_time = new_text
                            else:
                                current_item.setText("")
                                self.set_status_message(f"시간 형식이 잘못됨: {new_text}")
                                return True

                            day = self.days[current_col]
                            current_schedule = self.schedule_lists[self.current_list][day]
                            old_time = current_item.text() if current_item.text() else None
                            if old_time and old_time != formatted_time:
                                current_schedule[:] = [s for s in current_schedule if s["time"] != old_time]
                            else:
                                # 빈 셀 제거
                                current_schedule[:] = [s for s in current_schedule if s["time"]]

                            if not any(s["time"] == formatted_time for s in current_schedule):
                                desc = "수업시작" if current_row % 2 == 0 else "종료"
                                if current_row > 0 and self.edit_table.item(current_row - 1, current_col):
                                    prev_item = self.edit_table.item(current_row - 1, current_col)
                                    if prev_item and prev_item.text():
                                        prev_desc = next((s["desc"] for s in current_schedule if s["time"] == prev_item.text()), None)
                                        if prev_desc:
                                            desc = "종료" if "시작" in prev_desc else "수업시작"
                                current_schedule.append({"time": formatted_time, "desc": desc, "active": "예"})
                                current_schedule.sort(key=lambda x: x["time"])
                            self.is_inserting = False
                            self.save_config()
                            self.update_edit_table()
                            self.edit_table.setRowHeight(current_row, 30)
                            self.update_today_schedule()
                            self.set_status_message(f"시간 업데이트: {formatted_time}")

                        next_row = current_row + 1
                        if next_row >= self.edit_table.rowCount():
                            self.edit_table.insertRow(next_row)
                            for col in range(self.edit_table.columnCount()):
                                empty_item = QTableWidgetItem("")
                                empty_item.setTextAlignment(Qt.AlignCenter)
                                self.edit_table.setItem(next_row, col, empty_item)
                            self.edit_table.setRowHeight(next_row, 30)
                        self.edit_table.setCurrentCell(next_row, current_col)
                        self.edit_table.editItem(self.edit_table.item(next_row, current_col))
                        return True

            elif current_items and (Qt.Key_0 <= key <= Qt.Key_9):
                current_row = current_items[0].row()
                current_col = current_items[0].column()
                current_item = self.edit_table.item(current_row, current_col)
                if not self.edit_table.isPersistentEditorOpen(current_item):
                    self.edit_table.editItem(current_item)
                editor = self.edit_table.cellWidget(current_row, current_col)
                if editor:
                    editor.setText(event.text())
                    editor.setFocus()
                self.set_status_message(f"숫자 입력 시작: {event.text()}")
                return True

            elif current_items and key == Qt.Key_Colon:
                self.edit_table.setCurrentItem(current_items[0])
                editor = self.edit_table.cellWidget(self.edit_table.currentRow(), self.edit_table.currentColumn())
                if not editor:
                    self.edit_table.editItem(current_items[0])
                    editor = self.edit_table.cellWidget(self.edit_table.currentRow(), self.edit_table.currentColumn())
                if editor:
                    current_text = editor.text()
                    editor.setText(current_text + ":")
                    editor.setFocus()
                self.set_status_message(f"콜론 입력: :, 셀 값: {current_items[0].text()}")
                return True

            elif key == Qt.Key_Escape and current_items:
                current_row = current_items[0].row()
                current_col = current_items[0].column()
                current_item = self.edit_table.item(current_row, current_col)
                editor = self.edit_table.cellWidget(current_row, current_col)
                logger.debug(f"ESC pressed at {current_row}, {current_col}, text: '{current_item.text() if current_item else None}', editor: {bool(editor)}")

                if editor:
                    self.edit_table.closePersistentEditor(current_item)
                    old_time = current_item.text()
                    current_item.setText(old_time if old_time else "")
                if self.is_inserting:
                    day = self.days[current_col]
                    current_schedule = self.schedule_lists[self.current_list][day]
                    current_schedule[:] = [s for s in current_schedule if s["time"]]
                    self.is_inserting = False
                    self.update_edit_table()
                    self.set_status_message("삽입 취소")
                else:
                    self.set_status_message(f"편집 취소")
                return True


            elif key == Qt.Key_Backspace and current_items:
                current_row = current_items[0].row()
                current_col = current_items[0].column()
                current_item = self.edit_table.item(current_row, current_col)
                logger.debug(f"Backspace pressed at {current_row}, {current_col}, text: '{current_item.text() if current_item else None}'")
                if self.is_inserting and (not current_item.text() or current_item.text() == ""):
                    # 삽입된 빈 셀에서 Backspace → 빈 셀 삭제
                    day = self.days[current_col]
                    current_schedule = self.schedule_lists[self.current_list][day]
                    if current_row < len(current_schedule) and not current_schedule[current_row]["time"]:
                        del current_schedule[current_row]
                    self.is_inserting = False
                    self.update_edit_table()
                    self.set_status_message("빈 셀이 삭제되었습니다.")
                else:
                    # 데이터가 있는 셀 → 기존 delete_schedule 호출
                    self.delete_schedule()
                return True

        elif source == self.edit_table and event.type() == QEvent.FocusOut:
            current_row = self.edit_table.currentRow()
            current_col = self.edit_table.currentColumn()
            current_item = self.edit_table.item(current_row, current_col) if current_row >= 0 and current_col >= 0 else None
            editor = self.edit_table.cellWidget(current_row, current_col) if current_row >= 0 else None
            logger.debug(f"FocusOut detected at {current_row}, {current_col}, text: '{current_item.text() if current_item else None}', editor: {bool(editor)}")
            if editor and self.is_inserting and current_item.text():
                self.edit_table.closePersistentEditor(current_item)
                self.is_inserting = False
                self.update_edit_table()
                self.set_status_message(f"시간 입력 완료: {current_item.text()}")
            return True

        return super().eventFilter(source, event)

    def create_list(self):
        name, ok = QInputDialog.getText(self, "새 목록 생성", "새 시간표 목록 이름을 입력하세요:")
        if ok and name and name not in self.schedule_lists:
            self.schedule_lists[name] = {day: [] for day in self.days}
            self.schedule_combo.addItem(name)
            self.schedule_combo.setCurrentText(name)
            self.current_list = name
            self.update_edit_table()
            self.save_config()
            self.set_status_message(f"새 시간표 목록 '{name}'이 생성되었습니다.")
        elif name in self.schedule_lists:
            self.set_status_message("이미 존재하는 이름입니다.")

    def rename_list(self):
        current_name = self.schedule_combo.currentText()
        if current_name == "기본 스케줄":
            self.set_status_message("'기본 스케줄'은 이름을 변경할 수 없습니다.")
            return
        new_name, ok = QInputDialog.getText(self, "이름 변경", "새 이름을 입력하세요:", text=current_name)
        if ok and new_name and new_name not in self.schedule_lists:
            self.schedule_lists[new_name] = self.schedule_lists.pop(current_name)
            self.schedule_combo.setItemText(self.schedule_combo.currentIndex(), new_name)
            self.current_list = new_name
            self.save_config()
            self.set_status_message(f"시간표 목록 이름이 '{new_name}'으로 변경되었습니다.")
        elif new_name in self.schedule_lists:
            self.set_status_message("이미 존재하는 이름입니다.")

    def clone_list(self):
        current_name = self.schedule_combo.currentText()
        new_name, ok = QInputDialog.getText(self, "목록 복제", "복제된 시간표 이름을 입력하세요:", text=f"{current_name} 복사본")
        if ok and new_name and new_name not in self.schedule_lists:
            from copy import deepcopy
            self.schedule_lists[new_name] = deepcopy(self.schedule_lists[current_name])
            self.schedule_combo.addItem(new_name)
            self.schedule_combo.setCurrentText(new_name)
            self.current_list = new_name
            self.update_edit_table()
            self.save_config()
            self.set_status_message(f"'{current_name}'이 '{new_name}'으로 복제되었습니다.")
        elif new_name in self.schedule_lists:
            self.set_status_message("이미 존재하는 이름입니다.")

    def delete_list(self):
        current_name = self.schedule_combo.currentText()
        if current_name == "기본 스케줄":
            self.set_status_message("'기본 스케줄'은 삭제할 수 없습니다.")
            return
        reply = QMessageBox.question(self, "삭제 확인", f"'{current_name}'을 삭제하시겠습니까?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            del self.schedule_lists[current_name]
            self.schedule_combo.removeItem(self.schedule_combo.currentIndex())
            self.current_list = self.schedule_combo.currentText()
            self.update_edit_table()
            self.save_config()
            self.set_status_message(f"'{current_name}'이 삭제되었습니다.")

    def load_list(self, name):
        self.current_list = name
        self.update_edit_table()
        self.save_config()
        self.set_status_message(f"'{name}' 시간표가 로드되었습니다.")

    def reset_schedule(self):
        reply = QMessageBox.question(self, "초기화 확인", "현재 시간표를 초기화하시겠습니까?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.schedule_lists[self.current_list] = {day: [] for day in self.days}
            self.update_edit_table()
            self.update_today_schedule()
            self.save_config()
            self.set_status_message("시간표가 초기화되었습니다.")

    def save_schedule(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "시간표 저장", "", "JSON Files (*.json)")
        if file_name:
            try:
                # 현재 선택된 스케줄 리스트만 추출
                current_schedule = {
                    "schedule": self.schedule_lists[self.current_list]
                }
                with open(file_name, 'w', encoding='utf-8') as f:
                    json.dump(current_schedule, f, ensure_ascii=False, indent=4)
                self.set_status_message(f"시간표가 {file_name}에 저장되었습니다.")
                logger.info(f"Schedule saved to {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"저장 실패: {str(e)}")
                logger.error(f"Failed to save schedule: {str(e)}")

    def load_schedule(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "시간표 불러오기", "", "JSON Files (*.json)")
        if not file_name:
            return

        try:
            with open(file_name, 'r', encoding='utf-8') as f:
                config = json.load(f)
            logger.info(f"Loaded config from {file_name}: {json.dumps(config, ensure_ascii=False)}")

            # "schedule" 키 확인 및 데이터 검증
            if "schedule" not in config or not isinstance(config["schedule"], dict):
                raise ValueError("파일에 'schedule' 데이터가 없거나 형식이 잘못되었습니다.")

            schedule_data = config["schedule"]
            for day in self.days:
                if day not in schedule_data:
                    schedule_data[day] = []  # 누락된 요일은 빈 리스트로 채움
                if not isinstance(schedule_data[day], list):
                    raise ValueError(f"'{day}'의 데이터가 리스트 형식이 아닙니다.")
                for item in schedule_data[day]:
                    if not isinstance(item, dict) or "time" not in item or "desc" not in item or "active" not in item:
                        raise ValueError(f"'{day}'에 잘못된 항목이 있습니다: {item}")

            # 새 시간표 이름 입력받기
            new_name, ok = QInputDialog.getText(self, "시간표 이름", "불러온 시간표의 이름을 입력하세요:", text="새 시간표")
            if not ok or not new_name:
                logger.info("User canceled or provided no name")
                return

            # 중복 이름 처리
            if new_name in self.schedule_lists:
                i = 1
                base_name = new_name
                while f"{base_name} ({i})" in self.schedule_lists:
                    i += 1
                new_name = f"{base_name} ({i})"
                logger.info(f"Name '{base_name}' already exists, using '{new_name}'")

            # 시간표 추가 및 UI 업데이트
            self.schedule_lists[new_name] = schedule_data
            self.schedule_combo.addItem(new_name)
            self.schedule_combo.setCurrentText(new_name)
            self.current_list = new_name

            self.update_edit_table()
            self.update_today_schedule()
            self.set_status_message(f"시간표가 {file_name}에서 '{new_name}'으로 불러와졌습니다.")
            logger.info(f"Schedule loaded as '{new_name}' from {file_name}")

        except json.JSONDecodeError as e:
            logger.error(f"JSON parsing error: {str(e)}")
            QMessageBox.critical(self, "오류", f"JSON 파일 파싱 실패: {str(e)}")
        except ValueError as e:
            logger.error(f"Validation error: {str(e)}")
            QMessageBox.critical(self, "오류", f"데이터 검증 실패: {str(e)}")
        except Exception as e:
            logger.error(f"Unexpected error during load: {str(e)}")
            QMessageBox.critical(self, "오류", f"불러오기 실패: {str(e)}")

    def save_config(self):
        config = {
            "schedule_lists": self.schedule_lists,
            "current_list": self.current_list,
            "start_bell": self.start_bell,
            "end_bell": self.end_bell,
            "misc_bell": self.misc_bell,  # 추가
            "volume": self.volume_slider.value(),
            "autostart": self.autostart_checkbox.isChecked()
        }
        try:
            # 기존 설정과 비교
            should_log = True
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    old_config = json.load(f)
                if old_config == config:
                    should_log = False  # 변경사항 없으면 로그 생략

            if os.path.exists(config_file):
                backup_file = app_data_dir / 'config_backup.json'
                with open(config_file, 'r', encoding='utf-8') as f, open(backup_file, 'w', encoding='utf-8') as bf:
                    bf.write(f.read())
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)

            if should_log:
                logger.info("Config saved successfully")
        except Exception as e:
            logger.error(f"Failed to save config: {str(e)}")
            QMessageBox.critical(self, "오류", "설정 저장에 실패했습니다.")

    def load_config(self):
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            self.schedule_lists = config.get("schedule_lists", {"기본 스케줄": {day: [] for day in self.days}})
            # 빈 항목 제거
            for list_name in self.schedule_lists:
                for day in self.days:
                    self.schedule_lists[list_name][day] = [s for s in self.schedule_lists[list_name][day] if s["time"]]
            self.current_list = config.get("current_list", "기본 스케줄")
            if not self.current_list or self.current_list not in self.schedule_lists:
                self.current_list = "기본 스케줄"
            self.start_bell = config.get("start_bell", "start_bell.wav")
            self.end_bell = config.get("end_bell", "end_bell.wav")
            self.misc_bell = config.get("misc_bell", self.start_bell)  # 기본값은 start_bell
            self.autostart_state = config.get("autostart", False)
            self.volume_value = config.get("volume", 100)
        else:
            self.current_list = "기본 스케줄"
            self.start_bell = resource_path("start_bell.wav")
            self.end_bell = resource_path("end_bell.wav")
            self.misc_bell = self.start_bell  # 기본값 설정
            self.autostart_state = False
            self.volume_value = 100


    def toggle_autostart(self):
        import ctypes
        task_name = "ClassBellApp"
        app_path = os.path.join(os.path.dirname(sys.executable), "class_bell_app.exe")
        if not os.path.exists(app_path):
            if "PyCharm" in sys.executable:
                return
            QMessageBox.critical(self, "오류", "실행 파일을 찾을 수 없습니다.")
            return
        username = getpass.getuser()
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
            if not is_admin:
                QMessageBox.warning(self, "권한 오류", "자동 시작 설정을 위해 관리자 권한이 필요합니다.")
                self.autostart_checkbox.setChecked(False)
                return
            if self.autostart_checkbox.isChecked():
                cmd = [
                    'schtasks', '/create', '/tn', task_name,
                    '/tr', f'"{app_path}"',
                    '/sc', 'onlogon', '/rl', 'highest',
                    '/ru', username, '/delay', '0000:30', '/f'
                ]
                subprocess.run(cmd, check=True, capture_output=True, text=True)
                QMessageBox.information(self, "성공", "프로그램이 부팅 시 자동 실행되도록 설정되었습니다.")
            else:
                cmd = ['schtasks', '/delete', '/tn', task_name, '/f']
                subprocess.run(cmd, check=True, capture_output=True, text=True)
            self.save_config()
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "오류", f"작업 스케줄러 설정 실패: {e.stderr}")

    def check_time(self):
        current_time = datetime.now(self.seoul_tz)
        current_day_idx = current_time.weekday()
        current_day = self.days[current_day_idx]
        current_time_qt = QTime.currentTime()
        current_time_str = current_time_qt.toString("HH:mm")

        # 날짜 변경 감지
        if not hasattr(self, 'last_day_idx'):
            self.last_day_idx = current_day_idx
        if self.last_day_idx != current_day_idx:
            self.update_today_schedule()
            self.last_day_idx = current_day_idx
            logger.info(f"Day changed to {current_day}, updated today schedule")

        if hasattr(self, 'today_header'):
            self.today_header.setText(self.get_today_header())

        # 벨소리 체크 로직
        bell_triggered = False  # 벨이 울렸는지 추적
        for item in self.schedule_lists[self.current_list][current_day]:
            schedule_time = item["time"]
            if item["active"] == "예" and schedule_time == current_time_str and self.last_played_time != schedule_time:
                self.last_played_time = schedule_time
                if "수업시작" in item["desc"] or "교시 시작" in item["desc"]:
                    self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.start_bell))))
                    self.player.play()
                    logger.debug(f"Playing start bell at {current_time_str}")
                    bell_triggered = True
                elif "종료" in item["desc"] or "교시 종료" in item["desc"]:
                    self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.end_bell))))
                    self.player.play()
                    logger.debug(f"Playing end bell at {current_time_str}")
                    bell_triggered = True
                elif "기타시작" in item["desc"]:
                    self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.misc_bell))))
                    self.player.play()
                    logger.debug(f"Playing misc bell at {current_time_str}")
                    bell_triggered = True

        # 벨이 울린 후 즉시 갱신 (추가)
        if bell_triggered:
            self.update_today_schedule()

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.showNormal()
            self.activateWindow()

    def exit_program(self):
        self.save_config()
        self.tray_icon.hide()
        QApplication.quit()

    def handle_link_activated(self, link):
        if link == "showqr://support":
            qr_dialog = SupportQRDialog(self)
            qr_dialog.exec_()
        # 외부 링크는 setOpenExternalLinks(True)로 이미 처리됨

    def keyPressEvent(self, event):
        # Ctrl+C와 Ctrl+V 키 입력 감지
        if event.modifiers() == Qt.ControlModifier:
            if event.key() == Qt.Key_C:
                self.copy_schedule()
            elif event.key() == Qt.Key_V:
                self.paste_schedule()
        super().keyPressEvent(event)

    def copy_schedule(self):
        # 선택된 셀에서 일정 복사
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.set_status_message("복사할 셀이 선택되지 않았습니다.")
            return
        self.copied_data = []
        for item in selected_items:
            row = item.row()
            col = item.column()
            time = item.text()  # 셀에 표시된 시간
            if time:
                day = self.days[col]
                # schedule_lists에서 해당 일정의 설명 찾기
                schedule = next((s for s in self.schedule_lists[self.current_list][day]
                                 if s["time"] == time), None)
                if schedule:
                    self.copied_data.append({
                        "row": row,
                        "col": col,
                        "time": schedule["time"],
                        "desc": schedule["desc"]
                    })
        if self.copied_data:
            self.set_status_message(f"{len(self.copied_data)}개의 일정이 복사되었습니다.")
        else:
            self.set_status_message("복사할 유효한 일정이 없습니다.")

    def paste_schedule(self):
        # 복사된 일정을 선택된 위치에 붙여넣기
        if not self.copied_data:
            self.set_status_message("붙여넣을 데이터가 없습니다.")
            return
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.set_status_message("붙여넣을 위치가 선택되지 않았습니다.")
            return

        # 붙여넣기 기준 위치 (첫 번째 선택된 셀)
        base_row = selected_items[0].row()
        base_col = selected_items[0].column()

        # 복사된 데이터의 상대적 위치 기준 계산
        min_row = min(data["row"] for data in self.copied_data)
        min_col = min(data["col"] for data in self.copied_data)

        for data in self.copied_data:
            rel_row = data["row"] - min_row
            target_row = base_row + rel_row
            target_col = base_col  # 붙여넣기는 선택된 열(요일)에만 적용

            if target_row >= self.edit_table.rowCount() or target_col >= len(self.days):
                continue  # 테이블 범위를 벗어나면 무시

            day = self.days[target_col]
            time = data["time"]
            desc = data["desc"]

            # 기존 일정 확인 (빈 칸에만 붙여넣기)
            existing = next((s for s in self.schedule_lists[self.current_list][day]
                             if s["time"] == time), None)
            if not existing:
                self.schedule_lists[self.current_list][day].append({
                    "time": time,
                    "desc": desc,
                    "active": "예"
                })
                self.schedule_lists[self.current_list][day].sort(key=lambda x: x["time"])

        # 테이블 갱신
        self.update_edit_table()
        self.set_status_message(f"{len(self.copied_data)}개의 일정이 붙여넣어졌습니다.")

    def closeEvent(self, event):
        """창 닫기 버튼을 눌렀을 때 트레이로 최소화"""
        event.ignore()  # 기본 종료 이벤트 무시
        self.hide()  # 창 숨기기
        self.tray_icon.showMessage(
            "오늘은 v1.2",
            "프로그램이 트레이로 최소화되었습니다. 종료하려면 트레이 아이콘을 우클릭 후 '종료하기'를 선택하세요.",
            QSystemTrayIcon.Information,
            2000
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ClassBellApp()
    sys.exit(app.exec_())