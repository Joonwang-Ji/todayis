import sys
import getpass
import subprocess
import json
import os
import winreg
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
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon, QKeyEvent, QGuiApplication, QRegExpValidator
import pytz
import re
import requests  # 날씨 API 호출용

# 로그 파일 설정 (파일 상단에 배치)
app_data_dir = Path(os.getenv('APPDATA')) / '오늘은'
app_data_dir.mkdir(exist_ok=True)
log_file = app_data_dir / 'app.log'
config_file = app_data_dir / 'config.json'

handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=10, encoding="utf-8")
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger = logging.getLogger()
logger.setLevel(logging.INFO)  # 배포용: INFO로 변경
logger.handlers.clear()
logger.addHandler(handler)

def resource_path(relative_path):
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
        config_path = os.path.join(os.path.dirname(sys.executable), "config.json")
        config = {}
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
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
    def __init__(self, start_bell, end_bell, parent=None):
        super().__init__(parent)
        self.setWindowTitle("벨소리 설정")
        self.setFixedSize(400, 150)
        self.start_bell_input = QLineEdit(self)
        self.end_bell_input = QLineEdit(self)
        self.start_bell_input.setText(start_bell)
        self.end_bell_input.setText(end_bell)

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
        start_row.setAlignment(Qt.AlignVCenter)
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
        end_row.setAlignment(Qt.AlignVCenter)
        layout.addLayout(end_row)

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

    def get_bells(self):
        start_bell = self.start_bell_input.text().strip()
        end_bell = self.end_bell_input.text().strip()
        return start_bell, end_bell

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

class ClassBellApp(QMainWindow):
    def __init__(self):
        self.shared_mem = QSharedMemory("ClassBellAppUniqueKey")
        if not self.shared_mem.create(1):
            QMessageBox.information(None, "오늘은 v1.1", "프로그램이 이미 실행 중입니다.")
            sys.exit(0)

        super().__init__()
        self.setWindowTitle("오늘은 v1.1")
        self.setFixedSize(800, 700)
        self.setWindowIcon(QIcon(resource_path("icon.ico")))

        screen = QGuiApplication.primaryScreen().geometry()
        self.move((screen.width() - 800) // 2, (screen.height() - 700) // 2)

        self.days = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
        self.schedule_lists = {"기본 스케줄": {day: [] for day in self.days}}
        self.start_bell = resource_path("start_bell.wav")  # 기본값 번들링
        self.end_bell = resource_path("end_bell.wav")  # 기본값 번들링
        self.last_played_time = None

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_time)
        self.timer.start(1000)  # 500ms → 1000ms로 변경

        self.daily_timer = QTimer(self)
        self.daily_timer.timeout.connect(self.update_today_schedule)
        self.daily_timer.start(24 * 60 * 60 * 1000)

        self.seoul_tz = pytz.timezone("Asia/Seoul")
        self.player = QMediaPlayer()
        self.player.stateChanged.connect(self.on_player_state_changed)

        self.tray_icon = QSystemTrayIcon(QIcon(resource_path("icon.ico")), self)
        self.tray_icon.setToolTip("오늘은 v1.1")
        tray_menu = QMenu()
        tray_menu.addAction("보이기", self.show)
        tray_menu.addAction("종료하기", self.exit_program)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        self.tray_icon.activated.connect(self.on_tray_icon_activated)

        self.weather_data = "날씨 정보를 불러오는 중..."
        self.weather_timer = QTimer(self)
        self.weather_timer.timeout.connect(self.update_weather)
        self.weather_timer.start(24 * 60 * 60 * 1000)  # 10분 → 24시간으로 변경

        self.load_config()
        self.setup_ui()
        self.apply_style()
        self.edit_table.installEventFilter(self)
        self.update_edit_table()
        self.update_weather()
        self.show()

    def fetch_weather(self):
        logger.debug("Starting fetch_weather")
        cache_file = app_data_dir / "weather_cache.json"
        cache_duration = timedelta(hours=24)

        # 캐시 확인
        if os.path.exists(cache_file):
            try:
                with open(cache_file, "r", encoding="utf-8") as f:
                    cache_data = json.load(f)
                last_updated = datetime.fromisoformat(cache_data["timestamp"])
                if datetime.now() - last_updated < cache_duration:
                    self.weather_data = cache_data["weather_data"]
                    logger.info(f"Using cached weather: {self.weather_data}")
                    return
            except Exception as e:
                logger.error(f"Cache read error: {str(e)}")

        # API 호출 (최대 3회 재시도)
        api_key = "UVBfVbhomHosY6RfYywnTw3LYQ3IoKWeDgIpcEM%2Fs3zqYABIXGRSMEggQ37qCVDaWgRwasS5GSpDzYkV17zTtQ%3D%3D"
        retries = 3
        for attempt in range(retries):
            try:
                # IP로 위치 가져오기
                ip_response = requests.get("https://ipinfo.io/json", timeout=5)
                ip_data = ip_response.json()
                city_from_ip = ip_data.get("city", "Seoul")

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
                }  # 기존 city_map 유지
                city_info = city_map.get(city_from_ip, {"name": "서울특별시", "nx": 60, "ny": 127})
                city_name, nx, ny = city_info["name"], city_info["nx"], city_info["ny"]

                now = datetime.now()
                base_date = now.strftime("%Y%m%d")
                base_time_str = "0500"
                real_time = now - timedelta(minutes=30)
                real_date = real_time.strftime("%Y%m%d")
                real_hour = real_time.strftime("%H00")

                # 단기예보
                url_fcst = f"http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getVilageFcst?serviceKey={api_key}&numOfRows=1000&pageNo=1&base_date={base_date}&base_time={base_time_str}&nx={nx}&ny={ny}&dataType=JSON"
                response_fcst = requests.get(url_fcst, timeout=10)
                response_fcst.raise_for_status()
                data_fcst = response_fcst.json()
                if data_fcst["response"]["header"]["resultCode"] != "00":
                    raise ValueError(f"Forecast API Error: {data_fcst['response']['header']['resultMsg']}")

                # 초단기실황
                url_real = f"http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getUltraSrtNcst?serviceKey={api_key}&numOfRows=10&pageNo=1&base_date={real_date}&base_time={real_hour}&nx={nx}&ny={ny}&dataType=JSON"
                response_real = requests.get(url_real, timeout=10)
                response_real.raise_for_status()
                data_real = response_real.json()
                if data_real["response"]["header"]["resultCode"] != "00":
                    raise ValueError(f"Real-time API Error: {data_real['response']['header']['resultMsg']}")

                # 데이터 파싱 (별도 함수로 분리)
                temp, temp_min, temp_max, weather_desc, icon = self.parse_weather_data(data_real, data_fcst, base_date, now.hour)
                self.weather_data = f"{icon} {city_name} 날씨: {weather_desc}, {temp:.1f}°C (최저 {temp_min:.1f}°C / 최고 {temp_max:.1f}°C)"

                # 캐시 저장
                cache_data = {"timestamp": datetime.now().isoformat(), "weather_data": self.weather_data, "min_max": (temp_min, temp_max)}
                with open(cache_file, "w", encoding="utf-8") as f:
                    json.dump(cache_data, f, ensure_ascii=False)
                logger.info(f"Weather updated: {self.weather_data}")
                break
            except (requests.RequestException, ValueError) as e:
                logger.error(f"Weather fetch attempt {attempt + 1} failed: {str(e)}")
                if attempt == retries - 1:
                    self.weather_data = "⚠️ 날씨 정보를 가져올 수 없습니다."
                    logger.error("All retries failed")

    def parse_weather_data(self, data_real, data_fcst, base_date, current_hour):
        # 실황 데이터
        items_real = data_real["response"]["body"]["items"]["item"]
        temp = pty = r06 = wsd = reh = None
        for item in items_real:
            if item["category"] == "T1H":
                temp = float(item["obsrValue"])
            elif item["category"] == "PTY":
                pty = item["obsrValue"]
            elif item["category"] == "RN1":
                r06 = float(item["obsrValue"]) if item["obsrValue"] != "강수없음" else 0
            elif item["category"] == "WSD":
                wsd = float(item["obsrValue"])
            elif item["category"] == "REH":
                reh = float(item["obsrValue"])

        # 예보 데이터
        items_fcst = data_fcst["response"]["body"]["items"]["item"]
        all_temps = []
        sky = s06 = None
        for item in items_fcst:
            if item["fcstDate"] == base_date:
                if item["category"] == "TMP": all_temps.append(float(item["fcstValue"]))
                if int(item["fcstTime"][:2]) >= current_hour:
                    if item["category"] == "SKY" and sky is None: sky = item["fcstValue"]
                    if item["category"] == "PTY" and pty is None: pty = item["fcstValue"]
                    if item["category"] == "R06" and r06 is None: r06 = float(item["fcstValue"])
                    if item["category"] == "S06" and s06 is None: s06 = float(item["fcstValue"])
                    if item["category"] == "WSD" and wsd is None: wsd = float(item["fcstValue"])
                    if item["category"] == "REH" and reh is None: reh = float(item["fcstValue"])

        if temp is None or not all_temps:
            raise ValueError("Insufficient weather data")

        temp_min, temp_max = min(all_temps), max(all_temps)
        if temp < temp_min: temp_min = temp
        if temp > temp_max: temp_max = temp

        # 날씨 설명
        pty_map = {"0": "", "1": "비", "2": "비/눈", "3": "눈", "4": "소나기"}
        sky_map = {"1": "맑음", "3": "구름많음", "4": "흐림"}
        base_desc = pty_map.get(pty, sky_map.get(sky, "알 수 없음")) if pty != "0" else sky_map.get(sky, "알 수 없음")
        weather_desc = base_desc
        if pty == "1" and r06:
            weather_desc = "이슬비" if r06 < 1 else "약한 비" if r06 < 5 else "강한 비" if r06 > 20 else "비"
        elif pty == "3" and s06:
            weather_desc = "약한 눈" if s06 < 5 else "폭설" if s06 > 20 else "눈"
        elif pty == "4" and r06:
            weather_desc = "약한 소나기" if r06 < 5 else "강한 소나기" if r06 > 20 else "소나기"
        if wsd:
            if wsd > 14:
                weather_desc += " (강풍)"
            elif pty == "0" and sky != "1" and wsd > 7:
                weather_desc = f"바람부는 {weather_desc}"
        if reh and reh > 90 and wsd and wsd < 3 and pty == "0": weather_desc = "안개" if base_desc == "맑음" else f"{base_desc} 속 안개"

        icon = "☀️" if "맑음" in weather_desc and "안개" not in weather_desc else "☁️" if "흐림" in weather_desc or "구름" in weather_desc else "🌧️" if "비" in weather_desc or "소나기" in weather_desc else "❄️" if "눈" in weather_desc else "🌫️" if "안개" in weather_desc else "💨" if "바람" in weather_desc or "강풍" in weather_desc else "⚠️"
        return temp, temp_min, temp_max, weather_desc, icon

    def update_weather(self):
        self.fetch_weather()
        if hasattr(self, 'weather_label'):
            self.weather_label.setText(self.weather_data)

    def on_player_state_changed(self, state):
        if state == QMediaPlayer.StoppedState:
            pass  # 재생 완료 시 추가 처리 필요 시 여기에 로직 추가

    def apply_config(self, config):
        self.schedule_lists = config.get("schedule_lists", {"기본 스케줄": {day: [] for day in self.days}})
        self.current_list = config.get("current_list", "기본 스케줄")
        self.start_bell = config.get("start_bell", "start_bell.wav")
        self.end_bell = config.get("end_bell", "end_bell.wav")
        self.volume_slider.setValue(config.get("volume", 100))
        self.autostart_checkbox.setChecked(config.get("autostart", False))
        self.update_edit_table()

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
        self.today_table.setStyleSheet("""
            QTableWidget::item {
                padding:8px;
            }
        """)
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

        # 날씨 라벨 개선
        self.weather_label = QLabel(self.weather_data)
        self.weather_label.setAlignment(Qt.AlignCenter)
        self.weather_label.setFont(QFont("맑은 고딕", 10))  # 폰트 크기 줄임 (16 → 12)
        self.weather_label.setMaximumWidth(540)  # 헤더와 동일한 너비
        self.weather_label.setStyleSheet("""
                    background-color: #F7F9FA;  /* 부드러운 연한 회색-파랑 */
                    border: 1px solid #E0E4E8;  /* 얇고 깔끔한 테두리 */
                    border-radius: 8px;
                    padding: 8px;
                    color: #4A5A6A;  /* 차분한 다크 그레이-블루 */
                """)
        right_layout.addWidget(self.weather_label)
        right_layout.addSpacerItem(QSpacerItem(20, 20))


        right_layout.addSpacerItem(QSpacerItem(20, 10))
        volume_layout = QHBoxLayout()
        volume_label = QLabel("볼륨")
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setValue(self.volume_value)
        self.volume_slider.setMaximumWidth(400)
        self.volume_slider.valueChanged.connect(self.set_volume)
        self.volume_slider.setToolTip("볼륨 조절")
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

        right_layout.addSpacerItem(QSpacerItem(20, 20))

        info_text = (
            "이 프로그램은 대전소재의 기독교 대안학교 '노엠스쿨'(여자 중고등학교)의 교사가 학교의 필요에 의해 제작한 프로그램입니다.<br>"
            "무료로 사용하셔도 되며 아무쪼록 좋은 곳에 널리 쓰이길 바라며 배포합니다.<br>"
            "- 2025년 3월의 어느 오늘<br><br>"
            "* 프로그램 문의: <a href='mailto:todayis@j2w.kr'>todayis@j2w.kr</a><br>"
            "* 노엠스쿨: <a href='https://noemschool.org'>https://noemschool.org</a>"
        )
        self.info_label = QLabel(info_text, self.today_widget)
        self.info_label.setFont(QFont("맑은 고딕", 11))
        self.info_label.setAlignment(Qt.AlignLeft)
        self.info_label.setWordWrap(True)
        self.info_label.setMaximumWidth(540)
        self.info_label.setOpenExternalLinks(True)
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
        self.edit_table.setEditTriggers(QTableWidget.NoEditTriggers)
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
            self.edit_table.setColumnWidth(i, 101)
        self.edit_table.setMinimumHeight(500)
        self.edit_table.horizontalHeader().setFixedHeight(40)
        self.edit_table.verticalHeader().setDefaultSectionSize(25)
        self.edit_table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.edit_table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
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
        <h2>🌟 오늘은 v1.1 사용설명서</h2>
        <p>안녕하세요, 선생님! "오늘은"은 수업 시간에 맞춰 벨을 울리고 시간표를 관리하는 간단한 프로그램이에요. 처음 사용하셔도 쉽게 익힐 수 있도록 설명드릴게요!</p>

        <h3>1. 프로그램 시작하기</h3>
        <ul>
            <li><b>실행</b>: 바탕화면 아이콘을 더블클릭하거나, 컴퓨터를 켜면 자동으로 시작돼요 (설정 필요 시 아래 참고).</li>
            <li><b>창</b>: 프로그램이 열리면 세 개의 탭이 보여요: "오늘 시간표", "시간표 관리", "사용설명서" (지금 여기!).</li>
        </ul>

        <h3>2. 주요 기능</h3>
        <p>탭마다 할 수 있는 일을 알려드릴게요!</p>

        <h4>🎯 첫 번째 탭: 오늘 시간표</h4>
        <ul>
            <li><b>시간 확인</b>: 오늘 날짜와 시간이 커다랗게 표시돼요.</li>
            <li><b>오늘 일정</b>: 왼쪽에 오늘 울릴 시간과 "수업시작", "종료" 등이 보여요.</li>
            <li><b>🎵 벨소리 재생</b>: 버튼을 누르면 시작 벨을 테스트할 수 있어요.</li>
            <li><b>🔔 벨소리 변경</b>: 시작과 종료 벨을 원하는 소리로 바꿀 수 있어요 (mp3, wav 파일 선택).</li>
            <li><b>🔊 볼륨</b>: 슬라이더를 움직여 소리 크기를 조절해요 (0~100%).</li>
            <li><b>🚀 자동 시작</b>: 체크하면 컴퓨터 켤 때마다 프로그램이 자동 실행돼요.</li>
            <li><b>🚪 종료</b>: 버튼을 누르면 프로그램이 완전히 꺼져요.</li>
        </ul>

        <h4>📅 두 번째 탭: 시간표 관리</h4>
        <ul>
            <li><b>시간표 목록</b>: "기본 스케줄"이 기본이에요. 새로 만들거나 복제, 이름 변경, 삭제도 가능해요!</li>
            <ul>
                <li>📝 <b>목록생성</b>: 새 시간표를 만들어요.</li>
                <li>🔧 <b>이름변경</b>: 현재 시간표 이름을 바꿔요.</li>
                <li>📋 <b>복제</b>: 현재 시간표를 똑같이 복사해요.</li>
                <li>🗑️ <b>삭제</b>: 필요 없는 시간표를 지워요 ("기본 스케줄"은 못 지움).</li>
            </ul>
            <li><b>시간 입력</b>: 테이블에서 요일별 시간을 입력해요.</li>
            <ul>
                <li>더블클릭하거나 Enter 키로 셀을 편집해요.</li>
                <li>형식: "HH:MM" (예: 09:00) 또는 "HHMM" (예: 0900).</li>
                <li>입력 후 Enter → 다음 줄로 이동하며 계속 추가 가능.</li>
            </ul>
            <li><b>시간 설정</b>: 셀을 선택하고 우클릭하거나 단축키로 설정해요.</li>
            <ul>
                <li><b>S 키</b>: "수업시작"으로 설정 (🟡 노란색).</li>
                <li><b>D 키</b>: "종료"로 설정 (⚪ 회색).</li>
                <li><b>R 키</b>: "기타시작"으로 설정 (🟢 연두색).</li>
                <li><b>Backspace 키</b>: 선택한 시간 삭제.</li>
                <li>우클릭 메뉴: "삭제", "아래추가", "시작/종료/기타시작" 선택.</li>
            </ul>
            <li><b>🔄 반복 스케줄 생성</b>: 똑같은 패턴으로 여러 시간을 추가해요.</li>
            <ul>
                <li>시작 시간 (예: 09:00), 수업 시간 (예: 60분), 쉬는 시간 (예: 10분), 반복 횟수, 요일을 선택.</li>
                <li>예: 09:00 시작, 60분 수업, 10분 쉬고, 5번 반복 → 5교시 자동 생성!</li>
            </ul>
            <li><b>♻️ 스케줄 초기화</b>: 현재 시간표를 모두 지워요.</li>
            <li><b>💾 파일 저장 / 📂 파일 불러오기</b>: 시간표를 파일로 저장하거나 불러와요.</li>
        </ul>

        <h4>📖 세 번째 탭: 사용설명서</h4>
        <ul>
            <li>지금 보고 계신 이 설명서예요! 스크롤해서 읽어보세요.</li>
        </ul>

        <h3>3. 간단 사용법</h3>
        <ol>
            <li><b>시간표 만들기</b>: "시간표 관리" 탭 → 시간 입력 → S/D/R 키로 시작/종료 설정.</li>
            <li><b>벨 확인</b>: "오늘 시간표" 탭 → "벨소리 재생"으로 소리 테스트.</li>
            <li><b>자동 실행</b>: "윈도우 부팅시 자동 시작" 체크는 관리자 모드로 실행해야 설정돼요 (우클릭 → "관리자 권한으로 실행"). 일반 모드에서는 체크가 안 돼요.</li>
            <li><b>최소화</b>: 창을 최소화하면 오른쪽 아래 트레이에 🔔 아이콘이 생겨요. 더블클릭하면 다시 열림.</li>
        </ol>

        <h3>4. 유용한 팁</h3>
        <ul>
            <li>🕛 <b>시간</b>: 밤 12시는 "00:00"으로 입력하세요.</li>
            <li>🎶 <b>벨소리</b>: mp3나 wav 파일을 준비해서 "벨소리 변경"으로 설정해 보세요.</li>
            <li>💻 <b>컴퓨터</b>: 노트북이라면 전원을 연결해 두세요.</li>
            <li>🚀 <b>자동 시작</b>: 체크는 관리자 모드로 실행 시 설정 가능. 일반 모드에서는 상태만 보여요.</li>
            <li>❓ <b>문제 시</b>: "종료" 후 다시 켜보세요. 안 되면 <a href="mailto:todayis@j2w.kr">todayis@j2w.kr</a>로 문의!</li>
        </ul>

        <p>이 프로그램은 노엠스쿨 교사가 선생님들과 학생들들을 위해 만든 무료 도구예요. 편리하게 사용해 주세요! 🌈<br>
        노엠스쿨 홈페이지: <a href="https://noemschool.org">https://noemschool.org</a></p>
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
            if item["active"] == "예" and schedule_time >= current_time:
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
            sorted_schedule = sorted(self.schedule_lists[self.current_list][day], key=lambda x: x["time"])
            for row in range(max_rows):
                if row < len(sorted_schedule):
                    item = sorted_schedule[row]
                    table_item = QTableWidgetItem(item["time"])
                    table_item.setTextAlignment(Qt.AlignCenter)
                    if "수업시작" in item["desc"] or "교시 시작" in item["desc"]:
                        start_icon_path = resource_path("start_icon.png")
                        if os.path.exists(start_icon_path):
                            table_item.setIcon(QIcon(start_icon_path))
                        table_item.setBackground(QColor("#FFFF99"))
                    elif "종료" in item["desc"] or "교시 종료" in item["desc"]:
                        end_icon_path = resource_path("end_icon.png")
                        if os.path.exists(end_icon_path):
                            table_item.setIcon(QIcon(end_icon_path))
                        table_item.setBackground(QColor("#F5F5F5"))
                    elif "기타시작" in item["desc"]:
                        start_icon_path = resource_path("start_icon.png")
                        if os.path.exists(start_icon_path):
                            table_item.setIcon(QIcon(start_icon_path))
                        table_item.setBackground(QColor("#CCFFCC"))
                    self.edit_table.setItem(row, col, table_item)
                else:
                    empty_item = QTableWidgetItem("")
                    empty_item.setTextAlignment(Qt.AlignCenter)
                    self.edit_table.setItem(row, col, empty_item)
        self.edit_table.setStyleSheet("QTableWidget::item { padding: 0px; }")
        self.edit_table.blockSignals(False)

    def play_start_bell(self):
        self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.start_bell))))
        self.player.play()

    def set_volume(self, value):
        self.player.setVolume(value)
        self.volume_value.setText(f"{value}%")

    def show_context_menu(self, pos):
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.status_label.setText("선택된 셀이 없습니다.")
            return
        menu = QMenu(self)
        delete_action = QAction("삭제", self)
        delete_action.triggered.connect(self.delete_schedule)
        menu.addAction(delete_action)
        add_below_action = QAction("아래추가", self)
        add_below_action.triggered.connect(self.add_below_schedule)
        menu.addAction(add_below_action)
        set_start_action = QAction("시작으로 설정", self)
        set_start_action.triggered.connect(lambda: self.set_schedule_type(selected_items, "수업시작"))
        menu.addAction(set_start_action)
        set_end_action = QAction("종료로 설정", self)
        set_end_action.triggered.connect(lambda: self.set_schedule_type(selected_items, "종료"))
        menu.addAction(set_end_action)
        set_misc_action = QAction("기타시작으로 설정", self)
        set_misc_action.triggered.connect(lambda: self.set_schedule_type(selected_items, "기타시작"))
        menu.addAction(set_misc_action)
        menu.exec_(self.edit_table.viewport().mapToGlobal(pos))

    def set_schedule_type(self, items, desc_type):
        for item in items:
            row = item.row()
            col = item.column()
            day = self.days[col]
            time = item.text()
            if time and QTime.fromString(time, "HH:mm").isValid():
                found = False
                for i, schedule in enumerate(self.schedule_lists[self.current_list][day]):
                    if schedule["time"] == time:
                        self.schedule_lists[self.current_list][day][i]["desc"] = desc_type
                        found = True
                        break
                if not found:
                    self.schedule_lists[self.current_list][day].append({"time": time, "desc": desc_type, "active": "예"})
        self.update_edit_table()
        self.save_config()
        self.status_label.setText(f"{desc_type} 설정 완료")

    def delete_schedule(self):
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.status_label.setText("삭제할 셀이 선택되지 않았습니다.")
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
        self.status_label.setText("선택된 시간이 삭제되었습니다.")
        self.edit_table.blockSignals(False)

    def add_below_schedule(self):
        selected_items = self.edit_table.selectedItems()
        if not selected_items:
            self.status_label.setText("추가할 위치가 선택되지 않았습니다.")
            return
        item = selected_items[0]
        row = item.row()
        col = item.column()

        self.edit_table.blockSignals(True)
        self.edit_table.insertRow(row + 1)
        new_item = QTableWidgetItem("")
        new_item.setTextAlignment(Qt.AlignCenter)
        self.edit_table.setItem(row + 1, col, new_item)
        self.edit_table.setCurrentCell(row + 1, col)
        self.edit_table.editItem(self.edit_table.item(row + 1, col))
        self.edit_table.blockSignals(False)
        self.status_label.setText("아래에 새 시간이 추가되었습니다.")

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
                self.status_label.setText("잘못된 시작 시간입니다.")
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
            self.status_label.setText(f"선택된 요일에 반복 스케줄이 추가되었습니다.")

    def on_item_changed(self, item):
        row = item.row()
        col = item.column()
        new_time = item.text().strip()
        day = self.days[col]

        if not new_time:
            self.status_label.setText("빈 입력은 저장되지 않습니다.")
            return

        formatted_time = None
        if len(new_time) == 4 and new_time.isdigit():
            formatted_time = f"{new_time[:2]}:{new_time[2:]}"
            if not QTime.fromString(formatted_time, "HH:mm").isValid():
                item.setText("")
                self.status_label.setText(f"유효하지 않은 시간: {new_time}")
                return
            item.setText(formatted_time)
        elif new_time and QTime.fromString(new_time, "HH:mm").isValid():
            formatted_time = new_time
        else:
            item.setText("")
            self.status_label.setText(f"시간 형식이 잘못됨: {new_time}")
            return

        if formatted_time and row < len(self.schedule_lists[self.current_list][day]):
            old_time = self.schedule_lists[self.current_list][day][row]["time"]
            old_desc = self.schedule_lists[self.current_list][day][row]["desc"]
            if old_time != formatted_time:
                self.schedule_lists[self.current_list][day] = [s for s in self.schedule_lists[self.current_list][day] if s["time"] != old_time]
                if not any(s["time"] == formatted_time for s in self.schedule_lists[self.current_list][day]):
                    self.schedule_lists[self.current_list][day].append({"time": formatted_time, "desc": old_desc, "active": "예"})
                    self.schedule_lists[self.current_list][day].sort(key=lambda x: x["time"])
            self.save_config()
            self.status_label.setText(f"시간 수정: {formatted_time}")
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
                                    self.status_label.setText(f"유효하지 않은 시간: {new_text}")
                                    return True
                            elif new_text and QTime.fromString(new_text, "HH:mm").isValid():
                                formatted_time = new_text
                            else:
                                current_item.setText("")
                                self.status_label.setText(f"시간 형식이 잘못됨: {new_text}")
                                return True

                            day = self.days[current_col]
                            current_schedule = self.schedule_lists[self.current_list][day]
                            old_time = current_item.text() if current_item.text() else None
                            if old_time and old_time != formatted_time:
                                current_schedule[:] = [s for s in current_schedule if s["time"] != old_time]

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
                            self.save_config()
                            self.update_edit_table()
                            self.update_today_schedule()
                            self.status_label.setText(f"시간 업데이트: {formatted_time}")

                    next_row = current_row + 1
                    if next_row >= self.edit_table.rowCount():
                        self.edit_table.insertRow(next_row)
                        for col in range(self.edit_table.columnCount()):
                            empty_item = QTableWidgetItem("")
                            empty_item.setTextAlignment(Qt.AlignCenter)
                            self.edit_table.setItem(next_row, col, empty_item)
                    self.edit_table.setCurrentCell(next_row, current_col)
                    self.edit_table.editItem(self.edit_table.item(next_row, current_col))
                    return True

            elif key == Qt.Key_S and current_items:
                self.set_schedule_type(current_items, "수업시작")
                self.update_edit_table()
                self.status_label.setText("S 키: 수업시작 설정")
                return True

            elif key == Qt.Key_D and current_items:
                self.set_schedule_type(current_items, "종료")
                self.update_edit_table()
                self.status_label.setText("D 키: 종료 설정")
                return True

            elif key == Qt.Key_R and current_items:
                self.set_schedule_type(current_items, "기타시작")
                self.update_edit_table()
                self.status_label.setText("R 키: 기타시작 설정")
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
                self.status_label.setText(f"숫자 입력 시작: {event.text()}")
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
                self.status_label.setText(f"콜론 입력: :, 셀 값: {current_items[0].text()}")
                return True

            elif key == Qt.Key_Escape and current_items:
                current_row = current_items[0].row()
                current_col = current_items[0].column()
                current_item = self.edit_table.item(current_row, current_col)
                editor = self.edit_table.cellWidget(current_row, current_col)
                if editor and current_item:
                    old_time = current_item.text()
                    self.edit_table.closePersistentEditor(current_item)
                    current_item.setText(old_time if old_time else "")
                    self.status_label.setText(f"편집 취소: 이전 값 '{old_time}' 복원")
                return True

            elif key == Qt.Key_Backspace and current_items:
                self.delete_schedule()
                return True

        return super().eventFilter(source, event)

    def save_schedule(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "시간표 저장", "", "JSON Files (*.json)")
        if file_name:
            config = {
                "schedule_lists": self.schedule_lists,
                "current_list": self.current_list,
                "start_bell": self.start_bell,
                "end_bell": self.end_bell,
                "volume": self.volume_slider.value(),
                "autostart": self.autostart_checkbox.isChecked()
            }
            try:
                with open(file_name, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=4)
                self.status_label.setText(f"시간표가 {file_name}에 저장되었습니다.")
                logging.info(f"Schedule saved to {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"저장 실패: {str(e)}")
                logging.error(f"Failed to save schedule: {str(e)}")

    def load_schedule(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "시간표 불러오기", "", "JSON Files (*.json)")
        if file_name:
            try:
                with open(file_name, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.schedule_lists = config.get("schedule_lists", {"기본 스케줄": {day: [] for day in self.days}})
                self.current_list = config.get("current_list", "기본 스케줄")
                if self.current_list not in self.schedule_lists:
                    self.current_list = "기본 스케줄"
                self.schedule_combo.clear()
                self.schedule_combo.addItems(self.schedule_lists.keys())
                self.schedule_combo.setCurrentText(self.current_list)
                self.start_bell = config.get("start_bell", "start_bell.wav")
                self.end_bell = config.get("end_bell", "end_bell.wav")
                self.volume_slider.setValue(config.get("volume", 100))
                self.autostart_checkbox.setChecked(config.get("autostart", False))
                self.update_edit_table()
                self.update_today_schedule()
                self.save_config()
                self.status_label.setText(f"시간표가 {file_name}에서 불러와졌습니다.")
                logging.info(f"Schedule loaded from {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"불러오기 실패: {str(e)}")
                logging.error(f"Failed to load schedule: {str(e)}")

    def save_config(self):
        config = {
            "schedule_lists": self.schedule_lists,
            "current_list": self.current_list,
            "start_bell": self.start_bell,
            "end_bell": self.end_bell,
            "volume": self.volume_slider.value(),
            "autostart": self.autostart_checkbox.isChecked()
        }
        try:
            if os.path.exists(config_file):
                backup_file = app_data_dir / 'config_backup.json'
                with open(config_file, 'r', encoding='utf-8') as f, open(backup_file, 'w', encoding='utf-8') as bf:
                    bf.write(f.read())
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            logger.info("Config saved successfully")
        except Exception as e:
            logger.error(f"Failed to save config: {str(e)}")
            QMessageBox.critical(self, "오류", "설정 저장에 실패했습니다.")

    def load_config(self):
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            self.schedule_lists = config.get("schedule_lists", {"기본 스케줄": {day: [] for day in self.days}})
            self.current_list = config.get("current_list", "기본 스케줄")
            if not self.current_list or self.current_list not in self.schedule_lists:
                self.current_list = "기본 스케줄"
            self.start_bell = config.get("start_bell", "start_bell.wav")
            self.end_bell = config.get("end_bell", "end_bell.wav")
            self.autostart_state = config.get("autostart", False)
            self.volume_value = config.get("volume", 100)
        else:
            self.current_list = "기본 스케줄"
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
        current_time_str = current_time.strftime("%H:%M")

        self.today_header.setText(self.get_today_header())
        self.update_today_schedule()

        if self.player.state() == QMediaPlayer.PlayingState:
            return

        for item in self.schedule_lists[self.current_list][current_day]:
            if item["active"] != "예":
                continue
            event_time = QTime.fromString(item["time"], "HH:mm")
            if not event_time.isValid():
                continue
            time_diff = abs(event_time.secsTo(current_time_qt))
            if time_diff <= 2:
                if self.last_played_time == (current_time_str, item["desc"]):
                    continue
                if "수업시작" in item["desc"] or "교시 시작" in item["desc"]:
                    self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.start_bell))))
                    self.player.play()
                    self.last_played_time = (current_time_str, item["desc"])
                elif "종료" in item["desc"] or "교시 종료" in item["desc"]:
                    self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.end_bell))))
                    self.player.play()
                    self.last_played_time = (current_time_str, item["desc"])
                elif "기타시작" in item["desc"]:
                    self.player.setMedia(QMediaContent(QUrl.fromLocalFile(resource_path(self.start_bell))))
                    self.player.play()
                    self.last_played_time = (current_time_str, item["desc"])
                break

    def create_list(self):
        try:
            list_name, ok = QInputDialog.getText(self, "새 시간표 목록", "새로운 시간표 목록의 이름을 입력하세요:")
            if ok and list_name:
                if list_name in self.schedule_lists:
                    QMessageBox.warning(self, "오류", "이미 존재하는 이름입니다.")
                    return
                self.schedule_lists[list_name] = {day: [] for day in self.days}
                self.schedule_combo.addItem(list_name)
                self.schedule_combo.setCurrentText(list_name)
                self.current_list = list_name
                self.update_edit_table()
                self.save_config()
                self.status_label.setText(f"새 시간표 '{list_name}'이(가) 생성되었습니다.")
        except Exception as e:
            logger.error(f"Create list failed: {str(e)}")
            QMessageBox.critical(self, "오류", f"시간표 생성 실패: {str(e)}")

    def rename_list(self):
        if self.current_list == "기본 스케줄":
            QMessageBox.warning(self, "오류", "'기본 스케줄'은 이름을 변경할 수 없습니다.")
            return
        new_name, ok = QInputDialog.getText(self, "시간표 이름 변경", "새로운 이름을 입력하세요:", text=self.current_list)
        if ok and new_name:
            if new_name in self.schedule_lists:
                QMessageBox.warning(self, "오류", "이미 존재하는 이름입니다.")
                return
            self.schedule_lists[new_name] = self.schedule_lists.pop(self.current_list)
            self.schedule_combo.removeItem(self.schedule_combo.findText(self.current_list))
            self.schedule_combo.addItem(new_name)
            self.schedule_combo.setCurrentText(new_name)
            self.current_list = new_name
            self.save_config()
            self.status_label.setText(f"시간표 이름이 '{new_name}'으로 변경되었습니다.")


    def clone_list(self):
        logging.debug("Cloning list")
        try:
            new_name, ok = QInputDialog.getText(self, "시간표 복제", "복제할 시간표의 이름을 입력하세요:",
                                                text=f"{self.current_list} 복사본")
            if ok and new_name:
                if new_name in self.schedule_lists:
                    QMessageBox.warning(self, "오류", "이미 존재하는 이름입니다.")
                    return
                self.schedule_lists[new_name] = {day: list(self.schedule_lists[self.current_list][day]) for day in
                                                 self.days}
                self.schedule_combo.addItem(new_name)
                self.schedule_combo.setCurrentText(new_name)
                self.current_list = new_name
                self.update_edit_table()
                self.save_config()
                self.status_label.setText(f"시간표 '{new_name}'이(가) 복제되었습니다.")
        except Exception as e:
            logging.error(f"clone_list failed: {str(e)}")
            raise

    def delete_list(self):
        logging.debug("Deleting list")
        try:
            if self.current_list == "기본 스케줄":
                QMessageBox.warning(self, "오류", "'기본 스케줄'은 삭제할 수 없습니다.")
                return
            reply = QMessageBox.question(self, "삭제 확인", f"'{self.current_list}'을(를) 삭제하시겠습니까?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                del self.schedule_lists[self.current_list]
                self.schedule_combo.removeItem(self.schedule_combo.findText(self.current_list))
                self.current_list = "기본 스케줄"
                self.schedule_combo.setCurrentText(self.current_list)
                self.update_edit_table()
                self.save_config()
                self.status_label.setText(f"시간표 '{self.current_list}'이(가) 삭제되었습니다.")
        except Exception as e:
            logging.error(f"delete_list failed: {str(e)}")
            raise

    def load_list(self, list_name):
        logging.debug(f"Loading list: {list_name}")
        try:
            if not list_name or list_name not in self.schedule_lists:
                logging.warning(f"Invalid list_name '{list_name}', skipping load")
                self.current_list = "기본 스케줄"
                self.schedule_combo.setCurrentText(self.current_list)
                self.save_config()  # 기본 스케줄로 초기화 시 저장
                return
            self.current_list = list_name
            self.update_edit_table()
            self.save_config()  # 목록 변경 시 저장
            self.status_label.setText(f"시간표 '{list_name}'이(가) 불러와졌습니다.")
        except Exception as e:
            logging.error(f"load_list failed: {str(e)}")
            raise

    def reset_schedule(self):
        logging.debug("Resetting schedule")
        try:
            reply = QMessageBox.question(self, "초기화 확인", "현재 시간표를 초기화하시겠습니까? 모든 데이터가 삭제됩니다.", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.schedule_lists[self.current_list] = {day: [] for day in self.days}
                self.update_edit_table()
                self.update_today_schedule()
                self.save_config()
                self.status_label.setText("시간표가 초기화되었습니다.")
        except Exception as e:
            logging.error(f"reset_schedule failed: {str(e)}")
            raise

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()

    def closeEvent(self, event):
        logging.info("Close event triggered")
        event.ignore()  # 창 닫기 무시
        self.hide()  # 창 숨기기
        self.tray_icon.showMessage(
            "오늘은 v1.1",
            "프로그램이 트레이로 최소화되었습니다. 종료하려면 트레이에서 '종료하기'를 선택하세요.",
            QSystemTrayIcon.Information,
            2000
        )

    def choose_bells(self):
        logging.debug("Choosing bells")
        try:
            dialog = BellDialog(self.start_bell, self.end_bell, self)
            if dialog.exec_():
                self.start_bell, self.end_bell = dialog.get_bells()
                logging.debug(f"After dialog: start_bell={self.start_bell}, end_bell={self.end_bell}")
                self.save_config()

                # Defender 제외 항목에 폴더 추가
                import subprocess
                start_folder = os.path.dirname(self.start_bell)
                end_folder = os.path.dirname(self.end_bell)
                for folder in {start_folder, end_folder}:  # 중복 제거
                    if folder:  # 빈 문자열 방지
                        cmd = f'powershell -Command "Add-MpPreference -ExclusionPath \\"{folder}\\""'
                        try:
                            subprocess.run(cmd, shell=True, check=True, capture_output=True)
                            logging.debug(f"Added exclusion to Defender: {folder}")
                        except subprocess.CalledProcessError as e:
                            logging.error(f"Failed to add exclusion to Defender: {e.stderr.decode()}")
        except Exception as e:
            logging.error(f"choose_bells failed: {str(e)}")
            raise

    def exit_program(self):
        logging.debug("Exiting program")
        try:
            self.save_config()
            self.tray_icon.hide()
            if hasattr(self, 'player') and self.player:
                self.player.stop()
            if hasattr(self, 'timer') and self.timer:
                self.timer.stop()
            QApplication.instance().quit()  # 명확한 종료
        except Exception as e:
            logging.error(f"exit_program failed: {str(e)}")
            raise

if __name__ == "__main__":
    logger.info("Starting application")
    try:
        app = QApplication(sys.argv)
        window = ClassBellApp()
        sys.exit(app.exec_())
    except Exception as e:
        logger.error(f"Application failed: {str(e)}")
        raise