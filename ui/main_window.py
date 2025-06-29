import sys
import os
import win32com.client
import time
import warnings
from PySide6.QtWidgets import (QApplication, QMainWindow, QButtonGroup, 
                             QMessageBox, QTableWidgetItem, QHeaderView, 
                             QAbstractItemView, QLabel, QWidget, QLineEdit, 
                             QPushButton, QComboBox, QStackedWidget, QTableWidget)
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, Qt, QFileInfo, QTimer, QUrl
from PySide6.QtGui import QColor, QDoubleValidator, QIcon, QPixmap

from ui.widgets import DropZone
from ui.widgets import MultiClassPanel
from core.score_logic import ScoreLogic
from services.tts_manager import ITTSManager

def resource_path(relative_path):
    # main.py가 있는 폴더 기준으로 절대경로 반환
    base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self, logic: ScoreLogic, tts: ITTSManager, parent=None):
        super().__init__(parent)
        self.ui = None
        self.logic = logic
        self.stacked_widget = None
        self.tts = tts  # TTS 관리자 인스턴스
        self.multi_panel = MultiClassPanel()  # 이동반 패널 인스턴스 생성
        self.is_processing_student_number = False  # 중복 실행 방지 플래그
        
        # 성능 최적화를 위한 변수들
        self._update_timer = QTimer()
        self._update_timer.setSingleShot(True)
        self._update_timer.timeout.connect(self._delayed_update_table)
        self._pending_table_update = False
        self._signal_blocked = False
        self._cached_pink_color = QColor("#e0ffff")
        
        self.setup_ui()
        self.setup_connections()
        self.setWindowTitle("수행평가 점수 입력기 (by melderse 짐승농장)")
        self.setWindowIcon(QIcon(resource_path("icon_SC.png")))
        self.resize(623, 426)

    def setup_ui(self):
        """Sets up the UI."""
        # UI 파일 로드
        ui_file = QFile(resource_path("merged_ui.ui"))
        if not ui_file.open(QFile.ReadOnly):
            return
        loader = QUiLoader()
        self.ui = loader.load(ui_file)
        ui_file.close()
        
        # MainWindow에 UI 설정
        self.setCentralWidget(self.ui)
        
        # 스택 위젯 찾기
        self.stacked_widget = self.ui.findChild(QStackedWidget, "stackedWidget")
        
        # 테이블 최적화 설정
        if hasattr(self.ui, 'tableWidget'):
            table = self.ui.tableWidget
            table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
            table.verticalHeader().setVisible(False)
            
            # 성능 최적화 설정
            table.setAlternatingRowColors(True)
            table.setSortingEnabled(False)  # 정렬 비활성화로 성능 향상
            table.setUpdatesEnabled(True)
        
        # 배경 이미지 + 밝기 감소(흐림) 오버레이 적용
        bg_path = resource_path("background.png")
        pixmap = QPixmap(bg_path)
        # QLabel을 만들어서 배경으로 설정
        self.bg_label = QLabel(self.ui)
        self.bg_label.setPixmap(pixmap)
        self.bg_label.setScaledContents(True)
        self.bg_label.lower()  # 가장 뒤로 보내기
        self.bg_label.setGeometry(self.ui.rect())  # MainWidget 전체 크기로
        # MainWidget 크기 변경 시 배경도 같이 변경
        def resize_bg_label(event):
            self.bg_label.setGeometry(self.ui.rect())
            return super(type(self.ui), self.ui).resizeEvent(event)
        self.ui.resizeEvent = resize_bg_label
        
        # 라디오버튼 상태 초기화 (UI 로드 후)
        self.prev_radio_state = 1 if self.ui.radioButton_1.isChecked() else 2
        # 라디오버튼 자동 배타 복원
        self.ui.radioButton_1.setAutoExclusive(True)
        self.ui.radioButton_2.setAutoExclusive(True)

    def setup_connections(self):
        """Connects all signals to slots."""
        # --- Radio Buttons for Mode Change ---
        if hasattr(self.ui, 'radioButton_1'):
            self.ui.radioButton_1.toggled.connect(self.on_mode_changed)
        if hasattr(self.ui, 'radioButton_2'):
            self.ui.radioButton_2.toggled.connect(self.on_mode_changed)

        # --- Common Widgets ---
        if hasattr(self.ui, 'dropZone'):
            if isinstance(self.ui.dropZone, QLabel):
                self.replace_dropzone()
            else:
                 self.ui.dropZone.fileDropped.connect(self.on_files_dropped)

        if hasattr(self.ui, 'tableWidget'):
            # 셀 클릭만으로 행 선택 이벤트 처리 - 최적화된 연결
            self.ui.tableWidget.cellClicked.connect(self._on_cell_clicked_optimized)
            
        if hasattr(self.ui, 'save_button') and self.ui.save_button is not None:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                try:
                    self.ui.save_button.clicked.disconnect()
                except Exception:
                    pass
            self.ui.save_button.clicked.connect(self.save_to_excel)
        if hasattr(self.ui, 'session_combo'):
            self.ui.session_combo.currentTextChanged.connect(self._schedule_table_update)
        if hasattr(self.ui, 'pushButton_2') and self.ui.pushButton_2 is not None: # Clear Button
            self.ui.pushButton_2.clicked.connect(self.clear_table_and_data)

        # --- Widgets inside StackedWidget ---
        # Page 1: 이동반
        page_multi = self.stacked_widget.findChild(QWidget, "page_multi")
        if page_multi:
            score_input_multi = page_multi.findChild(QLineEdit, "scoreInput")
            if score_input_multi:
                score_input_multi.returnPressed.connect(self.on_multi_score_entered)
            student_number_input = page_multi.findChild(QLineEdit, "studentNumberInput")
            if student_number_input:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    try:
                        student_number_input.returnPressed.disconnect()
                    except Exception:
                        pass
                student_number_input.returnPressed.connect(self.on_multi_student_number_entered)
            student_table = page_multi.findChild(QTableWidget, "studentTable")
            if student_table:
                student_table.cellClicked.connect(self.on_multi_student_table_cell_clicked)

        # Page 2: 단일반
        page_single = self.stacked_widget.findChild(QWidget, "page_single")
        if page_single:
            text_edit_single = page_single.findChild(QLineEdit, "text_edit")
            if text_edit_single:
                text_edit_single.returnPressed.connect(self.on_score_entered)
            
            sound_button = page_single.findChild(QPushButton, "sound_toggle_button")
            if sound_button:
                sound_button.toggled.connect(self.on_sound_toggled)

    def _on_cell_clicked_optimized(self, row, col):
        """최적화된 셀 클릭 핸들러"""
        if not self._signal_blocked:
            self.on_row_selected()

    def _schedule_table_update(self):
        """테이블 업데이트를 지연시켜 성능을 최적화합니다."""
        if not self._pending_table_update:
            self._pending_table_update = True
            self._update_timer.start(50)  # 50ms 지연

    def _delayed_update_table(self):
        """지연된 테이블 업데이트"""
        self._pending_table_update = False
        self.update_table_view()

    def apply_table_selection_style(self, disable: bool):
        if hasattr(self.ui, 'tableWidget'):
            if disable:
                # 파란색 선택/포커스 선을 없앰
                self.ui.tableWidget.setStyleSheet("""
                    QTableWidget::item:selected, QTableWidget::item:focus {
                        background: transparent;
                        color: black;
                        border: none;
                        outline: none;
                    }
                """)
            else:
                # 단일반: light cyan 배경
                self.ui.tableWidget.setStyleSheet("""
                    QTableWidget::item:selected, QTableWidget::item:focus {
                        background: #e0ffff;
                        color: black;
                    }
                """)

    def on_mode_changed(self, checked):
        if not checked:
            return

        sender = self.sender()
        new_state = 1 if sender == self.ui.radioButton_1 else 2

        if new_state == self.prev_radio_state:
            return

        has_data = self.logic.files and (self.logic.student_data or (hasattr(self.ui, 'fileListbox') and self.ui.fileListbox.count() > 0))
        if has_data:
            reply = QMessageBox.question(self, "모드 전환 경고", "입력된 데이터가 사라집니다. 계속하시겠습니까?", QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                # 시그널 차단으로 성능 최적화
                self._signal_blocked = True
                if self.prev_radio_state == 1:
                    self.ui.radioButton_1.setChecked(True)
                else:
                    self.ui.radioButton_2.setChecked(True)
                self._signal_blocked = False
                return
            self.clear_table_and_data()

        self.prev_radio_state = new_state
        if not self.stacked_widget: return

        # 이동반(radioButton_2)일 때만 파란색 선택/포커스 선 제거
        if new_state == 2:  # 이동반
            self.apply_table_selection_style(True)
        else:
            self.apply_table_selection_style(False)

        if new_state == 1: # 단일반
            self.stacked_widget.setCurrentIndex(1)
        else: # 이동반
            self.stacked_widget.setCurrentIndex(0)

        if hasattr(self.ui, 'tableWidget') and self.ui.tableWidget.rowCount() > 0:
            current_row = self.ui.tableWidget.currentRow()
            if current_row < 0:
                current_row = 0
            self.update_student_info_labels(current_row)

    def replace_dropzone(self):
        """Replaces the QLabel dropZone with the custom DropZone widget."""
        if not hasattr(self.ui, 'dropZone'): return
        original_label = self.ui.dropZone
        parent_layout = original_label.parent().layout()
        if not parent_layout: return
        for i in range(parent_layout.count()):
            item = parent_layout.itemAt(i)
            if item.widget() == original_label:
                new_drop_zone = DropZone("엑셀 파일을 여기에 드롭하세요")
                new_drop_zone.setMinimumSize(original_label.minimumSize())
                new_drop_zone.setSizePolicy(original_label.sizePolicy())
                parent_layout.insertWidget(i, new_drop_zone)
                original_label.deleteLater()
                self.ui.dropZone = new_drop_zone
                new_drop_zone.fileDropped.connect(self.on_files_dropped)
                return

    def get_current_text_edit(self):
        """Returns the currently active score input QLineEdit."""
        current_page = self.stacked_widget.currentWidget()
        if not current_page: return None
        
        if current_page.objectName() == "page_single":
            text_edit = current_page.findChild(QLineEdit, "text_edit")
            return text_edit
        elif current_page.objectName() == "page_multi":
            score_input = current_page.findChild(QLineEdit, "scoreInput")
            return score_input
        return None

    def on_files_dropped(self, file_paths):
        """Handles multiple file drop event."""
        # 파일명 기준 정렬
        file_paths = sorted(file_paths, key=lambda x: os.path.basename(x))
        loaded_any = False
        for file_path in file_paths:
            success, message = self.logic.load_excel_data(file_path)
            if not success:
                QMessageBox.warning(self, "파일 로드 오류", f"{file_path}: {message}")
            else:
                loaded_any = True
        if loaded_any:
            self.update_ui_after_file_load(file_paths[0])

    def update_ui_after_file_load(self, file_path):
        """Updates the UI after a file is loaded."""
        # 파일 리스트 업데이트 최적화
        if hasattr(self.ui, 'fileListbox'):
            self.ui.fileListbox.clear()
            file_names = [QFileInfo(f['path']).fileName() for f in self.logic.files]
            self.ui.fileListbox.addItems(file_names)

        # 테이블 업데이트 최적화
        if hasattr(self.ui, 'tableWidget'):
            table = self.ui.tableWidget
            headers = self.logic.headers
            bcd_indices = [1, 2, 3]
            bcd_headers = [headers[i] if i < len(headers) else f"컬럼{i}" for i in bcd_indices]
            
            # 한 번에 설정
            table.setUpdatesEnabled(False)  # 업데이트 차단
            table.clear()
            table.setColumnCount(len(bcd_headers))
            table.setHorizontalHeaderLabels(bcd_headers)
            
            # 데이터 배치 처리
            student_data = self.logic.student_data
            table.setRowCount(len(student_data))
            
            # 배치로 아이템 생성
            for row_idx, row_data in enumerate(student_data):
                for col_idx, src_idx in enumerate(bcd_indices):
                    cell_data = str(row_data[src_idx]) if src_idx < len(row_data) else ''
                    item = QTableWidgetItem(cell_data)
                    item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(row_idx, col_idx, item)
            
            table.resizeColumnsToContents()
            table.horizontalHeader().setStretchLastSection(True)
            table.setUpdatesEnabled(True)  # 업데이트 재개
            
        self.setup_session_combobox(len(self.logic.headers))
        self.update_table_view()
        
        if self.ui.radioButton_2.isChecked():
            page_multi = self.stacked_widget.findChild(QWidget, "page_multi")
            if page_multi:
                student_table = page_multi.findChild(QTableWidget, "studentTable")
                if student_table:
                    student_table.clearContents()
                    student_table.setRowCount(0)
                student_number_input = page_multi.findChild(QLineEdit, "studentNumberInput")
                if student_number_input:
                    student_number_input.setFocus()
            return
            
        if hasattr(self.ui, 'tableWidget') and self.ui.tableWidget.rowCount() > 0:
            self.ui.tableWidget.selectRow(0)
            self.on_row_selected()

    def setup_session_combobox(self, max_column):
        """Sets up the session combobox."""
        if hasattr(self.ui, 'session_combo'):
            combo = self.ui.session_combo
            combo.clear()
            num_sessions = max_column - 4
            if num_sessions > 0:
                # 한 번에 아이템 추가
                session_items = [f"{i}회" for i in range(1, num_sessions + 1)]
                combo.addItems(session_items)
    
    def update_table_view(self):
        """최적화된 테이블 뷰 업데이트"""
        table = self.ui.tableWidget
        if not hasattr(self.ui, 'tableWidget') or table.columnCount() == 0:
            return
            
        if not hasattr(self.ui, 'session_combo'):
            return
            
        current_session_text = self.ui.session_combo.currentText()
        if not current_session_text:
            return
            
        try:
            session_number = int(current_session_text.replace("회", ""))
            score_col_index = session_number + 3
            headers = self.logic.headers
            
            if score_col_index < len(headers):
                # 업데이트 차단으로 성능 최적화
                table.setUpdatesEnabled(False)
                
                # 컬럼 수 설정
                table.setColumnCount(4)
                
                # 헤더 설정
                header_items = [headers[1], headers[2], headers[3], headers[score_col_index]]
                for i, h in enumerate(header_items):
                    table.setHorizontalHeaderItem(i, QTableWidgetItem(h))
                
                # 데이터 배치 업데이트
                student_data = self.logic.student_data
                for row_idx, row_data in enumerate(student_data):
                    # 반, 번호, 성명
                    for col, src in enumerate([1, 2, 3]):
                        val = row_data[src] if src < len(row_data) else ''
                        item = table.item(row_idx, col)
                        if not item:
                            item = QTableWidgetItem()
                            table.setItem(row_idx, col, item)
                        item.setText(str(val))
                        item.setTextAlignment(Qt.AlignCenter)
                    
                    # 점수
                    value = row_data[score_col_index] if score_col_index < len(row_data) else ''
                    item = table.item(row_idx, 3)
                    if not item:
                        item = QTableWidgetItem()
                        table.setItem(row_idx, 3, item)
                    item.setText(str(value))
                    item.setTextAlignment(Qt.AlignCenter)
                
                table.resizeColumnsToContents()
                table.horizontalHeader().setStretchLastSection(True)
                table.setUpdatesEnabled(True)
                
        except (ValueError, TypeError) as e:
            pass

    def on_row_selected(self):
        if not hasattr(self.ui, 'tableWidget'): return
        
        selected_rows = self.ui.tableWidget.selectionModel().selectedRows()
        if not selected_rows:
            self.update_student_info_labels(-1)
            return

        row_index = selected_rows[0].row()
        self.update_student_info_labels(row_index)

        # 이동반 모드일 때 번호와 성명을 자동으로 입력
        if self.ui.radioButton_2.isChecked():
            self._handle_multi_class_selection(row_index)

        # 점수 입력란에 해당 행의 점수 표시 및 포커스 (단일반 모드)
        self._handle_score_input_focus(row_index)

    def _handle_multi_class_selection(self, row_index):
        """이동반 모드에서의 행 선택 처리"""
        page_multi = self.stacked_widget.findChild(QWidget, "page_multi")
        if not page_multi:
            return
            
        table = self.ui.tableWidget
        if row_index < table.rowCount():
            number_item = table.item(row_index, 1)
            name_item = table.item(row_index, 2)
            
            if number_item and name_item:
                number = number_item.text()
                name = name_item.text()
                
                # 시그널 차단으로 성능 최적화
                student_number_input = page_multi.findChild(QLineEdit, "studentNumberInput")
                if student_number_input:
                    student_number_input.blockSignals(True)
                    student_number_input.setText(number)
                    student_number_input.blockSignals(False)
                
                student_name_label = page_multi.findChild(QLabel, "studentName")
                if not student_name_label:
                    student_name_label = page_multi.findChild(QLabel, "label_5")
                if student_name_label:
                    student_name_label.setText(name)
                
                score_input = page_multi.findChild(QLineEdit, "scoreInput")
                if score_input:
                    QTimer.singleShot(0, lambda: (score_input.setFocus(), score_input.selectAll()))

    def _handle_score_input_focus(self, row_index):
        """점수 입력 포커스 처리"""
        if not hasattr(self.ui, 'session_combo'):
            return
            
        current_session_text = self.ui.session_combo.currentText()
        score_to_edit = ""
        
        if current_session_text:
            try:
                session_number = int(current_session_text.replace("회", ""))
                score_col_index = session_number + 3
                student_data = self.logic.student_data
                
                if 0 <= row_index < len(student_data) and 0 <= score_col_index < len(student_data[row_index]):
                    score_to_edit = student_data[row_index][score_col_index]
            except (ValueError, TypeError, IndexError):
                pass 
        
        current_text_edit = self.get_current_text_edit()
        if current_text_edit:
            current_text_edit.setText(str(score_to_edit))
            QTimer.singleShot(0, lambda: (current_text_edit.setFocus(), current_text_edit.selectAll()))
    
    def update_student_info_labels(self, row_index):
        page_single = self.stacked_widget.findChild(QWidget, "page_single")
        if page_single:
            label_num_val = page_single.findChild(QLabel, "label_num_val")
            label_name = page_single.findChild(QLabel, "label_name")
            if label_num_val and label_name:
                self._update_labels(row_index, label_num_val, label_name)

        page_multi = self.stacked_widget.findChild(QWidget, "page_multi")
        if page_multi:
            self._update_multi_student_table(page_multi, row_index)

    def _update_multi_student_table(self, page_multi, row_index):
        """이동반 학생 테이블 업데이트 최적화"""
        student_table = page_multi.findChild(QTableWidget, "studentTable")
        if not student_table:
            return
            
        student_data = self.logic.student_data
        if row_index < 0 or row_index >= len(student_data):
            student_table.clearContents()
            student_table.setRowCount(0)
            return

        student_row = student_data[row_index]
        if len(student_row) < 4:
            student_table.clearContents()
            student_table.setRowCount(0)
            return

        class_text = student_row[1].strip()
        number_text = student_row[2].strip()
        name_text = student_row[3].strip()

        # 한 번에 설정
        student_table.setUpdatesEnabled(False)
        student_table.setRowCount(1)
        student_table.setColumnCount(2)
        student_table.setHorizontalHeaderLabels(["번호", "이름"])
        
        # 아이템 생성 및 설정
        number_display = f"{class_text}_{number_text}" if class_text else number_text
        items = [
            (QTableWidgetItem(number_display), 0, 0),
            (QTableWidgetItem(name_text), 0, 1)
        ]
        
        for item, row, col in items:
            item.setTextAlignment(Qt.AlignCenter)
            student_table.setItem(row, col, item)
            
        student_table.setUpdatesEnabled(True)

    def _update_labels(self, row_index, label_num_val, label_name):
        student_data = self.logic.student_data
        if row_index < 0 or row_index >= len(student_data):
            label_num_val.setText("")
            label_name.setText("")
            return
            
        student_row = student_data[row_index]
        if len(student_row) < 4:
            label_num_val.setText("")
            label_name.setText("")
            return
            
        class_text = student_row[1].strip()
        number_text = student_row[2].strip()
        name_text = student_row[3].strip()
        
        display_number = f"{class_text}_{number_text}" if class_text else number_text
        label_num_val.setText(display_number)
        label_name.setText(name_text)
        
        # TTS: 단일반 모드, 사운드 ON, 이름이 있을 때 마지막 한글자만 읽기
        if self.tts and not self.ui.radioButton_2.isChecked() and name_text:
            page_single = self.stacked_widget.findChild(QWidget, "page_single")
            sound_button = page_single.findChild(QPushButton, "sound_toggle_button") if page_single else None
            if sound_button and sound_button.isChecked():
                self.tts.speak_name(str(name_text[-1]))

    def on_score_entered(self):
        """최적화된 점수 입력 처리"""
        if not hasattr(self.ui, 'tableWidget'): 
            return

        text_edit = self.get_current_text_edit()
        current_row = self.ui.tableWidget.currentRow()
        
        if current_row < 0 or not text_edit: 
            return

        score_text = text_edit.text().strip()
        
        # 입력 검증
        try:
            float(score_text)
        except ValueError:
            QMessageBox.warning(self, "입력 오류", "숫자만 입력가능합니다.")
            return

        # TTS: 단일반 모드, 사운드 ON, 숫자 있을 때 읽기
        if self.tts and not self.ui.radioButton_2.isChecked() and score_text:
            page_single = self.stacked_widget.findChild(QWidget, "page_single")
            sound_button = page_single.findChild(QPushButton, "sound_toggle_button") if page_single else None
            if sound_button and sound_button.isChecked():
                self.tts.speak_name(str(score_text), rate=2)

        if not hasattr(self.ui, 'session_combo'): 
            return
            
        session_index = self.ui.session_combo.currentIndex()
        if session_index < 0: 
            return
        
        # 데이터 업데이트
        self.logic.update_score(current_row, session_index, score_text)

        # UI 업데이트 최적화
        table = self.ui.tableWidget
        target_col = 3
        
        # 아이템 업데이트
        item = table.item(current_row, target_col)
        if not item:
            item = QTableWidgetItem()
            table.setItem(current_row, target_col, item)
        item.setText(score_text)
        item.setTextAlignment(Qt.AlignCenter)

        # 배경색 변경 최적화
        for col in range(table.columnCount()):
            cell = table.item(current_row, col)
            if not cell:
                cell = QTableWidgetItem()
                table.setItem(current_row, col, cell)
            cell.setBackground(self._cached_pink_color)
        
        text_edit.clear()
        table.scrollToItem(item, QAbstractItemView.ScrollHint.EnsureVisible)
        
        # 다음 행으로 이동
        next_row = current_row + 1
        if next_row < table.rowCount():
            table.selectRow(next_row)
            QTimer.singleShot(0, self.on_row_selected)
        else:
            QMessageBox.information(self, "알림", "마지막 학생까지 점수 입력이 완료되었습니다.")

    def on_sound_toggled(self, checked):
        """Handles the sound toggle button state change."""
        button = self.sender()
        if button:
            if checked:
                button.setText("🔊")
                button.setStyleSheet("background-color: #0078d4; color: white; border-radius: 8px;")
            else:
                button.setText("🔇")
                button.setStyleSheet("background-color: #f8f8f8; color: #888; border-radius: 8px;")

    def save_to_excel(self):
        """Saves the data to an Excel file."""
        success, message = self.logic.save_to_excel()
        if success:
            QMessageBox.information(self, "저장 완료", message)
        else:
            QMessageBox.critical(self, "저장 오류", message)

    def clear_table_and_data(self):
        """Clears the table and loaded data."""
        self.logic.clear_data()
        if hasattr(self.ui, 'fileListbox'):
            self.ui.fileListbox.clear()
        if hasattr(self.ui, 'tableWidget'):
            table = self.ui.tableWidget
            table.setUpdatesEnabled(False)
            table.setRowCount(0)
            table.setColumnCount(0)
            table.setUpdatesEnabled(True)
        if hasattr(self.ui, 'session_combo'):
            self.ui.session_combo.clear()
        self.update_student_info_labels(-1)

    def on_multi_student_number_entered(self):
        """이동반 모드에서 학생번호 입력 처리 최적화"""
        if self.is_processing_student_number:
            return
        
        self.is_processing_student_number = True
        
        try:
            page_multi = self.stacked_widget.findChild(QWidget, "page_multi")
            if not page_multi:
                return
                
            student_number_input = page_multi.findChild(QLineEdit, "studentNumberInput")
            student_table = page_multi.findChild(QTableWidget, "studentTable")
            
            if not student_number_input or not student_table:
                return
                
            number = student_number_input.text().strip()
            if not number:
                student_table.clearContents()
                student_table.setRowCount(0)
                return

            # 검색 최적화 - 한 번의 순회로 처리
            results = []
            table = self.ui.tableWidget
            for row in range(table.rowCount()):
                num_item = table.item(row, 1)
                name_item = table.item(row, 2)
                if num_item and name_item and num_item.text() == number:
                    results.append([num_item.text(), name_item.text()])

            # 테이블 업데이트 최적화
            student_table.setUpdatesEnabled(False)
            student_table.setRowCount(len(results))
            student_table.setColumnCount(2)
            student_table.setHorizontalHeaderLabels(["번호", "이름"])
            
            for i, row in enumerate(results):
                for j, text in enumerate(row):
                    item = QTableWidgetItem(text)
                    student_table.setItem(i, j, item)
            
            student_table.resizeColumnsToContents()
            header = student_table.horizontalHeader()
            if header:
                header.setStretchLastSection(True)
            student_table.setUpdatesEnabled(True)

            # 단일 결과 처리
            self._handle_single_search_result(page_multi, results)
            
            if not results:
                QMessageBox.information(self, "검색 결과 없음", f"{number}번 학생을 찾을 수 없습니다.")
                
        finally:
            self.is_processing_student_number = False

    def _handle_single_search_result(self, page_multi, results):
        """단일 검색 결과 처리"""
        student_name_label = page_multi.findChild(QLabel, "studentName")
        if not student_name_label:
            student_name_label = page_multi.findChild(QLabel, "label_5")
            
        if student_name_label:
            if len(results) == 1:
                student_name_label.setText(results[0][1])
                
                def clear_and_focus():
                    student_number_input = page_multi.findChild(QLineEdit, "studentNumberInput")
                    score_input = page_multi.findChild(QLineEdit, "scoreInput")
                    
                    if student_number_input:
                        student_number_input.blockSignals(True)
                        student_number_input.clear()
                        student_number_input.blockSignals(False)
                    if score_input:
                        score_input.setFocus()
                        
                QTimer.singleShot(0, clear_and_focus)
            else:
                student_name_label.setText("")

    def on_multi_student_table_cell_clicked(self, row, col):
        """이동반 학생 테이블 셀 클릭 처리"""
        page_multi = self.stacked_widget.findChild(QWidget, "page_multi")
        if not page_multi:
            return
            
        student_table = page_multi.findChild(QTableWidget, "studentTable")
        student_name_label = page_multi.findChild(QLabel, "studentName")
        if not student_name_label:
            student_name_label = page_multi.findChild(QLabel, "label_5")
            
        if not student_table or not student_name_label:
            return

        name_item = student_table.item(row, 1)
        if name_item:
            student_name_label.setText(name_item.text())
            
            def clear_and_focus():
                student_number_input = page_multi.findChild(QLineEdit, "studentNumberInput")
                score_input = page_multi.findChild(QLineEdit, "scoreInput")
                
                if student_number_input:
                    student_number_input.blockSignals(True)
                    student_number_input.clear()
                    student_number_input.blockSignals(False)
                if score_input:
                    score_input.setFocus()
                    score_input.selectAll()
                    
            QTimer.singleShot(0, clear_and_focus)
        else:
            student_name_label.setText("")

    def on_multi_score_entered(self):
        """이동반 점수 입력 처리 최적화"""
        page_multi = self.stacked_widget.findChild(QWidget, "page_multi")
        if not page_multi:
            return
            
        score_input = page_multi.findChild(QLineEdit, "scoreInput")
        if not score_input:
            return
            
        score_text = score_input.text().strip()
        if not score_text:
            return

        student_name_label = page_multi.findChild(QLabel, "studentName")
        if not student_name_label:
            student_name_label = page_multi.findChild(QLabel, "label_5")
        if not student_name_label:
            QMessageBox.warning(self, "오류", "학생 이름 라벨을 찾을 수 없습니다.")
            return
        
        current_name = student_name_label.text().strip()
        if not current_name:
            QMessageBox.warning(self, "선택 오류", "학생을 먼저 선택하세요.")
            return
        
        student_number_input = page_multi.findChild(QLineEdit, "studentNumberInput")
        current_number = student_number_input.text().strip() if student_number_input else ""
        
        # 학생 찾기 및 점수 업데이트 최적화
        success = self._update_multi_student_score(current_number, current_name, score_text)
        
        if success:
            score_input.clear()
            if student_number_input:
                student_number_input.setFocus()
        else:
            QMessageBox.warning(self, "찾기 실패", f"학생 '{current_name}'을 tableWidget에서 찾을 수 없습니다.")

    def _update_multi_student_score(self, number, name, score):
        """이동반 학생 점수 업데이트"""
        if not hasattr(self.ui, 'tableWidget') or not hasattr(self.ui, 'session_combo'):
            return False
            
        table = self.ui.tableWidget
        session_index = self.ui.session_combo.currentIndex()
        
        if session_index < 0:
            QMessageBox.warning(self, "회차 오류", "회차를 선택하세요.")
            return False
        
        # 학생 찾기
        for r in range(table.rowCount()):
            num_item = table.item(r, 1)
            name_item = table.item(r, 2)
            
            if num_item and name_item:
                num_text = num_item.text().strip()
                name_text = name_item.text().strip()
                
                # 번호와 이름 매칭
                if ((number and num_text == number and name_text == name) or 
                    (not number and name_text == name)):
                    # 데이터 업데이트
                    self.logic.update_score(r, session_index, score)
                    
                    # UI 업데이트
                    score_col = 3
                    item = table.item(r, score_col)
                    if not item:
                        item = QTableWidgetItem()
                        table.setItem(r, score_col, item)
                    item.setText(score)
                    item.setTextAlignment(Qt.AlignCenter)
                    
                    # 배경색 및 포커스 적용 (단일반과 동일하게)
                    for col in range(table.columnCount()):
                        cell = table.item(r, col)
                        if not cell:
                            cell = QTableWidgetItem()
                            table.setItem(r, col, cell)
                        cell.setBackground(self._cached_pink_color)
                    table.selectRow(r)
                    return True
        
        return False