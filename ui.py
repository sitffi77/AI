"""
ui.py - Полный интерфейс с выводом симптомов и предпросмотром файлов
"""

import sys
import os
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import pandas as pd
from docx import Document
import openpyxl


class PainDiagnosisApp(QMainWindow):
    def __init__(self):
        super().__init__()
        from database import DatabaseManager
        from analyzer import EnhancedPainAnalyzer

        self.db = DatabaseManager()
        self.analyzer = EnhancedPainAnalyzer()
        self.current_patient_id = None
        self.patients_cache = {}

        # +++ ДОБАВЬТЕ ЭТИ АТРИБУТЫ +++
        self.last_analysis_results = None
        self.last_symptoms_text = ""

        self.initUI()
        self.load_patients_list()

    def display_general_results(self, results):
        """Отображение общих результатов анализа"""
        # Проверяем, есть ли необходимые ключи в results
        print(f"Результаты анализа:")

        if 'diagnosis' in results:
            diagnosis = results['diagnosis']
            print(f"Тип боли: {diagnosis.get('type', 'не определено')}")
            print(f"Подгруппа: {diagnosis.get('subgroup', 'не определена')}")
            print(f"Тяжесть: {diagnosis.get('severity', 'не определена')}")

        if 'confidence' in results:
            print(f"Достоверность: {results['confidence']:.1f}%")

        if 'recommendations' in results:
            print(f"Рекомендации:")
            for i, rec in enumerate(results['recommendations'], 1):
                print(f"  {i}. {rec}")

    def initUI(self):
        self.setWindowTitle('AI Диагностика Боли в Спине')
        self.setGeometry(100, 100, 1400, 900)

        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # ========== ЛЕВАЯ ПАНЕЛЬ ==========
        left_panel = QFrame()
        left_panel.setFrameStyle(QFrame.Panel | QFrame.Raised)
        left_layout = QVBoxLayout(left_panel)

        # 1. Панель пациентов
        patients_group = QGroupBox("Пациенты")
        patients_layout = QVBoxLayout()

        # Поиск
        search_layout = QHBoxLayout()
        self.patient_search = QLineEdit()
        self.patient_search.setPlaceholderText("Поиск по имени...")
        self.patient_search.textChanged.connect(self.search_patients)
        search_btn = QPushButton("🔍")
        search_btn.clicked.connect(self.search_patients)

        search_layout.addWidget(self.patient_search)
        search_layout.addWidget(search_btn)
        patients_layout.addLayout(search_layout)

        # Список пациентов
        self.patients_list = QListWidget()
        self.patients_list.setMinimumWidth(200)
        self.patients_list.itemClicked.connect(self.load_selected_patient)
        patients_layout.addWidget(self.patients_list)

        # Кнопки пациентов
        patient_btn_layout = QHBoxLayout()
        self.new_patient_btn = QPushButton("➕ Новый")
        self.new_patient_btn.clicked.connect(self.show_new_patient_dialog)
        self.delete_patient_btn = QPushButton("🗑️ Удалить")
        self.delete_patient_btn.clicked.connect(self.delete_selected_patient)

        patient_btn_layout.addWidget(self.new_patient_btn)
        patient_btn_layout.addWidget(self.delete_patient_btn)
        patients_layout.addLayout(patient_btn_layout)

        patients_group.setLayout(patients_layout)
        left_layout.addWidget(patients_group)

        # 2. Текущий пациент
        self.current_patient_group = QGroupBox("Текущий пациент")
        current_layout = QVBoxLayout()

        self.current_patient_label = QLabel("Не выбран")
        self.current_patient_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        current_layout.addWidget(self.current_patient_label)

        self.current_patient_group.setLayout(current_layout)
        left_layout.addWidget(self.current_patient_group)

        # 3. Симптомы пациента (НОВОЕ: список симптомов)
        self.symptoms_list_group = QGroupBox("Симптомы пациента")
        symptoms_list_layout = QVBoxLayout()

        self.patient_symptoms_list = QListWidget()
        self.patient_symptoms_list.setMinimumHeight(150)
        symptoms_list_layout.addWidget(self.patient_symptoms_list)

        # Кнопка очистки симптомов
        clear_symptoms_btn = QPushButton("Очистить симптомы")
        clear_symptoms_btn.clicked.connect(self.clear_patient_symptoms)
        symptoms_list_layout.addWidget(clear_symptoms_btn)

        self.symptoms_list_group.setLayout(symptoms_list_layout)
        left_layout.addWidget(self.symptoms_list_group)

        # 4. Панель ввода симптомов
        symptoms_input_group = QGroupBox("Ввод симптомов")
        symptoms_input_layout = QVBoxLayout()

        self.symptoms_text = QTextEdit()
        self.symptoms_text.setPlaceholderText("Введите симптомы здесь...\nКаждый симптом с новой строки.")
        self.symptoms_text.setMinimumHeight(100)
        symptoms_input_layout.addWidget(self.symptoms_text)

        # Быстрые шаблоны
        templates_layout = QHBoxLayout()
        template1_btn = QPushButton("Шаблон 1")
        template1_btn.clicked.connect(lambda: self.load_template(1))
        template2_btn = QPushButton("Шаблон 2")
        template2_btn.clicked.connect(lambda: self.load_template(2))
        template3_btn = QPushButton("Шаблон 3")
        template3_btn.clicked.connect(lambda: self.load_template(3))

        templates_layout.addWidget(template1_btn)
        templates_layout.addWidget(template2_btn)
        templates_layout.addWidget(template3_btn)
        symptoms_input_layout.addLayout(templates_layout)

        symptoms_input_group.setLayout(symptoms_input_layout)
        left_layout.addWidget(symptoms_input_group)

        # 5. Загрузка файлов
        file_group = QGroupBox("Загрузка файлов")
        file_layout = QVBoxLayout()

        # Кнопки загрузки
        load_btn_layout = QHBoxLayout()
        self.load_word_btn = QPushButton("📝 Word")
        self.load_word_btn.clicked.connect(self.load_word_file)
        self.load_excel_btn = QPushButton("📊 Excel")
        self.load_excel_btn.clicked.connect(self.load_excel_file)
        self.load_text_btn = QPushButton("📄 Текст")
        self.load_text_btn.clicked.connect(self.load_text_file)

        load_btn_layout.addWidget(self.load_word_btn)
        load_btn_layout.addWidget(self.load_excel_btn)
        load_btn_layout.addWidget(self.load_text_btn)
        file_layout.addLayout(load_btn_layout)

        # Предпросмотр файла (ИСПРАВЛЕНО: добавляем атрибут)
        file_layout.addWidget(QLabel("Предпросмотр:"))
        self.file_preview = QTextEdit()
        self.file_preview.setReadOnly(True)
        self.file_preview.setMaximumHeight(100)
        file_layout.addWidget(self.file_preview)

        file_group.setLayout(file_layout)
        left_layout.addWidget(file_group)

        # 6. Шкалы
        scales_group = QGroupBox("Шкалы оценки")
        scales_layout = QGridLayout()

        self.vas_input = QSpinBox()
        self.vas_input.setRange(0, 10)
        self.oswestry_input = QSpinBox()
        self.oswestry_input.setRange(0, 100)
        self.pain_detect_input = QSpinBox()
        self.pain_detect_input.setRange(0, 38)
        self.hads_anxiety_input = QSpinBox()
        self.hads_anxiety_input.setRange(0, 21)
        self.hads_depression_input = QSpinBox()
        self.hads_depression_input.setRange(0, 21)
        self.pcs_input = QSpinBox()
        self.pcs_input.setRange(0, 52)

        scales_layout.addWidget(QLabel("ВАШ (0-10):"), 0, 0)
        scales_layout.addWidget(self.vas_input, 0, 1)
        scales_layout.addWidget(QLabel("Освестри (0-100):"), 1, 0)
        scales_layout.addWidget(self.oswestry_input, 1, 1)
        scales_layout.addWidget(QLabel("PainDetect (0-38):"), 2, 0)
        scales_layout.addWidget(self.pain_detect_input, 2, 1)
        scales_layout.addWidget(QLabel("HADS Тревога (0-21):"), 0, 2)
        scales_layout.addWidget(self.hads_anxiety_input, 0, 3)
        scales_layout.addWidget(QLabel("HADS Депрессия (0-21):"), 1, 2)
        scales_layout.addWidget(self.hads_depression_input, 1, 3)
        scales_layout.addWidget(QLabel("PCS (0-52):"), 2, 2)
        scales_layout.addWidget(self.pcs_input, 2, 3)

        scales_group.setLayout(scales_layout)
        left_layout.addWidget(scales_group)

        # 7. Кнопка анализа
        self.analyze_btn = QPushButton("🔍 АНАЛИЗИРОВАТЬ")
        self.analyze_btn.clicked.connect(self.analyze_pain)
        self.analyze_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 16px;
                padding: 12px;
                border-radius: 6px;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.analyze_btn.setEnabled(False)
        left_layout.addWidget(self.analyze_btn)

        # Добавляем левую панель
        main_layout.addWidget(left_panel, 1)

        # ========== ПРАВАЯ ПАНЕЛЬ ==========
        right_panel = QFrame()
        right_panel.setFrameStyle(QFrame.Panel | QFrame.Raised)
        right_layout = QVBoxLayout(right_panel)

        # Вкладки результатов
        self.results_tabs = QTabWidget()

        # Вкладка 1: Диагноз
        diagnosis_tab = QWidget()
        diagnosis_layout = QVBoxLayout(diagnosis_tab)

        self.diagnosis_text = QTextEdit()
        self.diagnosis_text.setReadOnly(True)
        diagnosis_layout.addWidget(QLabel("<h3>Диагноз:</h3>"))
        diagnosis_layout.addWidget(self.diagnosis_text)

        self.results_tabs.addTab(diagnosis_tab, "📋 Диагноз")

        # Вкладка 2: Рекомендации
        treatment_tab = QWidget()
        treatment_layout = QVBoxLayout(treatment_tab)

        self.treatment_text = QTextEdit()
        self.treatment_text.setReadOnly(True)
        treatment_layout.addWidget(QLabel("<h3>Рекомендации:</h3>"))
        treatment_layout.addWidget(self.treatment_text)

        self.results_tabs.addTab(treatment_tab, "💊 Лечение")

        # Вкладка 3: Обнаруженные симптомы (НОВАЯ ВКЛАДКА)
        detected_symptoms_tab = QWidget()
        detected_layout = QVBoxLayout(detected_symptoms_tab)

        self.detected_symptoms_text = QTextEdit()
        self.detected_symptoms_text.setReadOnly(True)
        detected_layout.addWidget(QLabel("<h3>Обнаруженные симптомы:</h3>"))
        detected_layout.addWidget(self.detected_symptoms_text)

        self.results_tabs.addTab(detected_symptoms_tab, "📝 Симптомы")

        # Вкладка 4: История
        history_tab = QWidget()
        history_layout = QVBoxLayout(history_tab)

        self.history_table = QTableWidget()
        self.history_table.setColumnCount(6)
        self.history_table.setHorizontalHeaderLabels([
            "Дата", "Тип", "Подгруппа", "Достоверность", "Тяжесть", "Лечение"
        ])
        history_layout.addWidget(self.history_table)

        self.results_tabs.addTab(history_tab, "📅 История")

        right_layout.addWidget(self.results_tabs)
        main_layout.addWidget(right_panel, 2)

        # Загружаем первый пациент если есть
        if self.patients_list.count() > 0:
            self.patients_list.setCurrentRow(0)
            self.load_selected_patient(self.patients_list.item(0))

    # ========== МЕТОДЫ ДЛЯ ПАЦИЕНТОВ ==========

    def load_patients_list(self):
        """Загрузка списка пациентов"""
        try:
            patients = self.db.get_all_patients()
            self.patients_list.clear()

            for patient in patients:
                item_text = f"{patient['name']}"
                if patient.get('age'):
                    item_text += f", {patient['age']} лет"
                if patient.get('diagnosis'):
                    short_diagnosis = patient['diagnosis'][:30]
                    if len(patient['diagnosis']) > 30:
                        short_diagnosis += "..."
                    item_text += f"\n{short_diagnosis}"

                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, patient['id'])
                self.patients_list.addItem(item)

        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить пациентов:\n{str(e)}")

    def search_patients(self):
        """Поиск пациентов"""
        search_term = self.patient_search.text().strip()
        if not search_term:
            self.load_patients_list()
            return

        try:
            patients = self.db.search_patients(search_term)
            self.patients_list.clear()

            for patient in patients:
                item_text = f"{patient['name']}"
                if patient.get('age'):
                    item_text += f", {patient['age']} лет"

                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, patient['id'])
                self.patients_list.addItem(item)

        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Ошибка поиска:\n{str(e)}")

    def load_selected_patient(self, item):
        """Загрузка выбранного пациента"""
        if not item:
            return

        patient_id = item.data(Qt.UserRole)
        self.current_patient_id = patient_id

        try:
            patient = self.db.get_patient_by_id(patient_id)
            if patient:
                # Обновляем информацию
                info_text = f"{patient['name']}"
                if patient.get('age'):
                    info_text += f", {patient['age']} лет"
                if patient.get('gender'):
                    info_text += f" ({patient['gender']})"

                self.current_patient_label.setText(info_text)

                # Загружаем симптомы пациента в список (НОВОЕ)
                self.load_patient_symptoms_list(patient_id)

                # Загружаем шкалы
                scales = self.db.get_patient_scales(patient_id)
                if scales:
                    self.vas_input.setValue(scales.get('vas', 0))
                    self.oswestry_input.setValue(scales.get('oswestry', 0))
                    self.pain_detect_input.setValue(scales.get('pain_detect', 0))
                    self.hads_anxiety_input.setValue(scales.get('hads_anxiety', 0))
                    self.hads_depression_input.setValue(scales.get('hads_depression', 0))
                    self.pcs_input.setValue(scales.get('pcs', 0))

                # Загружаем историю
                self.load_diagnosis_history()

                # Активируем кнопку
                self.analyze_btn.setEnabled(True)

        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить пациента:\n{str(e)}")

    def load_patient_symptoms_list(self, patient_id):
        """Загрузка симптомов пациента в список (НОВЫЙ МЕТОД)"""
        try:
            symptoms = self.db.get_patient_symptoms(patient_id)
            self.patient_symptoms_list.clear()

            for symptom in symptoms:
                symptom_text = symptom.get('symptom_name', 'Нет названия')
                category = symptom.get('category', 'неизвестно')
                severity = symptom.get('severity', 0)

                item_text = f"{symptom_text}"
                if category != 'неизвестно':
                    item_text += f" [{category}]"
                if severity > 0:
                    item_text += f" (тяжесть: {severity})"

                item = QListWidgetItem(item_text)
                self.patient_symptoms_list.addItem(item)

            if not symptoms:
                self.patient_symptoms_list.addItem("Нет сохраненных симптомов")

        except Exception as e:
            print(f"Ошибка загрузки симптомов: {e}")

    def clear_patient_symptoms(self):
        """Очистка списка симптомов (только визуально)"""
        self.symptoms_text.clear()
        self.patient_symptoms_list.clear()
        self.patient_symptoms_list.addItem("Симптомы очищены")

    def show_new_patient_dialog(self):
        """Создание нового пациента"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Новый пациент")
        dialog.setModal(True)
        dialog.resize(400, 200)

        layout = QVBoxLayout(dialog)

        form_layout = QFormLayout()
        name_input = QLineEdit()
        age_input = QSpinBox()
        age_input.setRange(0, 120)
        gender_combo = QComboBox()
        gender_combo.addItems(['Мужской', 'Женский'])

        form_layout.addRow("ФИО *:", name_input)
        form_layout.addRow("Возраст:", age_input)
        form_layout.addRow("Пол:", gender_combo)

        layout.addLayout(form_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        if dialog.exec() == QDialog.Accepted:
            name = name_input.text().strip()
            if not name:
                QMessageBox.warning(self, "Ошибка", "Введите ФИО пациента")
                return

            try:
                patient_id = self.db.add_patient(
                    name=name,
                    age=age_input.value() if age_input.value() > 0 else None,
                    gender=gender_combo.currentText()
                )

                QMessageBox.information(self, "Успех", f"Пациент создан!\nID: {patient_id}")
                self.load_patients_list()

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось создать пациента:\n{str(e)}")

    def delete_selected_patient(self):
        """Удаление выбранного пациента"""
        if not self.current_patient_id:
            QMessageBox.warning(self, "Ошибка", "Выберите пациента для удаления")
            return

        reply = QMessageBox.question(
            self, "Подтверждение",
            "Удалить пациента и все его данные?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            try:
                self.db.delete_patient(self.current_patient_id)
                QMessageBox.information(self, "Успех", "Пациент удален")

                # Очищаем интерфейс
                self.current_patient_id = None
                self.current_patient_label.setText("Не выбран")
                self.patient_symptoms_list.clear()
                self.analyze_btn.setEnabled(False)

                self.load_patients_list()

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось удалить пациента:\n{str(e)}")

    # ========== МЕТОДЫ ЗАГРУЗКИ ФАЙЛОВ ==========

    def load_word_file(self):
        """Загрузка Word файла"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Выберите Word файл", "",
                "Word Files (*.docx *.doc);;All Files (*)"
            )

            if not file_path:
                return

            doc = Document(file_path)

            # Собираем весь текст
            full_text = []
            for para in doc.paragraphs:
                if para.text.strip():
                    full_text.append(para.text.strip())

            if not full_text:
                QMessageBox.warning(self, "Предупреждение", "Файл пуст или не содержит текста")
                return

            # Объединяем
            all_text = "\n".join(full_text)

            # Предпросмотр (первые 500 символов)
            preview = all_text[:500]
            if len(all_text) > 500:
                preview += "..."

            self.file_preview.setText(f"Файл: {os.path.basename(file_path)}\n\n{preview}")

            # Загружаем в поле симптомов (первые 10000 символов)
            symptoms_text = all_text[:10000]
            if len(all_text) > 10000:
                symptoms_text += "\n\n[Текст обрезан...]"

            self.symptoms_text.setText(symptoms_text)

            # Извлекаем шкалы
            self._extract_scales_from_text(all_text)

            QMessageBox.information(self, "Успех",
                                    f"Word файл загружен!\n"
                                    f"Символов: {len(all_text)}\n"
                                    f"Строк: {len(full_text)}")

        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить Word файл:\n{str(e)}")

    def load_excel_file(self):
        """Загрузка Excel файла"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Выберите Excel файл", "",
                "Excel Files (*.xlsx *.xls);;All Files (*)"
            )

            if not file_path:
                return

            # Читаем все листы
            df_dict = pd.read_excel(file_path, sheet_name=None)

            if not df_dict:
                QMessageBox.warning(self, "Предупреждение", "Файл пуст")
                return

            # Собираем данные из всех листов
            all_data = []
            for sheet_name, df in df_dict.items():
                if not df.empty:
                    # Преобразуем DataFrame в текст
                    for _, row in df.iterrows():
                        for cell in row:
                            if pd.notna(cell):
                                all_data.append(str(cell).strip())

            if not all_data:
                QMessageBox.warning(self, "Предупреждение", "Файл не содержит данных")
                return

            all_text = "\n".join(all_data)

            # Предпросмотр
            preview = all_text[:500]
            if len(all_text) > 500:
                preview += "..."

            self.file_preview.setText(f"Файл: {os.path.basename(file_path)}\n"
                                      f"Листов: {len(df_dict)}\n\n"
                                      f"{preview}")

            # Загружаем симптомы
            symptoms_text = all_text[:10000]
            self.symptoms_text.setText(symptoms_text)

            # Извлекаем шкалы
            self._extract_scales_from_text(all_text)

            QMessageBox.information(self, "Успех",
                                    f"Excel файл загружен!\n"
                                    f"Листов: {len(df_dict)}\n"
                                    f"Записей: {len(all_data)}")

        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить Excel файл:\n{str(e)}")

    def load_text_file(self):
        """Загрузка текстового файла"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Выберите текстовый файл", "",
                "Text Files (*.txt);;All Files (*)"
            )

            if not file_path:
                return

            with open(file_path, 'r', encoding='utf-8') as f:
                all_text = f.read()

            if not all_text.strip():
                QMessageBox.warning(self, "Предупреждение", "Файл пуст")
                return

            # Предпросмотр
            preview = all_text[:500]
            if len(all_text) > 500:
                preview += "..."

            self.file_preview.setText(f"Файл: {os.path.basename(file_path)}\n\n{preview}")

            # Загружаем симптомы
            symptoms_text = all_text[:10000]
            self.symptoms_text.setText(symptoms_text)

            QMessageBox.information(self, "Успех",
                                    f"Текстовый файл загружен!\n"
                                    f"Символов: {len(all_text)}")

        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def load_template(self, template_num):
        """Загрузка шаблона симптомов"""
        templates = {
            1: "Боль в пояснице\nБоль усиливается при сидении\nМышечное напряжение\nОграничение движений\nУтренняя скованность",
            2: "Боль в спине с иррадиацией в ногу\nОнемение в нижней конечности\nСлабость в стопе\nНарушение походки\nПокалывание",
            3: "Боль в нижней части спины\nЭмоциональная лабильность\nТревога\nДепрессия\nНарушение сна\nКатастрофизация боли"
        }

        if template_num in templates:
            self.symptoms_text.setText(templates[template_num])

    def _extract_scales_from_text(self, text):
        """Извлечение значений шкал из текста"""
        import re

        # ВАШ
        vas_match = re.search(r'ваш\s*[:=]?\s*(\d+)', text.lower())
        if vas_match:
            self.vas_input.setValue(int(vas_match.group(1)))

        # Освестри
        oswestry_match = re.search(r'освестри\s*[:=]?\s*(\d+)', text.lower())
        if oswestry_match:
            self.oswestry_input.setValue(int(oswestry_match.group(1)))

        # PainDetect
        paindetect_match = re.search(r'paindetect\s*[:=]?\s*(\d+)', text.lower())
        if paindetect_match:
            self.pain_detect_input.setValue(int(paindetect_match.group(1)))

        # HADS
        hads_a_match = re.search(r'hads.*тревог\w*\s*[:=]?\s*(\d+)', text.lower())
        if hads_a_match:
            self.hads_anxiety_input.setValue(int(hads_a_match.group(1)))

        hads_d_match = re.search(r'hads.*депресс\w*\s*[:=]?\s*(\d+)', text.lower())
        if hads_d_match:
            self.hads_depression_input.setValue(int(hads_d_match.group(1)))

        # PCS
        pcs_match = re.search(r'pcs\s*[:=]?\s*(\d+)', text.lower())
        if pcs_match:
            self.pcs_input.setValue(int(pcs_match.group(1)))

    # ========== АНАЛИЗ СИМПТОМОВ ==========

    def analyze_pain(self):
        """Анализ симптомов"""
        if not self.current_patient_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите пациента")
            return

        symptoms = self.symptoms_text.toPlainText().strip()
        if not symptoms:
            QMessageBox.warning(self, "Ошибка", "Введите симптомы для анализа")
            return

        scales = {
            'vas': self.vas_input.value(),
            'oswestry': self.oswestry_input.value(),
            'pain_detect': self.pain_detect_input.value(),
            'hads_anxiety': self.hads_anxiety_input.value(),
            'hads_depression': self.hads_depression_input.value(),
            'pcs': self.pcs_input.value()
        }

        try:
            # Анализ с поиском похожих симптомов
            analysis_results = self.analyzer.analyze_symptoms(symptoms, scales)

            # Отображаем детальный анализ симптомов
            self.display_detailed_symptoms_analysis(analysis_results)

            # +++ ВОТ ЭТУ СТРОКУ ДОБАВЬТЕ: Отображаем результаты в интерфейсе +++
            self.display_results(analysis_results, symptoms)
            # +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            # Сохраняем симптомы в БД
            if symptoms:
                symptoms_list = []
                for line in symptoms.split('\n'):
                    if line.strip():
                        symptoms_list.append({
                            'name': line.strip(),
                            'category': 'неизвестно',
                            'severity': 0
                        })

                if symptoms_list:
                    self.db.add_symptoms(self.current_patient_id, symptoms_list)
                    self.load_patient_symptoms_list(self.current_patient_id)

            # Сохраняем диагноз
            diagnosis_data = {
                'type': analysis_results['diagnosis']['type'],
                'subgroup': analysis_results['diagnosis']['subgroup'],
                'subtype': analysis_results['diagnosis'].get('subtype', ''),
                'confidence': analysis_results['confidence'],
                'recommended_treatment': 'консервативное',
                'treatment_model': '; '.join(analysis_results['recommendations']),
                'severity': analysis_results['diagnosis'].get('severity', 'умеренная')
            }

            self.db.add_diagnosis(self.current_patient_id, diagnosis_data)

            # Сохраняем шкалы
            self.db.add_scales(self.current_patient_id, scales)

            # Показываем общие результаты (в консоль)
            self.display_general_results(analysis_results)

            # Обновляем историю
            self.load_diagnosis_history()

            QMessageBox.information(self, "Успех",
                                    f"✅ Анализ завершен!\n"
                                    f"Достоверность: {analysis_results['confidence']:.1f}%\n"
                                    f"Сопоставлено симптомов: {analysis_results['statistics']['matched_symptoms_count']}/{analysis_results['statistics']['total_input_symptoms']}")

        except Exception as e:
            QMessageBox.warning(self, "Ошибка анализа",
                                f"Не удалось проанализировать симптомы:\n{str(e)}")

    def display_results(self, results, symptoms_text):
        """Отображение результатов анализа"""
        diagnosis = results['diagnosis']

        # 1. Диагноз
        diagnosis_html = f"""
        <div style='background-color: #f0f8ff; padding: 15px; border-radius: 10px;'>
            <h2>Диагностическое заключение</h2>
            <p><b>Тип боли:</b> <span style='color: #1976d2; font-weight: bold;'>
            {diagnosis.get('type', 'не определен').upper()}</span></p>
            <p><b>Подгруппа:</b> <span style='color: #388e3c; font-weight: bold;'>
            {diagnosis.get('subgroup', 'не определена').upper()}</span></p>
            <p><b>Подтип:</b> {diagnosis.get('subtype', 'Не определен')}</p>
            <p><b>Тяжесть:</b> {diagnosis.get('severity', 'умеренная').upper()}</p>
            <p><b>Достоверность:</b> 
            <span style='font-size: 18px; color: #4caf50; font-weight: bold;'>
            {results.get('confidence', 0):.1f}%</span></p>
        </div>
        """
        self.diagnosis_text.setHtml(diagnosis_html)

        # 2. Рекомендации
        treatment_html = "<h3>Рекомендации по лечению:</h3><ul>"
        recommendations = results.get('recommendations', [])
        if recommendations:
            for rec in recommendations:
                treatment_html += f"<li>{rec}</li>"
        else:
            treatment_html += "<li>Рекомендации не сгенерированы</li>"
        treatment_html += "</ul>"

        self.treatment_text.setHtml(treatment_html)

        # 3. Обнаруженные симптомы - УБРАТЬ ЭТОТ КОД!
        # Не нужно здесь перезаписывать detected_symptoms_text,
        # так как это уже сделано в display_detailed_symptoms_analysis()

        # Вместо этого, давайте просто выведем отладочную информацию:
        print(f"✓ Диагноз отображен: {diagnosis.get('type')}")
        print(f"✓ Рекомендаций: {len(recommendations)}")

    def load_diagnosis_history(self):
        """Загрузка истории диагнозов"""
        if not self.current_patient_id:
            return

        try:
            diagnoses = self.db.get_patient_diagnoses(self.current_patient_id)
            self.history_table.setRowCount(len(diagnoses))

            for row, diagnosis in enumerate(diagnoses):
                # Дата
                date_item = QTableWidgetItem(str(diagnosis.get('created_at', ''))[:19])
                self.history_table.setItem(row, 0, date_item)

                # Тип
                type_item = QTableWidgetItem(diagnosis.get('diagnosis_type', ''))
                self.history_table.setItem(row, 1, type_item)

                # Подгруппа
                subgroup_item = QTableWidgetItem(diagnosis.get('subgroup', ''))
                self.history_table.setItem(row, 2, subgroup_item)

                # Достоверность
                conf_item = QTableWidgetItem(f"{diagnosis.get('confidence', 0):.1f}%")
                self.history_table.setItem(row, 3, conf_item)

                # Тяжесть
                severity_item = QTableWidgetItem(diagnosis.get('severity', ''))
                self.history_table.setItem(row, 4, severity_item)

                # Лечение
                treatment_item = QTableWidgetItem(diagnosis.get('treatment_model', '')[:50])
                self.history_table.setItem(row, 5, treatment_item)

            self.history_table.resizeColumnsToContents()

        except Exception as e:
            print(f"Ошибка загрузки истории: {e}")

    def closeEvent(self, event):
        """Закрытие приложения"""
        self.db.close()
        event.accept()

    def display_detailed_symptoms_analysis(self, analysis_results):
        """Отображение детального анализа симптомов"""
        html = "<h3>🔍 Детальный анализ сопоставления симптомов</h3>"

        if not analysis_results.get('input_symptoms'):
            html += "<p>Симптомы не введены</p>"
            self.detected_symptoms_text.setHtml(html)
            return

        # Сводка статистики
        stats = analysis_results.get('statistics', {})
        html += f"""
        <div style='background-color: #f5f5f5; padding: 10px; border-radius: 5px; margin-bottom: 15px;'>
            <h4>📊 Сводная статистика:</h4>
            <p><b>Всего введено симптомов:</b> {stats.get('total_input_symptoms', 0)}</p>
            <p><b>Сопоставлено с базой:</b> {stats.get('matched_symptoms_count', 0)}</p>
            <p><b>Красные флаги:</b> {stats.get('red_flags_count', 0)}</p>
            <p><b>Специфические симптомы:</b> {stats.get('specific_count', 0)}</p>
            <p><b>Неспецифические симптомы:</b> {stats.get('nonspecific_count', 0)}</p>
            <p><b>Желтые флаги:</b> {stats.get('yellow_flags_count', 0)}</p>
            <p><b>Голубые флаги:</b> {stats.get('blue_flags_count', 0)}</p>
            <p><b>Черные флаги:</b> {stats.get('black_flags_count', 0)}</p>
        </div>
        """

        # Введенные симптомы
        html += "<h4>📝 Введенные симптомы:</h4><ul>"
        input_symptoms = analysis_results.get('input_symptoms', [])
        for i, symptom in enumerate(input_symptoms[:30], 1):  # Показываем первые 30
            html += f"<li>{symptom}</li>"
        html += "</ul>"

        if len(input_symptoms) > 30:
            html += f"<p>... и еще {len(input_symptoms) - 30} симптомов</p>"

        # Красные флаги (детально)
        if analysis_results.get('red_flags_found'):
            html += "<h4 style='color: #d32f2f;'>🔴 Обнаруженные красные флаги:</h4>"
            for item in analysis_results['red_flags_found']:
                best_match = item.get('best_match', ['Не найдено', 0])
                match_text = best_match[0] if isinstance(best_match, list) else best_match
                match_score = best_match[1] if isinstance(best_match, list) else 0

                html += f"""
                <div style='margin: 5px 0; padding: 8px; border-left: 4px solid #d32f2f; background-color: #ffebee;'>
                    <b>«{item.get('input', 'Неизвестно')}»</b><br>
                    <span style='color: #666;'>→ <b>{match_text}</b> ({match_score * 100:.1f}% совпадение)</span>
                </div>
                """

        # Специфические симптомы
        if analysis_results.get('specific_symptoms_found'):
            html += "<h4 style='color: #1976d2;'>🔵 Обнаруженные специфические симптомы:</h4>"
            for item in analysis_results['specific_symptoms_found']:
                best_match = item.get('best_match', ['Не найдено', 0])
                match_text = best_match[0] if isinstance(best_match, list) else best_match
                match_score = best_match[1] if isinstance(best_match, list) else 0

                html += f"""
                <div style='margin: 5px 0; padding: 8px; border-left: 4px solid #1976d2; background-color: #e3f2fd;'>
                    <b>«{item.get('input', 'Неизвестно')}»</b><br>
                    <span style='color: #666;'>→ <b>{match_text}</b> ({match_score * 100:.1f}% совпадение)</span>
                </div>
                """

        # Желтые флаги
        if analysis_results.get('yellow_flags_found'):
            html += "<h4 style='color: #f57c00;'>🟡 Обнаруженные желтые флаги:</h4>"
            for item in analysis_results['yellow_flags_found']:
                best_match = item.get('best_match', ['Не найдено', 0])
                match_text = best_match[0] if isinstance(best_match, list) else best_match
                match_score = best_match[1] if isinstance(best_match, list) else 0

                html += f"""
                <div style='margin: 5px 0; padding: 8px; border-left: 4px solid #f57c00; background-color: #fff3e0;'>
                    <b>«{item.get('input', 'Неизвестно')}»</b><br>
                    <span style='color: #666;'>→ <b>{match_text}</b> ({match_score * 100:.1f}% совпадение)</span>
                </div>
                """

        # Неспецифические симптомы
        if analysis_results.get('nonspecific_symptoms_found'):
            html += "<h4 style='color: #388e3c;'>🟢 Обнаруженные неспецифические симптомы:</h4>"
            count = len(analysis_results['nonspecific_symptoms_found'])
            html += f"<p>Найдено: {count} неспецифических симптомов</p>"

            # Показываем только первые 5 для экономии места
            for item in analysis_results['nonspecific_symptoms_found'][:5]:
                best_match = item.get('best_match', ['Не найдено', 0])
                match_text = best_match[0] if isinstance(best_match, list) else best_match

                html += f"<li><b>{item.get('input', 'Неизвестно')}</b> → {match_text}</li>"

            if count > 5:
                html += f"<p>... и еще {count - 5} неспецифических симптомов</p>"

        # Добавляем информацию о типе боли
        if analysis_results.get('diagnosis'):
            diagnosis = analysis_results['diagnosis']
            html += f"""
            <div style='background-color: #e8f5e9; padding: 10px; border-radius: 5px; margin-top: 15px;'>
                <h4>🧬 Итоговый диагноз:</h4>
                <p><b>Тип боли:</b> {diagnosis.get('type', 'не определен')}</p>
                <p><b>Подгруппа:</b> {diagnosis.get('subgroup', 'не определена')}</p>
                <p><b>Тяжесть:</b> {diagnosis.get('severity', 'не определена')}</p>
                <p><b>Достоверность анализа:</b> {analysis_results.get('confidence', 0):.1f}%</p>
            </div>
            """

        self.detected_symptoms_text.setHtml(html)


# ========== ТОЧКА ВХОДА ==========

if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication

    print("🚀 Запуск приложения диагностики боли...")

    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = PainDiagnosisApp()
    window.show()

    sys.exit(app.exec_())