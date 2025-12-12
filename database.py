"""
database.py - Полностью исправленная версия
"""

import pymysql
from pymysql import Error
from pymysql.cursors import DictCursor
from config import DB_CONFIG
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DatabaseManager:
    def __init__(self):
        self.connection = None
        self.connect()
        self.create_tables()

    def connect(self):
        try:
            self.connection = pymysql.connect(**DB_CONFIG)
            self.connection.autocommit = True
            logger.info("✅ Подключение к базе данных успешно!")
        except Error as e:
            logger.error(f"❌ Ошибка подключения: {e}")
            raise

    def create_tables(self):
        cursor = None
        try:
            cursor = self.connection.cursor()

            # Таблица пациентов
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS patients (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(100) NOT NULL,
                    age INT,
                    gender VARCHAR(10),
                    admission_date DATE,
                    discharge_date DATE,
                    diagnosis TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """)

            # Таблица симптомов - TEXT вместо VARCHAR!
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS symptoms (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    patient_id INT NOT NULL,
                    symptom_name TEXT,  -- ИСПРАВЛЕНО: TEXT
                    category VARCHAR(50),
                    severity INT DEFAULT 0,
                    notes TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (patient_id) 
                        REFERENCES patients(id) 
                        ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """)

            # Таблица диагнозов
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS diagnoses (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    patient_id INT NOT NULL,
                    diagnosis_type VARCHAR(50),
                    subgroup VARCHAR(50),
                    subtype VARCHAR(50),
                    confidence FLOAT,
                    recommended_treatment VARCHAR(50),
                    treatment_model TEXT,
                    severity VARCHAR(20),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (patient_id) 
                        REFERENCES patients(id) 
                        ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """)

            # Таблица шкал
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS patient_scales (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    patient_id INT NOT NULL,
                    vas INT DEFAULT 0,
                    oswestry INT DEFAULT 0,
                    pain_detect INT DEFAULT 0,
                    hads_anxiety INT DEFAULT 0,
                    hads_depression INT DEFAULT 0,
                    pcs INT DEFAULT 0,
                    measured_date DATE,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (patient_id) 
                        REFERENCES patients(id) 
                        ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """)

            self.connection.commit()
            logger.info("✅ Все таблицы созданы успешно!")

        except Error as e:
            logger.error(f"❌ Ошибка создания таблиц: {e}")
            raise
        finally:
            if cursor:
                cursor.close()

    # ========== МЕТОДЫ ДЛЯ ПАЦИЕНТОВ ==========

    def get_all_patients(self):
        cursor = self.connection.cursor(cursor=DictCursor)
        cursor.execute("""
            SELECT id, name, age, gender, admission_date, 
                   discharge_date, diagnosis, created_at
            FROM patients ORDER BY created_at DESC
        """)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_patient_by_id(self, patient_id):
        cursor = self.connection.cursor(cursor=DictCursor)
        cursor.execute("SELECT * FROM patients WHERE id = %s", (patient_id,))
        result = cursor.fetchone()
        cursor.close()
        return result

    def search_patients(self, search_term):
        cursor = self.connection.cursor(cursor=DictCursor)
        cursor.execute("""
            SELECT * FROM patients 
            WHERE name LIKE %s OR diagnosis LIKE %s
            ORDER BY name
        """, (f"%{search_term}%", f"%{search_term}%"))
        result = cursor.fetchall()
        cursor.close()
        return result

    def add_patient(self, name, age=None, gender=None, admission_date=None,
                   discharge_date=None, diagnosis=None):
        cursor = self.connection.cursor()
        cursor.execute("""
            INSERT INTO patients (name, age, gender, admission_date, discharge_date, diagnosis)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (name, age, gender, admission_date, discharge_date, diagnosis))
        patient_id = cursor.lastrowid
        self.connection.commit()
        cursor.close()
        return patient_id

    def delete_patient(self, patient_id):
        cursor = self.connection.cursor()
        cursor.execute("DELETE FROM patients WHERE id = %s", (patient_id,))
        self.connection.commit()
        cursor.close()
        return True

    # ========== МЕТОДЫ ДЛЯ СИМПТОМОВ ==========

    def add_symptoms(self, patient_id, symptoms_list):
        """Добавить симптомы - текст любой длины"""
        if not symptoms_list:
            return

        cursor = self.connection.cursor()
        try:
            for symptom in symptoms_list:
                cursor.execute("""
                    INSERT INTO symptoms (patient_id, symptom_name, category, severity, notes)
                    VALUES (%s, %s, %s, %s, %s)
                """, (
                    patient_id,
                    str(symptom.get('name', ''))[:10000],  # Безопасное обрезание
                    symptom.get('category', 'неизвестно'),
                    symptom.get('severity', 0),
                    symptom.get('notes', '')
                ))

            self.connection.commit()
            logger.info(f"✅ Симптомы добавлены для пациента {patient_id}")

        except Error as e:
            logger.error(f"❌ Ошибка добавления симптомов: {e}")
            raise
        finally:
            cursor.close()

    def get_patient_symptoms(self, patient_id):
        cursor = self.connection.cursor(cursor=DictCursor)
        cursor.execute("""
            SELECT symptom_name, category, severity, notes, created_at
            FROM symptoms WHERE patient_id = %s ORDER BY created_at DESC
        """, (patient_id,))
        result = cursor.fetchall()
        cursor.close()
        return result

    # ========== МЕТОДЫ ДЛЯ ДИАГНОЗОВ ==========

    def add_diagnosis(self, patient_id, diagnosis_data):
        """Добавить диагноз - ПРАВИЛЬНЫЙ вызов с 2 аргументами"""
        cursor = self.connection.cursor()
        try:
            cursor.execute("""
                INSERT INTO diagnoses 
                (patient_id, diagnosis_type, subgroup, subtype, confidence, 
                 recommended_treatment, treatment_model, severity)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                patient_id,
                diagnosis_data.get('type', 'неспецифическая'),
                diagnosis_data.get('subgroup', 'ноцицептивная'),
                diagnosis_data.get('subtype', ''),
                diagnosis_data.get('confidence', 0.0),
                diagnosis_data.get('recommended_treatment', ''),
                diagnosis_data.get('treatment_model', ''),
                diagnosis_data.get('severity', 'умеренная')
            ))

            diagnosis_id = cursor.lastrowid
            self.connection.commit()
            logger.info(f"✅ Диагноз добавлен: ID {diagnosis_id}")
            return diagnosis_id

        except Error as e:
            logger.error(f"❌ Ошибка добавления диагноза: {e}")
            raise
        finally:
            cursor.close()

    def get_patient_diagnoses(self, patient_id):
        cursor = self.connection.cursor(cursor=DictCursor)
        cursor.execute("""
            SELECT * FROM diagnoses 
            WHERE patient_id = %s ORDER BY created_at DESC
        """, (patient_id,))
        result = cursor.fetchall()
        cursor.close()
        return result

    # ========== МЕТОДЫ ДЛЯ ШКАЛ ==========

    def add_scales(self, patient_id, scales_data):
        cursor = self.connection.cursor()
        cursor.execute("""
            INSERT INTO patient_scales 
            (patient_id, vas, oswestry, pain_detect, hads_anxiety, hads_depression, pcs, measured_date)
            VALUES (%s, %s, %s, %s, %s, %s, %s, CURDATE())
        """, (
            patient_id,
            scales_data.get('vas', 0),
            scales_data.get('oswestry', 0),
            scales_data.get('pain_detect', 0),
            scales_data.get('hads_anxiety', 0),
            scales_data.get('hads_depression', 0),
            scales_data.get('pcs', 0)
        ))
        self.connection.commit()
        cursor.close()

    def get_patient_scales(self, patient_id):
        cursor = self.connection.cursor(cursor=DictCursor)
        cursor.execute("""
            SELECT * FROM patient_scales 
            WHERE patient_id = %s ORDER BY created_at DESC LIMIT 1
        """, (patient_id,))
        result = cursor.fetchone()
        cursor.close()
        return result

    def close(self):
        if self.connection:
            self.connection.close()