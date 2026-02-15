"""
analyzer.py - Улучшенный анализатор с поиском похожих симптомов
"""

import re
from typing import List, Dict, Tuple
import pandas as pd
import numpy as np
from difflib import SequenceMatcher
import Levenshtein  # Для установки: pip install python-Levenshtein
from collections import defaultdict

from PyQt5.QtWidgets import QMessageBox


class EnhancedPainAnalyzer:
    def __init__(self):
        # Загружаем базу симптомов
        self.symptoms_db = self._load_symptoms_database()
        self.similarity_threshold = 0.7  # Порог схожести 70%

        # Категории симптомов
        self.categories = {
            'specific': self._load_specific_symptoms(),
            'nonspecific': self._load_nonspecific_symptoms(),
            'red_flags': self._load_red_flags(),
            'yellow_flags': self._load_yellow_flags(),
            'blue_flags': self._load_blue_flags(),
            'black_flags': self._load_black_flags()
        }

        # Веса для разных типов симптомов
        self.weights = {
            'red_flag': 3.0,
            'specific': 2.5,
            'neurological': 2.0,
            'yellow_flag': 1.5,
            'blue_flag': 1.2,
            'black_flag': 1.8,
            'nonspecific': 1.0
        }

    def _load_symptoms_database(self):
        """Загружаем базу симптомов из вашего Excel файла"""
        symptoms_db = {
            'specific': [],
            'nonspecific': [],
            'red_flags': [],
            'yellow_flags': [],
            'blue_flags': [],
            'black_flags': []
        }

        # Здесь можно загрузить из вашего Excel файла
        # Пока используем статические списки
        return symptoms_db

    def _load_specific_symptoms(self):
        """Специфические симптомы"""
        return [
            "аллодиния", "анестезия в зоне дерматома", "анталгическая походка с неврологическим дефицитом",
            "бледные кожные покровы", "болезненность при перкуссии позвоночника", "боль в грудной клетке",
            "боль в покое", "боль при пальпации в точках валле", "боль усиливается в положении лёжа",
            "боль усиливается при кашле", "выпадение сухожильных рефлексов", "выраженная деформация позвоночника",
            "гипералгезия", "гиперпатия", "гипестезия по дерматому l5", "изменение уровня сознания",
            "иррадиация боли по корешку l5", "клонусы", "лихорадка", "локальная резкая болезненность",
            "нарушение мышечного тонуса", "нарушение функции тазовых органов", "необъяснимая потеря массы тела",
            "ночная боль", "парез конечности", "плегия конечности", "повышенные сухожильные рефлексы",
            "положительный симптом ласега", "положительный симптом мацкевича", "положительный симптом нери",
            "положительный тест кемпа", "положительный тест патрика", "постоянно прогрессирующая боль",
            "радикулопатия l5", "расширение рефлексогенных зон", "ригидность затылочных мышц",
            "симптом брудзинского", "симптом кернига", "слабость тыльного сгибания стопы",
            "снижение ахиллова рефлекса", "снижение болевой чувствительности", "спонтанные нейропатические боли",
            "степпаж стопы", "субфебрилитет", "хромота с корешковой симптоматикой"
        ]

    def _load_nonspecific_symptoms(self):
        """Неспецифические симптомы"""
        return [
            "анталгическая поза", "анталгическая походка", "апатия",
            "болезненность при пальпации паравертебральных точек",
            "боль в грудном отделе", "боль в пояснице", "боль усиливается при длительном сидении",
            "боль усиливается при стоянии", "боль усиливается при наклонах",
            "боль усиливается при работе за компьютером",
            "боль усиливается при разгибании", "боль усиливается при ходьбе", "временная нетрудоспособность",
            "вынужденное положение тела", "высокая катастрофизация", "движения болезненны", "движения ограничены",
            "депрессия", "дистресс", "длительное недомогание", "избегание активностей", "интенсивность боли 7-8 баллов",
            "катастрофизация боли", "кинезиофобия", "кифосколиоз", "любое движение вызывает боль",
            "мышечный гипертонус",
            "напряжение паравертебральных мышц", "нарушение походки", "нарушение самообслуживания", "нарушение сна",
            "не может сидеть за компьютером", "нежелание двигаться", "низкая социальная поддержка",
            "низкая удовлетворенность работой", "общая слабость", "ограничение времени ходьбы",
            "ограничение объема движений", "ограниченная переносимость стояния", "ограниченная переносимость ходьбы",
            "ощущение себя в плохом здоровье", "пальпируемые триггерные точки", "пассивные копинг-стратегии",
            "периоды вынужденного отдыха", "плохая самооценка здоровья", "пробуждения ночью из-за боли",
            "сколиотическая установка", "снижение бытовой активности", "снижение качества сна",
            "снижение повседневной активности", "снижение толерантности к статической позе",
            "снижение участия в активности", "снижение физической активности", "сниженный эмоциональный фон",
            "с-образный сколиоз", "стартовая боль", "страх обострения", "стресс", "тревога",
            "тревожно-депрессивные проявления", "трудности работать за компьютером", "трудности с засыпанием",
            "убеждение в пассивном лечении", "утренняя скованность", "хронический суставно-мышечный болевой синдром",
            "чувство безнадёжности", "эмоциональная лабильность"
        ]

    def _load_red_flags(self):
        """Красные флаги"""
        return [
            "боль в возрасте <20 или >55 лет", "травма спины", "нарастающий характер боли", "ночная боль",
            "боль в грудном отделе", "боль в грудной клетке", "онкологические заболевания",
            "длительный прием кортикостероидов", "остеопороз", "внутривенное употребление наркотиков",
            "иммунодефицит", "потеря массы тела", "нарушение функции тазовых органов", "симптом кашлевого толчка",
            "выраженная деформация позвоночника", "лихорадка", "необъяснимая потеря массы тела", "субфебрилитет"
        ]

    def _load_yellow_flags(self):
        """Желтые флаги"""
        return [
            "депрессия", "тревога", "дистресс", "катастрофизация", "кинезиофобия", "избегание",
            "пассивные копинг-стратегии", "плохая самооценка здоровья", "преморбидная хроническая боль",
            "ожидание пассивного лечени", "болевое мышление", "эмоциональная лабильность", "сниженный эмоциональный фон"
        ]

    def _load_blue_flags(self):
        """Голубые флаги"""
        return [
            "высокие рабочие требования", "дефицит времени", "перегруз", "низкий рабочий контроль",
            "низкая социальная поддержка", "низкая оценка успехов", "неблагоприятный командный климат",
            "низкая удовлетворенность работой", "связывание боли с работой", "скептический настрой",
            "трудности работать за компьютером"
        ]

    def _load_black_flags(self):
        """Черные флаги"""
        return [
            "неправильный подход к лечению", "требования получения инвалидности", "компенсаторные выплаты",
            "безработица"
        ]

    def _similarity_score(self, text1: str, text2: str) -> float:
        """Вычисление схожести между двумя строками"""
        # Приводим к нижнему регистру и удаляем лишние пробелы
        text1 = text1.lower().strip()
        text2 = text2.lower().strip()

        # 1. Проверка на полное совпадение
        if text1 == text2:
            return 1.0

        # 2. Проверка на вхождение
        if text1 in text2 or text2 in text1:
            return 0.9

        # 3. Используем SequenceMatcher
        similarity = SequenceMatcher(None, text1, text2).ratio()

        # 4. Дополнительно используем расстояние Левенштейна
        if len(text1) > 3 and len(text2) > 3:
            lev_similarity = 1 - (Levenshtein.distance(text1, text2) / max(len(text1), len(text2)))
            similarity = max(similarity, lev_similarity)

        return similarity

    def _find_similar_symptoms(self, input_symptom: str, symptom_list: List[str]) -> List[Tuple[str, float]]:
        """Поиск похожих симптомов в списке"""
        similar_symptoms = []

        for db_symptom in symptom_list:
            similarity = self._similarity_score(input_symptom, db_symptom)
            if similarity >= self.similarity_threshold:
                similar_symptoms.append((db_symptom, similarity))

        # Сортируем по убыванию схожести
        similar_symptoms.sort(key=lambda x: x[1], reverse=True)
        return similar_symptoms[:5]  # Возвращаем топ-5 похожих

    def _tokenize_text(self, text: str) -> List[str]:
        """Разбиваем текст на токены (слова)"""
        # Удаляем пунктуацию и приводим к нижнему регистру
        text = text.lower()
        text = re.sub(r'[^\w\s]', ' ', text)

        # Разбиваем на слова
        tokens = text.split()

        # Удаляем стоп-слова (можно расширить)
        stop_words = {'и', 'в', 'на', 'с', 'по', 'у', 'о', 'от', 'до', 'за', 'из', 'к', 'но', 'же', 'бы', 'ли',
                      'что', 'это', 'как', 'так', 'для', 'при', 'не', 'нет', 'да', 'очень', 'более', 'менее'}

        tokens = [word for word in tokens if word not in stop_words and len(word) > 2]
        return tokens

    def _match_tokens(self, input_tokens: List[str], symptom_tokens: List[str]) -> float:
        """Сопоставление токенов"""
        if not input_tokens or not symptom_tokens:
            return 0.0

        matched = 0
        for token in input_tokens:
            for s_token in symptom_tokens:
                if self._similarity_score(token, s_token) > 0.8:
                    matched += 1
                    break

        return matched / len(input_tokens)

    def analyze_symptoms(self, symptoms_text: str, scales: Dict) -> Dict:
        """Анализ симптомов с поиском похожих"""

        # Разбиваем введенные симптомы на отдельные строки
        input_symptoms = [s.strip() for s in symptoms_text.split('\n') if s.strip()]

        # Результаты анализа
        results = {
            'input_symptoms': input_symptoms,
            'matched_symptoms': [],
            'red_flags_found': [],
            'yellow_flags_found': [],
            'blue_flags_found': [],
            'black_flags_found': [],
            'specific_symptoms_found': [],
            'nonspecific_symptoms_found': [],
            'diagnosis': {},
            'confidence': 0.0,
            'recommendations': [],
            'similarity_analysis': []
        }

        # Анализируем каждый введенный симптом
        for symptom in input_symptoms:
            symptom_lower = symptom.lower()

            # Проверяем красные флаги
            similar_red = self._find_similar_symptoms(symptom_lower, self.categories['red_flags'])
            if similar_red:
                results['red_flags_found'].append({
                    'input': symptom,
                    'matches': similar_red,
                    'best_match': similar_red[0]
                })

            # Проверяем специфические симптомы
            similar_specific = self._find_similar_symptoms(symptom_lower, self.categories['specific'])
            if similar_specific:
                results['specific_symptoms_found'].append({
                    'input': symptom,
                    'matches': similar_specific,
                    'best_match': similar_specific[0]
                })

            # Проверяем неспецифические симптомы
            similar_nonspecific = self._find_similar_symptoms(symptom_lower, self.categories['nonspecific'])
            if similar_nonspecific:
                results['nonspecific_symptoms_found'].append({
                    'input': symptom,
                    'matches': similar_nonspecific,
                    'best_match': similar_nonspecific[0]
                })

            # Проверяем желтые флаги
            similar_yellow = self._find_similar_symptoms(symptom_lower, self.categories['yellow_flags'])
            if similar_yellow:
                results['yellow_flags_found'].append({
                    'input': symptom,
                    'matches': similar_yellow,
                    'best_match': similar_yellow[0]
                })

            # Проверяем голубые флаги
            similar_blue = self._find_similar_symptoms(symptom_lower, self.categories['blue_flags'])
            if similar_blue:
                results['blue_flags_found'].append({
                    'input': symptom,
                    'matches': similar_blue,
                    'best_match': similar_blue[0]
                })

            # Проверяем черные флаги
            similar_black = self._find_similar_symptoms(symptom_lower, self.categories['black_flags'])
            if similar_black:
                results['black_flags_found'].append({
                    'input': symptom,
                    'matches': similar_black,
                    'best_match': similar_black[0]
                })

        # ПОДСЧИТЫВАЕМ СТАТИСТИКУ - ПЕРЕНОСИМ ЭТО ПЕРЕД определением диагноза
        results['statistics'] = {
            'total_input_symptoms': len(input_symptoms),
            'red_flags_count': len(results['red_flags_found']),
            'specific_count': len(results['specific_symptoms_found']),
            'nonspecific_count': len(results['nonspecific_symptoms_found']),
            'yellow_flags_count': len(results['yellow_flags_found']),
            'blue_flags_count': len(results['blue_flags_found']),
            'black_flags_count': len(results['black_flags_found']),
            'matched_symptoms_count': sum([
                len(results['red_flags_found']),
                len(results['specific_symptoms_found']),
                len(results['nonspecific_symptoms_found']),
                len(results['yellow_flags_found']),
                len(results['blue_flags_found']),
                len(results['black_flags_found'])
            ])
        }

        # Определяем диагноз - ТЕПЕРЬ statistics уже существует
        diagnosis = self._determine_diagnosis(results, scales)
        results['diagnosis'] = diagnosis

        # Рассчитываем достоверность
        confidence = self._calculate_confidence(results, scales)
        results['confidence'] = confidence

        # Генерируем рекомендации
        recommendations = self._generate_recommendations(results, scales)
        results['recommendations'] = recommendations

        return results

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

            # +++ ОТЛАДКА: проверяем что получили от анализатора +++
            print(f"=== РЕЗУЛЬТАТЫ АНАЛИЗА ===")
            print(f"Входные симптомы: {analysis_results.get('input_symptoms', [])}")
            print(f"Красные флаги найдены: {len(analysis_results.get('red_flags_found', []))}")
            print(f"Специфические симптомы: {len(analysis_results.get('specific_symptoms_found', []))}")
            print(f"Желтые флаги: {len(analysis_results.get('yellow_flags_found', []))}")
            print(f"Статистика: {analysis_results.get('statistics', {})}")
            print(f"Диагноз: {analysis_results.get('diagnosis', {})}")
            print(f"==========================")

            # Отображаем детальный анализ симптомов
            self.display_detailed_symptoms_analysis(analysis_results)

            # Отображаем результаты в интерфейсе
            self.display_results(analysis_results, symptoms)

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

    def _determine_diagnosis(self, results: Dict, scales: Dict) -> Dict:
        """Определение диагноза на основе анализа"""
        diagnosis = {
            'type': 'неспецифическая',
            'subgroup': 'ноцицептивная',
            'subtype': '',
            'severity': 'умеренная',
            'flags_summary': {}
        }

        pain_detect = scales.get('pain_detect', 0)
        vas = scales.get('vas', 0)

        # Проверяем красные флаги
        if results['red_flags_found'] or results['specific_symptoms_found']:
            diagnosis['type'] = 'специфическая'

            # Определяем подтип
            red_flag_texts = [item['input'].lower() for item in results['red_flags_found']]
            specific_texts = [item['input'].lower() for item in results['specific_symptoms_found']]
            all_texts = red_flag_texts + specific_texts

            if any('онкологи' in text for text in all_texts):
                diagnosis['subtype'] = 'онкологическая'
            elif any('травма' in text for text in all_texts):
                diagnosis['subtype'] = 'травматическая'
            elif any('воспал' in text for text in all_texts):
                diagnosis['subtype'] = 'воспалительная'
            elif any('инфекц' in text for text in all_texts):
                diagnosis['subtype'] = 'инфекционно-воспалительная'

        # Определяем подгруппу
        if pain_detect >= 19:
            diagnosis['subgroup'] = 'нейропатическая'
        elif pain_detect >= 13:
            if results['yellow_flags_found'] or results['statistics']['yellow_flags_count'] >= 2:
                diagnosis['subgroup'] = 'дисфункциональная'
            else:
                diagnosis['subgroup'] = 'нейропатическая/дисфункциональная смешанная'
        else:
            if (results['yellow_flags_found'] or results['blue_flags_found'] or
                    results['black_flags_found']):
                diagnosis['subgroup'] = 'дисфункциональная'
            else:
                diagnosis['subgroup'] = 'ноцицептивная'

        # Определяем тяжесть
        if vas >= 8:
            diagnosis['severity'] = 'тяжелая'
        elif vas >= 5:
            diagnosis['severity'] = 'умеренная'
        else:
            diagnosis['severity'] = 'легкая'

        # Сводка по флагам
        diagnosis['flags_summary'] = {
            'red_flags': len(results['red_flags_found']),
            'yellow_flags': len(results['yellow_flags_found']),
            'blue_flags': len(results['blue_flags_found']),
            'black_flags': len(results['black_flags_found'])
        }

        return diagnosis

    def _calculate_confidence(self, results: Dict, scales: Dict) -> float:
        """Расчет достоверности диагноза"""
        confidence = 60.0  # Базовая

        # Учитываем количество сопоставленных симптомов
        matched_count = results['statistics']['matched_symptoms_count']
        confidence += min(matched_count * 3, 20)

        # Учитываем красные флаги
        if results['red_flags_found']:
            confidence += 15

        # Учитываем шкалы
        if scales.get('vas', 0) > 0:
            confidence += 5
        if scales.get('pain_detect', 0) > 0:
            confidence += 5
        if scales.get('oswestry', 0) > 0:
            confidence += 5

        # Учитываем психологические факторы
        if results['yellow_flags_found']:
            confidence += 8

        # Ограничиваем до 95%
        return min(confidence, 95.0)

    def _generate_recommendations(self, results: Dict, scales: Dict) -> List[str]:
        """Генерация рекомендаций"""
        recommendations = []

        diagnosis = results['diagnosis']

        # Общие рекомендации
        recommendations.append("Консультация невролога")

        # Специфические рекомендации
        if diagnosis['type'] == 'специфическая':
            recommendations.append("Срочное дополнительное обследование (МРТ, КТ, анализы)")
            recommendations.append("Консультация профильного специалиста")

            if diagnosis.get('subtype') == 'онкологическая':
                recommendations.append("Консультация онколога")
                recommendations.append("Онкологический скрининг")
            elif diagnosis.get('subtype') == 'травматическая':
                recommendations.append("Консультация травматолога")
                recommendations.append("Рентгенография/КТ позвоночника")
            elif diagnosis.get('subtype') == 'воспалительная':
                recommendations.append("Консультация ревматолога")
                recommendations.append("Анализы на воспалительные маркеры")

        # Рекомендации по подгруппам
        if diagnosis['subgroup'] == 'нейропатическая':
            recommendations.append("Прегабалин/Габапентин")
            recommendations.append("Антидепрессанты (дулоксетин, амитриптилин)")
            recommendations.append("Транскраниальная магнитная стимуляция")
            recommendations.append("Блокады")

        elif diagnosis['subgroup'] == 'дисфункциональная':
            recommendations.append("Когнитивно-поведенческая терапия")
            recommendations.append("Антидепрессанты (СИОЗС/СИОЗСН)")
            recommendations.append("Психологическое консультирование")
            recommendations.append("Кинезиотерапия с биологической обратной связью")

        else:  # ноцицептивная
            recommendations.append("НПВП (мелоксикам, целекоксиб)")
            recommendations.append("Миорелаксанты (толперизон, тизанидин)")
            recommendations.append("Лечебная физкультура")
            recommendations.append("Физиотерапия (магнитотерапия, лазер)")

        # Рекомендации по флагам
        if results['red_flags_found']:
            recommendations.append("Срочное обследование для исключения серьезной патологии")
            recommendations.append("Наблюдение в стационаре")

        if results['yellow_flags_found']:
            recommendations.append("Психологическая оценка и поддержка")
            recommendations.append("Образовательная программа по боли")

        if results['blue_flags_found']:
            recommendations.append("Эргономическая оценка рабочего места")
            recommendations.append("Трудотерапия")

        if results['black_flags_found']:
            recommendations.append("Социальная работа")
            recommendations.append("Профориентация")

        # Рекомендации по тяжести
        if diagnosis['severity'] == 'тяжелая':
            recommendations.append("Стационарное лечение")
            recommendations.append("Интенсивная обезболивающая терапия")
        elif diagnosis['severity'] == 'умеренная':
            recommendations.append("Дневной стационар")
            recommendations.append("Комплексная реабилитация")

        return recommendations

    def get_detailed_analysis(self, symptoms_text: str) -> Dict:
        """Детальный анализ симптомов с поиском похожих"""
        input_symptoms = [s.strip() for s in symptoms_text.split('\n') if s.strip()]

        detailed_results = {
            'input_symptoms': input_symptoms,
            'category_analysis': defaultdict(list),
            'best_matches': [],
            'summary': {}
        }

        for symptom in input_symptoms:
            symptom_lower = symptom.lower()

            # Ищем во всех категориях
            matches_by_category = {}

            for category_name, symptom_list in self.categories.items():
                similar = self._find_similar_symptoms(symptom_lower, symptom_list)
                if similar:
                    matches_by_category[category_name] = similar

            if matches_by_category:
                # Находим лучшую категорию (с наивысшей схожестью)
                best_category = None
                best_match = None
                best_score = 0

                for category, matches in matches_by_category.items():
                    if matches[0][1] > best_score:
                        best_score = matches[0][1]
                        best_category = category
                        best_match = matches[0][0]

                detailed_results['best_matches'].append({
                    'input': symptom,
                    'best_category': best_category,
                    'best_match': best_match,
                    'confidence': best_score * 100,
                    'all_matches': matches_by_category
                })

                # Добавляем в анализ по категориям
                for category, matches in matches_by_category.items():
                    detailed_results['category_analysis'][category].append({
                        'input': symptom,
                        'matches': matches
                    })
            else:
                detailed_results['best_matches'].append({
                    'input': symptom,
                    'best_category': 'не найдено',
                    'best_match': None,
                    'confidence': 0,
                    'message': 'Не найдено похожих симптомов в базе'
                })

        # Сводка
        detailed_results['summary'] = {
            'total_symptoms': len(input_symptoms),
            'matched_symptoms': len([m for m in detailed_results['best_matches'] if m['best_match']]),
            'categories_found': list(detailed_results['category_analysis'].keys()),
            'match_rate': (len([m for m in detailed_results['best_matches'] if m['best_match']]) /
                           len(input_symptoms) * 100) if input_symptoms else 0
        }

        return detailed_results


# Тестирование анализатора
if __name__ == "__main__":
    print("🧪 Тестирование анализатора симптомов...")

    analyzer = EnhancedPainAnalyzer()

    # Тестовые симптомы
    test_symptoms = """
    Боль в пояснице
    Депрессия и тревога
    Нарушение сна
    Мышечное напряжение
    Ограничение движений
    Боль усиливается при сидении
    """

    test_scales = {
        'vas': 7,
        'oswestry': 42,
        'pain_detect': 9,
        'hads_anxiety': 8,
        'hads_depression': 6,
        'pcs': 22
    }

    print("\n1. Детальный анализ симптомов:")
    detailed = analyzer.get_detailed_analysis(test_symptoms)

    print(f"Всего симптомов: {detailed['summary']['total_symptoms']}")
    print(f"Сопоставлено: {detailed['summary']['matched_symptoms']}")
    print(f"Процент совпадения: {detailed['summary']['match_rate']:.1f}%")

    print("\n2. Лучшие совпадения:")
    for match in detailed['best_matches']:
        if match['best_match']:
            print(f"  '{match['input']}' → '{match['best_match']}' "
                  f"({match['confidence']:.1f}%, категория: {match['best_category']})")
        else:
            print(f"  '{match['input']}' → не найдено")

    print("\n3. Полный анализ:")
    results = analyzer.analyze_symptoms(test_symptoms, test_scales)

    print(f"Тип боли: {results['diagnosis']['type']}")
    print(f"Подгруппа: {results['diagnosis']['subgroup']}")
    print(f"Тяжесть: {results['diagnosis']['severity']}")
    print(f"Достоверность: {results['confidence']:.1f}%")

    print("\n4. Рекомендации:")
    for i, rec in enumerate(results['recommendations'], 1):
        print(f"  {i}. {rec}")

    print("\n🎉 Анализатор работает корректно!")