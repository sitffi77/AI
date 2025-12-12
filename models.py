from dataclasses import dataclass
from datetime import datetime
from typing import List, Dict, Optional

@dataclass
class Symptom:
    id: int
    name: str
    category: str  # 'specific', 'nonspecific', 'red_flag', etc.
    description: str

@dataclass
class Diagnosis:
    id: int
    patient_id: int
    diagnosis_type: str  # 'specific', 'nonspecific'
    subgroup: str        # 'nociceptive', 'neuropathic', 'dysfunctional'
    subtype: str         # 'oncological', 'traumatic', 'inflammatory', etc.
    confidence: float    # 0-100%
    recommended_treatment: str  # 'conservative', 'surgical'
    date: datetime

@dataclass
class Patient:
    id: int
    name: str
    age: int
    gender: str
    symptoms: List[Symptom]
    medical_history: Dict
    scales: Dict  # ВАШ, Освестри, PainDetect, HADS, PCS
    diagnosis: Optional[Diagnosis] = None