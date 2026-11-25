"""Заполнение файла ПВД-УВД-ДВД-ГПА.xlsx"""


from datetime import datetime
from pathlib import Path
import re
import math

from openpyxl import load_workbook

from .reference_service import ReferenceService, ReferenceDataError



WORK_TYPES = {
    "Капитальный ремонт": "КР",
    "Средний ремонт": "СР",
    "Техническое обслуживание": "ТО",
    "Текущий ремонт": "ТР",
}



# Ячейка в ПВД-УВД-ДВД-ГПА.xlsx -> колонка в Тех-Места-ГПА.xlsx.
HEADER_CELL_MAP = {
    "D18": "B",
    "D19": "A",
    "D20": "L",
    "D21": "J",
    "D22": "Z",
    "C23": "M",
    "C24": "O",
    "E23": "AA",
    "E24": "AB",
    "D25": "C",
    "D26": "B",
    "D27": "AC",
    "D29": "AD",
    "F29": "AE",
    "C30": "AH",
    "C31": "AG",
}



class ExcelGenerationError(RuntimeError):
    """Срабатывает, когда выходной файл не может быть создан."""



class ExcelService:
    def __init__(self, template_path: Path, output_dir: Path, reference_service: ReferenceService):
        self.template_path = template_path
        self.output_dir = output_dir
        self.reference_service = reference_service

    def generate(self, work_type: str, station: str, equipment: str) -> Path:
        if work_type not in WORK_TYPES:
            raise ExcelGenerationError("Неизвестный вид работ.")
        if not station or not equipment:
            raise ExcelGenerationError("Необходимо выбрать КС и ГПА.")
        if not self.template_path.is_file():
            raise ExcelGenerationError(f"Файл-шаблон не найден: {self.template_path.name}")

        source_data = self.reference_service.find_equipment(station, equipment)

        try:
            workbook = load_workbook(self.template_path)
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet["D32"] = WORK_TYPES[work_type]

            for target_cell, source_column in HEADER_CELL_MAP.items():
                worksheet[target_cell] = None if source_column is None else source_data.get(source_column)

            value = worksheet["C31"].value
            result = value / 24000

            if work_type == "Капитальный ремонт":
                worksheet["D33"] = math.ceil(result)
            else:
                worksheet["D33"] = math.floor(result)

            self.output_dir.mkdir(parents=True, exist_ok=True)
            filename = self._build_filename(station, equipment, WORK_TYPES[work_type])
            output_path = self.output_dir / filename
            workbook.save(output_path)
            workbook.close()
        except ReferenceDataError:
            raise
        except Exception as exc:
            raise ExcelGenerationError("Не удалось сформировать итоговый Excel-файл.") from exc


        return output_path



    @staticmethod
    def _build_filename(station: str, equipment: str, work_type_code: str) -> str:
        safe_name = re.sub(r"[^\w\- ]+", "", f"{station} {equipment}", flags=re.UNICODE)
        safe_name = re.sub(r"\s+", "_", safe_name).strip("_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        return f"ВД_{work_type_code}_{safe_name}_{timestamp}.xlsx"
