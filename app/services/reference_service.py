"""Работа со справочником Тех-Места-ГПА.xlsx"""

from pathlib import Path
from typing import Any



class ReferenceDataError(RuntimeError):
    """Raised when the equipment reference cannot be read or is invalid."""



class ReferenceService:
    """Loads the technical-place workbook using pandas."""

    # Excel column letters are kept explicit to match the agreed workbook mapping.
    KS_COLUMN = "E"
    GPA_COLUMN = "G"

    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path
        self._cached_table = None
        self._cached_mtime = None

    def _read_table(self):
        import pandas as pd

        if not self.workbook_path.is_file():
            raise ReferenceDataError(f"Файл справочника не найден: {self.workbook_path.name}")

        current_mtime = self.workbook_path.stat().st_mtime
        if self._cached_table is not None and self._cached_mtime == current_mtime:
            return self._cached_table

        try:
            # Header is on Excel row 2. Read through AH because the agreed map
            # includes columns AG and AH, even if they are currently empty.
            table = pd.read_excel(
                self.workbook_path,
                sheet_name=0,
                header=1,
                usecols="A:AJ",
                engine="openpyxl",
            )
        except Exception as exc:
            raise ReferenceDataError("Не удалось прочитать файл техмест.") from exc


        if len(table.columns) < 36:
            # Pandas omits trailing empty columns in some workbook versions.
            for index in range(len(table.columns), 36):
                table[f"__empty_{index}"] = None

        self._cached_table = table
        self._cached_mtime = current_mtime
        return table


    @staticmethod
    def _clean_text(value: Any) -> str:
        if value is None:
            return ""
        try:
            import pandas as pd
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass
        return str(value).strip()



    def get_stations(self) -> list[str]:
        table = self._read_table()
        # Excel E is index 4 within the A:AH range.
        values = table.iloc[:, 4].map(self._clean_text)
        return sorted({value for value in values if value}, key=str.casefold)



    def get_equipment(self, station: str) -> list[str]:
        table = self._read_table()
        station_values = table.iloc[:, 4].map(self._clean_text)
        equipment_values = table.iloc[:, 6].map(self._clean_text)
        selected = equipment_values[station_values == station]
        return sorted({value for value in selected if value}, key=str.casefold)



    def find_equipment(self, station: str, equipment: str) -> dict[str, Any]:
        table = self._read_table()
        station_values = table.iloc[:, 4].map(self._clean_text)
        equipment_values = table.iloc[:, 6].map(self._clean_text)
        matches = table[(station_values == station) & (equipment_values == equipment)]

        if matches.empty:
            raise ReferenceDataError("Выбранный ГПА не найден для указанной КС.")
        if len(matches) > 1:
            raise ReferenceDataError("В справочнике найдено несколько строк для выбранного ГПА. Требуется уточнить запись.")

        row = matches.iloc[0]
        # Return values keyed by Excel column letters, matching the header map.
        result: dict[str, Any] = {}
        for letter, index in self._excel_column_indexes().items():
            if index >= len(row):
                result[letter] = None
            else:
                result[letter] = self._normalize_value(row.iloc[index])
        return result



    @staticmethod
    def _excel_column_indexes() -> dict[str, int]:
        # Index is relative to the A:AH dataframe range (A=0, B=1, ... AH=33).
        return {
            "A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6,
            "H": 7, "I": 8, "J": 9, "K": 10, "L": 11, "M": 12, "N": 13,
            "O": 14, "P": 15, "Q": 16, "R": 17, "S": 18, "T": 19,
            "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24, "Z": 25,
            "AA": 26, "AB": 27, "AC": 28, "AD": 29, "AE": 30, "AF": 31,
            "AG": 32, "AH": 33, "AI": 34, "AJ": 35,
        }



    @staticmethod
    def _normalize_value(value: Any) -> Any:
        import pandas as pd

        if value is None or pd.isna(value):
            return None
        # pandas Timestamp is compatible with openpyxl's datetime values.
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
        return value
