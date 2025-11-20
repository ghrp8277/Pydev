import pandas as pd
from pathlib import Path

class ExcelAdapter:
    def __init__(self, base_path: str = None):
        self.base_path = Path(base_path) if base_path else None

    def read(self, file_path: str, sheet_name: str = None):
        """
        엑셀 파일을 읽어 pandas DataFrame으로 반환한다.
        sheet_name이 None이면 모든 시트를 dict로 반환.
        """
        full_path = self._resolve_path(file_path)
        print(f"[ExcelAdapter] Reading from {full_path}, sheet='{sheet_name}'")

        if sheet_name:
            df = pd.read_excel(full_path, sheet_name=sheet_name)
            return df
        else:
            sheets = pd.read_excel(full_path, sheet_name=None)
            return sheets

    def write(self, file_path: str, dataframes: dict, mode: str = "w"):
        """
        여러 시트를 동시에 저장한다.
        dataframes: {sheet_name: DataFrame}
        mode: 'w' (덮어쓰기), 'a' (추가)
        """
        full_path = self._resolve_path(file_path)
        print(f"[ExcelAdapter] Writing to {full_path}, sheets={list(dataframes.keys())}")

        with pd.ExcelWriter(full_path, engine="openpyxl", mode=mode) as writer:
            for sheet, df in dataframes.items():
                df.to_excel(writer, sheet_name=sheet, index=False)

    def _resolve_path(self, path: str) -> Path:
        """base_path가 설정되어 있으면 결합해서 반환"""
        if self.base_path:
            return self.base_path / path
        return Path(path)

    def exists(self, file_path: str) -> bool:
        """파일 존재 여부 확인"""
        return self._resolve_path(file_path).exists()
