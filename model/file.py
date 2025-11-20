import pandas as pd
import openpyxl

#=================== File Control Function ===================#
def open_file(file_name):
    """
    Excel 파일을 열고 데이터 리스트로 반환
    """
    print("File Open")
    file_name = file_name.replace('\\', '\\\\')
    re_list = pd.read_excel(file_name, engine='openpyxl')
    if 'Row' in re_list.columns:
        re_list = re_list.drop('Row', axis=1)
    return re_list.values.tolist()

def load_full_excel(file_name):
    """
    통합 저장된 엑셀 파일에서 Step과 EIS 데이터를 분리하여 읽어오기.
    반환값:
        step_data_list, eis_data_list
    """
    wb = openpyxl.load_workbook(file_name, data_only=True)
    ws = wb.active

    step_data_list = []
    eis_data_list = []

    # --- 시트의 모든 행을 순회하면서 Step / EIS 영역 구분 ---
    for row in ws.iter_rows(values_only=True):
        # Step 구역: 첫 12열 (0~11)
        step_part = row[:12]
        # EIS 구역: 13열 이후 (12~)
        eis_part = row[13:19] if len(row) > 13 else None

        # Step 데이터 유효성 검사
        if any(cell is not None for cell in step_part):
            step_data_list.append(list(step_part))

        # EIS 데이터 유효성 검사
        if eis_part and any(cell is not None for cell in eis_part):
            eis_data_list.append(list(eis_part))

    return step_data_list, eis_data_list

def save_file(file_name, *args):
    """
    일반 테스트 데이터 저장
    """
    try:
        col_names = ["Step",
                    "Type",
                    "Mode",
                    "Mode value",
                    "End Type",
                    "Operator",
                    "End Value",
                    "Go to",
                    "Report Type",
                    "Report value",
                    "Note",
                    "Row"
                    ]
        col_names.append(len(args[0])+1)
        save_excel = pd.DataFrame(args[0], columns=col_names)
        save_excel.to_excel(file_name, index=False, engine='openpyxl')
        print(f"파일 저장 완료: {file_name}")

    except FileNotFoundError:
        print("파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"저장 중 오류 발생: {e}")


def save_eis(file_name, *args):
    """
    EIS 데이터 저장
    """
    try:
        col_names = [#"Step",
                     "Mode",
                     "Start_Frequency",
                     "Stop_Frequency",
                     "Amplitude",
                     "PointNumber"
        ]
        save_excel = pd.DataFrame(args[0], columns=col_names)
        save_excel.to_excel(file_name, index=False, engine='openpyxl')
        print(f"EIS 파일 저장 완료: {file_name}")

    except FileNotFoundError:
        print("❌ 파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"⚠️ 저장 중 오류 발생: {e}")

# ============================================================
# Step + EIS 통합 저장 (하나의 시트 내 구역 분리)
# ============================================================
def save_full_excel(file_name, step_data, eis_data):
    """
    하나의 엑셀 시트에 일반 Step 데이터와 EIS 데이터를 나란히 저장.
    step_data: 일반 Step 2D 리스트
    eis_data:  EIS Parameter 2D 리스트
    """
    try:
        # --- Step 데이터 구성 ---
        step_cols = [
            "Step", "Type", "Mode", "Mode value", "End Type", "Operator",
            "End Value", "Go to", "Report Type", "Report value", "Note", "Rows"
        ]
        step_cols.append(len(step_data)+1)
        df_steps = pd.DataFrame(step_data, columns=step_cols)

        # --- EIS 데이터 구성 ---
        eis_cols = ["Step",
                     "Mode",
                     "Start_Frequency",
                     "Stop_Frequency",
                     "Amplitude",
                     "PointNumber",
                     "EIS_Rows"
                     ]
        eis_cols.append(len(eis_data))
        df_eis = pd.DataFrame(eis_data, columns=eis_cols)

        # --- 엑셀 작성 ---
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_steps.to_excel(writer, index=False, startrow=0, startcol=0)
            df_eis.to_excel(writer, index=False, startrow=0, startcol=14)

        print(f"✅ Step + EIS 통합 엑셀 저장 완료: {file_name}")

    except FileNotFoundError:
        print("❌ 파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"⚠️ 통합 저장 중 오류 발생: {e}")