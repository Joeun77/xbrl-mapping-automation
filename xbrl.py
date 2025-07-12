import pandas as pd
import re

# 삼성전자 MAP 시트 불러오기
map_file = 'C:\\Users\\projo\\xbrl_auto\\삼성전자_사업보고서_XBRL_원문_MAP.xlsx'
map_df = pd.read_excel(map_file, sheet_name='MAP')
def classify_report(def_ko):
    if pd.isna(def_ko):
        return '기타'
    if '재무상태표' in def_ko:
        return '재무상태표'
    elif '포괄손익계산서' in def_ko or '손익계산서' in def_ko:
        return '포괄손익계산서'
    elif '현금흐름표' in def_ko:
        return '현금흐름표'
    elif '자본변동표' in def_ko:
        return '자본변동표'
    elif '주석' in def_ko or 'Notes' in def_ko:
        return '주석'
    else:
        return '기타'

map_df['Report_Type'] = map_df['DEFINITION_KO'].apply(classify_report)

def extract_note_number(def_ko):
    if pd.isna(def_ko):
        return None
    match = re.search(r'\[(\d+)\]', def_ko)
    return match.group(1) if match else None

map_df['Note_Number'] = map_df.apply(
    lambda x: extract_note_number(x['DEFINITION_KO']) if x['Report_Type'] == '주석' else None,
    axis=1
)
output_rows = []

for definition, group in map_df.groupby('DEFINITION_KO'):
    # Definition 행 추가
    output_rows.append({
        'prefix': 'Definition',
        'name': definition,
        'depth': '',
        'preferredLabel': '',
        'ko_label': '',
        'en_label': '',
        '행 데이터 타입': '',
        '차변/대변': '',
        '기간속성': '',
        'Definition': definition
    })

    # 그룹 내 데이터 추가
    depth = 0
    for _, row in group.iterrows():
        depth_value = 0 if str(row.get('NAME', '')).endswith('Abstract') or row.get('ABSTRACT') is True else depth + 1

        output_rows.append({
            'prefix': row.get('PREFIX', ''),
            'name': row.get('NAME', ''),
            'depth': depth_value,
            'preferredLabel': row.get('PREFERREDLABEL1', ''),
            'ko_label': row.get('LABEL_KO', ''),
            'en_label': row.get('LABEL_EN', ''),
            '행 데이터 타입': row.get('DATA_TYPE', ''),
            '차변/대변': row.get('BALANCE', ''),
            '기간속성': row.get('PERIOD_TYPE', ''),
            'Definition': definition
        })

        if depth_value != 0:
            depth = depth_value
            # DataFrame 변환
result_df = pd.DataFrame(output_rows)

# Excel 저장
result_df.to_excel('삼성전자_XBRL_맵핑결과.xlsx', index=False)