from model_data import *
import win32com.client as win32
import pandas as pd

outputFileName = r'd:\output.xlsx'

# input DB
dc = DataControl('SERVER')

curs = dc.conn.cursor()
sSql = f"select 현장코드, 분석월, 현장명p, 업체명, 출역일수, 소장출역, 직원출역, 확정일수 " \
       f"   from mdc_result ;"
curs.execute(sSql)
data = curs.fetchall()

columns = ['현장코드', '분석월', '현장명p', '업체명', '출역일수', '소장출역', '직원출역', '확정일수']
df = pd.DataFrame(data, columns=columns)
df['구분'] = df['업체명'].apply(lambda x: '당사' if '현대건설' in x else '협력업체')
df['대비'] = df.apply(lambda x: round(x.확정일수 / x.출역일수,2) if x.출역일수 and x.확정일수 else 0, axis=1)
df['비고'] = df['확정일수'].apply(lambda x: "◎ 퇴직공제부금 미등록" if x == 0 else "")
df = df.sort_values(['현장코드','분석월', '구분', '업체명'], ascending=[True, False, True, True], kind='mergesort')
df = df[['현장코드', '분석월', '구분', '현장명p', '업체명', '출역일수', '소장출역', '직원출역', '확정일수', '대비', '비고']]

df = df.astype({
       "출역일수": int, "소장출역": int, "직원출역": int, "확정일수": float, "대비": float,
})

with pd.ExcelWriter(outputFileName) as writer:
       if not df.empty:
              df.to_excel(writer, sheet_name="비교표")

os.startfile(outputFileName)