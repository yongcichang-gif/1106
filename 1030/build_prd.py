from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from pathlib import Path

out=Path(r'C:\Users\ncucc\Desktop\1030\Daymark_PRD.docx')
doc=Document()
sec=doc.sections[0]
sec.top_margin=Inches(.68); sec.bottom_margin=Inches(.68); sec.left_margin=Inches(.78); sec.right_margin=Inches(.78)
styles=doc.styles
for sn in ['Normal','Title','Heading 1','Heading 2']:
 st=styles[sn]; st.font.name='Aptos'; st._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
styles['Normal'].font.size=Pt(9.5); styles['Normal'].font.color.rgb=RGBColor(39,52,45)
styles['Normal'].paragraph_format.space_after=Pt(4)
styles['Title'].font.size=Pt(28); styles['Title'].font.bold=True; styles['Title'].font.color.rgb=RGBColor(29,75,57)
styles['Title'].paragraph_format.space_after=Pt(2)
for sn,size in [('Heading 1',16),('Heading 2',11)]:
 st=styles[sn]; st.font.size=Pt(size); st.font.bold=True; st.font.color.rgb=RGBColor(35,107,83); st.paragraph_format.space_before=Pt(10); st.paragraph_format.space_after=Pt(4); st.paragraph_format.keep_with_next=True

def shade(cell,fill):
 tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)
def set_cell(cell,text,bold=False,color=None,size=8.5):
 cell.text=''; p=cell.paragraphs[0]; p.paragraph_format.space_after=Pt(1)
 r=p.add_run(text); r.bold=bold; r.font.name='Aptos'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei'); r.font.size=Pt(size)
 if color:r.font.color.rgb=RGBColor(*color)
 cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def table(headers, rows, widths=None):
 t=doc.add_table(rows=1, cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
 for i,h in enumerate(headers):set_cell(t.rows[0].cells[i],h,True,(255,255,255));shade(t.rows[0].cells[i],'236B53')
 for ri,row in enumerate(rows):
  cells=t.add_row().cells
  for i,val in enumerate(row):
   set_cell(cells[i],val,bold=(i==0),size=8.2)
   if ri%2==1:shade(cells[i],'F2F6F1')
 if widths:
  for row in t.rows:
   for i,w in enumerate(widths):row.cells[i].width=Inches(w)
 doc.add_paragraph().paragraph_format.space_after=Pt(1)
 return t

p=doc.add_paragraph(style='Title');p.add_run('Daymark 產品需求文件')
p=doc.add_paragraph();p.paragraph_format.space_after=Pt(9)
r=p.add_run('個人與團隊任務管理原型');r.bold=True;r.font.size=Pt(12);r.font.color.rgb=RGBColor(92,111,99)
# compact metadata table
m=doc.add_table(rows=1,cols=3);m.alignment=WD_TABLE_ALIGNMENT.LEFT
for i,(a,b) in enumerate([('文件版本','1.0'),('更新日期','2026 年 10 月 30 日'),('產品狀態','互動式前端原型')]):
 c=m.rows[0].cells[i];shade(c,'EDF3EC');c.text='';p=c.paragraphs[0];p.paragraph_format.space_after=Pt(1);rr=p.add_run(a+'\n');rr.bold=True;rr.font.size=Pt(8);rr.font.color.rgb=RGBColor(35,107,83);rr2=p.add_run(b);rr2.font.size=Pt(9)

def H(text,level=1):doc.add_heading(text,level=level)
def P(text):doc.add_paragraph(text)
def bullets(items):
 for x in items:
  p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(2);p.add_run(x)
H('1. 產品概述')
P('Daymark 是面向個人與小型團隊的任務管理介面，將優先事項、截止日、負責人與通知選項集中呈現。產品目標是協助使用者快速辨認下一步、追蹤團隊工作狀態，並以輕量操作減少管理成本。')
H('2. 問題與目標')
P('任務分散時，使用者難以判斷先做什麼、哪些工作已逾期，以及團隊任務由誰負責。Daymark 以排序後的看板與日期提示降低這些資訊成本。')
table(['目標','原型中的對應方式'],[
('快速掌握工作量','顯示未完成、逾期與本週完成任務摘要。'),('辨識優先順序','依逾期狀態、截止日及重要性排序任務。'),('追蹤執行狀態','看板分為待處理、進行中與已完成。'),('管理協作資訊','顯示團隊成員及任務負責人，提供成員修改與移除操作。')],[1.65,5.7])
H('3. 目標使用者')
bullets(['個人工作者：需要整理每日待辦、截止日與優先級。','小型團隊成員或管理者：需要檢視任務狀態、負責人及基本成員資訊。'])
H('4. 核心使用情境')
bullets(['使用者進入優先任務頁面，掃描逾期及近期到期事項。','使用者新增任務，填入名稱、說明、截止日、重要性、負責人、進度與提醒偏好。','使用者將任務標記為完成，或以篩選檢視未完成與高重要性任務。','使用者查看今日行程、團隊成員或通知設定。','使用者在團隊成員頁修改名稱，或移除成員並確認相關影響。'])
H('5. 功能需求')
table(['編號','功能','需求說明','優先級'],[
('F-01','任務看板','以待處理、進行中、已完成三欄呈現本人任務及筆數。','P0'),('F-02','任務排序與狀態','依逾期、截止日、重要性排序；可切換完成狀態。','P0'),('F-03','新增任務','輸入名稱、說明、截止日期、重要性、負責人、進度及電子郵件提醒偏好。','P0'),('F-04','摘要指標','顯示待完成數、逾期數及本週完成摘要；逾期任務以醒目樣式標示。','P1'),('F-05','篩選與搜尋','支援全部、未完成、高重要性篩選及任務名稱搜尋。','P1'),('F-06','行程與行事曆','提供今日行程摘要與行事曆入口。現版為示意資料。','P1'),('F-07','團隊成員管理','列出團隊成員；支援修改名稱、移除成員及移除確認。修改後任務負責人名稱同步更新。','P1'),('F-08','通知設定','呈現預設提醒選項及啟用狀態；未連結郵件服務時需明確提示。','P1'),('F-09','專注時段','提供 25 分鐘專注時段入口；目前僅顯示開始提示，沒有計時或通知控制。','P2')],[.55,1.15,4.8,.65])
H('6. 主要頁面與資訊架構')
table(['頁面','主要內容'],[
('優先任務','摘要指標、任務篩選/排序、三欄看板、優先建議。'),('今日行程','當日時間安排與到期任務摘要。'),('行事曆','行事曆入口；目前呈現今日示意內容。'),('我的任務','本人任務清單/看板。'),('團隊成員','成員卡片、修改/移除操作及原型資料提示。'),('通知設定','電子郵件提醒說明與啟用選項。')],[1.4,5.95])
H('7. 資料欄位與規則')
bullets(['任務：識別碼、名稱、補充說明、截止日期、重要性（高/中/低）、狀態（待處理/進行中/已完成）、負責人、提醒偏好。','成員：識別碼、姓名、顯示字元與成員角色標籤。','任務日期以 YYYY-MM-DD 儲存；畫面提供「今天到期」、「明天到期」、「逾期 N 天」等相對提示。','建立任務時名稱與截止日期為必填；說明可留空。','移除成員的原型流程會先確認；若有指派任務，確認文案需揭露目前原型會一併刪除該成員的示範任務。正式產品應另行定義任務重新指派策略。'])
H('8. 非功能需求')
bullets(['介面採繁體中文，適配桌面與窄螢幕；窄螢幕的主要看板與成員卡片改為單欄呈現。','互動控制需提供明確名稱、可見狀態與取消/確認方式。','資料與操作需具備持久化、權限及稽核策略（正式版待定）；原型刷新後目前不保留變更。'])
H('9. 驗收條件')
for x in ['看板會依任務狀態分欄，並顯示每欄筆數。','新任務在欄位有效時建立，且在看板及摘要統計中反映。','篩選與搜尋可回饋符合條件的任務；完成狀態切換後看板更新。','團隊成員頁可修改成員名稱，並同步反映於相關任務卡片。','移除成員前要求確認；對有指派任務的成員清楚告知資料影響。','桌面及窄螢幕下頁面不應出現重疊、溢出或操作按鈕遭遮蔽。','尚未實作的郵件提醒、專注倒數、持久化與行事曆整合均需標示為原型或待建置。']:
 p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(2);p.add_run(x)
H('10. 範圍外與限制')
P('目前版本是單頁前端原型，資料為程式內示範資料。尚未連接後端、登入與權限、資料庫、郵件寄送、即時協作、真正的專注倒數、外部行事曆同步或跨裝置保存。刷新頁面會還原操作結果；任務新增和成員修改/移除僅在目前頁面工作階段生效。')
H('11. 待決事項')
bullets(['正式版的帳號、團隊邀請與角色權限模型。','移除成員時，指派任務應重新指派、保留為未指派，還是封存。','提醒服務、使用者時區、通知頻率與退訂流程。','本週完成數的統計週期與指標口徑。','專注時段是否需要可調長度、暫停/結束控制及通知靜音。','行事曆視圖、任務明細與多人協作的正式範圍。'])
# footer page number
for section in doc.sections:
 f=section.footer.paragraphs[0];f.alignment=WD_ALIGN_PARAGRAPH.RIGHT
 rr=f.add_run('Daymark PRD  |  ');rr.font.size=Pt(8);rr.font.color.rgb=RGBColor(120,136,126)
 fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');f._p.append(fld)
doc.core_properties.title='Daymark 產品需求文件'
doc.core_properties.subject='個人與團隊任務管理原型產品需求'
doc.core_properties.author='Daymark'
doc.save(out)
print(out)
