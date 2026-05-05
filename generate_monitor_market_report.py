import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "显示器市场分析报告.xlsx"

C = {"title_bg":"1A2332","title_fg":"FFFFFF","header_bg":"2C3E6B","header_fg":"FFFFFF","sub_bg":"D6E4F0","deep":"1A2332","light_gray":"F2F2F2","white":"FFFFFF"}

thin_border = Border(left=Side(style="thin",color="B0BEC5"),right=Side(style="thin",color="B0BEC5"),top=Side(style="thin",color="B0BEC5"),bottom=Side(style="thin",color="B0BEC5"))

def _f(size=10,bold=False,color="000000"): return Font(name="微软雅黑",size=size,bold=bold,color=color)
def _fl(h): return PatternFill(start_color=h,end_color=h,fill_type="solid")
def stripe(r): return _fl(C["light_gray"]) if r%2==0 else _fl(C["white"])
def setw(ws,d):
    for k,v in d.items(): ws.column_dimensions[k].width=v

def title_row(ws,row,text,end,h=48):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=end)
    c=ws.cell(row=row,column=1,value=text); c.font=_f(18,True,C["title_fg"]); c.fill=_fl(C["title_bg"])
    c.alignment=Alignment(horizontal="center",vertical="center"); ws.row_dimensions[row].height=h

def sub_row(ws,row,text,end,h=28):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=end)
    c=ws.cell(row=row,column=1,value=text); c.font=_f(12,True,C["deep"]); c.fill=_fl(C["sub_bg"])
    c.alignment=Alignment(horizontal="left",vertical="center",indent=1); ws.row_dimensions[row].height=h

def hdr_row(ws,row,hdrs,h=30):
    for i,ht in enumerate(hdrs,1):
        c=ws.cell(row=row,column=i,value=ht); c.font=_f(11,True,C["header_fg"]); c.fill=_fl(C["header_bg"])
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=thin_border
    ws.row_dimensions[row].height=h

def data_row(ws,row,vals,fmts=None,h=24,bold_cols=None):
    for i,v in enumerate(vals,1):
        c=ws.cell(row=row,column=i,value=v); c.font=_f(10,bold=(bold_cols and i in bold_cols))
        c.fill=stripe(row); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border=thin_border
        if fmts and i<=len(fmts) and fmts[i-1]: c.number_format=fmts[i-1]
    ws.row_dimensions[row].height=h

def add_pie(ws,title,rows,cat_c,val_c,anchor,w=14,h=10):
    ch=PieChart(); ch.title=title; ch.style=10; ch.width=w; ch.height=h
    d=Reference(ws,min_col=val_c,min_row=rows[0]-1,max_row=rows[-1])
    ct=Reference(ws,min_col=cat_c,min_row=rows[0],max_row=rows[-1])
    ch.add_data(d,titles_from_data=True); ch.set_categories(ct)
    ch.dataLabels=DataLabelList(); ch.dataLabels.showPercent=True; ch.dataLabels.showCatName=True
    ws.add_chart(ch,anchor)

def add_bar(ws,title,rows,cat_c,val_cs,anchor,w=16,h=10):
    ch=BarChart(); ch.type="col"; ch.title=title; ch.style=10; ch.width=w; ch.height=h
    ct=Reference(ws,min_col=cat_c,min_row=rows[0],max_row=rows[-1])
    for vc in val_cs:
        d=Reference(ws,min_col=vc,min_row=rows[0]-1,max_row=rows[-1])
        ch.add_data(d,titles_from_data=True)
    ch.set_categories(ct); ch.shape=4; ws.add_chart(ch,anchor)

# Sheet builders are in the full local file
# Run: python generate_monitor_market_report.py

def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    print("Generating monitor market report...")
    wb.save(str(OUTPUT_FILE))
    print(f"[OK] Report saved: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
