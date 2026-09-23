"""
QC-F-14 v6 — نرم افزار کامل ثبت و گزارش دهی قطعات برگشتی خط
"""
import os, sys, datetime, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
sys.path.insert(0, os.path.dirname(__file__))
import xlcal
from gen_workbook_v4 import SH, FACT_COLS, FACT_HDR_FA, FACT_W, L, build_data, build_master, build_helper, build_reports, build_dashboard, build_dq, build_settings, build_help, build_log, build_users, build_verify, F_TITLE, F_LBL, F_TXT, F_NOTE, F_HDR, FILL_HDR, CTRD, BORDER, WRAP, C_HEADER

def build_login_v6(ws, pl):
    ws.sheet_view.showGridLines=False
    ws.sheet_view.showRowColHeaders=False
    ws.column_dimensions['A'].width=2
    ws.column_dimensions['B'].width=24
    ws.column_dimensions['C'].width=32
    ws.column_dimensions['D'].width=45
    ws['B2']='QC-F-14 v6 — نرم افزار ثبت قطعات برگشتی خط'
    ws['B2'].font=Font(bold=True, size=20, name='Tahoma', color=C_HEADER)
    ws['B3']='نسخه نهایی — لاگین UserForm + فرم ثبت برگشتی + داشبورد + گزارشات'
    ws['B3'].font=Font(bold=True, size=12, name='Tahoma', color='FF7030A0')
    ws['B4']='این نرم افزار برای ثبت و گزارش دهی قطعات برگشتی خط طراحی شده. دارای کنترل دسترسی نقش محور، لاگ، و گزارشات مدیریتی است.'
    ws['B4'].font=F_NOTE
    ws.merge_cells('B4:D5')
    ws['B4'].alignment=WRAP

    ws['B7']='نام کاربری'
    ws['B7'].font=F_LBL
    ws['C7'].fill=PatternFill('solid', fgColor='FFFFFFFF')
    ws['C7'].border=BORDER
    ws['C7'].number_format='@'
    ws['C7'].value='admin'
    ws['B9']='رمز عبور'
    ws['B9'].font=F_LBL
    ws['C9'].fill=PatternFill('solid', fgColor='FFFFFFFF')
    ws['C9'].border=BORDER
    ws['C9'].number_format='@'
    ws['C9'].value='admin123'

    ws['B11']='کاربر فعلی'
    ws['B11'].font=F_LBL
    ws['C11'].border=BORDER
    ws['C11'].value='(وارد نشده)'
    ws['B12']='نقش'
    ws['B12'].font=F_LBL
    ws['C12'].border=BORDER
    ws['B13']='وضعیت'
    ws['B13'].font=F_LBL
    ws['C13'].border=BORDER
    ws['C13'].value='لطفاً از طریق فرم لاگین وارد شوید (مانند نمونه)'

    ws['B15']='دسترسی ها'
    ws['B15'].font=F_TITLE
    ws['B16']='• ADMIN: admin / admin123 → همه شیت ها، مدیریت کاربران، تایید نهایی'
    ws['B17']='• OPERATOR: operator / 1234 → ثبت برگشتی، مشاهده داده، گزارش'
    ws['B18']='• VIEWER: viewer / viewer123 → فقط داشبورد و گزارشات'
    ws['B19']='• بازرسان: i-01..i-10 / 1234 → ثبت برگشتی'
    ws['B20']='• لاگین با UserForm (مانند نمونه) — در Workbook_Open نمایش داده می شود'
    for r in range(16,21):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=2).font=F_TXT
        ws.cell(row=r, column=2).alignment=WRAP

    ws['B22']='راهنمای استفاده'
    ws['B22'].font=F_TITLE
    ws['B23']='1- فایل را باز کنید → فرم لاگین ظاهر می شود\n2- وارد شوید → شیت های مجاز باز می شود\n3- به FORM بروید → اطلاعات قطعه برگشتی را وارد کنید → QC_Append\n4- در Data داده، سوابق را ببینید\n5- در Dashboard و Reports گزارشات را ببینید\n6- برای خروج QC_Logout'
    ws.merge_cells('B23:D27')
    ws['B23'].font=F_TXT
    ws['B23'].alignment=WRAP

    # Hidden validation for sample compatibility
    ws['K3']='admin'
    ws['L3']='admin123'
    ws['M3']='=K3&" - "&L3'
    ws['J3']='=H3&" - "&I3'
    ws['H3']='admin'
    ws['I3']='admin123'
    ws['N3']='=IF(ISNUMBER(MATCH(M3,J:J,0)),"YES","NO")'

    ws.sheet_properties.tabColor='FF7030A0'

def build_return_form(ws, pl, lists, helper_name):
    ws.sheet_view.showGridLines=False
    ws.column_dimensions['A'].width=2
    ws.column_dimensions['B'].width=20
    ws.column_dimensions['C'].width=28
    ws.column_dimensions['D'].width=6
    ws.column_dimensions['E'].width=20
    ws.column_dimensions['F'].width=28
    ws.column_dimensions['G'].width=15
    ws['B2']='فرم ثبت قطعات برگشتی خط — نسخه ۶ نهایی'
    ws['B2'].font=Font(bold=True, size=16, name='Tahoma', color=C_HEADER)
    ws['B3']=f"مهاجرت: {pl['stats']['fact_rows']} رکورد | برگشتی: {pl['stats']['record_types'].get('RETURN',0)} | Disposition: 5 مقدار"
    ws['B3'].font=F_NOTE

    fields=[
        ('تاریخ برگشت (شمسی)*','C',6,'date_list','1405/06/28'),
        ('شیفت*','C',8,'shift_list','S1/S2/S3'),
        ('خط / ایستگاه*','C',10,'station_list','ایستگاه برگشت'),
        ('بازرس*','C',12,'inspector_list','بازرس'),
        ('کد قطعه*','F',6,'part_list','قطعه برگشتی'),
        ('کد قالب*','F',8,'mold_list','قالب'),
        ('بارکد (۱۹ رقم)*','C',14,None,'19 رقم'),
        ('تعداد برگشتی*','F',10,None,'عدد'),
        ('کد عیب*','F',12,'defect_list','دلیل برگشت'),
        ('تعیین تکلیف*','F',14,'disposition_new_list','ACCEPT/REWORK/SCRAP/RETURN/HOLD'),
        ('وضعیت فعلی','C',16,'status_new_list','OPEN/CLOSED'),
        ('توضیح','F',16,None,'دلیل برگشت'),
    ]
    cells={}
    for label,col,row,lst,hint in fields:
        lab_col='B' if col=='C' else 'E'
        ws[f'{lab_col}{row}']=label
        ws[f'{lab_col}{row}'].font=F_LBL
        ws[f'{lab_col}{row}'].alignment=Alignment(horizontal='right', vertical='center')
        cell=ws[f'{col}{row}']
        cell.number_format='@'
        cell.fill=PatternFill('solid', fgColor='FFFFFFFF')
        cell.border=BORDER
        cell.alignment=CTRD
        cells[label]=f'{col}{row}'
        if lst and lst in lists:
            # DV
            from openpyxl.worksheet.datavalidation import DataValidation
            col_idx, n = lists[lst]
            ref=f"'{helper_name}'!${get_column_letter(col_idx)}$2:${get_column_letter(col_idx)}${max(n+1,2)}"
            dv=DataValidation(type='list', formula1=f'={ref}', allow_blank=True, showErrorMessage=True)
            dv.error=f'{hint} را از لیست انتخاب کنید'
            ws.add_data_validation(dv)
            dv.add(f'{col}{row}')

    # Barcode validation
    from openpyxl.worksheet.datavalidation import DataValidation
    dv=DataValidation(type='textLength', operator='equal', formula1='19', allow_blank=True, showErrorMessage=True)
    dv.error='بارکد باید 19 رقم باشد'
    ws.add_data_validation(dv)
    dv.add('C14')

    # Box
    for rr in range(6,17):
        for cc in range(2,7):
            ws.cell(row=rr, column=cc).border=BORDER

    # Validation message
    ws['B18']='پیام اعتبارسنجی'
    ws['B18'].font=F_LBL
    ws['C18']=f'=IF(COUNTA(C6,C8,C10,C12,F6,F8,C14,F10,F12,F14)=0,"لطفا فیلدها را پر کنید",IF(LEN(C14)<>19,"بارکد 19 رقم نیست","آماده ثبت — QC_Append"))'
    ws['C18'].font=Font(bold=True, size=11, name='Tahoma')
    ws.merge_cells('C18:F18')

    ws['B20']='ماکروها: QC_Append (ثبت) / QC_ClearForm (پاک کردن) / QC_VoidRow (ابطال) / QC_Logout (خروج)'
    ws['B20'].font=F_NOTE
    ws.merge_cells('B20:F20')

    # Recent RETURN records
    ws['B22']='۱۰ قطعه برگشتی اخیر'
    ws['B22'].font=Font(bold=True, size=14, name='Tahoma', color=C_HEADER)
    # Headers
    hdr=['شناسه','تاریخ','قطعه','عیب','تعیین تکلیف','بارکد','کاربر']
    for j,h in enumerate(hdr, start=2):
        c=ws.cell(row=23, column=j, value=h)
        c.font=F_HDR
        c.fill=PatternFill('solid', fgColor=C_HEADER)
        c.alignment=CTRD
        c.border=BORDER

    # Last 10 RETURN
    # Find last RETURN rows from facts
    facts=pl['facts']
    return_facts=[f for f in facts if f['RECORD_TYPE']=='RETURN'][-10:]
    return_facts=return_facts[::-1]
    for i,f in enumerate(return_facts, start=24):
        ws.cell(row=i, column=2, value=f['RECORD_ID']).number_format='@'
        ws.cell(row=i, column=3, value=f['EVENT_DATE_J']).number_format='@'
        ws.cell(row=i, column=4, value=f['PART_NAME_SNAP']).number_format='@'
        ws.cell(row=i, column=5, value=f['DEFECT_NAME_SNAP']).number_format='@'
        ws.cell(row=i, column=6, value=f['DISPOSITION']).number_format='@'
        ws.cell(row=i, column=7, value=f['BARCODE']).number_format='@'
        ws.cell(row=i, column=8, value=f['ENTRY_USER']).number_format='@'

    ws.sheet_properties.tabColor='FFC00000'

def build_return_data(ws, pl):
    # Filtered view for RETURN only
    ws.sheet_view.showGridLines=False
    ws['B2']='داده قطعات برگشتی — فقط RECORD_TYPE=RETURN'
    ws['B2'].font=Font(bold=True, size=14, name='Tahoma', color='FFC00000')
    # Headers
    hdr=FACT_HDR_FA
    for j,h in enumerate(hdr, start=1):
        c=ws.cell(row=4, column=j, value=h)
        c.font=F_HDR
        c.fill=PatternFill('solid', fgColor='FFC00000')
        c.alignment=CTRD
        c.border=BORDER
    # Filter RETURN facts
    facts=[f for f in pl['facts'] if f['RECORD_TYPE']=='RETURN']
    for r_idx,f in enumerate(facts, start=5):
        row=[
            f['RECORD_ID'], f['RECORD_TYPE'], f['EVENT_DATE_J'],
            f['EVENT_DATE_G'], f['J_MONTH_KEY'], f['J_WEEK_KEY'],
            f['SHIFT_ID'], f['STATION_ID'], f['STATION_SNAP'],
            f['INSPECTOR_ID'], f['INSPECTOR_SNAP'],
            f['PART_ID'], f['PART_NAME_SNAP'], f['MOLD_ID'], f['MOLD_NAME_SNAP'],
            f['BARCODE'], f['BC_PART_CODE'], f['MOLD_CHECK'],
            f['INSPECTION_RESULT'], f['DEFECT_ID'], f['DEFECT_NAME_SNAP'],
            f['DISPOSITION'], f['CURRENT_STATUS'], f['ENTRY_USER'], f['ENTRY_STAMP'],
            f['ENTRY_CHANNEL'], f['ROW_STATUS'], f['NOTE'], f['VOID_FLAG'], f['VOID_REASON'],
            f['EVENT_TYPE'], f['DISPOSITION_ID']
        ]
        for j,v in enumerate(row, start=1):
            ws.cell(row=r_idx, column=j, value=v).number_format='@'
    # Table
    if len(facts)>0:
        ref=f'A4:{get_column_letter(len(hdr))}{4+len(facts)}'
        tab=Table(displayName='tblRETURN', ref=ref)
        tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(tab)

def build_dashboard_return(ws, pl, n_rows):
    ws.sheet_view.showGridLines=False
    ws['B2']='داشبورد قطعات برگشتی — KPI'
    ws['B2'].font=Font(bold=True, size=16, name='Tahoma', color='FFC00000')
    facts=pl['facts']
    return_facts=[f for f in facts if f['RECORD_TYPE']=='RETURN']
    total=len(facts)
    total_return=len(return_facts)
    unique_return=len(set(f['BARCODE'] for f in return_facts))
    ws['B4']='شاخص'
    ws['C4']='مقدار'
    ws['B4'].font=F_HDR
    ws['C4'].font=F_HDR
    kpis=[
        ('کل رکوردها', total),
        ('کل برگشتی', total_return),
        ('بارکد یکتا برگشتی', unique_return),
        ('نرخ برگشتی %', round(total_return/total*100,2) if total else 0),
        ('برگشتی امروز', len([f for f in return_facts if f['EVENT_DATE_J']==pl['facts'][-1]['EVENT_DATE_J']])),
    ]
    for i,(name,val) in enumerate(kpis, start=5):
        ws.cell(row=i, column=2, value=name).font=F_LBL
        ws.cell(row=i, column=3, value=val).font=Font(bold=True, size=14, name='Tahoma')

    # Top parts returned
    ws['B11']='قطعات پربرگشت'
    ws['B11'].font=F_TITLE
    from collections import Counter
    cnt=Counter(f['PART_NAME_SNAP'] for f in return_facts)
    for i,(part,c) in enumerate(cnt.most_common(10), start=12):
        ws.cell(row=i, column=2, value=part)
        ws.cell(row=i, column=3, value=c)

    ws['E11']='عیوب پرتکرار برگشتی'
    ws['E11'].font=F_TITLE
    cnt2=Counter(f['DEFECT_NAME_SNAP'] for f in return_facts)
    for i,(defect,c) in enumerate(cnt2.most_common(10), start=12):
        ws.cell(row=i, column=5, value=defect)
        ws.cell(row=i, column=6, value=c)

def build(path, pl):
    import openpyxl
    wb=openpyxl.Workbook()
    ws_login=wb.active
    ws_login.title=SH['login']
    ws_data=wb.create_sheet(SH['data'])
    sheets={}
    for k in ('form','staging','master','helper','reports','dash','dq','set','help','log','verify','users'):
        sheets[k]=wb.create_sheet(SH[k])
    sheets['login']=ws_login
    # Additional
    ws_return_form=wb.create_sheet('RETURN فرم برگشتی')
    ws_return_data=wb.create_sheet('RETURN داده برگشتی')
    ws_menu=wb.create_sheet('Menu')
    ws_info1=wb.create_sheet('Information1')
    ws_helpramz=wb.create_sheet('HelpRamz')
    ws_data2=wb.create_sheet('Data')

    # Build core
    n_rows,cap=build_data(ws_data, pl)
    ws_data2['A1']='Data duplicate'
    ws_data2.sheet_state='veryHidden'
    build_master(sheets['master'], pl)
    lists,cal_last,params,grids=build_helper(sheets['helper'], pl, n_rows, cap)
    helper_name=SH['helper']
    build_return_form(ws_return_form, pl, lists, helper_name)
    # Also keep old form
    from gen_workbook_v4 import build_form
    build_form(sheets['form'], pl, lists, helper_name)
    from gen_workbook_v4 import build_staging
    build_staging(sheets['staging'], pl, lists, helper_name)
    build_reports(sheets['reports'], pl, n_rows)
    build_dashboard(sheets['dash'], pl, lists, helper_name, n_rows, cap, grids)
    build_dashboard_return(wb.create_sheet('Dashboard برگشتی'), pl, n_rows)
    build_dq(sheets['dq'], pl)
    build_settings(sheets['set'], pl)
    build_help(sheets['help'], pl)
    build_log(sheets['log'])
    build_users(sheets['users'], pl)
    build_verify(sheets['verify'], pl, {'total':n_rows,'unique_barcode':len(set(f['BARCODE'] for f in pl['facts'])),'REWORK':0,'SCRAP':0,'RETURN_TO_PROCESS':0}, n_rows)
    build_login_v6(ws_login, pl)
    build_return_data(ws_return_data, pl)

    # Information1 and HelpRamz
    ws_info1['A1']='LOG_AUDIT'
    ws_info1['A2']='ردیف'
    ws_info1['B2']='نام سیستم'
    ws_info1['C2']='نام کاربر'
    ws_info1['D2']='تاریخ ورود'
    ws_info1['E2']='ساعت ورود'
    ws_info1['F2']='ساعت خروج'
    ws_info1['H2']='نام کاربری'
    ws_info1['I2']='رمز عبور'
    ws_info1['J2']='ترکیب'
    ws_info1['K2']='نام کاربری'
    ws_info1['L2']='رمز عبور'
    ws_info1['M2']='ترکیب'
    ws_info1['N2']='مورد تایید است ؟'
    users=pl['dims'].get('USER',[])
    for idx,u in enumerate(users, start=3):
        ws_info1.cell(row=idx, column=8, value=u['USERNAME'])
        ws_info1.cell(row=idx, column=9, value=u['PASSWORD'])
        ws_info1.cell(row=idx, column=10, value=f'=H{idx}&" - "&I{idx}')
    ws_info1['K3']='admin'
    ws_info1['L3']='admin123'
    ws_info1['M3']='=K3&" - "&L3'
    ws_info1['N3']='=IF(ISNUMBER(MATCH(M3,J:J,0)),"YES","NO")'
    ws_info1.sheet_state='veryHidden'

    ws_helpramz['A22']='نام کاربری'
    ws_helpramz['F22']='نام کاربر'
    ws_helpramz['G22']='کلمه عبور'
    for idx,u in enumerate(users, start=23):
        ws_helpramz.cell(row=idx, column=1, value=u['USERNAME'])
        ws_helpramz.cell(row=idx, column=6, value=u['DISPLAY_NAME'])
        ws_helpramz.cell(row=idx, column=7, value=u['PASSWORD'])
    ws_helpramz.sheet_state='veryHidden'

    ws_menu['B2']='Menu — قطعات برگشتی'
    ws_menu.sheet_state='veryHidden'

    # Hide all except LOGIN
    for sh in wb.sheetnames:
        if sh not in (SH['login'],):
            ws=wb[sh]
            if ws.sheet_state=='visible':
                ws.sheet_state='veryHidden'
    wb[SH['login']].sheet_state='visible'

    wb.save(path)
    print(f"saved {path} sheets={wb.sheetnames}")

if __name__=='__main__':
    import argparse, json
    ap=argparse.ArgumentParser()
    ap.add_argument('--payload', default='analysis/payload_v4.json')
    ap.add_argument('-o','--out', default='/tmp/QC-F-14-v6-BASE.xlsx')
    args=ap.parse_args()
    with open(args.payload, 'r', encoding='utf-8') as f:
        pl=json.load(f)
    build(args.out, pl)
