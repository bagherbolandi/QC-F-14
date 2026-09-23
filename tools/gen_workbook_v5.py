"""
QC-F-14 v5 — like sampel file: Menu visible only, Login UserForm, full-screen, role visibility
"""
import os, sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
sys.path.insert(0, os.path.dirname(__file__))
import xlcal, xlmigrate as M
from gen_workbook_v4 import SH, FACT_COLS, build_data, build_master, build_helper, build_form, build_staging, build_reports, build_dashboard, build_dq, build_settings, build_help, build_verify, build_log, build_users, build_login, F_TITLE, F_LBL, F_TXT, F_NOTE, F_HDR, FILL_HDR, CTRD, BORDER, WRAP

C_HEADER='FF1F3864'
C_OK='FF375623'

def build_login_v5(ws, pl):
    ws.sheet_view.showGridLines=False
    ws.sheet_view.showRowColHeaders=False
    ws.column_dimensions['A'].width=2
    ws.column_dimensions['B'].width=22
    ws.column_dimensions['C'].width=32
    ws.column_dimensions['D'].width=40
    ws.column_dimensions['E'].width=15
    ws['B2']='QC-F-14 — سامانه کنترل کیفیت'
    ws['B2'].font=Font(bold=True, size=20, name='Tahoma', color=C_HEADER)
    ws['B3']='LOGIN — ورود به سامانه (نسخه ۵ مشابه نمونه)'
    ws['B3'].font=Font(bold=True, size=14, name='Tahoma', color='FF7030A0')
    ws['B4']='این فایل دارای لاگین UserForm مانند نمونه است. در Workbook_Open فرم لاگین نمایش داده می‌شود، ریبون و تب‌ها مخفی می‌شوند، فقط LOGIN دیده می‌شود.'
    ws['B4'].font=F_NOTE
    ws.merge_cells('B4:D5')
    ws['B4'].alignment=WRAP

    ws['B6']='نام کاربری'
    ws['B6'].font=F_LBL
    ws['C6'].fill=PatternFill('solid', fgColor='FFFFFFFF')
    ws['C6'].border=BORDER
    ws['C6'].number_format='@'
    ws['C6'].value='admin'
    ws['B8']='رمز عبور'
    ws['B8'].font=F_LBL
    ws['C8'].fill=PatternFill('solid', fgColor='FFFFFFFF')
    ws['C8'].border=BORDER
    ws['C8'].number_format='@'
    ws['C8'].value='admin123'

    ws['B10']='کاربر فعلی'
    ws['B10'].font=F_LBL
    ws['C10'].border=BORDER
    ws['C10'].value='(وارد نشده)'
    ws['B11']='نقش'
    ws['B11'].font=F_LBL
    ws['C11'].border=BORDER
    ws['B12']='وضعیت'
    ws['B12'].font=F_LBL
    ws['C12'].border=BORDER
    ws['C12'].value='لطفاً از طریق فرم لاگین وارد شوید'

    # Hidden validation area like sampel Information1
    ws['K2']='Validation area (like sampel)'
    ws['K3'].value='admin'  # will be overwritten by UserForm
    ws['L3'].value='admin123'
    ws['M3'].value='=K3&" - "&L3'
    # J column will be filled from USERS later via formula, but we put placeholder
    ws['J2']='ترکیب'
    ws['J3'].value='=H3&" - "&I3'
    ws['H2']='نام کاربری'
    ws['I2']='رمز عبور'
    ws['H3']='admin'
    ws['I3']='admin123'
    ws['N3'].value='=IF(ISNUMBER(MATCH(M3,J:J,0)),"YES","NO")'
    ws['N2']='تایید؟'

    ws['B14']='راهنما'
    ws['B14'].font=F_TITLE
    ws['B15']='• ADMIN: admin / admin123 → همه شیت‌ها'
    ws['B16']='• OPERATOR: operator / 1234 → FORM + STAGING + Data + Master + Help + Dashboard + Reports'
    ws['B17']='• VIEWER: viewer / viewer123 → Dashboard + Reports + Help + Data'
    ws['B18']='• بازرس‌ها: i-01 / 1234 ...'
    ws['B19']='• لاگین با UserForm (مانند نمونه) — در Workbook_Open نمایش داده می‌شود. اگر بسته شد، فایل بسته می‌شود.'
    ws['B20']='• خروج: QC_Logout — از Developer → Macros اجرا کنید.'
    for r in range(15,21):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=2).font=F_TXT
        ws.cell(row=r, column=2).alignment=WRAP

    ws['B22']='وضعیت شیت‌ها'
    ws['B22'].font=F_TITLE
    ws['B23']='پس از لاگین موفق، شیت‌های مجاز Visible و بقیه VeryHidden می‌شوند. ریبون مخفی است (CloseMenu) — برای نمایش OpenMenu را اجرا کنید.'
    ws.merge_cells('B23:D24')
    ws['B23'].font=F_NOTE
    ws['B23'].alignment=WRAP

    ws.sheet_properties.tabColor='FF7030A0'

def build_information1(ws, pl):
    ws.sheet_view.showGridLines=False
    ws['A1']='اطلاعات ورود کاربران (LOG_AUDIT + Validation)'
    ws['A1'].font=F_TITLE
    # Headers like sampel
    ws['A2']='ردیف'
    ws['B2']='نام سیستم'
    ws['C2']='نام کاربر'
    ws['D2']='تاریخ ورود'
    ws['E2']='ساعت ورود'
    ws['F2']='ساعت خروج'
    ws['G2']='ردیف'
    ws['H2']='نام کاربری'
    ws['I2']='رمز عبور'
    ws['J2']='ترکیب'
    ws['K2']='نام کاربری'
    ws['L2']='رمز عبور'
    ws['M2']='ترکیب'
    ws['N2']='مورد تایید است ؟'
    ws['P2']='نام شیت ها'
    ws['Q2']='ردیف'
    ws['R2']='نام کاربر'
    ws['S2']='نام شیت'
    # Fill users for validation
    users=pl['dims'].get('USER',[])
    for idx, u in enumerate(users, start=3):
        ws.cell(row=idx, column=8, value=u['USERNAME'])
        ws.cell(row=idx, column=9, value=u['PASSWORD'])
        ws.cell(row=idx, column=10, value=f'=H{idx}&" - "&I{idx}')
    # Validation cells
    ws['K3']='admin'
    ws['L3']='admin123'
    ws['M3']='=K3&" - "&L3'
    ws['N3']='=IF(ISNUMBER(MATCH(M3,J:J,0)),"YES","NO")'
    # Audit log sample
    ws['A3']='=IF(B3="","",COUNT($A$2:A2)+1)'
    ws['B3']='system'
    ws['C3']='admin'
    ws['D3']='2026-09-19'
    ws['E3']='11:00:00'
    # Sheet access mapping for OPERATOR etc.
    ws['P3']='Menu'
    ws['P4']='List'
    ws['P5']='Index'
    # For QC we map to our sheets
    qc_sheets=['LOGIN ورود','FORM فرم ثبت','Data داده','STAGING ورود اضطراری','Master اطلاعات پایه','Dashboard داشبورد','Reports گزارش‌ها','Help راهنما','LOG_AUDIT','USERS کاربران']
    for i, sname in enumerate(qc_sheets, start=3):
        ws.cell(row=i, column=16, value=sname)
    # Access per user - admin gets all
    ws['R3']='admin'
    ws['S3']='admin'
    # For operator, list allowed
    allowed_operator=['LOGIN ورود','FORM فرم ثبت','STAGING ورود اضطراری','Data داده','Master اطلاعات پایه','Help راهنما','Dashboard داشبورد','Reports گزارش‌ها']
    for i, sname in enumerate(allowed_operator, start=4):
        ws.cell(row=i, column=18, value='operator')
        ws.cell(row=i, column=19, value=sname)
    allowed_viewer=['LOGIN ورود','Dashboard داشبورد','Reports گزارش‌ها','Help راهنما','Data داده']
    start=len(allowed_operator)+4
    for i, sname in enumerate(allowed_viewer, start=start):
        ws.cell(row=i, column=18, value='viewer')
        ws.cell(row=i, column=19, value=sname)

    ws.sheet_state='veryHidden'

def build_helpramz(ws, pl):
    ws['A1']='USER'
    ws['B1']='Pass 1'
    ws['D1']='Pass 2'
    ws['E1']='Password'
    ws['F1']='نام کاربر'
    ws['G1']='کلمه عبور'
    # Top active users (like sampel)
    users=pl['dims'].get('USER',[])
    for idx, u in enumerate(users, start=2):
        ws.cell(row=idx, column=1, value=u['USERNAME'])
        ws.cell(row=idx, column=2, value=f'=IF(A{idx}="","",IF(A{idx}=0,"",VLOOKUP(A{idx},$A$23:$G$42,7,0)))')
        ws.cell(row=idx, column=4, value=f'=IF(A{idx}=$F$1,$H$1,B{idx})')
        ws.cell(row=idx, column=5, value=f'=IF(D{idx}=0,B{idx},D{idx})')
    # Master list A23:G42
    ws['A22']='نام کاربری'
    ws['B22']='تصویر کاربر'
    ws['D22']='امضای کاربر'
    ws['E22']='شماره تماس'
    ws['F22']='نام کاربر'
    ws['G22']='کلمه عبور'
    for idx, u in enumerate(users, start=23):
        ws.cell(row=idx, column=1, value=u['USERNAME'])
        ws.cell(row=idx, column=5, value='')
        ws.cell(row=idx, column=6, value=u['DISPLAY_NAME'])
        ws.cell(row=idx, column=7, value=u['PASSWORD'])
    ws['B43']='=COUNTA(A23:A42)'
    ws.sheet_state='veryHidden'

def build(path, pl):
    wb=openpyxl.Workbook()
    ws_login=wb.active
    ws_login.title=SH['login']
    # Create QC sheets
    ws_data=wb.create_sheet(SH['data'])
    sheets={}
    for k in ('form','staging','master','helper','reports','dash','dq','set','help','log','verify','users'):
        sheets[k]=wb.create_sheet(SH[k])
    sheets['login']=ws_login
    # Additional sampel-compatible sheets
    ws_menu=wb.create_sheet('Menu')
    ws_info1=wb.create_sheet('Information1')
    ws_helpramz=wb.create_sheet('HelpRamz')
    ws_data2=wb.create_sheet('Data')
    ws_list=wb.create_sheet('List')
    ws_index=wb.create_sheet('Index')
    ws_home=wb.create_sheet('Home')
    # Build
    n_rows,cap=build_data(ws_data, pl)
    # duplicate Data to Data sheet - simple copy without table to avoid duplicate name
    ws_data2.sheet_view.showGridLines=False
    ws_data2['A1']='Data duplicate for sampel compatibility'
    ws_data2.sheet_state='veryHidden'
    build_master(sheets['master'], pl)
    lists,cal_last,params,grids=build_helper(sheets['helper'], pl, n_rows, cap)
    helper_name=SH['helper']
    build_form(sheets['form'], pl, lists, helper_name)
    build_staging(sheets['staging'], pl, lists, helper_name)
    build_reports(sheets['reports'], pl, n_rows)
    model,filt=build_dashboard(sheets['dash'], pl, lists, helper_name, n_rows, cap, grids)
    build_dq(sheets['dq'], pl)
    set_refs=build_settings(sheets['set'], pl)
    build_help(sheets['help'], pl)
    build_log(sheets['log'])
    build_users(sheets['users'], pl)
    build_verify(sheets['verify'], pl, model, n_rows)
    build_login_v5(ws_login, pl)
    build_information1(ws_info1, pl)
    build_helpramz(ws_helpramz, pl)

    # Menu sheet duplicate of login
    ws_menu.sheet_view.showGridLines=False
    ws_menu['B2']='QC-F-14 — Menu (مانند نمونه)'
    ws_menu['B2'].font=F_TITLE
    ws_menu['B3']='این شیت Menu است — در نسخه نمونه فقط Menu دیده می‌شود. در نسخه ۵ ما LOGIN را Visible می‌کنیم.'
    ws_menu.sheet_state='veryHidden'

    # List, Index, Home minimal
    for ws in [ws_list, ws_index, ws_home]:
        ws['A1']='Sample compatible sheet'
        ws.sheet_state='veryHidden'

    # Hide all except LOGIN
    for sh in wb.sheetnames:
        if sh not in (SH['login'],):
            ws=wb[sh]
            if ws.sheet_state=='visible':
                ws.sheet_state='veryHidden'
    # Make LOGIN visible
    wb[SH['login']].sheet_state='visible'

    wb.save(path)
    print(f"saved {path} sheets={wb.sheetnames}")

if __name__=='__main__':
    import argparse, json, os
    ap=argparse.ArgumentParser()
    ap.add_argument('--payload', default='analysis/payload_v4.json')
    ap.add_argument('-o','--out', default='/tmp/QC-F-14-v5-BASE.xlsx')
    args=ap.parse_args()
    with open(args.payload, 'r', encoding='utf-8') as f:
        pl=json.load(f)
    build(args.out, pl)
