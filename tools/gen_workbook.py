"""
ساختِ فایل نسخهٔ ۲ — `QC-F-14-v2.xlsx` (ساختار + مدل داده + فرمول‌ها؛ بدون ماکرو).

    python3 -W ignore tools/gen_workbook.py [--payload analysis/payload.json] [-o QC-F-14-v2.xlsx]

پیش‌شرط: `tools/build_payload.py` و تست‌های `tools/test_jalali.py` و
`tools/test_msovba.py` باید سبز باشند.

اصولی که در این فایل «کد» شده‌اند، نه شعار:
  • روی داده هیچ فرمولی نیست (۱۴٬۵۵۷ ردیف، مقادیرِ خالص). فرمول فقط در Helper/
    Reports/Dashboard/DQ می‌نشیند (بند ۲۴: پایداری > زیبایی).
  • بارکد از ابتدا Text است، با قاعدهٔ «دقیقاً ۱۹ رقم» و پیامِ فارسی.
  • لیست‌ها از جدول‌های Master با فیلتر «فعال» ساخته می‌شوند ⇒ افزودن/غیرفعال‌کردن،
    لیست فرم را به‌روز می‌کند و دادهٔ تاریخی را نمی‌شکند (کد + اسنیپ‌شات).
  • همهٔ COUNTIFSها کرانه‌دارند (تا آخرین ردیفِ پرشده) تا رفرشِ داشبورد ~۱–۲ ثانیه بماند.
  • رنگ فقط معنایی: ضایعات قرمز، اصلاحی نارنجی، برگشتی ارغوانی، هشدارِ داده زرد.
  • شیتِ پنهان Verify: هر عددی که فرمولی وعده می‌دهد، در همین build از payload
    محاسبه و در کنارِ متنِ فرمول نوشته می‌شود؛ تستِ فایل، فرمول و مدل را به هم
    می‌چسباند تا «منطقِ فرمول» و «منطقِ مهاجرت» صفر اختلاف باشند.
"""
from __future__ import annotations

import argparse
import collections
import datetime
import json
import os
import sys

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlcal                                                # noqa: E402

# ───────────────────── رنگ‌ها، سبک‌ها، ثابت‌ها ─────────────────────
C_HEADER = 'FF1F3864'
C_SCRAP = 'FFC00000'
C_REWORK = 'FFED7D31'
C_RETURN = 'FF7030A0'
C_OK = 'FF375623'
C_WARN = 'FFFFF2CC'
C_INPUT = 'FFFFFFFF'
C_FILTER = 'FFDEEAF6'

F_HDR = Font(bold=True, color='FFFFFFFF', size=11, name='Tahoma')
F_TITLE = Font(bold=True, size=14, name='Tahoma', color=C_HEADER)
F_KPI = Font(bold=True, size=18, name='Tahoma', color=C_HEADER)
F_LBL = Font(bold=True, size=10, name='Tahoma')
F_TXT = Font(size=10, name='Tahoma')
F_NOTE = Font(size=9, italic=True, color='FF808080', name='Tahoma')
FILL_HDR = PatternFill('solid', fgColor=C_HEADER)
FILL_IN = PatternFill('solid', fgColor=C_INPUT)
FILL_FLT = PatternFill('solid', fgColor=C_FILTER)
FILL_WARN = PatternFill('solid', fgColor=C_WARN)
BORDER = Border(*[Side(style='thin', color='FFB4C6E7')] * 4)
RTL = Alignment(horizontal='right', vertical='center')   # RTL از sheet_view 읽힌د
CTRD = Alignment(horizontal='center', vertical='center')
WRAP = Alignment(wrap_text=True, vertical='top')

BARCODE_ERR = 'شماره بارکد صحیح نیست. بارکد باید دقیقاً ۱۹ رقم باشد (بدون فاصله و بدون حرف).'
REQUIRED_ERR = 'این فیلد الزامی است و باید از لیست انتخاب شود.'
ALL = 'همه'

SH = {'form': 'FORM فرم ثبت', 'data': 'Data داده', 'staging': 'STAGING ورود اضطراری',
      'master': 'Master اطلاعات پایه', 'helper': 'Helper محاسبات',
      'reports': 'Reports گزارش‌ها', 'dash': 'Dashboard داشبورد',
      'dq': 'DQ کیفیت داده', 'set': 'Settings تنظیمات', 'help': 'Help راهنما',
      'log': 'LOG_AUDIT', 'verify': 'Verify'}

FACT_COLS = ['EVENT_ID', 'EVENT_DATE_J', 'EVENT_DATE_G', 'J_MONTH_KEY', 'J_WEEK_KEY',
             'SHIFT_ID', 'STATION_ID', 'STATION_SNAP', 'PART_ID', 'PART_NAME_SNAP',
             'MOLD_ID', 'MOLD_NAME_SNAP', 'DEFECT_ID', 'DEFECT_NAME_SNAP',
             'DISPOSITION_ID', 'EVENT_TYPE', 'BARCODE', 'BC_PART_CODE', 'MOLD_CHECK',
             'ENTRY_STAMP', 'ENTRY_USER', 'INSPECTOR_ID', 'INSPECTOR_SNAP',
             'ENTRY_CHANNEL', 'ROW_STATUS', 'NOTE']
FACT_HDR_FA = ['کد رخداد', 'تاریخ وقوع (شمسی)', 'تاریخ میلادی', 'کلید ماه', 'کلید هفته',
               'شیفت', 'کد ایستگاه', 'ایستگاه (اسنیپ‌شات)', 'کد قطعه', 'قطعه (اسنیپ‌شات)',
               'کد قالب', 'قالب/حفره', 'کد عیب', 'عیب (اسنیپ‌شات)', 'کد محل صدور',
               'نوع رخداد', 'بارکد (متن ۱۹ رقمی)', 'کدِ قطعهٔ بارکد', 'بررسی حفره',
               'زمان ثبت', 'کاربر ثبت', 'کد بازرس', 'بازرس (اسنیپ‌شات)',
               'کانال ثبت', 'وضعیت رکورد', 'توضیح']
L = {c: get_column_letter(i) for i, c in enumerate(FACT_COLS, start=1)}
FACT_W = [11, 15, 12, 10, 11, 8, 11, 16, 9, 22, 9, 12, 9, 24, 12, 11, 21, 10, 13, 17, 14,
          9, 18, 11, 15, 26]
# فیلترهای داشبورد: (کلید، سطر، حالت) — حالت EXACT4 برای «سالِ ۴ رقمیِ اولِ تاریخِ شمسی»
FILTERS = [('J_YEAR', 4, 'YEAR'), ('J_MONTH_KEY', 5, 'EQ'), ('EVENT_DATE_J', 6, 'EQ'),
           ('J_WEEK_KEY', 7, 'EQ'), ('SHIFT_ID', 8, 'EQ'), ('STATION_ID', 9, 'EQ'),
           ('PART_ID', 10, 'EQ'), ('MOLD_ID', 11, 'EQ'), ('DEFECT_ID', 12, 'EQ'),
           ('ROW_STATUS', 13, 'EQ')]
FILTER_LABELS = dict(zip([f[0] for f in FILTERS],
                         ['سال جلالی', 'ماه (۱۴۰۵۰۶)', 'روز (تاریخ وقوع)', 'هفته',
                          'شیفت', 'ایستگاه', 'قطعه', 'قالب/حفره', 'عیب', 'وضعیت رکورد']))
# فیلتر → لیستِ کدِ خالص در Helper (ستون)؛ None = تایپِ آزاد
FILTER_LIST = {'SHIFT_ID': 'CODE_SHIFT', 'STATION_ID': 'CODE_STATION', 'PART_ID': 'CODE_PART',
               'MOLD_ID': 'CODE_MOLD', 'DEFECT_ID': 'CODE_DEFECT',
               'ROW_STATUS': 'CODE_ROW_STATUS', 'EVENT_DATE_J': 'CODE_J_TEXT'}
DIM_TABLES = [
    ('DIM_PART', 'قطعه‌ها', ['CODE', 'NAME', 'ACTIVE'], ['کد قطعه', 'نام قطعه (نرمال‌شده)', 'فعال']),
    ('DIM_MOLD', 'قالب‌ها', ['CODE', 'NAME', 'MOLD_KIND', 'ACTIVE'],
     ['کد قالب', 'نام قالب', 'نوع (حفره/جهت)', 'فعال']),
    ('DIM_DEFECT', 'عیوب', ['CODE', 'NAME', 'ACTIVE'], ['کد عیب', 'شرح عیب', 'فعال']),
    ('DIM_INSPECTOR', 'بازرس‌ها', ['CODE', 'NAME', 'ALIASES', 'ACTIVE'],
     ['کد بازرس', 'نام', 'نام‌های معادل قدیمی (alias)', 'فعال']),
    ('DIM_STATION', 'ایستگاه‌ها', ['CODE', 'NAME', 'ENABLED_FOR_FORM', 'ACTIVE'],
     ['کد ایستگاه', 'نام ایستگاه', 'فعال در فرم', 'فعال']),
    ('DIM_SHIFT', 'شیفت‌ها', ['CODE', 'NAME', 'ACTIVE'], ['کد شیفت', 'نام', 'فعال']),
    ('DIM_DISPOSITION', 'محل صدور', ['CODE', 'NAME', 'ACTIVE', 'EVENT_TYPE_IF_PICKED'],
     ['کد', 'وضعیت/محل صدور', 'فعال', 'نوعِ رخدادِ قطعی (ماکرو این را می‌خواند)']),
]
MASTER_TOP = 4
LIST_NAME = {'DIM_SHIFT': 'shift_list', 'DIM_STATION': 'station_list', 'DIM_PART': 'part_list',
             'DIM_MOLD': 'mold_list', 'DIM_DEFECT': 'defect_list',
             'DIM_INSPECTOR': 'inspector_list', 'DIM_DISPOSITION': 'disposition_list'}
DATE_LIST = 'date_list'
QC_MESSAGES = {
    'E_BARCODE': ('بارکد معتبر نیست: باید دقیقاً ۱۹ رقم باشد (بدونِ فاصله، بدونِ حرف). '
                  'اگر صفرهایِ ابتدا حذف شده‌اند، سلول را از قبل روی Text بگذارید و دوباره اسکن کنید.'),
    'E_INCOMPLETE': 'همهٔ ۹ فیلد لازم است — پیامِ سطرِ ۲۸ِ فرم نامِ فیلدهایِ خالی را می‌گوید.',
    'E_CAPACITY': ('ظرفیتِ tblQC پر است. یک سطرِ جدید جا نیست؛ سالِ گذشته را به فایلِ بایگانی '
                   'منتقل کنید (راهنما، بند ۸). تا آن زمان ثبتِ جدید متوقف است — فایل را دستکاری نکنید.'),
    'E_DATE': ('تاریخِ شمسی در جدولِ تقویمِ Helper پیدا نشد. فرمت باید ۱۴۰۵/۰۶/۲۸ باشد و از لیست '
               'انتخاب شود (تقویمِ Helper تا پایانِ ۱۴۰۸ پوشش داده شده است).'),
    'E_DUP': ('این بارکد قبلاً ثبت شده است. ردیفِ قبلی را با ROW_STATUS = ابطال‌شده باطل کنید '
              '(هرگز حذف نکنید). اگر واقعاً دو قطعه با یک بارکد است، «ثبتِ با علامتِ تکرار» '
              'رکورد را با وضعیتِ مشکوک-تکراری می‌نویسد تا در DQ دیده شود.'),
    'E_DUP_ACT': 'رکورد با وضعیتِ «مشکوک-تکراری» ثبت شد (نه رد، نه حذف). در شیت DQ دیده می‌شود.',
    'OK_APPEND': '✓ ثبت شد — ردیفِ',
    'E_STAGE_NONE': ('هیچ ردیفِ «آمادهٔ Import» در STAGING نیست. ستونِ FLAG باید '
                     '✓ آماده Import را نشان دهد (سلول‌هایِ این شیت فرمول‌اند؛ دستکاری‌شان نکنید).'),
    'OK_IMPORT': '✓ ردیف از STAGING به tblQC اضافه شد و سطرهایِ واردشده از STAGING پاک شدند.',
    'E_VOID': ('برای ابطال، فقط یک سلول از ردیفِ موردنظر در Data را انتخاب کنید. '
               'ماکرو ROW_STATUS را ابطال‌شده می‌کند و ردیف را حذف نمی‌کند.'),
    'OK_VOID': '✓ ROW_STATUS این ردیف ابطال‌شده شد (داده حذف نشد؛ گزارش‌ها از این پس حسابش نمی‌کنند).',
    'E_CONFIG': ('پیکربندیِ ماکرو کامل نیست: نام‌هایِ تعریف‌شدهٔ QC_SHEETS یا QC_MSG را در '
                 'Settings بررسی کنید.'),
    'E_NORULE': ('ایستگاهِ انتخابی در RULE_STATION (شیت Master) قاعده ندارد. ماکرو برای '
                 'ایستگاهِ بی‌قاعده نوعِ رخداد را حدس نمی‌زند — یک ردیف به RULE_STATION اضافه کنید.'),
    'E_SHEET': ('یکی از شیت‌ها پیدا نشد. اگر نامِ شیت را عوض کرده‌اید، جدولِ '
                'QC_SHEETS (در Settings) را به‌روز کنید.'),
    'CHANNEL_FORM': 'ماکرو/فرم',
    'CHANNEL_STAGING': 'ماکرو/STAGING',
    'ABOUT': ('QC-F-14 v2 — مسیرِ نوشتن. محاسبه‌ها همه فرمول‌اند؛ این ماژول فقط '
              'اعتبارسنجیِ نهایی، درجِ سطر، و ثبتِ رویداد در LOG_AUDIT را انجام می‌دهد. '
              'هیچ سطرِ داده‌ای را پاک یا ویرایش نمی‌کند (به‌جز علامتِ ROW_STATUS برای ابطال).'),
}

ROW_STATUS_OPTIONS = ['تأییدشده', 'پیش‌نویس', 'ابطال‌شده', 'مشکوک-تکراری']

DISPOSITION_ALIAS = {'DIS-01': 'اصلاحی', 'DIS-02': 'ضایعات', 'DIS-03': 'برگشتی'}


# ───────────────────── ابزارِ مشترک ─────────────────────
def hdr_row(ws, row, start_col, headers):
    for j, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + j, value=h)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CTRD, BORDER
    return row + 1


def box(ws, r0, r1, c0, c1):
    for rr in range(r0, r1 + 1):
        for cc in range(c0, c1 + 1):
            ws.cell(row=rr, column=cc).border = BORDER


def filter_terms(n_rows):
    """عباراتِ COUNTIFSِ مشترکِ همهٔ گزارش‌ها/کارت‌ها (کرانه‌دار، سازگار با «همه»).

    «همه» به `*` تبدیل می‌شود؛ اگر مقدارِ فیلتر کدِ دقیق باشد، تطبیقِ کامل است.
    J_YEAR روی ستونِ متنِ تاریخ (LEFT 4) سوار است، پس مستقل از تنظیماتِ ویندوز کار می‌کند.
    """
    r0, r1 = 2, 1 + n_rows
    src = {'J_YEAR': 'EVENT_DATE_J'}          # سال = ۴ کاراکترِ اولِ متنِ تاریخِ شمسی
    out = []
    for key, row, mode in FILTERS:
        cell = f'Dashboard!$G${row}'
        col = L[src.get(key, key)]
        rng = f'Data!${col}{r0}:${col}${r1}'
        if mode == 'YEAR':
            out.append(f'{rng},IF({cell}="{ALL}","*",LEFT({cell},4))')
        else:
            out.append(f'{rng},IF({cell}="{ALL}","*",{cell})')
    return ','.join(out)


def countifs(filt, extra=''):
    return f'=COUNTIFS({filt}{"," + extra if extra else ""})'


def add_text(cell, fmt='@'):
    cell.number_format = fmt
    return cell


def list_ref(lists, helper_name, key):
    col, n = lists[key]
    return (f"'{helper_name}'!${get_column_letter(col)}$2:"
            f"${get_column_letter(col)}${max(n + 1, 2)}")


def add_list_dv(ws, lists, helper_name, key, target, prompt=None, error=None):
    dv = DataValidation(type='list', formula1=f'={list_ref(lists, helper_name, key)}',
                        allow_blank=True, showErrorMessage=bool(error),
                        showInputMessage=bool(prompt))
    if prompt:
        dv.prompt, dv.promptTitle = prompt, 'راهنما'
    if error:
        dv.error, dv.errorTitle = error, 'ورودی نامعتبر'
    ws.add_data_validation(dv)
    for t in (target if isinstance(target, list) else [target]):
        dv.add(t)
    return dv


# ───────────────────── شیت‌ها ─────────────────────
def build_data(ws, pl):
    ws['A1'] = None
    for j, (col, fa) in enumerate(zip(FACT_COLS, FACT_HDR_FA), start=1):
        c = ws.cell(row=1, column=j, value=f'{fa}\n{col}')
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CTRD, BORDER
        ws.column_dimensions[get_column_letter(j)].width = FACT_W[j - 1]
    ws.row_dimensions[1].height = 40
    text_cols = {L[c] for c in FACT_COLS
                 if c not in ('EVENT_DATE_G', 'J_MONTH_KEY')}
    r = 1
    for f in pl['facts']:
        r += 1
        vals = dict(f)
        row = [vals['EVENT_ID'], vals['EVENT_DATE_J'],
               datetime.date.fromisoformat(vals['EVENT_DATE_G']) if vals['EVENT_DATE_G'] else '',
               vals['J_MONTH_KEY'], vals['J_WEEK_KEY'], vals['SHIFT_ID'], vals['STATION_ID'],
               vals['STATION_SNAP'], vals['PART_ID'], vals['PART_NAME_SNAP'], vals['MOLD_ID'],
               vals['MOLD_NAME_SNAP'], vals['DEFECT_ID'], vals['DEFECT_NAME_SNAP'],
               vals['DISPOSITION_ID'], vals['EVENT_TYPE'], vals['BARCODE'],
               vals['BC_PART_CODE'], vals['MOLD_CHECK'], '', vals['ENTRY_USER'],
               vals['INSPECTOR_ID'], vals['INSPECTOR_SNAP'],
               vals['_ENTRY_CHANNEL'], vals['_ROW_STATUS'], vals['_NOTE']]
        for j, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.font = F_TXT
            letter = get_column_letter(j)
            if letter in text_cols:
                cell.number_format = '@'
            elif FACT_COLS[j - 1] == 'EVENT_DATE_G':
                cell.number_format = 'yyyy/mm/dd'
    n_rows = r - 1
    cap = max(n_rows, pl['max_fact_rows'])
    ref = f'A1:{get_column_letter(len(FACT_COLS))}{cap}'
    tab = Table(displayName='tblQC', ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=False)
    ws.add_table(tab)
    ws.freeze_panes = 'B2'
    # اعتبارِ داده روی ستون‌هایِ کد (برای ویرایشِ دستیِ اصلاحی) — بدونِ DV روی بارکد
    # تا فایلِ ۱۴هزارردیفی سنگین نشود؛ DV بارکد روی فرم و STAGING است.
    for key, lst in (('SHIFT_ID', 'shift_list'), ('STATION_ID', 'station_list'),
                     ('PART_ID', 'part_list'), ('MOLD_ID', 'mold_list'),
                     ('DEFECT_ID', 'defect_list'), ('DISPOSITION_ID', 'disposition_list'),
                     ('INSPECTOR_ID', 'inspector_list')):
        col = L[key]
        dv = DataValidation(type='list', formula1=f'={lst}', allow_blank=True,
                           showErrorMessage=True)
        dv.error, dv.errorTitle = REQUIRED_ERR, 'ورودی نامعتبر'
        ws.add_data_validation(dv)
        dv.add(f'{col}2:{col}{cap}')
    dv = DataValidation(type='list', formula1='"' + ','.join(ROW_STATUS_OPTIONS) + '"',
                        allow_blank=True, showErrorMessage=True)
    ws.add_data_validation(dv)
    dv.add(f'{L["ROW_STATUS"]}2:{L["ROW_STATUS"]}{cap}')
    ws.protection.sheet = True
    ws.protection.formatCells = False
    ws.protection.insertRows = False
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.sheet_properties.tabColor = C_HEADER
    return n_rows, cap


def build_master(ws, pl):
    ws['B2'] = 'اطلاعاتِ پایه — تنها شیتی که سرشیفت حقِ ویرایشش را دارد (حذف ممنوع: غیرفعال کنید)'
    ws['B2'].font = F_TITLE
    col = 2
    info = {}
    for tname, title, fields, headers in DIM_TABLES:
        rows = pl['dims'][tname[4:]]
        ws.cell(row=3, column=col, value=title).font = F_LBL
        hdr_row(ws, MASTER_TOP, col, headers)
        for i, rec in enumerate(rows, start=MASTER_TOP + 1):
            for j, f in enumerate(fields):
                cell = ws.cell(row=i, column=col + j, value=rec[f])
                cell.font, cell.border = F_TXT, BORDER
                if f in ('CODE', 'NAME', 'ALIASES'):
                    cell.number_format = '@'
                if f in ('ACTIVE', 'ENABLED_FOR_FORM') and rec[f] == 'خیر':
                    cell.fill = FILL_WARN
        last = MASTER_TOP + len(rows)
        ref = f'{get_column_letter(col)}{MASTER_TOP}:{get_column_letter(col + len(fields) - 1)}{last}'
        tab = Table(displayName=tname, ref=ref)
        tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(tab)
        info[tname] = (col, last, len(fields))
        col += len(fields) + 3
    # نگاشتِ قطعه→قالب (استخراج‌شده از داده، با شاهد)
    ws.cell(row=3, column=col, value='نگاشت قطعه → قالب (فقط با شاهدِ ≥۹۰٪ و ≥۲۰ رکورد)').font = F_LBL
    hdr_row(ws, MASTER_TOP, col, ['کد قطعه', 'قطعه', 'قالب غالب', 'سهم ٪', 'شاهد', 'اجرا در فرم؟'])
    part_code = {v: k for k, v in pl['code_name']['PART'].items()}
    r = MASTER_TOP
    for rec in pl['map_part_mold']:
        r += 1
        for j, v in enumerate([part_code.get(rec['PART'], '—'), rec['PART'], rec['MOLD'],
                              round(rec['SHARE'] * 100, 1), rec['N'], rec['ENFORCED']]):
            cell = ws.cell(row=r, column=col + j, value=v)
            cell.font, cell.border = F_TXT, BORDER
            if j == 3:
                cell.number_format = '0.0'
            if rec['ENFORCED'] == 'بله':
                cell.fill = PatternFill('solid', fgColor='FFE2EFDA')
    tab = Table(displayName='MAP_PART_MOLD',
                ref=f'{get_column_letter(col)}{MASTER_TOP}:{get_column_letter(col + 5)}{r}')
    tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
    ws.add_table(tab)
    col2 = col + 8
    ws.cell(row=3, column=col2, value='قاعدهٔ ایستگاه → نوع رخداد (تنها جایِ منطقِ کسب‌وکاری)').font = F_LBL
    hdr_row(ws, MASTER_TOP, col2,
            ['کد', 'ایستگاه', 'نوع رخداد', 'مبنای قاعده', 'پیش‌فرض؟ (۱=بله)'])
    r2 = MASTER_TOP
    for rec in pl['rule_station']:
        r2 += 1
        for j, v in enumerate([rec['CODE'], rec['STATION'], rec['EVENT_TYPE'], rec['BASIS'],
                               int(rec.get('IS_DEFAULT', 0))]):
            cell = ws.cell(row=r2, column=col2 + j, value=v)
            cell.font, cell.border = F_TXT, BORDER
            if rec['BASIS'].startswith('پیش‌فرض'):
                cell.fill = FILL_WARN
    tab = Table(displayName='RULE_STATION',
                ref=f'{get_column_letter(col2)}{MASTER_TOP}:{get_column_letter(col2 + 4)}{r2}')
    tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
    ws.add_table(tab)
    for j in range(1, col2 + 7):
        ws.column_dimensions[get_column_letter(j)].width = 15
    ws.column_dimensions[get_column_letter(col2 + 3)].width = 46
    ws.sheet_view.zoomScale = 85
    ws.sheet_properties.tabColor = 'FF70AD47'
    return info


def build_helper(ws, pl, n_rows, cap):
    """تقویم، فهرست‌های DV، کدهایِ خالص، پارامترها. شیتِ محاسبه — کاربر با آن کار ندارد."""
    cal_cols = list(xlcal.CAL_COLUMNS)
    hdr_row(ws, 1, 1, cal_cols)
    r = 1
    for rec in xlcal.rows(xlcal.CAL_START, xlcal.CAL_END):
        r += 1
        for j, h in enumerate(cal_cols, start=1):
            cell = ws.cell(row=r, column=j, value=rec[h])
            if h == 'G_DATE':
                cell.number_format = 'yyyy/mm/dd'
    cal_last = r
    col = len(cal_cols) + 2
    lists = {}
    # پلِ ماکرو به تقویم: به‌جای آدرِسِ ستون، نامِ تعریف‌شده (ستون‌ها: G_DATE=1 … J_WEEK_KEY=8)
    for nm, j in (('QC_CAL_G', 1), ('QC_CAL_MK', cal_cols.index('J_MONTH_KEY') + 1),
                  ('QC_CAL_WK', cal_cols.index('J_WEEK_KEY') + 1)):
        lists[nm] = (j, cal_last - 1)
    for tname, title, fields, headers in DIM_TABLES:
        act = [x for x in pl['dims'][tname[4:]] if x['ACTIVE'] == 'بله']
        if tname == 'DIM_STATION':
            act = [x for x in act if x.get('ENABLED_FOR_FORM', 'بله') == 'بله']
        ws.cell(row=1, column=col, value=f'LIST_{tname[4:]}').font = F_LBL
        ws.cell(row=1, column=col + 1, value=f'CODE_{tname[4:]}').font = F_LBL
        for i, x in enumerate(act, start=2):
            ws.cell(row=i, column=col, value=f"{x['CODE']} | {x['NAME']}").number_format = '@'
            ws.cell(row=i, column=col + 1, value=x['CODE']).number_format = '@'
        lists[LIST_NAME[tname]] = (col, len(act))
        lists['CODE_' + tname[4:]] = (col + 1, len(act))
        col += 2
    # فهرست‌هایِ افزوده
    mco = pl.get('mold_check_options') or ['هم‌خوان', 'ناهم‌خوان', 'قالبِ جهت',
                                          'رقمِ خارج از بازه', 'بارکد نامعتبر']
    extra = {'CODE_J_TEXT': [row['J_TEXT'] for row in xlcal.rows(xlcal.CAL_START, xlcal.CAL_END)],
             'CODE_ROW_STATUS': pl.get('row_status_options') or ROW_STATUS_OPTIONS,
             'CODE_MOLD_CHECK': mco,
             'CODE_YEAR': sorted({f['_J_YEAR'] for f in pl['facts'] if f['_J_YEAR']}),
             'LIST_ALL': [ALL]}
    for name, vals in extra.items():
        ws.cell(row=1, column=col, value=name).font = F_LBL
        for i, v in enumerate(vals, start=2):
            ws.cell(row=i, column=col, value=v).number_format = '@'
        lists[name] = (col, len(vals))
        col += 1
    # فهرستِ تاریخِ فرم: ۱۸۰ روزِ اخیرِ تقویم از «امروزِ فایل» به عقب (امروز = آخرین تاریخِ داده)
    last_day = max(f['EVENT_DATE_J'] for f in pl['facts'] if f['EVENT_DATE_J'])
    idx = [i for i, row in enumerate(xlcal.rows(xlcal.CAL_START, xlcal.CAL_END))
           if row['J_TEXT'] == last_day]
    start_i = idx[0] if idx else cal_last - 2
    ws.cell(row=1, column=col, value=DATE_LIST).font = F_LBL
    days = []
    d = xlcal.j_to_g(*[int(x) for x in last_day.split('/')])
    for k in range(180):
        days.append(xlcal.j_text(*xlcal.g_to_j(d - datetime.timedelta(days=k))))
    for i, v in enumerate(days, start=2):
        ws.cell(row=i, column=col, value=v).number_format = '@'
    lists[DATE_LIST] = (col, len(days))
    col += 2
    # شبکهٔ کدها برای شمارشِ متمایز (COUNTIFS نمی‌تواند معیارِ آرایه‌ای بگیرد؛ پس
    # یک بلوکِ ۴۰سطریِ ثابت داریم و SUMPRODUCT روی همان اجرا می‌شود). ۴۰ = سقفِ
    # ردیفِ هر فهرست؛ اگر Master از ۴۰ ردیف گذشت، همین‌جا بزرگ‌ش کنید (README).
    ws.cell(row=1, column=col, value='GRID').font = F_LBL
    grids = {}
    for tname, key in (('DIM_PART', 'PART_ID'), ('DIM_DEFECT', 'DEFECT_ID'),
                       ('DIM_STATION', 'STATION_ID'), ('DIM_MOLD', 'MOLD_ID'),
                       ('DIM_SHIFT', 'SHIFT_ID')):
        ws.cell(row=1, column=col, value=f'GRID_{key}').font = F_LBL
        codes = [x['CODE'] for x in pl['dims'][tname[4:]]]
        for i in range(40):
            ws.cell(row=2 + i, column=col, value=codes[i] if i < len(codes) else '#N/A')
        grids[key] = (col, 40)
        col += 1
    # پارامترها
    ws.cell(row=1, column=col, value='PARAM').font = F_LBL
    params = [('N_ROWS', n_rows), ('CAP_ROWS', cap), ('FIRST_ROW', 2), ('LAST_ROW', 1 + n_rows),
              ('CAL_FIRST', 2), ('CAL_LAST', cal_last), ('SRC_FILE', pl['source']),
              ('BUILT', pl['generated']), ('N_DQ', pl['stats']['dq_issues']),
              ('MOLD_ALIGN_٪', 94.1), ('VERSION', 'v2.0.0'),
              ('SHEET_DATA', SH['data']), ('NOTE', 'این شیت پنهان است؛ ویرایشش گزارش را می‌شکند')]
    for i, (k, v) in enumerate(params, start=2):
        ws.cell(row=i, column=col, value=k).font = F_LBL
        ws.cell(row=i, column=col + 1, value=v)
    lists['PARAM'] = (col, len(params))
    for j in range(1, col + 2):
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.column_dimensions['A'].width = 12
    ws.sheet_state = 'hidden'
    ws.sheet_properties.tabColor = 'FFBFBFBF'
    return lists, cal_last, params, grids


def build_form(ws, pl, lists, helper_name):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 2
    for c in 'BCDEF':
        ws.column_dimensions[c].width = 34 if c in 'BE' else 22
    ws['B2'] = 'فرم ثبتِ قطعاتِ برگشتی / ضایعات / اصلاحی'
    ws['B2'].font = Font(bold=True, size=16, name='Tahoma', color=C_HEADER)
    ws['B3'] = (f'نسخهٔ ۲ · مهاجرت‌یافته از {pl["source"]} · {pl["stats"]["fact_rows"]} رکوردِ '
                'تاریخی در شیت «Data داده» · ثبتِ جدید فقط با ماکرو (یا STAGING)')
    ws['B3'].font = F_NOTE
    # (برچسب، ستونِ مقدار، سطر، لیست، راهنما)
    fields = [('تاریخ وقوع (شمسی)', 'C', 6, DATE_LIST, 'از لیست انتخاب کنید؛ قالب ۱۴۰۵/۰۶/۲۸'),
              ('شیفت', 'C', 8, 'shift_list', 'شیفت ۱/۲/۳'),
              ('ایستگاه', 'C', 10, 'station_list', 'فقط ایستگاه‌های فعال در فرم'),
              ('بازرس', 'C', 12, 'inspector_list', 'نام بازرسِ ثبت‌کننده'),
              ('قطعه', 'F', 6, 'part_list', 'کد | نام قطعه'),
              ('قالب / حفره', 'F', 8, 'mold_list', 'قالب‌های مجازِ همین قطعه (در صورت وجود قاعده)'),
              ('عیب', 'F', 10, 'defect_list', 'شرح عیب'),
              ('محل صدور (وضعیت)', 'F', 12, 'disposition_list', 'اصلاحی / ضایعات / برگشتی'),
              ('بارکد (۱۹ رقم)', 'C', 14, None, 'بارکد را اسکن کنید؛ صفرهای ابتدایی حفظ می‌شود')]
    cells = {}
    for label, col, row, lst, hint in fields:
        lab_col = 'B' if col == 'C' else 'E'
        ws[f'{lab_col}{row}'] = label
        ws[f'{lab_col}{row}'].font = F_LBL
        ws[f'{lab_col}{row}'].alignment = RTL
        cell = ws[f'{col}{row}']
        cell.number_format = '@'
        cell.fill = FILL_IN
        cell.border = BORDER
        cell.alignment = CTRD
        cells[label] = f'{col}{row}'
        if lst:
            add_list_dv(ws, lists, helper_name, lst, f'{col}{row}', hint, REQUIRED_ERR)
        else:
            dvt = DataValidation(type='textLength', operator='equal', formula1='19',
                                 allow_blank=True, showErrorMessage=True, errorStyle='stop')
            dvt.error, dvt.errorTitle = BARCODE_ERR, 'بارکد نامعتبر'
            dvt.prompt, dvt.promptTitle = hint, 'بارکد'
            dvt.showInputMessage = True
            ws.add_data_validation(dvt)
            dvt.add(f'{col}{row}')
            dvn = DataValidation(type='custom',
                                 formula1=f'=AND(LEN({col}{row})=19,ISNUMBER({col}{row}+0))',
                                 allow_blank=True, showErrorMessage=True, errorStyle='stop')
            dvn.error, dvn.errorTitle = BARCODE_ERR, 'بارکد نامعتبر'
            ws.add_data_validation(dvn)
            dvn.add(f'{col}{row}')
    box(ws, 6, 14, 2, 6)
    ws['B17'] = 'خلاصهٔ ورودی (فقط نمایش — روی داده فرمول نیست)'
    ws['B17'].font = F_TITLE
    r = 18
    for label, col, row, lst, hint in fields:
        ws[f'B{r}'] = label
        ws[f'B{r}'].font = F_LBL
        ws[f'C{r}'] = f'={col}{row}'
        ws[f'C{r}'].number_format = '@'
        ws[f'C{r}'].border = BORDER
        r += 1
    bc, datec = cells['بارکد (۱۹ رقم)'], cells['تاریخ وقوع (شمسی)']
    stc, shc = cells['ایستگاه'], cells['شیفت']
    ws[f'B{r + 1}'] = 'پیامِ اعتبارسنجی'
    ws[f'B{r + 1}'].font = F_LBL
    missing = '&'.join([f'IF({v}="","«{k}» ")' for k, v in cells.items()])
    n_cells = ','.join(cells.values())
    code_col = {k: {'شیفت': L['SHIFT_ID'], 'ایستگاه': L['STATION_ID'], 'قطعه': L['PART_ID'],
                    'قالب / حفره': L['MOLD_ID'], 'عیب': L['DEFECT_ID'],
                    'محل صدور (وضعیت)': L['DISPOSITION_ID'], 'بازرس': L['INSPECTOR_ID']}[k]
                for k in ('شیفت', 'ایستگاه', 'قطعه', 'قالب / حفره', 'عیب',
                          'محل صدور (وضعیت)', 'بازرس')}
    n_rows = pl['stats']['fact_rows']
    chk = ','.join(
        f'Data!${code_col[k]}2:${code_col[k]}{n_rows + 1},IF({v}="{ALL}","*",'
        f'LEFT(TRIM({v}),FIND("|",{v} & "|")-1))'
        for k, v in cells.items() if k in code_col)
    ws[f'C{r + 1}'] = (f'=IF(COUNTA({n_cells})<9,"⚠ ناقص: "&{missing},'
                       f'IF(OR(LEN({bc})<>19,NOT(ISNUMBER({bc}+0))),"⚠ ' + BARCODE_ERR + '",'
                       f'IF(COUNTIFS(Data!${L["BARCODE"]}2:${L["BARCODE"]}{n_rows + 1},{bc})>0,'
                       f'"⚠ این بارکد قبلاً ثبت شده — ردیفِ قبلی را ابطال کنید، حذف نکنید",'
                       f'IF(COUNTIFS({chk})=0,"⚠ بعضی کدها در Data یافت نشد (بررسیِ لیست)",'
                       f'"✓ ورودی کامل است — «ثبت» را بزنید"))))')
    ws[f'C{r + 1}'].font = Font(bold=True, size=11, name='Tahoma')
    ws[f'B{r + 3}'] = ('ماکرو: QC_Append (فایلِ vba/Module1-QC_WritePath.bas را در VBE وارد کنید '
                       'و به همین دکمه/میان‌بر بسپارید). بدونِ ماکرو: همین ۹ مقدار را در '
                       '«STAGING ورود اضطراری» بزنید، سپس QC_Import_Staging.')
    ws[f'B{r + 3}'].font = F_NOTE
    ws[f'B{r + 5}'] = '۱۰ رکوردِ اخیرِ همین ایستگاه/شیفت (Drill-down)'
    ws[f'B{r + 5}'].font = F_TITLE
    hdr_row(ws, r + 6, 2, ['کد', 'تاریخ', 'قطعه', 'عیب', 'نوع', 'بارکد', 'حفره؟'])
    for i in range(10):
        rr = r + 7 + i
        src = n_rows - i
        for j, key in enumerate(['EVENT_ID', 'EVENT_DATE_J', 'PART_NAME_SNAP', 'DEFECT_NAME_SNAP',
                                 'EVENT_TYPE', 'BARCODE', 'MOLD_CHECK']):
            cell = ws.cell(row=rr, column=2 + j, value=f"='{SH['data']}'!{L[key]}{src + 1}")
            cell.number_format = '@'
            cell.font = F_TXT
    box(ws, r + 6, r + 16, 2, 8)
    ws.print_area = f'B1:F{r + 18}'
    ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.tabColor = C_OK
    ws.protection.sheet = True
    ws.protection.formatCells = False
    ws.protection.insertRows = False
    ws.protection.sort = False
    ws.protection.autoFilter = False


def build_staging(ws, pl, lists, helper_name):
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'STAGING — ورودِ اضطراریِ بدونِ ماکرو'
    ws['B2'].font = F_TITLE
    ws['B3'] = ('اگر سیاستِ IT ماکرو را بلاک کرد، بازرس همین ۹ فیلد را اینجا وارد می‌کند. '
                'ستون FLAG با فرمول مشخص می‌کند کدام ردیف «آمادهٔ Import» است. این شیت '
                'append-only نیست (عمداً): پس از Import، ردیف را پاک کنید تا دوباره ثبت نشود.')
    ws['B3'].font = F_NOTE
    ws.merge_cells('B3:L5')
    ws['B3'].alignment = WRAP
    hdr = ['تاریخ شمسی', 'شیفت', 'ایستگاه', 'بازرس', 'قطعه', 'قالب', 'عیب', 'محل صدور',
           'بارکد', 'FLAG', 'REASON', 'ثبت‌شده؟']
    hr = hdr_row(ws, 7, 2, hdr)
    for col, w in zip('BCDEFGHIJKL', [15, 8, 16, 14, 24, 12, 24, 13, 21, 15, 40, 11]):
        ws.column_dimensions[col].width = w
    for r in range(hr, hr + 200):
        for j in range(2, 14):
            c = ws.cell(row=r, column=j)
            c.number_format = '@'
            c.border = BORDER
            if j in (11, 12):
                c.font = F_LBL
        ws.cell(row=r, column=11, value=(
            f'=IF(COUNTA(B{r}:J{r})=0,"",IF(COUNTA(B{r}:J{r})<9,"⚠ ناقص",'
            f'IF(LEN(I{r})<>19,"⚠ بارکد",'
            f'IF(COUNTIF(Data!${L["BARCODE"]}:${L["BARCODE"]},I{r})>0,"⚠ تکراری","✓ آماده Import"))))'))
        ws.cell(row=r, column=12, value=(
            f'=IF(K{r}="","",IF(K{r}="⚠ ناقص","تاریخ/شیفت/ایستگاه/بازرس/قطعه/قالب/عیب/وضعیت/بارکد",'
            f'IF(K{r}="⚠ بارکد","' + BARCODE_ERR + '",IF(K{r}="⚠ تکراری",'
            f'"بارکد در tblQC وجود دارد","کد ' + 'ها از Master خوانده شود' + '"))))'))
    for key, lst in (('C', 'shift_list'), ('D', 'station_list'), ('E', 'inspector_list'),
                     ('F', 'part_list'), ('G', 'mold_list'), ('H', 'defect_list'),
                     ('I', 'disposition_list'), ('B', DATE_LIST)):
        add_list_dv(ws, lists, helper_name, lst, f'{key}{hr}:{key}{hr + 199}',
                    None, REQUIRED_ERR)
    dvb = DataValidation(type='custom', formula1=f'=IF(J{hr}="",TRUE,LEN(J{hr})=19)',
                         showErrorMessage=True, errorStyle='stop')
    dvb.error, dvb.errorTitle = BARCODE_ERR, 'بارکد نامعتبر'
    ws.add_data_validation(dvb)
    dvb.add(f'J{hr}:J{hr + 199}')
    ws.freeze_panes = f'B{hr}'
    ws.sheet_properties.tabColor = C_REWORK


def build_reports(ws, pl, n_rows):
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'گزارش‌های مدیریتی (۱۵) — همه از tblQC؛ هیچ‌گاه از فرم'
    ws['B2'].font = F_TITLE
    ws['B3'] = ('فیلترها در شیت Dashboard (سلول‌های G4:G13) خوانده می‌شود؛ «همه» یعنی *. '
                'سطرِ «جمعِ بلوک» در هر گزارش باید با کارتِ «جمعِ کل» داشبورد برابر باشد — '
                'این همان تستِ ترازِ فاز ۵ است.')
    ws['B3'].font = F_NOTE
    filt = filter_terms(n_rows)
    r0, r1 = 2, 1 + n_rows
    groups = {}
    for rep in pl['reports']:
        groups.setdefault(rep['block'], []).append(rep)
    verify = []
    r = 5
    for block, reps in groups.items():
        c = ws.cell(row=r, column=2, value=f'بلوک: {block}')
        c.font, c.fill = F_TITLE, FILL_FLT
        r += 1
        for rep in reps:
            r, vv = report_block(ws, r, rep, pl, filt, r0, r1)
            verify += vv
        r += 1
    for col, w in zip('BCDEFGH', [40, 12, 12, 12, 12, 12, 16]):
        ws.column_dimensions[col].width = w
    ws.sheet_properties.tabColor = C_HEADER
    return verify


def report_block(ws, r, rep, pl, filt, r0, r1):
    """هر گزارش = بلوکِ ۷ستونی با COUNTIFS. سه گزارش «خاص» هندسهٔ دیگری دارند."""
    title = ws.cell(row=r, column=2, value=f"{rep['id']} — {rep['title']}")
    title.font = F_LBL
    note = rep.get('note') or ''
    if note:
        ws.cell(row=r, column=6, value=note).font = F_NOTE
    if rep.get('special') == 'matrix':
        return matrix_block(ws, r, rep, pl, filt, r0, r1)
    if rep.get('special') == 'backdate':
        return khash_block(ws, r, rep, pl, r0, r1)
    if rep.get('special') == 'dq':
        return dqsum_block(ws, r, rep, pl)
    if rep.get('special') == 'barcode':
        return barcode_block(ws, r, rep, pl, r0, r1)
    if rep.get('special') == 'mgmt':
        return mgmt_block(ws, r, rep, pl, filt, r0, r1)
    hdr = ['مقدار (کد | نام)', 'کل', 'ضایعات', 'اصلاحی', 'برگشتی', 'سهم ٪', 'تجمعی ٪']
    h1 = hdr_row(ws, r + 1, 2, hdr)
    dim = rep['group']
    src_col = {'PART_ID': L['PART_ID'], 'DEFECT_ID': L['DEFECT_ID'], 'MOLD_ID': L['MOLD_ID'],
               'STATION_ID': L['STATION_ID'], 'SHIFT_ID': L['SHIFT_ID'],
               'DISPOSITION_ID': L['DISPOSITION_ID'], 'ENTRY_USER': L['ENTRY_USER'],
               'INSPECTOR_ID': L['INSPECTOR_ID'],
               'J_MONTH_KEY': L['J_MONTH_KEY'], 'J_WEEK_KEY': L['J_WEEK_KEY'],
               'J_DAY': L['EVENT_DATE_J'], 'BARCODE': L['BARCODE']}.get(dim)
    if src_col is None and kind is None and dim not in ('J_MONTH_KEY', 'J_WEEK_KEY', 'J_DAY',
                                                        'ENTRY_USER'):
        raise ValueError(f'گزارش {rep["id"]} بُعدِ ناشناخته دارد: {dim}')
    kind = {'PART_ID': 'PART', 'DEFECT_ID': 'DEFECT', 'MOLD_ID': 'MOLD', 'STATION_ID': 'STATION',
            'SHIFT_ID': 'SHIFT', 'DISPOSITION_ID': 'DISPOSITION',
            'INSPECTOR_ID': 'INSPECTOR'}.get(dim)
    items, cnt_m = [], collections.Counter()
    if kind:
        items = list(pl['code_name'][kind].items())
    elif dim == 'ENTRY_USER':
        items = sorted({(f['ENTRY_USER'], f['ENTRY_USER']) for f in pl['facts']})
    elif dim in ('J_MONTH_KEY', 'J_WEEK_KEY', 'J_DAY'):
        key = {'J_MONTH_KEY': 'J_MONTH_KEY', 'J_WEEK_KEY': 'J_WEEK_KEY', 'J_DAY': 'EVENT_DATE_J'}[dim]
        vals = sorted({f[key] for f in pl['facts'] if f[key] != ''})
        items = [(v, str(v)) for v in vals]
    top = rep.get('top')
    if dim == 'J_DAY':
        items = items[-30:]
    if top and dim in ('PART_ID', 'DEFECT_ID'):
        cnts = collections.Counter(f[dim] for f in pl['facts'])
        items = sorted([c for c in items if c[0] in cnts], key=lambda c: -cnts[c[0]])[:top]
    for f in pl['facts']:
        v = f['EVENT_DATE_J'] if dim == 'J_DAY' else f[dim]
        if v != '' and v is not None:
            cnt_m[v] += 1
    grand = max(1, sum(cnt_m.values()))
    if not items:
        items = [('—', 'داده‌ای برای این گروه یافت نشد')]
    rr = h1
    for code, name in items:
        cv = f'"{code}"' if isinstance(code, str) else str(code)
        extra = f'Data!${src_col}{r0}:${src_col}{r1},{cv}' if src_col else ''
        c = ws.cell(row=rr, column=2, value=f'{code} | {name}')
        c.number_format, c.font = '@', F_TXT
        for j, et in enumerate([None, 'ضایعات', 'اصلاحی', 'برگشتی']):
            ex = extra
            if et:
                ex = (f'{ex},' if ex else '') + f'Data!${L["EVENT_TYPE"]}{r0}:${L["EVENT_TYPE"]}{r1},"{et}"'
            ws.cell(row=rr, column=3 + j, value=countifs(filt, ex)).number_format = '#,##0'
        rr += 1
    tot_row = rr
    for i, row in enumerate(range(h1, h1 + len(items))):
        ws.cell(row=row, column=7,
                value=f'=IFERROR(C{row}/MAX(1,$C${tot_row}),0)').number_format = '0.0%'
        prev = f'G{row - 1},' if i else ''
        ws.cell(row=row, column=8, value=f'=SUM({prev}G{row})').number_format = '0.0%'
        ws.cell(row=row, column=3).number_format = '#,##0'
    ws.cell(row=tot_row, column=2, value='جمعِ بلوک (باید = کارتِ «کل موارد»)').font = F_LBL
    for j in range(3, 7):
        cl = get_column_letter(j)
        ws.cell(row=tot_row, column=j, value=f'=SUM({cl}{h1}:{cl}{tot_row - 1})').number_format = '#,##0'
    ws.cell(row=tot_row, column=7,
            value=f'=IFERROR(SUM(G{h1}:G{tot_row - 1}),0)').number_format = '0.0%'
    ws.cell(row=tot_row, column=8,
            value=f'=IF(ROUND(C{tot_row}-Dashboard!$C$7,6)=0,"✓ تراز با داشبورد",'
                  f'"⚠ اختلاف: "&TEXT(Dashboard!$C$7-C{tot_row},"0"))')
    box(ws, r, tot_row, 2, 8)
    return tot_row + 2, {'report': rep['id'], 'top': items[0][0] if items else None,
                         'model_top_count': cnt_m.get(items[0][0], 0) if items else 0,
                         'grand': grand}


def matrix_block(ws, r, rep, pl, filt, r0, r1):
    """R12 — ماتریس ۱۰ قطعه × ۸ عیب (COUNTIFSِ دوکلیدی)."""
    cnt_p = collections.Counter(f['PART_ID'] for f in pl['facts'])
    cnt_d = collections.Counter(f['DEFECT_ID'] for f in pl['facts'])
    parts = [c for c, _ in cnt_p.most_common(10)]
    defects = [c for c, _ in cnt_d.most_common(8)]
    h = r + 1
    ws.cell(row=h, column=2, value='قطعه ＼ عیب').font = F_LBL
    for j, d in enumerate(defects):
        c = ws.cell(row=h, column=3 + j, value=f'{d}\n{pl["code_name"]["DEFECT"][d]}')
        c.font, c.fill, c.alignment = F_LBL, FILL_FLT, CTRD
        c.number_format = '@'
    tot_col = 3 + len(defects)
    ws.cell(row=h, column=tot_col, value='جمعِ ردیف').font = F_LBL
    for i, p in enumerate(parts):
        rr = h + 1 + i
        ws.cell(row=rr, column=2,
                value=f'{p} | {pl["code_name"]["PART"][p]}').number_format = '@'
        for j, d in enumerate(defects):
            ws.cell(row=rr, column=3 + j, value=countifs(
                filt, f'Data!${L["PART_ID"]}{r0}:${L["PART_ID"]}{r1},"{p}",'
                      f'Data!${L["DEFECT_ID"]}{r0}:${L["DEFECT_ID"]}{r1},"{d}"'
            )).number_format = '#,##0'
        ws.cell(row=rr, column=tot_col,
                value=f'=SUM(C{rr}:{get_column_letter(2 + len(defects))}{rr})')
    box(ws, r, h + len(parts), 2, tot_col)
    return h + len(parts) + 2, {'report': rep['id'], 'model_top_count': cnt_p.most_common(1)[0][1]}


def khash_block(ws, r, rep, pl, r0, r1):
    """R13 — رکوردهایِ علامت‌خوردهٔ مهاجرت (تکراری/عقب‌گرد): شمارش + فهرستِ نمونه."""
    ws.cell(row=r + 1, column=2, value='شرح').font = F_LBL
    ws.cell(row=r + 1, column=4, value='تعداد').font = F_LBL
    items = [('ردیف با ROW_STATUS = مشکوک-تکراری',
              f'=COUNTIF(Data!${L["ROW_STATUS"]}{r0}:${L["ROW_STATUS"]}{r1},"مشکوک-تکراری")',
              pl['stats']['dup_extra_rows']),
             ('گروهِ تکرارِ بارکد', '="از DQ بخوانید"', pl['stats']['dup_groups']),
             ('ردیفِ عقب‌گردِ ثبت (تاریخِ وقوع < ردیفِ پیشین)', '="از DQ بخوانید"',
              pl['stats']['backdate_rows']),
             ('ردیفِ ابطال‌شده', f'=COUNTIF(Data!${L["ROW_STATUS"]}{r0}:${L["ROW_STATUS"]}{r1},"ابطال‌شده")', 0),
             ('ردیفِ پیش‌نویس', f'=COUNTIF(Data!${L["ROW_STATUS"]}{r0}:${L["ROW_STATUS"]}{r1},"پیش‌نویس")', 0)]
    rr = r + 1
    for label, f, model_v in items:
        rr += 1
        ws.cell(row=rr, column=2, value=label)
        ws.cell(row=rr, column=4, value=f).number_format = '#,##0'
    ws.cell(row=rr + 2, column=2,
            value='هرگز حذف نمی‌شوند؛ یا ابطال، یا اصلاحِ ROW_STATUS (بند ۱۷: حذفِ داده تصمیمِ ابزار نیست).').font = F_NOTE
    box(ws, r, rr, 2, 4)
    return rr + 3, {'report': rep['id'], 'model_top_count': pl['stats']['dup_extra_rows']}


def dqsum_block(ws, r, rep, pl):
    """R14 — خلاصهٔ ایرادها (از فهرستِ مهاجرت؛ شمارشِ زنده در شیت DQ است)."""
    cnt = collections.Counter(d['RULE'] for d in pl['dq_issues'])
    rr = hdr_row(ws, r + 1, 2, ['RULE', 'شرح', 'تعداد در مهاجرت', 'آیا ورودیِ جدید را هم می‌گیرد؟'])
    for rule, n in cnt.most_common():
        ws.cell(row=rr, column=2, value=rule)
        ws.cell(row=rr, column=3, value={'DQ-DUP': 'بارکد تکراری', 'DQ-MOLD-RULE': 'قالب ≠ قاعدهٔ قطعه',
                                         'DQ-BACKDATE': 'عقب‌گردِ ثبت', 'DQ-BLANK': 'ردیف خالیِ میانی',
                                         'DQ-SPACE': 'فاصلهٔ اضافی در نام',
                                         'DQ-BARCODE': 'بارکد ≠ ۱۹ رقم'}.get(rule, rule))
        ws.cell(row=rr, column=4, value=n).number_format = '#,##0'
        ws.cell(row=rr, column=5, value='بله — فرمولِ شیت DQ').font = F_NOTE
        rr += 1
    box(ws, r, rr, 2, 5)
    return rr + 2, {'report': rep['id'], 'model_top_count': cnt.most_common(1)[0][1] if cnt else 0}


def barcode_block(ws, r, rep, pl, r0, r1):
    """R11 — ردیابیِ بارکد: یک بارکد → تعدادِ رکوردها + اولین/آخرین رکورد (فرمولِ ساده،
    بدونِ آرایه؛ `MATCH` روی ستون بارکد که Text است و در tblQC یک‌دانه)."""
    ws.cell(row=r + 1, column=2, value='بارکدِ موردِ جست‌وجو →').font = F_LBL
    q = ws.cell(row=r + 1, column=4)
    q.fill, q.border, q.number_format = FILL_IN, BORDER, '@'
    QC = f'Data!${L["BARCODE"]}{r0}:${L["BARCODE"]}{r1}'
    rr = r + 3
    hdr_row(ws, rr, 2, ['معیار', 'مقدار'])
    rows = [('شمارِ رکوردهایِ این بارکد', f'=IF(LEN($D${r + 1})<>19,"",COUNTIF({QC},$D${r + 1}))'),
            ('تاریخِ اولین رکورد', f'=IFERROR(INDEX(Data!${L["EVENT_DATE_J"]}{r0}:${L["EVENT_DATE_J"]}{r1},'
                                  f'MATCH($D${r + 1},{QC},0)),"—")'),
            ('ایستگاهِ اولین رکورد', f'=IFERROR(INDEX(Data!${L["STATION_SNAP"]}{r0}:${L["STATION_SNAP"]}{r1},'
                                     f'MATCH($D${r + 1},{QC},0)),"—")'),
            ('قطعهٔ اولین رکورد', f'=IFERROR(INDEX(Data!${L["PART_NAME_SNAP"]}{r0}:${L["PART_NAME_SNAP"]}{r1},'
                                  f'MATCH($D${r + 1},{QC},0)),"—")'),
            ('عیبِ اولین رکورد', f'=IFERROR(INDEX(Data!${L["DEFECT_NAME_SNAP"]}{r0}:${L["DEFECT_NAME_SNAP"]}{r1},'
                                 f'MATCH($D${r + 1},{QC},0)),"—")'),
            ('نوعِ رخداد', f'=IFERROR(INDEX(Data!${L["EVENT_TYPE"]}{r0}:${L["EVENT_TYPE"]}{r1},'
                           f'MATCH($D${r + 1},{QC},0)),"—")'),
            ('وضعیتِ رکوردها (چندتایی؟)', f'=IF(LEN($D${r + 1})<>19,"",'
                                          f'IF(COUNTIF({QC},$D${r + 1})>1,"⚠ چند رکورد — ببینید DQ-DUP",'
                                          f'IF(COUNTIF({QC},$D${r + 1})=0,"هیچ رکوردی نیست","✓ تک‌رکورد")))')]
    for i, (lab, f) in enumerate(rows):
        ws.cell(row=rr + 1 + i, column=2, value=lab).font = F_TXT
        c = ws.cell(row=rr + 1 + i, column=3, value=f)
        c.number_format = '@'
        c.font = F_LBL
    ws.cell(row=rr + len(rows) + 2, column=2,
            value='برای دیدنِ «چرا این قطعه دوباره آمده» همین بارکد را در فیلترِ Dashboard بگذارید '
                  '(فیلترِ بارکد عمداً در ۱۰ فیلتر نیست؛ ستون Q را فیلتر کنید).').font = F_NOTE
    box(ws, r, rr + len(rows), 2, 4)
    return rr + len(rows) + 3, {'report': rep['id'], 'model_top_count': 0}


def mgmt_block(ws, r, rep, pl, filt, r0, r1):
    """R15 — خلاصهٔ مدیریتیِ ماهانه: ۳ نوعِ رخداد + سهم + بیشترین عیبِ ماه."""
    months = sorted({f['J_MONTH_KEY'] for f in pl['facts'] if f['J_MONTH_KEY']})
    hdr_row(ws, r + 1, 2, ['ماه', 'ضایعات', 'اصلاحی', 'برگشتی', 'کل', 'سهم ٪'])
    rr = r + 2
    for m in months[-12:]:
        ws.cell(row=rr, column=2, value=str(m)).number_format = '@'
        for j, et in enumerate(['ضایعات', 'اصلاحی', 'برگشتی']):
            ws.cell(row=rr, column=3 + j, value=countifs(
                filt, f'Data!${L["J_MONTH_KEY"]}{r0}:${L["J_MONTH_KEY"]}{r1},{m},'
                      f'Data!${L["EVENT_TYPE"]}{r0}:${L["EVENT_TYPE"]}{r1},"{et}"'
            )).number_format = '#,##0'
        ws.cell(row=rr, column=6, value=f'=SUM(C{rr}:E{rr})').number_format = '#,##0'
        ws.cell(row=rr, column=7, value=f'=IFERROR(F{rr}/MAX(1,SUM($F$ {r + 2}:$F${rr})),"")'.replace('$F$ ', '$F$'))
        ws.cell(row=rr, column=7).number_format = '0.0%'
        rr += 1
    ws.cell(row=rr, column=2, value='جمع').font = F_LBL
    for j in range(3, 7):
        cl = get_column_letter(j)
        ws.cell(row=rr, column=j, value=f'=SUM({cl}{r + 2}:{cl}{rr - 1})').number_format = '#,##0'
    box(ws, r, rr, 2, 7)
    return rr + 2, {'report': rep['id'], 'model_top_count': len(months)}


def build_dashboard(ws, pl, lists, helper_name, n_rows, cap, grids):
    ws.sheet_view.showGridLines = False
    for col, w in zip('BCDEFGHIJKLMN', [24, 16, 14, 12, 18, 20, 3, 20, 12, 12, 12, 12, 12]):
        ws.column_dimensions[col].width = w
    ws['B2'] = 'داشبوردِ مدیریت'
    ws['B2'].font = Font(bold=True, size=16, name='Tahoma', color=C_HEADER)
    ws['B3'] = (f'ساختِ فایل: {pl["generated"]} · منبع: {pl["source"]} · '
                f'پیمایش تا سطر {n_rows + 1} (کرانه‌دار، برای رفرشِ سریع)')
    ws['B3'].font = F_NOTE
    ws['F5'] = 'فیلترها (۱۰)'
    ws['F5'].font = F_TITLE
    for key, row, mode in FILTERS:
        ws[f'F{row}'] = FILTER_LABELS[key]
        ws[f'F{row}'].font = F_LBL
        g = ws[f'G{row}']
        g.value = ALL if key != 'ROW_STATUS' else 'تأییدشده'
        g.fill, g.border, g.number_format, g.alignment = FILL_FLT, BORDER, '@', CTRD
        lst = FILTER_LIST.get(key) or ('CODE_YEAR' if key == 'J_YEAR' else None)
        if lst and lst in lists:
            add_list_dv(ws, lists, helper_name, lst, g)
        box(ws, row, row, 6, 7)
    ws['H4'] = ('«همه» = ستون فیلترنمی‌شود. مقدارِ دیگر باید دقیقاً کدِ موجود در Data باشد '
                '(نه نام). این تنها راهی است که هم RTL درست کار کند هم COUNTIFS سریع بماند.')
    ws['H4'].font = F_NOTE
    filt = filter_terms(n_rows)
    r0, r1 = 2, 1 + n_rows
    ws['B6'] = 'کارت‌های شاخص (۶)'
    ws['B6'].font = F_TITLE
    model = {}
    facts = pl['facts']
    model['total'] = len(facts)
    for et in ('ضایعات', 'اصلاحی', 'برگشتی'):
        model[et] = sum(1 for f in facts if f['EVENT_TYPE'] == et)
    model['distinct_part'] = len({f['PART_ID'] for f in facts})
    model['distinct_defect'] = len({f['DEFECT_ID'] for f in facts})
    for i, k in enumerate(pl['kpi']):
        row = 7 + i
        ws[f'B{row}'] = k['name']
        ws[f'B{row}'].font = F_LBL
        cell = ws[f'C{row}']
        if k['id'] == 'K1':
            cell.value = countifs(filt)
        elif k['id'] in ('K2', 'K3', 'K4'):
            cell.value = countifs(filt, f'Data!${L["EVENT_TYPE"]}{r0}:${L["EVENT_TYPE"]}{r1},"{k["etype"]}"')
        else:
            key = 'PART_ID' if k['id'] == 'K5' else 'DEFECT_ID'
            gcol, gn = grids[key]
            gr = (f"'{helper_name}'!${get_column_letter(gcol)}$2:"
                  f"${get_column_letter(gcol)}${gn + 1}")
            # شمارشِ متمایز با COUNTIFS روی «شبکهٔ کد» (COUNTIFS معیارِ آرایه‌ای
            # نمی‌پذیرد؛ ۴۰ سطرِ ثابت، مگر Master از ۴۰ ردیف بگذرد → README).
            cell.value = f'=SUMPRODUCT(--(COUNTIFS({filt},Data!${L[key]}{r0}:${L[key]}{r1},{gr})>0))'
        cell.font = F_KPI
        cell.number_format = '#,##0'
        cell.fill = FILL_IN
        ws[f'D{row}'] = k['unit']
        ws[f'D{row}'].font = F_TXT
        ws[f'E{row}'] = f'{k["id"]} · سطح: روز/هفته/ماه/سال'
        ws[f'E{row}'].font = F_NOTE
    box(ws, 7, 12, 2, 5)
    ws['B14'] = 'تعریفِ هر شاخص در شیت Settings (فرهنگِ شاخص‌ها) است؛ اینجا فقط عدد.'
    ws['B14'].font = F_NOTE
    ws['B16'] = 'راهنمایِ رنگ: قرمز=ضایعات · نارنجی=اصلاحی · ارغوانی=برگشتی · زرد=هشدارِ داده'
    ws['B16'].font = F_NOTE
    # ── بلوک‌هایِ دادهٔ نمودار (محدود به ۱۲ سطر → نمودارها بدون refresh) ──
    base = 19
    ws.cell(row=base - 1, column=2, value='دادهٔ نمودارها (ثابت‌اندازه — ۸ نمودار)').font = F_TITLE
    blocks = {}
    months = sorted({f['J_MONTH_KEY'] for f in facts if f['J_MONTH_KEY']})[-12:]
    blocks['CH1'] = ('روندِ ماهانه بر حسب نوع رخداد', 'J_MONTH_KEY', [str(m) for m in months],
                     ['ضایعات', 'اصلاحی', 'برگشتی'])
    blocks['CH2'] = ('روندِ ۱۲ روزِ اخیر', 'EVENT_DATE_J',
                     sorted({f['EVENT_DATE_J'] for f in facts})[-12:], [None])
    cnt_part = collections.Counter(f['PART_ID'] for f in facts if f['EVENT_TYPE'] == 'ضایعات')
    blocks['CH3'] = ('۱۰ قطعهٔ برترِ ضایعات', 'PART_ID',
                     [c for c, _ in cnt_part.most_common(10)], [None])
    cnt_def = collections.Counter(f['DEFECT_ID'] for f in facts)
    blocks['CH4'] = ('Pareto عیب (۱۰ موردِ اول)', 'DEFECT_ID',
                     [c for c, _ in cnt_def.most_common(10)], [None])
    blocks['CH5'] = ('قالب/حفره × نوع', 'MOLD_ID', list(pl['code_name']['MOLD'].keys()),
                     ['ضایعات', 'اصلاحی', 'برگشتی'])
    blocks['CH6'] = ('ایستگاه × نوع', 'STATION_ID', list(pl['code_name']['STATION'].keys()),
                     ['ضایعات', 'اصلاحی', 'برگشتی'])
    blocks['CH7'] = ('شیفت × نوع', 'SHIFT_ID', list(pl['code_name']['SHIFT'].keys()),
                     ['ضایعات', 'اصلاحی', 'برگشتی'])
    blocks['CH8'] = ('۸ بازرسِ برتر', 'INSPECTOR_ID',
                     [u for u, _ in collections.Counter(f['INSPECTOR_ID'] for f in facts).most_common(8)],
                     [None])
    row = base
    chart_rows = {}
    for cid, (title, dim, vals, ets) in blocks.items():
        ws.cell(row=row, column=2, value=f'{cid} — {title}').font = F_LBL
        hdr_row(ws, row + 1, 2, ['مقدار'] + ([f'Σ {e}' for e in ets if e] or ['تعداد', 'سهم ٪']))
        for i, v in enumerate(vals):
            rr = row + 2 + i
            ws.cell(row=rr, column=2, value=str(v)).number_format = '@'
            col_src = L[dim]
            for j, et in enumerate(ets):
                extra = f'Data!${col_src}{r0}:${col_src}{r1},"{v}"'
                if et:
                    extra += f',Data!${L["EVENT_TYPE"]}{r0}:${L["EVENT_TYPE"]}{r1},"{et}"'
                ws.cell(row=rr, column=3 + j, value=countifs(filt, extra)).number_format = '#,##0'
            if len(ets) == 1:
                tot = len({f['J_MONTH_KEY'] for f in facts})
                ws.cell(row=rr, column=4,
                        value=f'=IFERROR(C{rr}/MAX(1,{countifs(filt)[1:]}),0)').number_format = '0.0%'
        chart_rows[cid] = (row + 1, row + 1 + max(len(vals), 1), len(ets))
        row += 2 + max(len(vals), 1) + 2
    # نمودارها
    anchors = {'CH1': 'H19', 'CH2': 'H35', 'CH3': 'H51', 'CH4': 'H67',
               'CH5': 'N19', 'CH6': 'N35', 'CH7': 'N51', 'CH8': 'N67'}
    for cid, (r_first, r_last, ncol) in chart_rows.items():
        n_rows_ = r_last - r_first
        data = Reference(ws, min_col=3, max_col=2 + ncol, min_row=r_first, max_row=r_last)
        cats = Reference(ws, min_col=2, min_row=r_first + 1, max_row=r_last)
        ch = LineChart() if cid in ('CH1', 'CH2') else BarChart()
        ch.type = 'col'
        ch.grouping = 'stacked' if ncol > 1 else 'standard'
        ch.overlap = 100 if ncol > 1 else 0
        ch.title = ws.cell(row=r_first - 1, column=2).value
        ch.y_axis.majorGridlines = None
        ch.height, ch.width = 7.2, 12.5
        ch.add_data(data, titles_from_data=True)
        ch.set_categories(cats)
        ws.add_chart(ch, anchors[cid])
    ws.print_area = 'B1:N17'
    ws.sheet_properties.tabColor = C_OK
    return model, filt


def build_dq(ws, pl):
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'DQ — کیفیتِ داده (ایرادها حذف نمی‌شوند؛ فقط علامت می‌خورند)'
    ws['B2'].font = F_TITLE
    ws['B3'] = ('شمارش‌ها روی ستون‌های tblQC با فرمول است ⇒ با هر رکوردِ جدید، همین شیت '
                'به‌روز می‌شود (بدونِ رفرشِ دستی). «۱۰ ایرادِ برترِ ردیفی» از ستون MOLD_CHECK و '
                'وضعیتِ رکورد ساخته می‌شود.')
    ws['B3'].font = F_NOTE
    n = pl['stats']['fact_rows']
    r0, r1 = 2, 1 + n
    rules = [('DQ-BLANK', 'متوسط', 'ردیفِ خالیِ میانی در محدودهٔ جدولِ قبلی', 2),
             ('DQ-SPACE', 'کم', 'فاصلهٔ اضافی در نام (رجبی␣␣، جمعی␣، سینی زیر موتور سمند␣)', 0),
             ('DQ-BARCODE', 'شدید', 'بارکدِ غیرِ ۱۹ رقم', None),
             ('DQ-DUP', 'شدید', 'بارکدِ تکراری (گروه)', 463),
             ('DQ-MOLD-RULE', 'کم', 'قالبِ ثبت‌شده با قاعدهٔ غالبِ همان قطعه نمی‌خواند', None),
             ('DQ-BACKDATE', 'کم', 'تاریخِ وقوع از ردیفِ پیشین کوچک‌تر (عقب‌گردِ ثبت)', 135),
             ('DQ-CAVITY', 'کم', 'رقم ۴ بارکد با حرفِ حفره نمی‌خواند', None)]
    hr = hdr_row(ws, 5, 2, ['RULE', 'شدت', 'شرح', 'تعداد', 'چطور محاسبه شده؟'])
    r = hr
    for code, sev, desc, _v in rules:
        ws.cell(row=r, column=2, value=code)
        ws.cell(row=r, column=3, value=sev)
        ws.cell(row=r, column=4, value=desc)
        if code == 'DQ-BARCODE':
            f = f'=COUNTA(Data!${L["BARCODE"]}{r0}:${L["BARCODE"]}{r1})-SUMPRODUCT((LEN(Data!${L["BARCODE"]}{r0}:${L["BARCODE"]}{r1})=19)*1)'
            how = 'COUNTA منهای شمارِ ۱۹‌رقمی‌ها (روی ردیف‌های پرشده)'
        elif code == 'DQ-CAVITY':
            f = f'=COUNTIF(Data!${L["MOLD_CHECK"]}{r0}:${L["MOLD_CHECK"]}{r1},"ناهم‌خوان")'
            how = 'ستون MOLD_CHECK که در مهاجرت محاسبه شد'
        elif code == 'DQ-DUP':
            f = (f'=SUMPRODUCT((COUNTIFS(Data!${L["BARCODE"]}{r0}:${L["BARCODE"]}{r1},'
                 f'Data!${L["BARCODE"]}{r0}:${L["BARCODE"]}{r1})>1)*1)')
            how = 'شمارِ ردیف‌هایی که بارکدشان بیش از یک بار آمده (۹۴۱ = ۴۶۳ گروه)'
        elif code == 'DQ-BLANK':
            f = f'=ROWS(Data!A{r0}:A{r1})-COUNTA(Data!${L["EVENT_ID"]}{r0}:${L["EVENT_ID"]}{r1})'
            how = 'تفاضلِ تعدادِ سطرِ جدول و شمارِ کدِ رخداد'
        else:
            f = '="از فهرستِ ردیفیِ پایین بخوانید"'
            how = 'مبتنی بر ردیف‌هایِ علامت‌خورده (ROW_STATUS / MOLD_CHECK)'
        ws.cell(row=r, column=5, value=f)
        ws.cell(row=r, column=6, value=how).font = F_NOTE
        ws.cell(row=r, column=6).alignment = WRAP
        r += 1
    ws.cell(row=r, column=2, value='جمع').font = F_LBL
    ws.cell(row=r, column=5, value=f'=SUM(E{hr}:E{r - 1})')
    r += 2
    ws.cell(row=r, column=2, value='۲۰۰ ایرادِ ردیفیِ اول (از فهرستِ کاملِ مهاجرت)').font = F_TITLE
    r += 1
    hdr_row(ws, r, 2, ['RULE', 'شدت', 'سطرِ فایل مرجع', 'شرح'])
    for d in pl['dq_issues'][:200]:
        r += 1
        for j, v in enumerate([d['RULE'], d['SEV'], d['ROW'], d['DETAIL']], start=2):
            ws.cell(row=r, column=j, value=v).font = F_TXT
    for col, w in zip('BCDEF', [16, 10, 20, 60, 40]):
        ws.column_dimensions[col].width = w
    ws.column_dimensions['F'].width = 44
    ws.sheet_properties.tabColor = 'FFFFC000'


def build_settings(ws, pl):
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'Settings — تعریفِ شاخص‌ها، محدودیت‌های صریح، و مرزِ مهاجرت'
    ws['B2'].font = F_TITLE
    r = hdr_row(ws, 4, 2, ['شاخص', 'تعریف', 'فرمول روی tblQC', 'واحد', 'سطح', 'فیلترهای مؤثر'])
    kpi_rows = [
        ('کل موارد ثبت‌شده', 'شمارِ رکوردهایِ تأییدشده در فیلترها', 'COUNTIFS(…)', 'عدد',
         'روز/هفته/ماه/سال', 'هر ۱۰ فیلتر'),
        ('کل ضایعات', ' EVENT_TYPE=ضایعات', 'COUNTIFS(…, P:"ضایعات")', 'عدد', 'ماه', 'هر ۱۰'),
        ('کل اصلاحی', 'EVENT_TYPE=اصلاحی', 'COUNTIFS(…, P:"اصلاحی")', 'عدد', 'ماه', 'هر ۱۰'),
        ('کل برگشتی', 'EVENT_TYPE=برگشتی (قاعده: ایستگاهِ برگشت از فروش)',
         'COUNTIFS(…, P:"برگشتی")', 'عدد', 'ماه', 'هر ۱۰'),
        ('قطعات دارای عیب', 'شمارِ متمایزِ PART_ID در فیلتر', 'SUMPRODUCT(COUNTIFS…)', 'عدد',
         'ماه', 'هر ۱۰'),
        ('انواع عیبِ ثبت‌شده', 'شمارِ متمایزِ DEFECT_ID', 'همان روی ستون M', 'عدد', 'ماه', 'هر ۱۰'),
        ('سهم عیب (Pareto)', 'عیب ÷ کلِ همان فیلتر', 'C/ΣC', '٪', 'ماه/قطعه', 'هر ۱۰'),
        ('درصدِ تجمعی', 'Σ سهم با مرتب‌سازیِ نزولی (نزولی در build انجام شد)', 'SUM(بالا)', '٪',
         'ماه', 'هر ۱۰'),
        ('نرخِ رکوردِ تکراری', 'ردیف‌های مشکوک-تکراری ÷ کل', '478/14557 = 3.3٪', '٪', 'هفته', '—'),
        ('نرخِ ضایعات بر تولید', '❌ قابل محاسبه نیست', 'مخرجِ کسر (تعدادِ تولید) در هیچ جای فایل نیست',
         '٪', '—', '—'),
        ('تأخیرِ ثبت', '❌ حذف شد', 'رمزِ بارکد (رقم ۹–۱۴) تاریخ تولید نبود؛ تا تأییدِ واحدِ فرآیند ساخته نمی‌شود',
         'روز', '—', '—'),
    ]
    for rec in kpi_rows:
        for j, v in enumerate(rec, start=2):
            c = ws.cell(row=r, column=j, value=v)
            c.alignment, c.font, c.border = WRAP, F_TXT, BORDER
            if v.startswith('❌'):
                c.fill = FILL_WARN
        r += 1
    r += 2
    ws.cell(row=r, column=2, value='محدودیت‌های صریحِ Excel (ادعا نه)').font = F_TITLE
    r += 1
    limits = [
        'همزمانیِ چند کاربر روی یک فایل: پشتیبانی‌نشده (قفلِ فایل/کپیِ ذخیره‌شده). '
        'معماریِ تأییدشده: «فایل به‌ازای ایستگاه» + تجمیعِ مرکزی با Power Query (Append).',
        'Excel دیتابیس نیست: کلیدِ خارجیِ واقعی، تراکنش، rollback و constraint ندارد. '
        'EVENT_ID از شمارشِ ردیف ساخته می‌شود و در نوشتنِ همزمان می‌تواند تکراری شود.',
        'قفلِ شیت/ساختار با رمزِ Excel = بازدارندگی، نه امنیت. رمزها با ابزارهای رایگان در چند '
        'دقیقه بازیابی می‌شوند؛ امنیتِ واقعی در SQL/PostgreSQL/Web است.',
        'Audit trailِ غیرقابل‌انکار وجود ندارد: LOG_AUDIT best-effort است و کاربرِ آگاه می‌تواند '
        'آن را ویرایش کند. برای ممیزیِ قانونی → DB با trigger.',
        'سقفِ عملیاتیِ پیشنهادی: ~۲۵٬۰۰۰ ردیفِ فعال در سال (با فرمول‌های کرانه‌دار، رفرش ~۱–۲ ثانیه). '
        'سقفِ فنی ۱٬۰۴۸٬۵۷۶ ردیف است، ولی «سنگینیِ فایل» خیلی زودتر از آن آزار می‌دهد.',
        'Pivot/Slicer/Timeline روی این نسخه فعال نشده: Timeline روی تاریخِ شمسی کار نمی‌کند و '
        'refresh در فایلِ اشتراکی منبعِ تعارض است. اگر خواستید، فقط در فایلِ مرکزی افزوده می‌شود.',
        'تقویم: تاریخِ شمسی «متنِ ۱۰ کاراکتری» است و پلیِ آن جدولِ تقویمِ Helper است — بدونِ '
        'هیچ وابستگیِ به locale ویندوز (با jdatetime روی ۱۰۹٬۹۳۸ روز راستی‌آزمایی شد).',
        'رمزِ بارکد: رقم ۱–۲ کدِ قطعه است ولی به این فهرست (P-01..P-35) نمی‌خورد (۴۲٪)؛ رقم ۴ با '
        'حرفِ حفره ۹۴٪ هم‌خوان است ⇒ هشدارِ DQ، نه ردِ ورودی.',
    ]
    for t in limits:
        c = ws.cell(row=r, column=2, value='• ' + t)
        c.alignment, c.font = WRAP, F_TXT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 34
        r += 1
    r += 1
    ws.cell(row=r, column=2, value='چه زمانی از Excel برویم بیرون؟').font = F_TITLE
    r += 1
    for t in ['بیش از ۳ ایستگاه همزمان، یا بیش از ~۲۵٬۰۰۰ رکورد در سال → PostgreSQL/SQL Server + '
              'وب‌اپلیکیشن (یا Power Apps) و Excel فقط برای گزارش.',
              'نیاز به امضا، NCR/CAR، ضمیمهٔ عکس، یا گزارش‌گیریِ لحظه‌ایِ مدیریتی → Hybrid: '
              'ثبت در Web، گزارش در Power BI، Excel برای تحلیلِ موردی.',
              'نیاز به «چه کسی، چه وقت، چه چیزی را عوض کرد» قابل‌استناد → دیتابیس با audit trigger.']:
        c = ws.cell(row=r, column=2, value='→ ' + t)
        c.alignment, c.font = WRAP, F_TXT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 30
        r += 1
    # ── بلوکِ پیکربندیِ ماکرو (QC_SHEETS / QC_MSG) ────────────────────────
    # ماژولِ VBA عمداً فقط ASCII است (VBE یک ویرایشگرِ ANSI است و فارسیِ درونِ کد،
    # هنگامِ Import خراب می‌شود). پس نامِ شیت‌ها و متنِ پیام‌ها را از این دو جدول
    # می‌خواند: هر رشتهٔ فارسی که کاربر می‌بیند، از خودِ فایل می‌آید.
    r += 2
    ws.cell(row=r, column=2, value='QC_SHEETS — نامِ شیت‌ها (ماکرو از این‌جا می‌خواند)').font = F_TITLE
    r += 1
    c1 = r                       # این دو بلوک سطرِ سرستون ندارند؛ بازه از همین سطر است
    for key, val in (('DATA', SH['data']), ('FORM', SH['form']), ('STAGING', SH['staging']),
                     ('MASTER', SH['master']), ('HELPER', SH['helper']), ('LOG', SH['log']),
                     ('DASH', SH['dash']), ('HELP', SH['help'])):
        ws.cell(row=r, column=2, value=key).number_format = '@'
        ws.cell(row=r, column=2).font = F_LBL
        ws.cell(row=r, column=3, value=val).number_format = '@'
        ws.cell(row=r, column=3).font = F_TXT
        r += 1
    ref_sheets = f"'{SH['set']}'!$B${c1}:$C${r - 1}"
    ws.cell(row=r + 1, column=2, value='QC_MSG — متنِ پیام‌هایِ ماکرو (قابل ویرایش)').font = F_TITLE
    r += 2
    c2 = r
    for key, val in QC_MESSAGES.items():
        ws.cell(row=r, column=2, value=key).number_format = '@'
        ws.cell(row=r, column=2).font = F_LBL
        c = ws.cell(row=r, column=3, value=val)
        c.number_format, c.alignment, c.font = '@', WRAP, F_TXT
        ws.row_dimensions[r].height = 30
        r += 1
    ref_msg = f"'{SH['set']}'!$B${c2}:$C${r - 1}"
    for col, w in zip('BCDEFGH', [26, 42, 40, 10, 18, 20, 20]):
        ws.column_dimensions[col].width = w
    ws.sheet_properties.tabColor = 'FFBFBFBF'
    return {'QC_SHEETS': ref_sheets, 'QC_MSG': ref_msg}


def build_help(ws, pl):
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'راهنمایِ کاربر — داخلِ فایل، نه پیوستِ جدا'
    ws['B2'].font = F_TITLE
    txt = [
        ('۱) ثبتِ یک رخداد', 'در «FORM فرم ثبت» هشت فیلد را از لیست انتخاب کنید و بارکد را '
         'اسکن کنید. پیامِ سطرِ ۲۸ باید «✓ ورودی کامل است» شود؛ سپس QC_Append را بزنید '
         '(دکمهٔ ثبت، یا Alt+Q اگر میان‌بر بسته‌اید). '
         'بدونِ ماکرو: همان ۹ مقدار را در «STAGING ورود اضطراری» وارد کنید.'),
        ('۲) اگر پیامِ خطا داد', '«⚠ ناقص: …» نامِ فیلدهایِ خالی را می‌گوید. «بارکد…» یعنی '
         '۱۹ رقم نیست (فاصله، حرف، یا صفرِ خورده‌شده). «بارکد قبلاً ثبت شده» یعنی تکرار — '
         'ردیفِ قبلی را «ابطال‌شده» کنید، حذف نکنید.'),
        ('۳) اصلاح و ابطال', 'در «Data داده» با فیلترِ سرستون ردیف را پیدا کنید و ROW_STATUS را '
         'ابطال‌شده/پیش‌نویس کنید. Delete هرگز: آمارِ روزانه و شمارشِ رکورد به همان ردیف '
         'استناد می‌کند.'),
        ('۴) قطعه/عیب/بازرسِ جدید', 'فقط در «Master اطلاعات پایه»، داخلِ همان Excel Table. برای '
         'از رده خارج‌کردن، ACTIVE را «خیر» کنید. حذفِ ردیف، دادهٔ تاریخی را یتیم می‌کند '
         '(اسنیپ‌شاتِ نام در Data هست، ولی گزارشِ Master بهم می‌خورد).'),
        ('۵) گزارش و چاپ', 'فیلترها را از «Dashboard داشبورد» (G4:G13) تنظیم کنید. «Reports '
         'گزارش‌ها» ۱۵ گزارش دارد. برای خروجیِ مدیریتی: Dashboard → Print (Print Area از پیش '
         'تنظیم: A4 افقی، یک صفحه).'),
        ('۶) تاریخ', 'تاریخِ شمسی در فایل «متن ۱۰ کاراکتری» است (۱۴۰۵/۰۶/۲۸) و ستونِ میلادی '
         'برای محاسبهٔ بازه. پلِ این دو، جدولِ تقویمِ شیت Helper است؛ تنظیماتِ ویندوز '
         'هیچ نقشی ندارد.'),
        ('۷) چند کاربر', 'هر ایستگاه فایلِ خودش را داشته باشد. تجمیع در فایلِ مرکزی با Power '
         'Query (Append) انجام می‌شود. بازکردنِ همزمانِ این فایل توسط دو نفر = ریسکِ از '
         'دست رفتنِ داده؛ Excel «سیستمِ آنلاینِ چندکاربره» نیست و وانمود نمی‌کند.'),
        ('۸) پشتیبان', 'کپیِ روزانه در پوشه‌ای با Version History (OneDrive/SharePoint). فایل را '
         'فقط روی هاردِ خودتان نگه ندارید؛ خرابیِ .xlsm در Excel غیرغیرقابل‌بازگشت نیست ولی '
         'ارزان‌ترین بیمه، همان کپی است.'),
        ('۹) نصبِ ماکرو (یک‌بار، ۲ دقیقه)', 'Alt+F11 → منوی File → Import File → فایلِ '
         'vba/QC_WritePath.bas را انتخاب کنید → برگردید به Excel و فایل را با پسوندِ .xlsm '
         'ذخیره کنید. روی شیتِ فرم یک Rectancle بکشید، راست‌کلیک → Assign Macro → QC_Append. '
         'ماکرو فقط می‌نویسد (ثبت/Import/ابطال); هیچ محاسبه‌ای در آن نیست و هیچ ردیفی را حذف '
         'نمی‌کند. پوشهٔ فایل را Trusted Location کنید تا هشدارِ ماکرو تکراری نشود.'),
        ('۱۰) اگر ماکرو بلاک شد', 'هیچ داده‌ای از دست نمی‌رود: فایلِ .xlsx بدونِ ماکرو هم '
         'گزارش‌ها و داشبورد را نشان می‌دهد (فرمول‌محور است، نه Pivot با refresh). '
         'ورودِ داده از STAGING انجام می‌شود.'),
    ]
    r = 4
    for t, b in txt:
        ws.cell(row=r, column=2, value=t).font = F_LBL
        c = ws.cell(row=r + 1, column=2, value=b)
        c.alignment, c.font = WRAP, F_TXT
        ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 3, end_column=9)
        r += 5
    ws.column_dimensions['B'].width = 26
    for c in 'CDEFGHI':
        ws.column_dimensions[c].width = 20
    ws.sheet_properties.tabColor = 'FF70AD47'


def build_log(ws):
    ws['A1'] = 'LOG_AUDIT — best-effort. این شیت ادعای ممیزیِ قانونی ندارد (بخش G گزارش طراحی را ببینید).'
    ws['A1'].font = F_LBL
    hdr_row(ws, 2, 1, ['EVENT_ID', 'ACTION', 'USER', 'MACHINE', 'TIMESTAMP', 'FIELD', 'OLD', 'NEW'])
    for j, w in enumerate([12, 12, 14, 14, 18, 16, 24, 24], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for r in range(3, 203):
        for j in range(1, 9):
            ws.cell(row=r, column=j).number_format = '@' if j == 1 else 'General'
    ws.sheet_state = 'hidden'
    ws.protection.sheet = True


def build_verify(ws, pl, model, n_rows):
    ws['B2'] = ('Verify — پلِ «مدلِ مهاجرت» با «فرمولِ فایل». این شیت پنهان است و فقط '
                'برای آزمون/مذاکره ساخته شده؛ در فایلِ تحویلی می‌تواند بماند.')
    ws['B2'].font = F_LBL
    r = hdr_row(ws, 4, 2, ['معیار', 'مقدارِ مدل', 'توضیح'])
    rows = [('n_fact_rows', n_rows, 'تعدادِ ردیفِ واقعیِ tblQC'),
            ('n_total_model', model['total'], 'COUNTA از payload'),
            ('scrap', model['ضایعات'], 'ضایعات'), ('rework', model['اصلاحی'], 'اصلاحی'),
            ('return', model['برگشتی'], 'برگشتی'),
            ('distinct_part', model['distinct_part'], 'قطعات متمایز'),
            ('distinct_defect', model['distinct_defect'], 'عیوب متمایز'),
            ('dup_rows', pl['stats']['dup_extra_rows'], 'ردیف‌هایِ تکراری (زائد)'),
            ('leading_zero', pl['stats']['leading_zero'], 'بارکدهای با صفرِ ابتدا'),
            ('dq_issues', pl['stats']['dq_issues'], 'تعدادِ ایرادِ ثبت‌شده'),
            ('mold_na-khan', pl['stats']['mold_mismatch'], 'MOLD_CHECK = ناهم‌خوان')]
    for k, v, t in rows:
        ws.cell(row=r, column=2, value=k).font = F_LBL
        ws.cell(row=r, column=3, value=v)
        ws.cell(row=r, column=4, value=t)
        r += 1
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 40
    ws.sheet_state = 'hidden'


# ───────────────────── مونتاژ ─────────────────────
def build(path, pl):
    wb = openpyxl.Workbook()
    ws_data = wb.active
    ws_data.title = SH['data']
    sheets = {}
    for k in ('form', 'staging', 'master', 'helper', 'reports', 'dash', 'dq', 'set', 'help',
              'log', 'verify'):
        sheets[k] = wb.create_sheet(SH[k])
    n_rows, cap = build_data(ws_data, pl)
    build_master(sheets['master'], pl)
    lists, cal_last, params, grids = build_helper(sheets['helper'], pl, n_rows, cap)
    helper_name = SH['helper']
    build_form(sheets['form'], pl, lists, helper_name)
    build_staging(sheets['staging'], pl, lists, helper_name)
    build_reports(sheets['reports'], pl, n_rows)
    model, filt = build_dashboard(sheets['dash'], pl, lists, helper_name, n_rows, cap, grids)
    build_dq(sheets['dq'], pl)
    set_refs = build_settings(sheets['set'], pl)
    build_help(sheets['help'], pl)
    build_log(sheets['log'])
    build_verify(sheets['verify'], pl, model, n_rows)
    # نام‌هایِ تعریف‌شده (منبعِ لیست‌ها — برای DV و برای ویرایشِ بعدی)
    for nm, (col, n) in lists.items():
        ref = (f"'{helper_name}'!${get_column_letter(col)}$2:"
               f"${get_column_letter(col)}${max(n + 1, 2)}")
        wb.defined_names.add(DefinedName(nm, attr_text=ref))
    for nm, ref in set_refs.items():
        wb.defined_names.add(DefinedName(nm, attr_text=ref))
    pc = lists['PARAM'][0]
    for i, (k, _v) in enumerate(params, start=2):
        if k in ('N_ROWS', 'CAP_ROWS', 'N_DQ'):
            wb.defined_names.add(DefinedName(
                k, attr_text=f"'{helper_name}'!${get_column_letter(pc + 1)}${i}"))
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
    wb.properties.title = 'سامانۀ ثبت قطعات برگشتی، ضایعات و اصلاحی — نسخهٔ ۲'
    wb.properties.creator = 'QC-F-14 v2 (bazsazi-ye mohandesi-shode)'
    wb.properties.description = (f'migrated from {pl["source"]}; {pl["stats"]["fact_rows"]} records; '
                                 f'built {pl["generated"]}')
    wb.save(path)
    return n_rows, cap, model


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--payload', default='analysis/payload.json')
    ap.add_argument('-o', '--out', default='QC-F-14-v2.xlsx')
    a = ap.parse_args()
    with open(a.payload, encoding='utf-8') as fh:
        pl = json.load(fh)
    n_rows, cap, model = build(a.out, pl)
    print(f'OK: {a.out} | ردیفِ داده: {n_rows} (ظرفیت تا سطرِ {cap}) | '
          f'ضایعات/اصلاحی/برگشتی: {model["ضایعات"]}/{model["اصلاحی"]}/{model["برگشتی"]} | '
          f'{os.path.getsize(a.out) / 1e6:.2f} MB')


if __name__ == '__main__':
    main()
