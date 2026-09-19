"""payload v4 = v3 + USERS table for login"""
import argparse, collections, datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlcal, xlmigrate as M

FACT_COLUMNS_V4 = [
    ('RECORD_ID','QC-00000001'),
    ('RECORD_TYPE','INSPECTION/REINSPECTION/RETURN'),
    ('EVENT_DATE_J','1405/06/28'),
    ('EVENT_DATE_G','تاریخ میلادی'),
    ('J_MONTH_KEY','140506'),
    ('J_WEEK_KEY','1405-W26'),
    ('SHIFT_ID','S1'),
    ('STATION_ID','ST-07'),
    ('STATION_SNAP','نام ایستگاه'),
    ('INSPECTOR_ID','I-01'),
    ('INSPECTOR_SNAP','نام بازرس'),
    ('PART_ID','P-01'),
    ('PART_NAME_SNAP','نام قطعه'),
    ('MOLD_ID','M-01'),
    ('MOLD_NAME_SNAP','نام قالب'),
    ('BARCODE','19 رقمی Text'),
    ('BC_PART_CODE','کد قطعه بارکد'),
    ('MOLD_CHECK','هم‌خوان/ناهم‌خوان'),
    ('INSPECTION_RESULT','OK/NOK'),
    ('DEFECT_ID','D-01'),
    ('DEFECT_NAME_SNAP','نام عیب'),
    ('DISPOSITION','ACCEPT/REWORK/SCRAP/RETURN_TO_PROCESS/HOLD'),
    ('CURRENT_STATUS','OPEN/.../VOID'),
    ('ENTRY_USER','کاربر'),
    ('ENTRY_STAMP','زمان ثبت'),
    ('ENTRY_CHANNEL','کانال'),
    ('ROW_STATUS','legacy'),
    ('NOTE','توضیح'),
    ('VOID_FLAG','TRUE/FALSE'),
    ('VOID_REASON','دلیل ابطال'),
    ('EVENT_TYPE','legacy'),
    ('DISPOSITION_ID','legacy'),
]

MAX_FACT_ROWS=20000
MOLD_CHECK_OPTIONS=['هم‌خوان','ناهم‌خوان','قالبِ جهت','رقمِ خارج از بازه','بارکد نامعتبر']
ROW_STATUS_OPTIONS=['تأییدشده','پیش‌نویس','ابطال‌شده','مشکوک-تکراری']
RECORD_TYPE_OPTIONS=['INSPECTION','REINSPECTION','RETURN']
RESULT_OPTIONS=['OK','NOK']
DISPOSITION_OPTIONS=['ACCEPT','REWORK','SCRAP','RETURN_TO_PROCESS','HOLD']
STATUS_OPTIONS=['OPEN','UNDER_REWORK','WAITING_REINSPECTION','CLOSED','VOID']
ROLE_OPTIONS=['ADMIN','OPERATOR','VIEWER']

REPORTS_V4 = [
    dict(id='R01', title='روند روزانه (۳۰ روز اخیر)', group='J_DAY', axis='EVENT_DATE_J', chart='line', block='روند'),
    dict(id='R02', title='روند ماهانه — تفکیک نوع رخداد', group='J_MONTH_KEY', axis='J_MONTH_KEY', chart='line', block='روند'),
    dict(id='R03', title='روند هفتگی', group='J_WEEK_KEY', axis='J_WEEK_KEY', chart='line', block='روند'),
    dict(id='R04', title='شیفت × نوع رکورد', group='SHIFT_ID', axis='SHIFT_ID', chart='stacked', block='تولید'),
    dict(id='R05', title='ایستگاه × نوع رکورد', group='STATION_ID', axis='STATION_ID', chart='stacked', block='تولید'),
    dict(id='R06', title='۱۵ قطعه برتر', group='PART_ID', axis='PART_ID', chart='bar', block='قطعه', top=15),
    dict(id='R07', title='Pareto عیب', group='DEFECT_ID', axis='DEFECT_ID', chart='pareto', block='عیب', top=15),
    dict(id='R08', title='قالب × نوع', group='MOLD_ID', axis='MOLD_ID', chart='stacked', block='قطعه'),
    dict(id='R09', title='بازرس × تعداد', group='INSPECTOR_ID', axis='INSPECTOR_ID', chart='bar', block='بازرس', top=21),
    dict(id='R10', title='Disposition', group='DISPOSITION', axis='DISPOSITION', chart='bar', block='تولید'),
    dict(id='R11', title='Barcode History', group='BARCODE', axis='BARCODE', chart=None, block='بارکد', special='barcode'),
    dict(id='R12', title='ماتریس قطعه × عیب', group='PART_X_DEFECT', axis=None, chart='matrix', block='عیب', special='matrix'),
    dict(id='R13', title='تکراری/عقب‌گرد', group='ENTRY_VS_EVENT', axis=None, chart=None, block='کیفیت', special='backdate'),
    dict(id='R14', title='کیفیت داده', group='DQ_RULE', axis=None, chart='bar', block='کیفیت', special='dq'),
    dict(id='R15', title='خلاصه مدیریتی', group='J_MONTH_KEY', axis='J_MONTH_KEY', chart='stacked', block='مدیریتی', special='mgmt'),
]

KPI_V4 = [
    dict(id='K1', name='کل رکوردها', unit='عدد', agg='count'),
    dict(id='K2', name='بارکد یکتا', unit='عدد', agg='distinct_barcode'),
    dict(id='K3', name='NOK', unit='عدد', agg='count'),
    dict(id='K4', name='OK', unit='عدد', agg='count'),
    dict(id='K5', name='Rework', unit='عدد', agg='count'),
    dict(id='K6', name='Scrap', unit='عدد', agg='count'),
    dict(id='K7', name='Return', unit='عدد', agg='count'),
]

FILTERS_V4=['J_YEAR','J_MONTH_KEY','EVENT_DATE_J','J_WEEK_KEY','SHIFT_ID','STATION_ID','PART_ID','MOLD_ID','DEFECT_ID','RECORD_TYPE','DISPOSITION','CURRENT_STATUS','INSPECTION_RESULT']

def name_of(dims,kind,code):
    for row in dims[kind]:
        if row['CODE']==code:
            return row['NAME']
    return '—'

def build(ref_path, cache=None, verbose=True):
    if cache and os.path.exists(cache):
        import pickle
        with open(cache,'rb') as f:
            ref=pickle.load(f)
    else:
        ref=M.load_reference(ref_path)
        if cache:
            os.makedirs(os.path.dirname(cache), exist_ok=True)
            import pickle
            with open(cache,'wb') as f:
                pickle.dump(ref,f,protocol=4)
    dims=M.build_masters(ref['masters'])
    rows,dq,stats_raw,part_mold,station_status,_=M.analyze(ref,dims)
    stats=dict(stats_raw)

    dims['RECORDTYPE']=[
        {'CODE':'RT-01','NAME':'INSPECTION','ACTIVE':'بله','DESC':'بازرسی اولیه'},
        {'CODE':'RT-02','NAME':'REINSPECTION','ACTIVE':'بله','DESC':'بازرسی مجدد'},
        {'CODE':'RT-03','NAME':'RETURN','ACTIVE':'بله','DESC':'برگشت از فروش'},
    ]
    dims['RESULT']=[
        {'CODE':'RES-01','NAME':'OK','ACTIVE':'بله'},
        {'CODE':'RES-02','NAME':'NOK','ACTIVE':'بله'},
    ]
    dims['DISPOSITION_NEW']=[
        {'CODE':'DP-01','NAME':'ACCEPT','ACTIVE':'بله','LEGACY':''},
        {'CODE':'DP-02','NAME':'REWORK','ACTIVE':'بله','LEGACY':'اصلاحی'},
        {'CODE':'DP-03','NAME':'SCRAP','ACTIVE':'بله','LEGACY':'ضایعات'},
        {'CODE':'DP-04','NAME':'RETURN_TO_PROCESS','ACTIVE':'بله','LEGACY':'برگشتی'},
        {'CODE':'DP-05','NAME':'HOLD','ACTIVE':'بله','LEGACY':''},
    ]
    dims['STATUS_NEW']=[
        {'CODE':'STT-01','NAME':'OPEN','ACTIVE':'بله'},
        {'CODE':'STT-02','NAME':'UNDER_REWORK','ACTIVE':'بله'},
        {'CODE':'STT-03','NAME':'WAITING_REINSPECTION','ACTIVE':'بله'},
        {'CODE':'STT-04','NAME':'CLOSED','ACTIVE':'بله'},
        {'CODE':'STT-05','NAME':'VOID','ACTIVE':'بله'},
    ]
    # USERS table - default users
    dims['USER']=[
        {'CODE':'U-01','USERNAME':'admin','PASSWORD':'admin123','ROLE':'ADMIN','DISPLAY_NAME':'مدیر سیستم','ACTIVE':'بله'},
        {'CODE':'U-02','USERNAME':'operator','PASSWORD':'1234','ROLE':'OPERATOR','DISPLAY_NAME':'بازرس خط','ACTIVE':'بله'},
        {'CODE':'U-03','USERNAME':'viewer','PASSWORD':'viewer123','ROLE':'VIEWER','DISPLAY_NAME':'مدیر تولید','ACTIVE':'بله'},
        # Add inspector-based users (username = inspector code lower, password = 1234)
        # We'll add first 5 inspectors as operators
    ]
    # Add inspector users
    for ins in dims.get('INSPECTOR',[])[:5]:
        code=ins['CODE']
        uname=code.lower()
        # avoid duplicate admin
        if uname in ('admin','operator','viewer'):
            continue
        dims['USER'].append({'CODE':f'U-{len(dims["USER"])+1:02d}','USERNAME':uname,'PASSWORD':'1234','ROLE':'OPERATOR','DISPLAY_NAME':ins['NAME'],'ACTIVE':'بله'})

    disp_map={}
    for r in dims['DISPOSITION']:
        n=r['NAME']
        if n=='اصلاحی':
            disp_map[r['CODE']]='REWORK'
        elif n=='ضایعات':
            disp_map[r['CODE']]='SCRAP'
        elif n=='برگشتی':
            disp_map[r['CODE']]='RETURN_TO_PROCESS'
        else:
            disp_map[r['CODE']]='HOLD'

    rule_station=M.event_type_rules(station_status, dims)
    etype_by_station={r['STATION']:r['EVENT_TYPE'] for r in rule_station}
    pick_rule={'DIS-02':'ضایعات','DIS-03':'برگشتی'}

    # Record_Type by barcode history
    sorted_rows=sorted(rows, key=lambda x: (x['G_DATE'] or datetime.date.min, x['_row']))
    first_seen={}
    record_type_by_index={}
    for r in sorted_rows:
        orig_idx=rows.index(r)
        bc=r['BARCODE']
        st=r['STATION']
        if st=='برگشت از فروش':
            rt='RETURN'
        elif bc in first_seen:
            rt='REINSPECTION'
        else:
            rt='INSPECTION'
        record_type_by_index[orig_idx]=rt
        if bc not in first_seen:
            first_seen[bc]=r

    facts=[]
    prev_g=None
    backdate=0
    cavity_of_mold={r['CODE']:r['NAME'] for r in dims['MOLD'] if r['MOLD_KIND']=='حفره'}
    for i,r in enumerate(rows, start=1):
        bc=r['BARCODE']
        sem=M.barcode_semantics(bc)
        g=r['G_DATE']
        if g is None:
            jy=jm=jd=None
            j_text_v=''
            mkey=wkey=''
        else:
            jy,jm,jd=xlcal.g_to_j(g)
            j_text_v=xlcal.j_text(jy,jm,jd)
            mkey=jy*100+jm
            wkey=xlcal.week_key(jy,jm,jd)
        mold_name=name_of(dims,'MOLD',r['MOLD_ID'])
        guess=sem.get('BC_CAVITY_GUESS','')
        cavity=cavity_of_mold.get(r['MOLD_ID'],'')
        if not sem['_ok']:
            mold_check='بارکد نامعتبر'
        elif not cavity:
            mold_check='قالبِ جهت'
        elif not guess:
            mold_check='رقمِ خارج از بازه'
        else:
            mold_check='هم‌خوان' if cavity==guess else 'ناهم‌خوان'
        rec_type=record_type_by_index.get(i-1,'INSPECTION')
        insp_result='NOK'
        disp_new=disp_map.get(r['DISPOSITION_ID'],'HOLD')
        if rec_type=='RETURN':
            disp_new='RETURN_TO_PROCESS'
        curr_status='CLOSED'
        station_snap=name_of(dims,'STATION',r['STATION_ID'])
        event_type_legacy=(pick_rule.get(r['DISPOSITION_ID']) or etype_by_station.get(station_snap) or 'اصلاحی')
        if prev_g is not None and g is not None and g < prev_g:
            backdate+=1
        if g is not None:
            prev_g=g
        facts.append({
            'RECORD_ID':f'QC-{i:08d}',
            'RECORD_TYPE':rec_type,
            'EVENT_DATE_J':j_text_v,
            'EVENT_DATE_G':g.isoformat() if g else '',
            'J_MONTH_KEY':mkey,
            'J_WEEK_KEY':wkey,
            'SHIFT_ID':r.get('SHIFT_CODE') or f"S{r['SHIFT_RAW']}" if str(r.get('SHIFT_RAW') or '').isdigit() else 'S?',
            'STATION_ID':r['STATION_ID'],
            'STATION_SNAP':station_snap,
            'INSPECTOR_ID':r.get('INSPECTOR_ID') or '—',
            'INSPECTOR_SNAP':name_of(dims,'INSPECTOR',r.get('INSPECTOR_ID')),
            'PART_ID':r['PART_ID'],
            'PART_NAME_SNAP':name_of(dims,'PART',r['PART_ID']),
            'MOLD_ID':r['MOLD_ID'],
            'MOLD_NAME_SNAP':mold_name,
            'BARCODE':bc,
            'BC_PART_CODE':sem['BC_PART_CODE'],
            'MOLD_CHECK':mold_check,
            'INSPECTION_RESULT':insp_result,
            'DEFECT_ID':r['DEFECT_ID'],
            'DEFECT_NAME_SNAP':name_of(dims,'DEFECT',r['DEFECT_ID']),
            'DISPOSITION':disp_new,
            'CURRENT_STATUS':curr_status,
            'ENTRY_USER':'MIGRATION',
            'ENTRY_STAMP':'',
            'ENTRY_CHANNEL':'مهاجرت',
            'ROW_STATUS':'تأییدشده',
            'NOTE':'',
            'VOID_FLAG':'FALSE',
            'VOID_REASON':'',
            'EVENT_TYPE':event_type_legacy,
            'DISPOSITION_ID':r['DISPOSITION_ID'],
            '_J_YEAR':jy,'_J_MONTH':jm,'_J_DAY':jd,'_SRC_ROW':r['_row'],
        })
    by_bc=collections.defaultdict(list)
    for idx,f in enumerate(facts):
        if f['BARCODE']:
            by_bc[f['BARCODE']].append(idx)
    dup_groups=0
    for bc,idxs in by_bc.items():
        if len(idxs)>1:
            dup_groups+=1
            sorted_idxs=sorted(idxs, key=lambda ii: facts[ii]['EVENT_DATE_G'])
            for k,idx in enumerate(sorted_idxs):
                if k==0:
                    continue
                else:
                    facts[idx]['CURRENT_STATUS']='OPEN'
                    facts[idx]['ROW_STATUS']='مشکوک-تکراری'
                    if facts[idx]['RECORD_TYPE']!='RETURN':
                        facts[idx]['RECORD_TYPE']='REINSPECTION'
            dq.append({'RULE':'DQ-DUP','SEV':'متوسط','ROW':facts[idxs[0]]['_SRC_ROW'],'DETAIL':f'بارکد {bc} در {len(idxs)} ردیف'})
    for f in facts:
        if f['RECORD_TYPE']=='RETURN':
            f['DISPOSITION']='RETURN_TO_PROCESS'

    part_mold_c=collections.defaultdict(collections.Counter)
    for r in ref['records']:
        if r.get('_blank'): continue
        part_mold_c[r['PART']][r['MOLD']]+=1
    map_part_mold=M.mold_rules(part_mold_c, dims)

    stats.update(dict(
        fact_rows=len(facts),
        dq_issues=len(dq),
        dup_groups=dup_groups,
        backdate_rows=backdate,
        record_types=dict(collections.Counter(f['RECORD_TYPE'] for f in facts)),
        dispositions=dict(collections.Counter(f['DISPOSITION'] for f in facts)),
        results=dict(collections.Counter(f['INSPECTION_RESULT'] for f in facts)),
        statuses=dict(collections.Counter(f['CURRENT_STATUS'] for f in facts)),
    ))
    code_name={}
    for kind in ('PART','MOLD','DEFECT','INSPECTOR','STATION','DISPOSITION','SHIFT','RECORDTYPE','RESULT','DISPOSITION_NEW','STATUS_NEW'):
        if kind in dims:
            code_name[kind]={r['CODE']:r['NAME'] for r in dims[kind]}
    if 'USER' in dims:
        code_name['USER']={r['CODE']:r['USERNAME'] for r in dims['USER']}

    payload=dict(
        generated=datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        source=os.path.basename(ref_path),
        fact_columns=[c for c,_ in FACT_COLUMNS_V4],
        fact_headers=dict(FACT_COLUMNS_V4),
        facts=facts,
        dims={k:v for k,v in dims.items() if not k.startswith('_')},
        map_part_mold=map_part_mold,
        rule_station=rule_station,
        dq_issues=dq,
        reports=REPORTS_V4,
        kpi=KPI_V4,
        filters=FILTERS_V4,
        mold_check_options=MOLD_CHECK_OPTIONS,
        row_status_options=ROW_STATUS_OPTIONS,
        record_type_options=RECORD_TYPE_OPTIONS,
        result_options=RESULT_OPTIONS,
        disposition_options=DISPOSITION_OPTIONS,
        status_options=STATUS_OPTIONS,
        role_options=ROLE_OPTIONS,
        code_name=code_name,
        stats=stats,
        max_fact_rows=MAX_FACT_ROWS,
        calendar=dict(start=xlcal.CAL_START.isoformat(), end=xlcal.CAL_END.isoformat()),
    )
    if verbose:
        print(f"facts {len(facts)} dup {dup_groups} types {stats['record_types']}")
    return payload

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('src', nargs='?', default='FORM-QC-F-14.xlsm')
    ap.add_argument('-o','--out', default='analysis/payload_v4.json')
    ap.add_argument('--cache', default='analysis/cache/ref.pkl')
    a=ap.parse_args()
    pl=build(a.src, cache=a.cache)
    with open(a.out,'w',encoding='utf-8') as fh:
        json.dump(pl,fh,ensure_ascii=False)
    print(f"written {a.out} {os.path.getsize(a.out)} bytes")

if __name__=='__main__':
    main()
