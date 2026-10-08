import openpyxl,sys
wb=openpyxl.load_workbook(sys.argv[1],data_only=True)
for s in sys.argv[2].split(','):
    ws=wb[s]; print('==',s)
    for r in ws.iter_rows(min_row=3,values_only=True):
        if r[0] and r[0]!='ID' and len(str(r[0]))<=9 and not str(r[0])[0] in '一二三四五六':
            v=[ (round(x,3) if isinstance(x,float) else x) for x in r[3:9]]
            print(r[0], (r[1] or '')[:28], v if s!='Checks' else r[2])
