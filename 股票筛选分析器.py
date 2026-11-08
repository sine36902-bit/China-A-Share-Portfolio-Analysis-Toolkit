"""本地教学分析；Python 3.11+；pip install pandas numpy openpyxl matplotlib。无网络调用。"""
from pathlib import Path
import re, json, math, threading, queue, tempfile, os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import numpy as np

DEFAULTS=[('601398','工商银行','银行业'),('601288','农业银行','银行业'),('601988','中国银行','银行业'),('601939','建设银行','银行业'),('601328','交通银行','银行业'),('600036','招商银行','银行业'),('601166','兴业银行','银行业'),('002371','北方华创','半导体行业'),('002049','紫光国微','半导体行业'),('600584','长电科技','半导体行业'),('600703','三安光电','半导体行业')]
METHODS=[('收益率','相邻参考交易日均有有效价格且均未确认停牌时 r=P/P前-1；不补值，不重复加股息。'),('平均与风险','历史日算术平均；样本方差、协方差均除以n-1；标准差为方差平方根。'),('累计与年化','累计=P末/P初-1；年化=(P末/P初)^(365.2425/实际天数)-1；采用区间有效端点，不连乘删缺失后的收益率。年化按几何平均思想计算，365.2425是本程序日计数约定。'),('配对','仅在共同有效日期上计算双方均值、样本方差、标准差、协方差和ρ。'),('解释','相关性较低说明历史线性联动较弱，不代表收益最高、组合风险最低或已经找到最优投资组合。'),('筛选用途','仅缩小研究范围；最终结合截至2023年底已公开的公司基本面资料决定，不自动选定股票。'),('公式符号','rₜ是日简单收益率，r̄是样本均值，n是有效样本数，s是样本标准差；Σ对有效日期求和。配对表的两股均值、风险、协方差与ρ全部基于同一组共同日期；年度表限于该年。'),('宽表说明','日收益率宽表保留首行日期及纯股票代码，数值为小数；采用 rₜ=Pₜ/Pₜ₋₁−1，不跨缺口。'),('空值','少于2样本不计算样本风险；零标准差相关系数留空。空白不是零。'),('清洗提示','日涨跌绝对值超过20%仅提示；不删极值、不缩尾。缺口原因待核实，不推断停牌。缺失比例超过5%提示可比性不足。'),('边界','没有前一参考交易日价格则首日收益率为空；临时日期并集无法发现所有股票同时缺失的交易日。')]


def code_str(x):
    s=str(x).strip().removesuffix('.0')
    return s.zfill(6) if re.fullmatch(r'\d{1,6}',s) else ''


def read_csv(p):
    for enc in ['utf-8-sig','gb18030']:
        try:return pd.read_csv(p,dtype=str,encoding=enc)
        except UnicodeDecodeError:pass
    raise ValueError(f'无法识别CSV编码：{p}')


def parse_dates(s):
    def one(v):
        if pd.isna(v):return pd.NaT
        if isinstance(v,(int,float)):
            return pd.Timestamp('1899-12-30')+pd.Timedelta(days=float(v)) if 20000<float(v)<100000 else pd.NaT
        try:return pd.to_datetime(str(v).strip(),errors='raise').normalize()
        except Exception:return pd.NaT
    return s.map(one)


def import_files(paths):
    sources=[]
    for path in paths:
        p=Path(path)
        if p.suffix.lower()=='.csv':
            tables={p.stem:read_csv(p)}
            side=p.parent/'下载说明.csv'
            meta=read_csv(side) if side.exists() else pd.DataFrame()
        else:
            tables=pd.read_excel(p,sheet_name=None,dtype=object)
            meta=tables.get('下载说明',pd.DataFrame())
        for sheet,df in tables.items():
            if sheet=='下载说明':continue
            if '日期' not in df:raise ValueError(f'{p.name}/{sheet} 没有“日期”列')
            m=re.search(r'(股票|指数)[_：:]([0-9]{6})(?:_(.*))?',sheet)
            c=m[2] if m else ''
            kind=m[1] if m else '股票'
            name=m[3] if m else ''
            if not c:
                match=re.search(r'(?<!\d)(\d{6})(?!\d)',sheet)
                c=match[1] if match else ''
            if not c and '代码' in df:
                codes=df['代码'].map(code_str).unique()
                if len(codes)==1:c=codes[0]
            if not c:raise ValueError(f'{p.name}/{sheet} 无法识别代码；请命名为 股票_六位代码_名称 或加入唯一的“代码”列。')
            rows=meta.copy()
            if '代码' in rows:
                rows=rows[rows['代码'].map(code_str)==c]
                if '证券类型' in rows:rows=rows[rows['证券类型'].astype(str)==kind]
            else:rows=pd.DataFrame()
            md=rows.iloc[0].fillna('').to_dict() if len(rows)==1 else {}
            name=str(md.get('证券名称',name or c))
            basis=str(md.get('复权或指数口径','未知'))
            if basis not in ['后复权','前复权','不复权']:basis='未知'
            cols={}
            for col in df:
                plain=re.sub(r'[（(].*','',str(col))
                if plain in ['收盘','收盘价']:cols[col]='价格'
                if plain in ['最高','最高价']:cols[col]='最高'
                if plain in ['最低','最低价']:cols[col]='最低'
                if plain in ['停牌','停牌状态']:cols[col]='停牌'
            df=df.rename(columns=cols)
            if '价格' not in df and kind=='股票':raise ValueError(f'{sheet} 没有收盘价列')
            df=df.copy();df['来源文件']=str(p.resolve());df['来源工作表']=sheet;df['原始行号']=np.arange(len(df))+2
            df['原始日期']=df['日期'];df['日期']=parse_dates(df['日期'])
            sources.append(dict(code=c,name=name,kind=kind,basis=basis,period=str(md.get('行情周期','未知')),df=df,meta=md,path=str(p.resolve()),sheet=sheet,manual=''))
    return sources


def clean_stock(code,sources,calendar,start,end,resolutions):
    raw=pd.concat([s['df'] for s in sources],ignore_index=True)
    logs=[];invalid=0;duplicates=0;conflicts=0
    def log(row,reason):
        logs.append({'代码':code,'日期':row.get('原始日期',''),'来源文件':row.get('来源文件',''),'来源工作表':row.get('来源工作表',''),'原始行号':row.get('原始行号',''),'原始价格':str(row.get('价格','')),'原因':reason})
    values={}; flags={}; origins={};halts={}
    for _,row in raw[raw['日期'].isna()].iterrows():log(row,'日期无法解析，移入异常表')
    valid=raw[raw['日期'].notna()]
    for d,group in valid.groupby('日期',sort=True):
        group=group.copy(); signatures=[]; kept=[]
        for idx,row in group.iterrows():
            sig=tuple(str(row.get(k,'')) for k in ['价格','最高','最低','停牌'])
            if sig in signatures:duplicates+=1;log(row,'完全重复记录，删除')
            else:signatures.append(sig);kept.append(idx)
        group=group.loc[kept]
        vals=pd.to_numeric(group['价格'],errors='coerce')
        bad=~np.isfinite(vals) | (vals<=0)
        invalid+=int(bad.sum())
        for _,row in group[bad].iterrows():log(row,'无效价格：空白、非数值、无穷或非正')
        notes=[]
        # 非完全相同记录一律保守列为冲突，避免高低价或停牌字段互相矛盾。
        chosen=group.iloc[0]
        price=float(vals.iloc[0]) if not bad.iloc[0] else np.nan
        if len(group)>1:
            conflicts+=1;notes.append('重复日期冲突')
            for _,r in group.iterrows():log(r,'冲突候选：默认该日期缺失')
            rid=resolutions.get((code,str(d.date())))
            matches=group[group.apply(lambda r:f"{r['来源文件']}|{r['来源工作表']}|{r['原始行号']}"==rid,axis=1)] if rid else pd.DataFrame()
            if len(matches)==1:
                chosen=matches.iloc[0];price=pd.to_numeric(chosen['价格'],errors='coerce')
                if not np.isfinite(price) or price<=0:price=np.nan
                notes.append('用户明确选择记录');log(chosen,'用户明确保留此冲突候选')
            else:price=np.nan
        if pd.isna(price):notes.append('价格无效或冲突')
        high=pd.to_numeric(chosen.get('最高',np.nan),errors='coerce');low=pd.to_numeric(chosen.get('最低',np.nan),errors='coerce')
        if (pd.notna(high) and pd.notna(low) and high<low) or (pd.notna(price) and ((pd.notna(high) and price>high) or (pd.notna(low) and price<low))):
            notes.append('高低收盘关系异常，仅提示');log(chosen,notes[-1])
        halt=str(chosen.get('停牌','')).strip().lower() in ['是','停牌','true','1','1.0']
        if halt:notes.append('已确认停牌，排除相关单日收益')
        values[d]=price;flags[d]='；'.join(notes);halts[d]=halt
        origins[d]='；'.join(f"{r['来源文件']}|{r['来源工作表']}|行{r['原始行号']}" for _,r in group.iterrows())
    outside=set(values)-set(calendar)
    for d in sorted(outside):
        for _,row in valid[valid['日期']==d].iterrows():log(row,'不在参考交易日历内，未纳入日收益计算')
    detail=pd.DataFrame(index=calendar)
    detail['价格']=pd.Series(values,dtype=float);detail['已确认停牌']=pd.Series(halts).reindex(calendar,fill_value=False)
    detail['来源记录']=pd.Series(origins);detail['异常标记']=pd.Series(flags).reindex(calendar).fillna('缺少日期记录，原因待核实')
    effective=detail['价格'].mask(detail['已确认停牌'])
    detail['日收益率']=effective/effective.shift(1)-1
    detail['日收益率']=detail['日收益率'].replace([np.inf,-np.inf],np.nan)
    for d in detail.index[detail['日收益率'].abs()>0.2]:detail.loc[d,'异常标记']+='；极端波动超过20%，仅提示'
    detail.loc[detail['日收益率'].isna(),'异常标记']+='；当日或前一参考交易日无有效价格/停牌，收益率留空'
    detail=detail.loc[start:end].copy();detail.index.name='日期'
    missing=detail['价格'].isna();run=0;longest=0
    for v in missing:run=run+1 if v else 0;longest=max(longest,run)
    prices=detail['价格'].dropna()
    summary={'代码':code,'原始记录数':len(raw),'清洗后有效价格数':len(prices),'有效收益率数':detail['日收益率'].count(),'重复记录数':duplicates,'冲突日期数':conflicts,'缺失日期数':sum(d not in values for d in detail.index),'无效价格记录数':invalid,'最大连续价格缺口':longest,'参考交易日数':len(detail),'缺失价格比例':missing.mean() if len(detail) else np.nan,'实际开始日期':str(prices.index.min().date()) if len(prices) else '', '实际结束日期':str(prices.index.max().date()) if len(prices) else '', '质量提示':'缺失超过5%，样本可比性不足' if len(detail) and missing.mean()>.05 else '不代表完整十年样本','统计范围说明':'有效价格/收益/缺口为分析区间；原始、重复、冲突及无效记录数为全部导入记录'}
    return detail,summary,logs


def check_data(sources,roster,settings,calendar=None,resolutions=None):
    if not any(i['选中'] for i in roster.values()):raise ValueError('请至少勾选一只股票')
    start,end=pd.Timestamp(settings['开始']),pd.Timestamp(settings['结束'])
    if start>end:raise ValueError('开始日期不能晚于结束日期')
    if settings['行业一']==settings['行业二']:raise ValueError('请选择两个不同的行业')
    if calendar is None:
        indices=[s for s in sources if s['kind']=='指数' and s['code']=='000300']
        pool=indices or [s for s in sources if s['kind']=='股票']
        calendar=pd.DatetimeIndex(sorted(set(d for s in pool for d in s['df']['日期'].dropna())))
        calendar_note='沪深300导入日期（仍需核查覆盖范围）' if indices else '导入股票日期并集：不能识别所有股票同时缺失的交易日'
    else:calendar_note='用户导入本地交易日历'
    calendar=pd.DatetimeIndex(calendar).dropna().normalize().unique().sort_values()
    calendar=calendar[calendar<=end]
    if len(calendar)==0:raise ValueError('没有参考交易日期')
    details={};summaries=[];logs=[];status=[]
    for code,info in roster.items():
        if not info['选中']:continue
        ss=[s for s in sources if s['code']==code and s['kind']=='股票']
        state='已匹配';reason=''
        if not ss:state='缺少数据'
        elif any(s['basis']=='未知' or s['period']=='未知' for s in ss):state='口径待确认';reason='在“来源口径”中逐项手动指定复权和日线，并记录理由'
        elif any(s['basis']!=settings['复权'] for s in ss):state='口径不一致';reason='不得混用；请取消勾选或导入一致口径的数据'
        elif any(s['period']!='日线' for s in ss):state='非日线';reason='本程序只计算日收益率'
        else:
            d,q,l=clean_stock(code,ss,calendar,start,end,resolutions or {})
            q.update({'名称':info['名称'],'行业':info['行业']});summaries.append(q);logs.extend(l)
            if d['价格'].count()==0:state='无有效价格'
            else:details[code]=d
        status.append({'代码':code,'名称':info['名称'],'行业':info['行业'],'状态':state,'说明':reason})
    count1=sum(roster[c]['行业']==settings['行业一'] for c in details)
    count2=sum(roster[c]['行业']==settings['行业二'] for c in details)
    return dict(details=details,quality=pd.DataFrame(summaries),logs=pd.DataFrame(logs),status=pd.DataFrame(status),calendar_note=calendar_note,pairs=count1*count2,roster=roster,settings=settings,sources=sources)


def stats(r):
    r=r.dropna();n=len(r)
    return {'样本数':n,'日算术平均收益率':r.mean() if n else np.nan,'日方差':r.var(ddof=1) if n>=2 else np.nan,'日标准差':r.std(ddof=1) if n>=2 else np.nan}


def analyze(check):
    details=check['details'];roster=check['roster'];cfg=check['settings'];annual=[];individual=[];pairs=[]
    for c,d in details.items():
        base={'代码':c,'名称':roster[c]['名称'],'行业':roster[c]['行业']};s=dict(base,**stats(d['日收益率']));p=d['价格'].dropna()
        s.update({'实际开始日期':p.index[0].date(),'实际结束日期':p.index[-1].date(),'累计收益率':p.iloc[-1]/p.iloc[0]-1 if len(p)>1 else np.nan,'年化几何平均收益率':(p.iloc[-1]/p.iloc[0])**(365.2425/(p.index[-1]-p.index[0]).days)-1 if len(p)>1 else np.nan})
        individual.append(s)
        for year in range(pd.Timestamp(cfg['开始']).year,pd.Timestamp(cfg['结束']).year+1):annual.append(dict(base,年份=year,**stats(d.loc[d.index.year==year,'日收益率'])))
    first=[c for c in details if roster[c]['行业']==cfg['行业一']];second=[c for c in details if roster[c]['行业']==cfg['行业二']]
    for a in first:
        for b in second:
            common=pd.concat([details[a]['日收益率'].rename('甲'),details[b]['日收益率'].rename('乙')],axis=1).dropna();n=len(common);x=stats(common['甲']);y=stats(common['乙'])
            cov=common['甲'].cov(common['乙']) if n>=2 else np.nan
            rho=cov/(x['日标准差']*y['日标准差']) if n>=2 and x['日标准差']>0 and y['日标准差']>0 else np.nan
            row={'股票一':a,'名称一':roster[a]['名称'],'股票二':b,'名称二':roster[b]['名称'],'共同样本数':n,'共同开始日期':common.index.min().date() if n else '', '共同结束日期':common.index.max().date() if n else '', '相关系数ρ':rho,'协方差':cov,'说明':'样本不足或零方差时相应指标留空；历史共同样本比较'}
            for suffix,st in [('一',x),('二',y)]:
                for key,val in st.items():row[key+suffix]=val
            pairs.append(row)
    pair=pd.DataFrame(pairs)
    tables={'分析设置与方法说明':pd.DataFrame(list(cfg.items())+[('交易日历',check['calendar_note']),('实际配对数',len(pairs))]+METHODS,columns=['项目','说明']),'个股收益和风险汇总':pd.DataFrame(individual),'分年度风险统计':pd.DataFrame(annual),'数据质量摘要':check['quality'],'匹配状态':check['status'],'清洗异常日志':check['logs']}
    if pair.empty:
        pair=pd.DataFrame(columns=['股票一','名称一','股票二','名称二','共同样本数','相关系数ρ','协方差'])
    ranked=pair.sort_values('相关系数ρ',kind='stable',na_position='last').reset_index(drop=True)
    ranked.insert(0,'排名',ranked['相关系数ρ'].rank(method='min'))
    tables['配对筛选排名']=ranked
    tables['相关性最低前五组']=ranked[ranked['相关系数ρ'].notna()].head(5).copy()
    tables['待进一步研究的组合']=ranked.iloc[:0].copy()
    for name,col in [('跨行业相关系数矩阵','相关系数ρ'),('跨行业协方差矩阵','协方差')]:
        matrix=pair.pivot(index='股票一',columns='股票二',values=col).reindex(index=first,columns=second)
        matrix.index=[f'{c} {roster[c]["名称"]}' for c in matrix.index]
        matrix.columns=[f'{c} {roster[c]["名称"]}' for c in matrix.columns]
        matrix.index.name='股票代码与名称';tables[name]=matrix.reset_index()
    wide=pd.DataFrame({c:d['日收益率'] for c,d in details.items()})
    wide.index.name='日期';tables['日收益率宽表']=wide.reset_index()
    tables['股票代码与名称对应']=pd.DataFrame([{'代码':c,'名称':roster[c]['名称'],'行业':roster[c]['行业'],'复权方式':cfg['复权'],'研究开始':cfg['开始'],'研究结束':cfg['结束']} for c in details])
    meta=[]
    for src in check['sources']:meta.append({'代码':src['code'],'名称':src['name'],'证券类型':src['kind'],'来源文件':src['path'],'工作表':src['sheet'],'确认复权':src['basis'],'确认周期':src['period'],'手动确认记录':src['manual'],'原始元数据':json.dumps(src['meta'],ensure_ascii=False,default=str)})
    tables['原始来源和复权说明']=pd.DataFrame(meta)
    for c,d in details.items():tables[f'明细_{c}_{roster[c]["名称"]}'[:31]]=d.reset_index()
    return tables


def metric_formula(sheet, column):
    """文本公式说明，不是Excel计算公式；不改变底层数值和数据接口。"""
    if sheet == '跨行业相关系数矩阵' and column != '股票代码与名称':
        return 'ρ₁₂ = Cov(r₁,r₂)/(s₁s₂)；每格使用该两股共同有效日期'
    if sheet == '跨行业协方差矩阵' and column != '股票代码与名称':
        return 'Cov(r₁,r₂) = Σ[(r₁ₜ−r̄₁)(r₂ₜ−r̄₂)]/(n−1)；每格共同样本'
    if column == '相关系数ρ':return 'ρ₁₂ = Cov(r₁,r₂)/(s₁s₂)；任一s=0时无法计算'
    if column == '协方差':return 'Cov(r₁,r₂) = Σ[(r₁ₜ−r̄₁)(r₂ₜ−r̄₂)]/(n−1)'
    if '方差' in column:return 's² = Σ(rₜ−r̄)²/(n−1)；n≥2'
    if '标准差' in column:return 's = √s²；n≥2'
    if '算术平均收益率' in column:return 'r̄ = Σrₜ/n；只计有效日收益率'
    if column == '累计收益率':return 'R = P末/P初−1；采用区间内有效价格端点'
    if column == '年化几何平均收益率':return 'g = (P末/P初)^(1/T)−1；T=实际持有天数/365.2425'
    if column == '日收益率':return 'rₜ = Pₜ/Pₜ₋₁−1；相邻参考交易日均有效，不跨缺口'
    if column == '共同样本数':return 'n = 两股均有有效日收益率的共同日期数'
    if column.startswith('样本数'):return 'n = 有效日收益率个数；配对表只用共同日期'
    return ''


def export_excel(tables,path):
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.utils import get_column_letter
    path=Path(path);fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=path.parent);os.close(fd)
    try:
        with pd.ExcelWriter(tmp,engine='openpyxl',datetime_format='yyyy-mm-dd',date_format='yyyy-mm-dd') as writer:
            for name,df in tables.items():
                name=re.sub(r'[\\/*?:\[\]]','_',name)[:31]
                formulas=[metric_formula(name,str(col)) for col in df.columns]
                # 宽表作为程序间接口，首行仍是“日期＋纯代码”，不插入说明行。
                formula_row=any(formulas) and name!='日收益率宽表'
                header_row=2 if formula_row else 1
                df.to_excel(writer,sheet_name=name,index=False,startrow=header_row-1)
                ws=writer.sheets[name]
                ws.freeze_panes=f'A{header_row+1}'
                if len(df.columns):ws.auto_filter.ref=f'A{header_row}:{get_column_letter(len(df.columns))}{max(ws.max_row,header_row)}'
                if formula_row:
                    ws.row_dimensions[1].height=90
                    for j,formula in enumerate(formulas,1):
                        cell=ws.cell(1,j,formula or ('计算说明' if j==1 else ''))
                        cell.data_type='s';cell.font=Font(size=10,color='24476B')
                        cell.fill=PatternFill('solid',fgColor='EAF1F8')
                        cell.alignment=Alignment(wrap_text=True,vertical='center')
                for cell in ws[header_row]:
                    cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='24476B')
                    cell.alignment=Alignment(wrap_text=True,vertical='center')
                ws.row_dimensions[header_row].height=32
                for j,column in enumerate(ws.columns):
                    header=str(column[header_row-1].value)
                    width=42 if any(x in header for x in ['说明','标记','来源','原因','元数据']) else 22
                    if formula_row and formulas[j]:width=32
                    ws.column_dimensions[column[0].column_letter].width=width
                    for cell in column[header_row:]:
                        if isinstance(cell.value,str):cell.data_type='s'
                        if isinstance(cell.value,(float,int)):
                            if '方差' in header or name=='跨行业协方差矩阵':cell.number_format='0.00000000E+00'
                            elif any(x in header for x in ['收益率','标准差','比例']):cell.number_format='0.0000%'
                            elif '相关系数' in header or name=='跨行业相关系数矩阵':cell.number_format='0.000000'
                        cell.alignment=Alignment(vertical='top')
        os.replace(tmp,path)
    finally:
        if os.path.exists(tmp):os.unlink(tmp)


def heatmap_paths(excel_path):
    p=Path(excel_path)
    return [p.with_name(p.stem+'_'+name+'.png') for name in ['相关系数热力图','协方差热力图']]


def export_heatmaps(tables,excel_path):
    # 使用非交互绘图对象，可在后台线程导出，不干扰Tk窗口。
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
    from matplotlib import font_manager, rc_context
    available={f.name for f in font_manager.fontManager.ttflist}
    fonts=[f for f in ['Microsoft YaHei','PingFang SC','Heiti TC','Heiti SC','Arial Unicode MS','SimHei','Noto Sans CJK SC','WenQuanYi Micro Hei'] if f in available]
    if not fonts:raise ValueError('未找到中文字体。请安装思源黑体（Noto Sans CJK SC）后重试；Excel可单独保留。')
    settings=tables['分析设置与方法说明'].set_index('项目')['说明'].to_dict()
    palette=LinearSegmentedColormap.from_list('蓝白红',['#24588E','#F8FAFC','#BD4D56'])
    palette.set_bad('#E2E7ED')
    saved=[]
    with rc_context({'font.family':fonts[0],'axes.unicode_minus':False}):
        for sheet,path in zip(['跨行业相关系数矩阵','跨行业协方差矩阵'],heatmap_paths(excel_path)):
            matrix=tables[sheet].set_index(tables[sheet].columns[0])
            if matrix.empty:continue
            values=matrix.to_numpy(dtype=float);finite=values[np.isfinite(values)]
            correlation='相关系数' in sheet
            limit=1.0 if correlation else (float(np.max(np.abs(finite))) if finite.size else 1.0)
            if limit==0:limit=1e-8
            fig=Figure(figsize=(max(9,2.0*len(matrix.columns)+3),max(6.5,.65*len(matrix)+3)),dpi=150,facecolor='#F5F7FA')
            FigureCanvasAgg(fig)
            ax=fig.add_axes([.21,.22,.61,.59]);ax.set_facecolor('#F5F7FA')
            im=ax.imshow(np.ma.masked_invalid(values),cmap=palette,norm=TwoSlopeNorm(vmin=-limit,vcenter=0,vmax=limit),aspect='auto')
            def label(v):
                parts=str(v).split(' ',1)
                return parts[1]+'\n'+parts[0] if len(parts)==2 else str(v)
            ax.set_xticks(range(len(matrix.columns)),[label(c) for c in matrix.columns],fontsize=10)
            ax.set_yticks(range(len(matrix.index)),[label(c) for c in matrix.index],fontsize=10)
            ax.tick_params(axis='both',length=0,pad=10)
            ax.set_xticks(np.arange(-.5,len(matrix.columns),1),minor=True)
            ax.set_yticks(np.arange(-.5,len(matrix),1),minor=True)
            ax.grid(which='minor',color='white',linewidth=2);ax.tick_params(which='minor',length=0)
            for spine in ax.spines.values():spine.set_visible(False)
            for i in range(len(matrix)):
                for j in range(len(matrix.columns)):
                    value=values[i,j]
                    text='缺失' if not np.isfinite(value) else (f'{value:.3f}' if correlation else f'{value:.2e}')
                    ax.text(j,i,text,ha='center',va='center',fontsize=10,color='white' if np.isfinite(value) and abs(value)/limit>.62 else '#24364B')
            cb=fig.colorbar(im,cax=fig.add_axes([.86,.22,.025,.59]))
            cb.set_label('相关系数ρ' if correlation else '协方差（日收益率）',fontsize=10,labelpad=10)
            cb.outline.set_visible(False)
            if not correlation:
                from matplotlib.ticker import ScalarFormatter
                formatter=ScalarFormatter(useMathText=True);formatter.set_powerlimits((0,0));cb.formatter=formatter;cb.update_ticks()
            group1={'银行业':'银行','半导体行业':'半导体'}.get(str(settings.get('行业一','银行业')),str(settings.get('行业一','银行业')))
            group2={'银行业':'银行','半导体行业':'半导体'}.get(str(settings.get('行业二','半导体行业')),str(settings.get('行业二','半导体行业')))
            title=f'{group1}与{group2}股票收益率的'+('相关系数' if correlation else '协方差')
            fig.text(.08,.94,title,fontsize=20,fontweight='bold',color='#19334D')
            fig.text(.08,.885,f"{settings.get('开始','')} — {settings.get('结束','')}  |  {settings.get('复权','')}  |  {settings.get('行业一','')} × {settings.get('行业二','')}",fontsize=11,color='#526478')
            source_names=[]
            for raw in tables.get('原始来源和复权说明',pd.DataFrame()).get('原始元数据',pd.Series(dtype=str)):
                try:
                    source=json.loads(str(raw)).get('数据来源','')
                    if source and source not in source_names:source_names.append(source)
                except (ValueError,TypeError):pass
            source_text='、'.join(source_names) or '本地导入文件（详见Excel来源说明）'
            fig.text(.08,.845,'数据来源：'+source_text+'；不填补缺失值；灰色表示无法计算。',fontsize=9,color='#526478')
            years=f"{str(settings.get('开始','2014'))[:4]}—{str(settings.get('结束','2023'))[:4]}"
            fig.text(.08,.09,f'根据{years}年两只股票共同有效交易日的日收益率计算。',fontsize=10,color='#526478')
            fig.text(.08,.055,'相关系数较低，不代表投资组合风险一定最低。',fontsize=10,color='#526478')
            fd,tmp=tempfile.mkstemp(suffix='.png',dir=path.parent);os.close(fd)
            try:fig.savefig(tmp,dpi=300,facecolor=fig.get_facecolor());os.replace(tmp,path)
            finally:
                if os.path.exists(tmp):os.unlink(tmp)
                fig.clear()
            saved.append(str(path))
    return saved


def export_bundle(tables,path):
    export_excel(tables,path)
    try:
        images=export_heatmaps(tables,path)
        return '导出成功：'+str(path)+'；热力图：'+('；'.join(images) if images else '没有跨行业配对，未生成')
    except Exception as exc:
        return 'Excel已保存：'+str(path)+'；热力图导出失败：'+str(exc)+'（请检查matplotlib依赖及中文字体）'


class App:
    def __init__(self,root):
        self.root=root;root.title('股票筛选分析器（本地数据）');root.geometry('1100x820')
        self.sources=[];self.calendar=None;self.resolutions={};self.checked=None;self.tables=None;self.q=queue.Queue();self.busy=False
        self.roster={c:{'名称':n,'行业':g,'选中':True} for c,n,g in DEFAULTS}
        bar=ttk.Frame(root);bar.pack(fill='x',padx=10,pady=8)
        for label,cmd in [('选择数据文件',self.load),('来源口径确认',self.basis),('导入本地交易日历',self.load_calendar),('添加股票',self.add),('删除股票',self.delete),('修改行业',self.industry),('冲突记录选择',self.resolve)]:ttk.Button(bar,text=label,command=cmd).pack(side='left',padx=2)
        ttk.Label(root,text='双击名单切换勾选。加载/编辑后必须重新“检查数据”；未知复权及周期必须手动确认。').pack(anchor='w',padx=12)
        self.tree=ttk.Treeview(root,columns=['选中','代码','名称','行业','状态'],show='headings',height=12)
        for col in self.tree['columns']:self.tree.heading(col,text=col);self.tree.column(col,width=170)
        self.tree.pack(fill='x',padx=10);self.tree.bind('<Double-1>',self.toggle)
        form=ttk.Frame(root);form.pack(fill='x',padx=10,pady=8)
        self.v={k:tk.StringVar(value=v) for k,v in {'开始':'2014-01-01','结束':'2023-12-31','行业一':'银行业','行业二':'半导体行业','复权':'后复权'}.items()}
        for k in self.v:
            ttk.Label(form,text=k).pack(side='left')
            if k=='复权':ttk.Combobox(form,textvariable=self.v[k],values=['后复权','前复权','不复权'],width=9,state='readonly').pack(side='left')
            else:ttk.Entry(form,textvariable=self.v[k],width=12).pack(side='left',padx=3)
            self.v[k].trace_add('write',lambda *_:self.invalidate())
        buttons=ttk.Frame(root);buttons.pack(fill='x',padx=10)
        for label,cmd in [('检查数据',self.check),('开始分析',self.run),('查看筛选结果',self.show_results),('导出Excel及热力图',self.export)]:ttk.Button(buttons,text=label,command=cmd).pack(side='left',padx=5)
        self.status=tk.StringVar(value='请选择本地Excel或CSV文件，先检查，再确认分析。');ttk.Label(root,textvariable=self.status,wraplength=1060).pack(anchor='w',padx=12,pady=8)
        panel=ttk.Frame(root);panel.pack(fill='both',expand=True,padx=10,pady=8)
        self.text=tk.Text(panel,height=18,wrap='none');self.text.grid(row=0,column=0,sticky='nsew')
        sy=ttk.Scrollbar(panel,command=self.text.yview);sy.grid(row=0,column=1,sticky='ns')
        sx=ttk.Scrollbar(panel,orient='horizontal',command=self.text.xview);sx.grid(row=1,column=0,sticky='ew')
        self.text.configure(yscrollcommand=sy.set,xscrollcommand=sx.set);panel.rowconfigure(0,weight=1);panel.columnconfigure(0,weight=1)
        self.refresh();root.after(100,self.poll)
    def invalidate(self):self.checked=None;self.tables=None;self.research=set()
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        states={r['代码']:r['状态'] for r in self.checked['status'].to_dict('records')} if self.checked else {}
        for c,i in self.roster.items():self.tree.insert('', 'end',iid=c,values=['✓' if i['选中'] else '',c,i['名称'],i['行业'],states.get(c,'待检查' if any(s['code']==c for s in self.sources) else '缺少数据')])
    def task(self,fn,done):
        if self.busy:return
        self.busy=True;self.status.set('处理中，请稍候……')
        def worker():
            try:self.q.put((done,fn(),None))
            except Exception as e:self.q.put((None,None,str(e)))
        threading.Thread(target=worker,daemon=True).start()
    def poll(self):
        try:
            cb,val,error=self.q.get_nowait();self.busy=False
            if error:self.status.set('操作失败：'+error);messagebox.showerror('错误',error)
            else:cb(val)
        except queue.Empty:pass
        self.root.after(100,self.poll)
    def load(self):
        if self.busy:return
        paths=filedialog.askopenfilenames(filetypes=[('行情文件','*.xlsx *.csv')])
        if not paths:return
        def done(ss):
            keys={(s['path'],s['sheet']) for s in ss};self.sources=[s for s in self.sources if (s['path'],s['sheet']) not in keys]+ss
            for s in ss:
                if s['kind']=='股票':
                    if s['code'] not in self.roster:self.roster[s['code']]={'名称':s['name'],'行业':'未分组','选中':True}
                    else:self.roster[s['code']]['名称']=s['name']
            self.invalidate();self.refresh();self.status.set(f'已导入 {len(self.sources)} 个来源。请编辑名单、确认口径，然后检查数据。')
        self.task(lambda:import_files(paths),done)
    def toggle(self,event=None):
        if self.busy:return
        for c in self.tree.selection():self.roster[c]['选中']=not self.roster[c]['选中']
        self.invalidate();self.refresh()
    def add(self):
        if self.busy:return
        value=simpledialog.askstring('添加股票','输入六位代码，以逗号、空格或换行分隔')
        if not value:return
        codes=re.split(r'[,，\s]+',value.strip())
        if any(not re.fullmatch(r'\d{6}',c) for c in codes):messagebox.showerror('错误','请使用六位代码');return
        group=simpledialog.askstring('行业','输入行业名称')
        if not group:return
        for c in codes:
            name=next((s['name'] for s in self.sources if s['code']==c and s['kind']=='股票'),c)
            self.roster[c]={'名称':name,'行业':group,'选中':True}
        self.invalidate();self.refresh()
    def delete(self):
        if self.busy:return
        for c in self.tree.selection():self.roster.pop(c,None)
        self.invalidate();self.refresh()
    def industry(self):
        if self.busy:return
        group=simpledialog.askstring('修改行业','输入选中行的新行业')
        if group:
            for c in self.tree.selection():self.roster[c]['行业']=group
            self.invalidate();self.refresh()
    def basis(self):
        if self.busy:return
        win=tk.Toplevel(self.root);win.title('逐来源核对口径');win.geometry('900x450')
        box=tk.Listbox(win,width=120);box.pack(fill='both',expand=True)
        for i,s in enumerate(self.sources):box.insert('end',f'{i+1}. {Path(s["path"]).name}/{s["sheet"]} | {s["basis"]} | {s["period"]}')
        def edit():
            if not box.curselection():return
            s=self.sources[box.curselection()[0]]
            value=simpledialog.askstring('手动确认','填写复权方式：后复权、前复权或不复权。必须依据源文件或提供者说明，不能猜测。',parent=win)
            if value not in ['后复权','前复权','不复权']:return
            reason=simpledialog.askstring('确认依据','请填写确认依据，并确认原始数据为日线；若无法确认请取消。',parent=win)
            if not reason:return
            s['basis']=value;s['period']='日线';s['manual']=reason+'；用户确认日线；'+str(pd.Timestamp.now());self.invalidate();self.refresh();win.destroy()
        ttk.Button(win,text='修改选中来源的口径并记录依据',command=edit).pack()
    def load_calendar(self):
        if self.busy:return
        p=filedialog.askopenfilename(filetypes=[('本地交易日历','*.csv *.xlsx')])
        if not p:return
        try:
            df=read_csv(p) if p.lower().endswith('.csv') else pd.read_excel(p)
            self.calendar=pd.DatetimeIndex(parse_dates(df['日期']).dropna())
            if '是否交易日' in df:self.calendar=pd.DatetimeIndex(parse_dates(df.loc[df['是否交易日'].astype(str).isin(['1','True','是','交易日']),'日期']).dropna())
            self.invalidate();self.status.set('已载入本地交易日历；若无“是否交易日”列，则每行都必须是真实交易日。')
        except Exception as e:messagebox.showerror('日历错误',str(e))
    def resolve(self):
        if self.busy:return
        c=simpledialog.askstring('冲突记录','输入股票六位代码');d=simpledialog.askstring('冲突日期','日期 YYYY-MM-DD')
        if not c or not d:return
        win=tk.Toplevel(self.root);win.title('选择明确保留的原始记录');box=tk.Listbox(win,width=140);box.pack();ids=[]
        for s in self.sources:
            if s['code']!=c:continue
            for _,r in s['df'].iterrows():
                if pd.notna(r['日期']) and str(r['日期'].date())==d:
                    rid=f"{r['来源文件']}|{r['来源工作表']}|{r['原始行号']}";ids.append(rid);box.insert('end',f'{rid} | 收盘={r.get("价格")}')
        def choose():
            if box.curselection():self.resolutions[(c,d)]=ids[box.curselection()[0]];self.invalidate();win.destroy()
        ttk.Button(win,text='明确保留选中记录（否则维持缺失）',command=choose).pack()
    def check(self):
        if self.busy:return
        try:
            cfg={k:v.get() for k,v in self.v.items()}
            for k in ['开始','结束']:
                if not re.fullmatch(r'\d{4}-\d{2}-\d{2}',cfg[k]):raise ValueError('日期格式 YYYY-MM-DD')
                pd.Timestamp(cfg[k])
            roster={c:dict(i) for c,i in self.roster.items()}
        except Exception as e:messagebox.showerror('设置错误',str(e));return
        def done(result):
            self.checked=result;self.tables=None;self.refresh();self.text.delete('1.0','end');self.text.insert('end',result['status'].to_string(index=False)+'\n\n'+result['quality'].to_string(index=False)+'\n\n'+result['logs'].to_string(index=False));self.status.set(f'检查完成：实际有效股票 {len(result["details"])} 只，可生成 {result["pairs"]} 组；{result["calendar_note"]}。请检查清洗摘要后开始分析。')
        self.task(lambda:check_data(self.sources,roster,cfg,self.calendar,self.resolutions),done)
    def run(self):
        if self.busy:return
        if self.checked:
            try:
                changed=any(str(self.checked['settings'][k]) != v.get() for k,v in self.v.items())
            except ValueError:changed=True
            if changed:self.invalidate()
        if not self.checked:messagebox.showwarning('请先检查','设置变更后请重新检查数据');return
        if self.checked['status']['状态'].isin(['口径待确认','口径不一致','非日线']).any():messagebox.showerror('口径未通过','请逐来源确认口径，或取消勾选不一致股票后重新检查。');return
        if not self.checked['details']:messagebox.showerror('没有数据','没有可用股票');return
        warning='将分析实际可用股票，缺少数据股票会跳过。\n'+f'实际配对数：{self.checked["pairs"]}。\n'
        if self.checked['settings']['复权']=='不复权':warning+='注意：不复权仅为价格变化率，不是包含分红调整的收益率！\n'
        if not messagebox.askyesno('确认清洗与分析',warning+'已核对清洗摘要，继续？'):return
        def done(t):
            self.tables=t;self.research=set();self.status.set('筛选分析完成，可比较排序并勾选待进一步研究的配对。');self.show_results()
        self.task(lambda:analyze(self.checked),done)
    def update_research(self):
        if self.tables is None:return
        pairs=self.tables['配对筛选排名']
        keys=getattr(self,'research',set())
        mask=[(row['股票一'],row['股票二']) in keys for _,row in pairs.iterrows()]
        self.tables['待进一步研究的组合']=pairs.loc[mask].copy()

    def show_results(self):
        if self.tables is None:
            messagebox.showwarning('尚未分析','请先检查数据并开始分析');return
        snapshot=self.tables
        win=tk.Toplevel(self.root);win.title('比较收益、风险与相关性');win.geometry('1200x760')
        ttk.Label(win,text='相关性较低说明历史线性联动较弱，不代表收益最高、组合风险最低或已经找到最优投资组合。',wraplength=1150).pack(anchor='w',padx=8,pady=5)
        ttk.Label(win,text='筛选仅用于缩小范围，最终需结合截至2023年底已公开的基本面资料。双击任意配对可勾选/取消待研究。').pack(anchor='w',padx=8)
        tabs=ttk.Notebook(win);tabs.pack(fill='both',expand=True)
        def table(parent,df):
            frame=ttk.Frame(parent);frame.pack(fill='both',expand=True)
            tree=ttk.Treeview(frame,columns=list(df.columns),show='headings')
            tree.grid(row=0,column=0,sticky='nsew')
            for col in df:
                tree.heading(col,text=col);tree.column(col,width=130,stretch=False)
            sx=ttk.Scrollbar(frame,orient='horizontal',command=tree.xview);sx.grid(row=1,column=0,sticky='ew')
            sy=ttk.Scrollbar(frame,command=tree.yview);sy.grid(row=0,column=1,sticky='ns')
            tree.configure(xscrollcommand=sx.set,yscrollcommand=sy.set)
            frame.rowconfigure(0,weight=1);frame.columnconfigure(0,weight=1)
            return tree
        def fill(tree,df):
            tree.delete(*tree.get_children())
            for i,(_,row) in enumerate(df.iterrows()):
                vals=[]
                for col,value in row.items():
                    if pd.isna(value):text='无法计算' if col=='相关系数ρ' else ''
                    elif isinstance(value,(float,np.floating)):
                        if '方差' in col:text=f'{value:.8e}'
                        elif '收益率' in col or '标准差' in col:text=f'{value:.4%}'
                        else:text=f'{value:.6g}'
                    else:text=str(value)
                    vals.append(text)
                tree.insert('','end',iid=str(i),values=vals)
        one=ttk.Frame(tabs);tabs.add(one,text='个股比较')
        opts=ttk.Frame(one);opts.pack(fill='x')
        metric=tk.StringVar(value='日算术平均收益率');order=tk.StringVar(value='降序')
        ttk.Label(opts,text='排序指标').pack(side='left')
        ttk.Combobox(opts,textvariable=metric,values=['日算术平均收益率','年化几何平均收益率','日方差','日标准差'],state='readonly').pack(side='left')
        ttk.Combobox(opts,textvariable=order,values=['升序','降序'],state='readonly',width=8).pack(side='left')
        individual=snapshot['个股收益和风险汇总'];it=table(one,individual)
        def sort_one():fill(it,individual.sort_values(metric.get(),ascending=order.get()=='升序',na_position='last',kind='stable'))
        ttk.Button(opts,text='排序',command=sort_one).pack(side='left');sort_one()
        five=ttk.Frame(tabs);tabs.add(five,text='相关性最低前5组')
        top=snapshot['相关性最低前五组'];fill(table(five,top),top)
        alltab=ttk.Frame(tabs);tabs.add(alltab,text='全部配对与勾选')
        pairs=snapshot['配对筛选排名'];display=pairs.copy();display.insert(0,'待研究','')
        pt=table(alltab,display)
        chosen=ttk.Frame(tabs);tabs.add(chosen,text='待进一步研究')
        ct=table(chosen,pairs)
        def refresh_pairs():
            display['待研究']=['✓' if (r['股票一'],r['股票二']) in getattr(self,'research',set()) else '' for _,r in pairs.iterrows()]
            fill(pt,display);self.update_research();fill(ct,self.tables['待进一步研究的组合'])
        def toggle(event):
            if self.tables is not snapshot:
                messagebox.showwarning('结果已过期','设置或数据已更新，请关闭此窗口后重新分析。',parent=win);return
            item=pt.identify_row(event.y)
            if not item:return
            row=pairs.iloc[int(item)];key=(row['股票一'],row['股票二'])
            if key in self.research:self.research.remove(key)
            else:self.research.add(key)
            refresh_pairs()
        pt.bind('<Double-1>',toggle);refresh_pairs()

    def export(self):
        if self.busy:return
        if self.tables is None:messagebox.showwarning('尚未分析','先检查数据并开始分析');return
        p=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile='股票候选筛选结果.xlsx',filetypes=[('Excel','*.xlsx')],confirmoverwrite=True)
        if not p:return
        if Path(p).resolve() in {Path(s['path']).resolve() for s in self.sources}:messagebox.showerror('禁止覆盖','不能覆盖原始输入文件');return
        existing=[str(x) for x in heatmap_paths(p) if x.exists()]
        if existing and not messagebox.askyesno('确认覆盖热力图','以下图片已存在，是否覆盖？\n'+'\n'.join(existing)):return
        self.update_research()
        self.task(lambda:export_bundle(self.tables,p),lambda result:self.status.set(result))

if __name__=='__main__':
    root=tk.Tk();App(root);root.mainloop()
