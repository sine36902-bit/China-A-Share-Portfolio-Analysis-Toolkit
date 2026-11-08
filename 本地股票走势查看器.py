"""Python 3.11+；pip install pandas numpy openpyxl matplotlib。仅本地文件。"""
from pathlib import Path
import os, re, json, tempfile, queue, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib import font_manager, rcParams
import matplotlib.dates as mdates

RANGES={'近1个月':1,'近3个月':3,'近6个月':6,'近1年':12,'近3年':36,'近5年':60}
FREQ={'日':None,'周':'W-SUN','月':'M','年':'Y'}
COLORS=['#2563A6','#D97732','#238B78','#8B5DA7','#C84F68','#A58A26','#3498AC','#6E7B45','#735A4F','#6979B8','#42A1CE','#E29B59','#67AB87','#AF85BE','#D98695','#C5B45C','#72B9BC','#9AA879','#A78B7D','#939FCA']

def date_values(series):
    def parse(x):
        if pd.isna(x):return pd.NaT
        try:
            if isinstance(x,(int,float)):
                return (pd.Timestamp('1899-12-30')+pd.Timedelta(days=x)).normalize() if 20000<x<100000 else pd.NaT
            return pd.Timestamp(str(x)).normalize()
        except Exception:return pd.NaT
    return series.map(parse)

def code(x):
    s=str(x).strip().removesuffix('.0')
    return s.zfill(6) if re.fullmatch(r'\d{1,6}',s) else ''

def read_files(paths):
    result=[]
    for p in paths:
        tables=pd.read_excel(p,sheet_name=None,dtype=object)
        meta=tables.get('下载说明',pd.DataFrame())
        for sheet,frame in tables.items():
            if sheet=='下载说明':continue
            match=re.match(r'股票_(\d{6})_(.+)',sheet)
            if not match:continue
            c,name=match.groups();md={}
            if '代码' in meta:
                rows=meta[meta['代码'].map(code)==c]
                if '证券类型' in rows:rows=rows[rows['证券类型']=='股票']
                if len(rows)==1:md=rows.iloc[0].fillna('').to_dict();name=str(md.get('证券名称',name))
            rename={}
            for col in frame:
                plain=re.sub(r'[（(].*','',str(col))
                if plain in ['收盘','收盘价']:rename[col]='收盘价'
                if plain in ['最高','最高价']:rename[col]='最高价'
                if plain in ['最低','最低价']:rename[col]='最低价'
            frame=frame.rename(columns=rename)
            if '日期' not in frame or '收盘价' not in frame:raise ValueError(f'{sheet} 缺少日期或收盘价')
            frame=frame.copy();frame['原始日期']=frame['日期'];frame['日期']=date_values(frame['日期'])
            frame['来源记录']=[f'{Path(p).resolve()} | {sheet} | 行{i+2}' for i in range(len(frame))]
            for col in ['收盘价','最高价','最低价']:
                if col in frame:
                    frame['原始'+col]=frame[col].astype(str)
                    v=pd.to_numeric(frame[col],errors='coerce');frame[col]=v.where(np.isfinite(v)&(v>0))
            basis=str(md.get('复权或指数口径',md.get('复权方式','未知')))
            if basis not in ['不复权','前复权','后复权']:basis='未知'
            result.append(dict(code=c,name=name,basis=basis,period=str(md.get('行情周期','未知')),source=str(md.get('数据来源','本地文件，来源未注明')),meta=md,frame=frame,path=str(Path(p).resolve()),sheet=sheet,confirmation=''))
    if not result:raise ValueError('没有找到“股票_六位代码_名称”格式的有效工作表')
    return result

def clean(parts):
    raw=pd.concat([p['frame'] for p in parts],ignore_index=True);logs=[];rows=[];duplicates=0;conflicts=0
    def log(r,why):logs.append({'日期':str(r.get('原始日期','')),'来源记录':r['来源记录'],'原因':why,'原始价格':str(r.get('原始收盘价',''))})
    for _,r in raw[raw['日期'].isna()].iterrows():log(r,'日期无法解析，排除')
    fields=[x for x in ['收盘价','最高价','最低价'] if x in raw]
    for day,g in raw.dropna(subset=['日期']).groupby('日期',sort=True):
        unique=g.drop_duplicates(subset=fields);duplicates+=len(g)-len(unique)
        if len(g)>len(unique):log(g.iloc[0],f'完全相同价格记录去重 {len(g)-len(unique)} 条')
        row={'日期':day,'来源记录':'；'.join(g['来源记录']),'异常标记':''}
        if len(unique)>1:
            conflicts+=1;row.update({f:np.nan for f in fields});row['异常标记']='同日价格冲突，保留缺口'
            for _,r in g.iterrows():log(r,'同日价格冲突，未平均或择取；整日价格留空')
        else:
            r=unique.iloc[0];row.update({f:r[f] for f in fields})
            if pd.isna(r['收盘价']):row['异常标记']='无效收盘价，保留缺口';log(r,'无效、非正或缺失收盘价')
            for f in fields:
                if f!='收盘价' and pd.isna(r[f]):log(r,f+'无效或缺失')
            h=r.get('最高价',np.nan);l=r.get('最低价',np.nan);v=r['收盘价']
            if h<l or v>h or v<l:row['异常标记']+='；高低收盘关系异常，仅提示';log(r,'高低收盘关系异常，不修改价格')
        rows.append(row)
    if not rows:raise ValueError('没有可解析日期')
    data=pd.DataFrame(rows).set_index('日期').sort_index()
    return data,logs,{'原始记录数':len(raw),'完全重复记录数':duplicates,'冲突日期数':conflicts,'有效收盘价数':data['收盘价'].count()}

def bounds(latest,earliest,choice,start,end):
    if choice=='自定义':
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}',start) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}',end):raise ValueError('日期格式必须为 YYYY-MM-DD')
        a,b=pd.Timestamp(start),pd.Timestamp(end)
    elif choice=='全部数据':a,b=earliest,latest
    else:a,b=latest-pd.DateOffset(months=RANGES[choice]),latest
    if a>b:raise ValueError('开始日期不能晚于结束日期')
    return a,b

def sample(series,freq):
    if freq=='日':return series.copy()
    # 每个周期保留最后一个有效日期；全空周期保留最后一个缺失日期。
    chosen=[]
    for _,group in series.groupby(series.index.to_period(FREQ[freq])):
        valid=group.dropna();chosen.append(valid.index[-1] if len(valid) else group.index[-1])
    return series.loc[chosen]

def extrema(frame,mode):
    if mode=='关闭':return []
    result=[]
    for col,op,label in ([('收盘价','max','最高收盘价'),('收盘价','min','最低收盘价')] if mode=='收盘价最高最低点' else [('最高价','max','区间最高价'),('最低价','min','区间最低价')]):
        if col not in frame:raise ValueError('缺少最高/最低价列，不能使用区间最高最低价标注')
        values=frame[col].dropna()
        if values.empty:raise ValueError(f'{col}在区间内没有有效数据，无法标注')
        value=values.max() if op=='max' else values.min();days=values.index[values==value]
        result.append({'标注':label,'日期':days[0],'价格':value,'全部出现日期':'、'.join(d.strftime('%Y-%m-%d') for d in days)})
    return result

def prepare(stocks,choice,start,end,freq,mode,calendar=None):
    if not stocks:raise ValueError('请选择股票')
    if len({s['basis'] for s in stocks})!=1:raise ValueError('多股比较必须使用相同复权口径')
    if any(s['basis']=='未知' for s in stocks):raise ValueError('必须先确认复权口径')
    latest=min(s['data'].index.max() for s in stocks);earliest=min(s['data'].index.min() for s in stocks)
    a,b=bounds(latest,earliest,choice,start,end)
    reference=pd.DatetimeIndex(sorted(set(d for s in stocks for d in s['data'].index))) if calendar is None else calendar
    reference=reference[(reference>=a)&(reference<=b)]
    frames=[s['data'].reindex(reference) for s in stocks]
    common=pd.concat([f['收盘价'] for f in frames],axis=1).dropna().index
    multi=len(stocks)>1
    if multi and len(common)<2:raise ValueError('不足两个共同有效日期，无法进行多股比较')
    if multi:
        a,b=common[0],common[-1];frames=[f.loc[a:b] for f in frames]
    items=[]
    for stock,frame in zip(stocks,frames):
        valid=frame['收盘价'].dropna()
        if valid.empty:raise ValueError(stock['name']+'在所选区间内没有有效收盘价')
        first,last=valid.index[0],valid.index[-1];base=valid.iloc[0]
        points=extrema(frame,mode)
        series=frame['收盘价']/base*100 if multi else frame['收盘价']
        sampled=sample(series,freq)
        # 在已知缺失日期插入NaN，使周期抽样也不跨越已知缺口连线。
        plot=sampled.reindex(sampled.index.union(series.index[series.isna()])).sort_index()
        change=valid.iloc[-1]/base-1 if len(valid)>=2 else np.nan
        items.append(dict(stock=stock,frame=frame,display=sampled,plot=plot,points=points,base=base,summary={'代码':stock['code'],'名称':stock['name'],'实际开始':first,'实际结束':last,'起始收盘价':base,'结束收盘价':valid.iloc[-1],'区间涨跌幅':change,'有效日期数':len(valid),'来源':stock['source'],'复权口径':stock['basis']}))
    return dict(items=items,multi=multi,freq=freq,mode=mode,request_start=a,request_end=b,calendar_note='使用本地交易日历，已知缺失日期处断线' if calendar is not None else '无完整交易日历；以所选股票日期并集识别已知缺口，无法完整核实缺失交易日')

def setup_font():
    available={f.name for f in font_manager.fontManager.ttflist}
    fonts=[f for f in ['Microsoft YaHei','PingFang SC','Heiti TC','Arial Unicode MS','SimHei','Noto Sans CJK SC'] if f in available]
    if not fonts:raise ValueError('未找到中文字体，请安装思源黑体 Noto Sans CJK SC')
    rcParams['font.family']=fonts[0];rcParams['axes.unicode_minus']=False

def draw(result):
    setup_font();items=result['items'];multi=result['multi']
    fig=Figure(figsize=(16 if multi else 12,max(8,5+.3*len(items)) if multi else 7.5),dpi=110,facecolor='#F7F9FC')
    ax=fig.add_axes([.07,.25,.53,.47] if multi else [.09,.24,.86,.48]);ax.set_facecolor('white')
    basis=items[0]['stock']['basis'];title='股票走势比较（起点＝100）' if multi else items[0]['stock']['name']+' '+items[0]['stock']['code']+' 收盘价走势'
    fig.text(.09,.95,title,fontsize=19,fontweight='bold',color='#19334D')
    fig.text(.09,.91,basis+' | '+('股价涨跌幅，不含分红；除权除息可能影响走势' if basis=='不复权' else '复权价格涨跌幅，不代表实际成交价格或实际账户收益'),fontsize=10,color='#51657B')
    summaries=[]
    for i,item in enumerate(items):
        s=item['summary'];color=COLORS[i%len(COLORS)]
        ax.plot(item['plot'].index,item['plot'].values,color=color,linewidth=1.7,linestyle='--' if i>=10 else '-',marker='o' if len(item['display'])<15 else None,markersize=4,label=s['名称']+' '+s['代码'])
        change='无法计算（不足两个有效日期）' if pd.isna(s['区间涨跌幅']) else f"{s['区间涨跌幅']:+.2%}"
        summaries.append(f"{s['名称']} {s['代码']}  {s['实际开始']:%Y-%m-%d}—{s['实际结束']:%Y-%m-%d}  起始 {s['起始收盘价']:.6g} → 结束 {s['结束收盘价']:.6g}  涨跌幅 {change}")
        # 标签固定排在坐标上/下方，与数据点连线，避免标签随极值位置相互覆盖。
        for j,point in enumerate(item['points']):
            y=point['价格']/item['base']*100 if multi else point['价格']
            ax.scatter([point['日期']],[y],marker='^' if j==0 else 'v',s=55,color=color,zorder=5)
            xpos=(i+.5)/len(items) if multi else (.26 if j==0 else .74)
            ypos=1.04 if j==0 else -.19
            text=f"{s['名称']} {point['标注']}\n{point['日期']:%Y-%m-%d}  {point['价格']:.6g}元"
            ax.annotate(text,xy=(point['日期'],y),xytext=(xpos,ypos),textcoords='axes fraction',ha='center',va='bottom' if j==0 else 'top',fontsize=9,color=color,annotation_clip=False,arrowprops={'arrowstyle':'-','color':color,'alpha':.6})
    if multi:
        first=items[0]['summary']
        fig.text(.09,.865,f"共同区间 {first['实际开始']:%Y-%m-%d} — {first['实际结束']:%Y-%m-%d}  |  共 {len(items)} 只股票  |  起点＝100",fontsize=11,color='#33465A')
        tabax=fig.add_axes([.64,.25,.34,.47]);tabax.axis('off')
        rows=[]
        for item in items:
            summary=item['summary']
            rows.append([summary['名称']+' '+summary['代码'],f"{summary['起始收盘价']:.6g}",f"{summary['结束收盘价']:.6g}",f"{summary['区间涨跌幅']:+.2%}"])
        table=tabax.table(cellText=rows,colLabels=['股票 / 代码','起始价格','结束价格','涨跌幅'],colWidths=[.46,.18,.18,.18],cellLoc='center',bbox=[0,0,1,1])
        table.auto_set_font_size(False);table.set_fontsize(8 if len(items)>14 else 9)
        for (row,col),cell in table.get_celld().items():
            cell.set_edgecolor('#E8EDF3');cell.set_linewidth(.6)
            if row==0:cell.set_facecolor('#E5EDF6');cell.set_text_props(color='#19334D',weight='bold')
            else:
                cell.set_facecolor('#FFFFFF' if row%2 else '#F2F6FA')
                if col==0:cell.set_text_props(color=COLORS[(row-1)%len(COLORS)],weight='bold')
        tabax.set_title('颜色对应左侧曲线 · 价格单位：元',fontsize=10,color='#51657B',pad=12)
    else:fig.text(.09,.865,'\n'.join(summaries),fontsize=9,color='#33465A',va='top')
    ax.set_ylabel('归一化价格（起点＝100）' if multi else ('收盘价（元）' if basis=='不复权' else '复权价格（元）'))
    locator=mdates.AutoDateLocator(minticks=4,maxticks=8);ax.xaxis.set_major_locator(locator);ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'));ax.tick_params(axis='x',labelsize=8)
    ax.grid(alpha=.15);ax.margins(x=.04,y=.18)
    for side in ['top','right']:ax.spines[side].set_visible(False)
    # 多股使用右侧同色摘要表作为图例，避免遮挡曲线。
    source='；'.join(dict.fromkeys(x['stock']['source'] for x in items))
    fig.text(.09,.10,'来源：'+source+' | 显示频率：'+result['freq']+'（周期最后有效交易日，先截区间再抽样）',fontsize=9,color='#51657B')
    note='极值根据区间内原始日数据计算。' if result['mode']!='关闭' else '多股比较按共同有效起止日期归一化。' if multi else ''
    if result['mode']=='区间最高价最低价':note+='最高价、最低价不一定在收盘价折线上。'
    fig.text(.09,.065,note,fontsize=9,color='#51657B')
    fig.text(.09,.03,'缺失交易日可能无法全部识别。缩放仅用于查看局部走势；如需重新计算涨跌幅和极值，请修改日期范围后重新生成图表。',fontsize=9,color='#51657B')
    return fig

def export_data(result,path):
    with pd.ExcelWriter(path,engine='openpyxl',datetime_format='yyyy-mm-dd') as w:
        pd.DataFrame([i['summary'] for i in result['items']]).to_excel(w,sheet_name='指标摘要',index=False)
        pd.DataFrame([{'项目':k,'说明':str(result[k])} for k in ['freq','mode','request_start','request_end','calendar_note']]).replace({'freq':'显示频率','mode':'标注方式','request_start':'请求开始','request_end':'请求结束','calendar_note':'日期检查说明'}).to_excel(w,sheet_name='图表设置',index=False)
        for n,i in enumerate(result['items'],1):
            prefix=i['stock']['code']+'_'+str(n)
            pd.DataFrame({'日期':i['display'].index,'显示价格':i['display'].values}).to_excel(w,sheet_name=prefix+'_图表数据',index=False)
            i['frame'].reset_index(names='日期').to_excel(w,sheet_name=prefix+'_日数据',index=False)
            pd.DataFrame(i['points']).to_excel(w,sheet_name=prefix+'_极值',index=False)
            pd.DataFrame(i['stock']['logs']).to_excel(w,sheet_name=prefix+'_检查日志',index=False)
            pd.DataFrame([{'原始元数据':json.dumps(p['meta'],ensure_ascii=False,default=str),'来源文件':p['path'],'工作表':p['sheet'],'人工确认':p['confirmation']} for p in i['stock']['parts']]).to_excel(w,sheet_name=prefix+'_来源',index=False)
        for ws in w.book:
            ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width=24
                for cell in col:
                    if isinstance(cell.value,str):cell.data_type='s'
                    if col[0].value=='区间涨跌幅' and isinstance(cell.value,(int,float)):cell.number_format='0.00%'

class App:
    def __init__(self,root):
        self.root=root;root.title('本地股票走势查看器');root.geometry('1280x940');self.stocks=[];self.ids=[];self.chosen=set();self.calendar=None;self.result=None;self.fig=None;self.busy=False;self.q=queue.Queue()
        top=ttk.Frame(root,padding=8);top.pack(fill='x')
        ttk.Button(top,text='导入Excel文件',command=self.load).pack(side='left')
        ttk.Button(top,text='可选：导入交易日历',command=self.load_calendar).pack(side='left',padx=5)
        self.search=tk.StringVar();ttk.Label(top,text='搜索名称/代码').pack(side='left');ttk.Entry(top,textvariable=self.search,width=22).pack(side='left');self.search.trace_add('write',lambda *_:self.filter())
        self.multi=tk.BooleanVar();ttk.Checkbutton(top,text='多股比较',variable=self.multi,command=self.change_mode).pack(side='left',padx=8)
        selection=ttk.Frame(root);selection.pack(fill='x',padx=10)
        self.count=tk.StringVar(value='已勾选0只')
        ttk.Label(selection,textvariable=self.count).pack(side='left')
        ttk.Button(selection,text='勾选搜索结果',command=self.select_visible).pack(side='left',padx=6)
        ttk.Button(selection,text='清空勾选',command=self.clear_selection).pack(side='left')
        chooser=ttk.Frame(root);chooser.pack(fill='x',padx=10)
        self.check_canvas=tk.Canvas(chooser,height=115,highlightthickness=0)
        scroll=ttk.Scrollbar(chooser,orient='vertical',command=self.check_canvas.yview)
        self.check_canvas.configure(yscrollcommand=scroll.set);self.check_canvas.pack(side='left',fill='x',expand=True);scroll.pack(side='right',fill='y')
        self.check_frame=ttk.Frame(self.check_canvas)
        window=self.check_canvas.create_window((0,0),window=self.check_frame,anchor='nw')
        self.check_frame.bind('<Configure>',lambda e:self.check_canvas.configure(scrollregion=self.check_canvas.bbox('all')))
        self.check_canvas.bind('<Configure>',lambda e:self.check_canvas.itemconfigure(window,width=e.width))
        form=ttk.Frame(root,padding=8);form.pack(fill='x')
        self.range=tk.StringVar(value='全部数据');self.start=tk.StringVar();self.end=tk.StringVar();self.freq=tk.StringVar(value='日');self.mode=tk.StringVar(value='收盘价最高最低点')
        for title,var,choices in [('时间范围',self.range,list(RANGES)+['全部数据','自定义']),('显示频率',self.freq,list(FREQ)),('极值标注',self.mode,['收盘价最高最低点','区间最高价最低价','关闭'])]:
            ttk.Label(form,text=title).pack(side='left');cb=ttk.Combobox(form,textvariable=var,values=choices,width=18,state='readonly');cb.pack(side='left',padx=4)
            if title=='极值标注':self.modebox=cb
        for title,var in [('开始',self.start),('结束',self.end)]:ttk.Label(form,text=title).pack(side='left');ttk.Entry(form,textvariable=var,width=12).pack(side='left')
        bar=ttk.Frame(root);bar.pack(fill='x',padx=10)
        for title,fn in [('生成图表',self.generate),('导出PNG/SVG',self.export_image),('导出数据Excel',self.export_excel)]:ttk.Button(bar,text=title,command=fn).pack(side='left',padx=4)
        self.status=tk.StringVar(value='先导入文件。近期区间以文件最新日期为终点；修改选项后点击“生成图表”。');ttk.Label(root,textvariable=self.status,wraplength=1200).pack(anchor='w',padx=10)
        self.panel=ttk.Frame(root);self.panel.pack(fill='both',expand=True)
        self.details=tk.Text(root,height=5,wrap='word');self.details.pack(fill='x',padx=10)
        root.after(100,self.poll)
    def selected(self):return [self.stocks[i] for i in sorted(self.chosen) if i<len(self.stocks)]
    def filter(self):
        query=self.search.get().strip();self.ids=[i for i,s in enumerate(self.stocks) if query in s['code'] or query in s['name']]
        for widget in self.check_frame.winfo_children():widget.destroy()
        self.checkvars={}
        for n,i in enumerate(self.ids):
            stock=self.stocks[i];var=tk.BooleanVar(value=i in self.chosen);self.checkvars[i]=var
            text=f"{stock['name']} {stock['code']} | {stock['basis']}"
            ttk.Checkbutton(self.check_frame,text=text,variable=var,command=lambda idx=i:self.toggle_stock(idx)).grid(row=n//3,column=n%3,sticky='w',padx=8,pady=4)
        for col in range(3):self.check_frame.columnconfigure(col,weight=1)
        self.count.set(f'已勾选 {len(self.chosen)} 只（搜索不会清除已有勾选，最多20只）')
    def toggle_stock(self,i):
        if self.checkvars[i].get() and i not in self.chosen and len(self.chosen)>=20:
            self.checkvars[i].set(False);messagebox.showwarning('数量提示','每次最多勾选20只股票');return
        if self.checkvars[i].get():self.chosen.add(i)
        else:self.chosen.discard(i)
        if len(self.chosen)>1 and not self.multi.get():self.multi.set(True);self.mode.set('关闭')
        self.filter();self.update_options()
    def select_visible(self):
        if len(self.chosen.union(self.ids))>20:messagebox.showwarning('数量提示','最多比较20只，请缩小搜索范围或逐个勾选');return
        self.chosen.update(self.ids)
        if len(self.chosen)>1 and not self.multi.get():self.multi.set(True);self.mode.set('关闭')
        self.filter();self.update_options()
    def clear_selection(self):
        self.chosen.clear();self.filter();self.update_options()
    def update_options(self):
        ss=self.selected();ok=bool(ss) and all(all(c in s['data'] and s['data'][c].notna().any() for c in ['最高价','最低价']) for s in ss)
        self.modebox['values']=['收盘价最高最低点','区间最高价最低价','关闭'] if ok else ['收盘价最高最低点','关闭']
        if not ok and self.mode.get()=='区间最高价最低价':self.mode.set('收盘价最高最低点')
        if ss:
            self.start.set(str(min(s['data'].index.min() for s in ss).date()));self.end.set(str(min(s['data'].index.max() for s in ss).date()))
        if not ok:self.status.set('所选数据缺少有效最高价/最低价列，已禁用区间最高最低价选项。')
    def change_mode(self):
        if not self.multi.get() and len(self.chosen)>1:self.chosen={min(self.chosen)}
        self.mode.set('关闭' if self.multi.get() else '收盘价最高最低点');self.filter();self.update_options()
    def load(self):
        if self.busy:return
        paths=filedialog.askopenfilenames(filetypes=[('Excel文件','*.xlsx')])
        if not paths:return
        self.busy=True;self.status.set('正在读取本地文件……')
        def worker():
            try:self.q.put((read_files(paths),None))
            except Exception as e:self.q.put((None,str(e)))
        threading.Thread(target=worker,daemon=True).start()
    def poll(self):
        try:
            parts,error=self.q.get_nowait();self.busy=False
            if error:messagebox.showerror('读取失败',error)
            else:
                try:self.accept(parts)
                except Exception as e:messagebox.showerror('导入失败',str(e))
        except queue.Empty:pass
        self.root.after(100,self.poll)
    def accept(self,parts):
        accepted=[]
        for p in parts:
            if p['basis']=='未知':
                b=simpledialog.askstring('确认复权口径',p['sheet']+'：请输入不复权、前复权或后复权；无法确认请取消')
                if b not in ['不复权','前复权','后复权']:continue
                reason=simpledialog.askstring('确认依据','请填写口径依据（不能根据价格猜测）')
                if not reason:continue
                p['basis']=b;p['confirmation']='复权确认：'+reason
            if p['period']!='日线':
                if p['period'] not in ['未知','']:messagebox.showwarning('跳过',p['sheet']+'不是日线');continue
                if not messagebox.askyesno('确认周期',p['sheet']+'没有周期元数据，确认其为日线吗？'):continue
                p['confirmation']+='；用户确认日线'
            accepted.append(p)
        # 追加导入；同一文件同一工作表重新导入时替换旧版本。
        keys={(p['path'],p['sheet']) for p in accepted};existing=[p for s in self.stocks for p in s['parts'] if (p['path'],p['sheet']) not in keys]
        groups={}
        for p in existing+accepted:groups.setdefault((p['code'],p['basis']),[]).append(p)
        stocks=[];notices=[]
        for (c,b),group in groups.items():
            sets=[group]
            if len(group)>1:
                yes=messagebox.askyesno('合并确认',f'{c} {b} 有多个来源。确认复权口径和调整基准一致后才能合并。\n确认合并？选“否”将作为独立数据集保留。')
                if not yes:sets=[[p] for p in group]
                else:
                    for p in group:p['confirmation']+='；用户确认同口径且调整基准一致后合并'
            for group in sets:
                d,logs,q=clean(group)
                if not d['收盘价'].notna().any():notices.append(c+'无有效收盘价，跳过');continue
                stocks.append(dict(code=c,name=group[0]['name'],basis=b,source='；'.join(dict.fromkeys(p['source'] for p in group)),data=d,parts=group,logs=logs,quality=q))
                notices.append(c+'：'+str(q))
        self.stocks=stocks;self.chosen={0} if stocks else set();self.filter();self.update_options();self.result=None
        self.status.set('导入完成，请选择并生成图表。冲突日期保持缺口；没有完整日历时不能核实所有缺失交易日。')
        self.details.delete('1.0','end');self.details.insert('end','\n'.join(notices))
        if any(s['quality']['冲突日期数'] for s in stocks):messagebox.showwarning('存在价格冲突','重叠日期价格不一致，未直接拼接该日期价格；已设为缺失，详见检查日志。')
    def load_calendar(self):
        p=filedialog.askopenfilename(filetypes=[('本地交易日历','*.xlsx *.csv')])
        if not p:return
        try:
            d=pd.read_excel(p) if p.endswith('.xlsx') else pd.read_csv(p)
            if '是否交易日' in d:d=d[d['是否交易日'].astype(str).isin(['1','1.0','是','True'])]
            self.calendar=pd.DatetimeIndex(date_values(d['日期']).dropna().unique()).sort_values();self.status.set('已载入交易日历；无是否交易日列时，每行须为真实交易日。请重新生成图表。')
        except Exception as e:messagebox.showerror('日历错误',str(e))
    def generate(self):
        try:
            ss=self.selected()
            if not self.multi.get():ss=ss[:1]
            if self.multi.get() and len(ss)<2:raise ValueError('多股比较请勾选至少两只股票')
            if len(ss)>20:raise ValueError('每次最多比较20只股票，请取消部分勾选')
            if len(ss)>3 and self.mode.get()!='关闭':raise ValueError('为避免极值标签重叠，开启标注时最多比较3只股票')
            result=prepare(ss,self.range.get(),self.start.get(),self.end.get(),self.freq.get(),self.mode.get(),self.calendar)
            fig=draw(result)
            for child in self.panel.winfo_children():child.destroy()
            self.fig=fig;self.result=result
            canvas=FigureCanvasTkAgg(fig,master=self.panel);toolbar=NavigationToolbar2Tk(canvas,self.panel);toolbar.update();canvas.get_tk_widget().pack(fill='both',expand=True);canvas.draw()
            self.details.delete('1.0','end')
            for item in result['items']:
                self.details.insert('end',str(item['summary'])+'\n')
                for point in item['points']:self.details.insert('end',point['标注']+'全部出现日期：'+point['全部出现日期']+'（图中标注首次）\n')
            self.status.set('图表已更新。仅展示实际已有数据，实际覆盖见图上方；不足所选年限不补造。缩放/平移不改变统计区间；修改选项后需重新生成。')
        except Exception as e:messagebox.showerror('无法生成图表',str(e))
    def target(self,ext):
        if not self.result:messagebox.showwarning('尚未生成','请先生成图表');return None
        items=self.result['items'];a=items[0]['summary']['实际开始'];b=items[0]['summary']['实际结束'];codes='_'.join(i['stock']['code'] for i in items)
        name=f'{codes}_{a:%Y%m%d}_{b:%Y%m%d}_{self.result["freq"]}'
        p=filedialog.asksaveasfilename(initialfile=name+ext,defaultextension=ext,filetypes=[('图片','*.png *.svg')] if ext=='.png' else [('Excel','*.xlsx')],confirmoverwrite=True)
        if p and Path(p).resolve() in {Path(part['path']).resolve() for s in self.stocks for part in s['parts']}:raise ValueError('禁止覆盖原始输入文件')
        return p
    def export_image(self):
        try:
            p=self.target('.png')
            if p:
                if Path(p).suffix.lower() not in ['.png','.svg']:raise ValueError('图片格式必须为PNG或SVG')
                self.fig.savefig(p,dpi=300,facecolor=self.fig.get_facecolor());self.status.set('已导出当前视图：'+p+'；统计区间保持不变。')
        except Exception as e:messagebox.showerror('导出失败',str(e))
    def export_excel(self):
        try:
            p=self.target('.xlsx')
            if p:export_data(self.result,p);self.status.set('数据及指标已导出：'+p)
        except Exception as e:messagebox.showerror('导出失败',str(e))

if __name__=='__main__':
    root=tk.Tk();App(root);root.mainloop()
