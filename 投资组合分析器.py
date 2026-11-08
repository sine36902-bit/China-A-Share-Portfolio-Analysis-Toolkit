"""Python 3.11+；pip install pandas numpy openpyxl matplotlib xlsxwriter baostock。价格仅来自本地；API只核实交易日历和交易状态。"""
from pathlib import Path
import json,re,math,tempfile,os,threading,queue
import tkinter as tk
from tkinter import ttk,filedialog,messagebox,simpledialog
import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg,NavigationToolbar2Tk
from matplotlib.ticker import PercentFormatter
from matplotlib import font_manager,rcParams

STEPS={'101组（步长1%）':100,'1001组（步长0.1%）':1000,'10001组（步长0.01%）':10000}
NOTICE='预期收益率由历史算术平均收益率估计；图示仅适用于所选两只股票、当前样本及权重约束，不代表未来收益保证。'

def parse_date(x):
    if pd.isna(x):return pd.NaT
    try:
        if isinstance(x,(float,int)):
            return (pd.Timestamp('1899-12-30')+pd.Timedelta(days=x)).normalize() if 20000<x<100000 else pd.NaT
        return pd.Timestamp(str(x)).normalize()
    except Exception:return pd.NaT

def normalize_code(x):return str(x).strip().removesuffix('.0').zfill(6)

def read_tables(path):
    book=pd.ExcelFile(path);result={}
    for name in book.sheet_names:
        preview=pd.read_excel(book,name,header=None,nrows=5)
        row=0
        if name.startswith(('明细_','股票_')):
            found=[i for i,r in preview.iterrows() if '日期' in r.values]
            if not found:raise ValueError(f'{name}：前5行未找到“日期”表头，请提供日期及价格列，不能自动猜测')
            row=found[0]
        result[name]=pd.read_excel(book,name,header=row,dtype=object)
    return result

def import_workbook(path):
    tabs=read_tables(path);items=[]
    mapping=tabs.get('股票代码与名称对应',pd.DataFrame());meta=tabs.get('下载说明',pd.DataFrame())
    clean=[n for n in tabs if n.startswith('明细_')]
    names=clean or [n for n in tabs if n.startswith('股票_')]
    for sheet in names:
        match=re.match(r'(?:明细|股票)_(\d{6})_(.+)',sheet)
        if not match:raise ValueError(f'工作表名无法识别：{sheet}；要求 明细_代码_名称 或 股票_代码_名称')
        code,name=match.groups();df=tabs[sheet].copy();md={};basis='未知';source='本地文件，来源待核对'
        table=mapping if clean else meta
        if '代码' in table:
            rows=table[table['代码'].map(normalize_code)==code]
            if len(rows)==1:md=rows.iloc[0].fillna('').to_dict();name=str(md.get('名称',md.get('证券名称',name)))
        basis=str(md.get('复权方式',md.get('复权或指数口径','未知')))
        if basis not in ['后复权','前复权','不复权']:basis='未知'
        period=str(md.get('行情周期','日线' if clean else '未知'))
        if period not in ['日线','未知']:raise ValueError(f'{sheet}不是日线价格，不能据此计算所有频率')
        source=str(md.get('数据来源',source))
        audit=[]
        for tab in ['原始来源和复权说明','分析设置与方法说明','数据质量摘要','清洗异常日志']:
            if tab in tabs:
                extra=tabs[tab]
                if '代码' in extra:extra=extra[extra['代码'].map(normalize_code)==code]
                audit.append({'表':tab,'内容':extra.to_json(orient='records',force_ascii=False,default_handler=str)})
        if clean and '原始来源和复权说明' in tabs:
            original=tabs['原始来源和复权说明'];sources=[]
            if '代码' in original:
                for _,r in original[original['代码'].map(normalize_code)==code].iterrows():
                    try:
                        raw=json.loads(str(r.get('原始元数据','{}')))
                        if raw.get('数据来源'):sources.append(raw['数据来源'])
                    except (ValueError,TypeError):pass
                if sources:source='；'.join(dict.fromkeys(sources))
        price=next((c for c in ['价格','收盘价（元）','收盘价','收盘'] if c in df),None)
        if price is None or '日期' not in df:raise ValueError(f'{sheet}缺少明确的日期/收盘价字段；支持“价格”“收盘价（元）”“收盘价”“收盘”')
        df=df.rename(columns={price:'价格'});df['原始日期']=df['日期'];df['日期']=df['日期'].map(parse_date)
        df['原始价格']=df['价格'].astype(str);v=pd.to_numeric(df['价格'],errors='coerce');df['价格']=v.where(np.isfinite(v)&(v>0))
        df['来源位置']=[f'{Path(path).resolve()} | {sheet} | 行{i+2+(1 if clean else 0)}' for i in range(len(df))]
        halt=next((c for c in ['已确认停牌','停牌状态','停牌'] if c in df),None)
        df['停牌']=df[halt].astype(str).str.lower().isin(['true','1','1.0','是','停牌']) if halt else False
        df['交易状态']=df['tradestatus'].astype(str).str.replace('.0','',regex=False) if 'tradestatus' in df else (df['交易状态'].astype(str).replace({'正常交易':'1','停牌':'0','股票停牌':'0'}) if '交易状态' in df else '')
        df['缺失分类']=df['缺失原因'].astype(str) if '缺失原因' in df else '原因未知'
        logs=[];records=[]
        for _,r in df[df['日期'].isna()].iterrows():logs.append({'来源位置':r['来源位置'],'日期':str(r['原始日期']),'原因':'无法解析日期'})
        for day,group in df.dropna(subset=['日期']).groupby('日期'):
            cols=[c for c in ['价格','停牌','交易状态','最高价（元）','最低价（元）','缺失分类'] if c in group]
            unique=group.drop_duplicates(cols)
            if len(unique)<len(group):logs.append({'日期':day,'来源位置':'；'.join(group['来源位置']),'原因':f'删除完全重复记录{len(group)-len(unique)}条'})
            r=unique.iloc[0];value=r['价格'];reason=''
            if len(unique)>1:value=np.nan;reason='同日冲突价格，默认缺失'
            elif pd.isna(value):reason='无效或缺失价格：'+r['缺失分类']
            if r['停牌']:reason+='；已确认停牌，无交易报价不用于收益边界'
            if reason:logs.append({'日期':day,'来源位置':'；'.join(group['来源位置']),'原因':reason})
            records.append({'日期':day,'价格':value,'停牌':bool(r['停牌']),'交易状态':str(r['交易状态']) if len(unique)==1 else '', '原因':reason,'来源位置':'；'.join(group['来源位置'])})
        if not records:continue
        data=pd.DataFrame(records).set_index('日期').sort_index()
        items.append(dict(code=code,name=name,basis=basis,source=source,period=period,base=str(md.get('调整基准','未注明')),path=str(Path(path).resolve()),sheet=sheet,data=data,logs=logs,audit=audit,meta=md,confirmation=''))
    if not items:raise ValueError('未找到清洗价格明细或原始股票行情工作表。仅有收益率宽表不能重新计算月/年收益，请补充价格。')
    return items

def periods(stock,calendar,start,end,freq,coverage=None,exploratory=False):
    calendar=pd.DatetimeIndex(calendar).dropna().normalize().unique().sort_values()
    calendar=calendar[calendar<=end]  # 永不使用研究期之后价格
    data=stock['data'];rows=[]
    def make(label,left,right,extra=''):
        reason=[]
        if extra:reason.append(extra)
        prices=[];actual=[]
        for d in [left,right]:
            if pd.isna(d):prices.append(np.nan);actual.append(pd.NaT);reason.append('无法确认所需交易日');continue
            if d not in data.index:prices.append(np.nan);actual.append(pd.NaT);reason.append(f'{d:%Y-%m-%d}缺少价格，原因未知');continue
            r=data.loc[d];v=r['价格'];actual.append(d)
            if bool(r['停牌']):v=np.nan;reason.append(f'{d:%Y-%m-%d}已确认停牌报价')
            if pd.isna(v):reason.append(f'{d:%Y-%m-%d}无有效边界价格：{r["原因"]}')
            prices.append(v)
        interior=calendar[(calendar>left)&(calendar<right)] if pd.notna(left) and pd.notna(right) else pd.DatetimeIndex([])
        missing=sum(d not in data.index or pd.isna(data.loc[d,'价格']) for d in interior)
        stopped=sum(d in data.index and bool(data.loc[d,'停牌']) for d in interior)
        value=prices[1]/prices[0]-1 if not reason and all(pd.notna(v) for v in prices) else np.nan
        rows.append({'期间':label,'所需期初日期':left,'所需期末日期':right,'实际期初价格日期':actual[0],'实际期末价格日期':actual[1],'期初价格':prices[0],'期末价格':prices[1],'收益率':value,'内部缺失日期数':missing,'内部确认停牌日期数':stopped,'排除原因':'；'.join(reason),'内部提示':'有缺口/停牌，不代表日数据完整' if missing or stopped else ''})
    if freq=='日':
        for i,d in enumerate(calendar):
            if start<=d<=end:make(str(d.date()),calendar[i-1] if i else pd.NaT,d)
    else:
        if coverage is None and not exploratory:raise ValueError('月/年收益必须导入可靠交易日历并确认覆盖范围，不能用股票日期并集冒充期末交易日。')
        unit='M' if freq=='月' else 'Y'
        for p in pd.period_range(start,end,freq=unit):
            if p.start_time<start or p.end_time.normalize()>end:
                make(str(p),pd.NaT,pd.NaT,'研究区间未覆盖完整月份/年份');continue
            prev=p-1;extra=''
            if coverage is not None and (coverage[0]>prev.start_time or coverage[1]<p.end_time.normalize()):extra='已确认日历覆盖范围不足，不能确认两端期末交易日'
            def last(period):
                dates=calendar[(calendar>=period.start_time)&(calendar<=period.end_time.normalize())]
                return dates[-1] if len(dates) else pd.NaT
            make(str(p),last(prev),last(p),extra)
    if exploratory:
        for row in rows:
            row['内部提示']='文件日期试算：期末交易日未核实；不能视为已确认的完整月/年收益。'+row['内部提示']
    return pd.DataFrame(rows,columns=['期间','所需期初日期','所需期末日期','实际期初价格日期','实际期末价格日期','期初价格','期末价格','收益率','内部缺失日期数','内部确认停牌日期数','排除原因','内部提示'])

def estimate(first,second):
    joined=first.set_index('期间').join(second.set_index('期间'),lsuffix='1',rsuffix='2',how='outer')
    joined['共同纳入']=joined['收益率1'].notna()&joined['收益率2'].notna()
    common=joined[joined['共同纳入']];n=len(common)
    if n<2:return joined,pd.DataFrame(),None
    r1=common['收益率1'].astype(float);r2=common['收益率2'].astype(float)
    m1=r1.mean();m2=r2.mean();a=r1-m1;b=r2-m2
    detail=pd.DataFrame({'股票1收益率':r1,'股票1偏差':a,'股票1偏差平方':a*a,'股票2收益率':r2,'股票2偏差':b,'股票2偏差平方':b*b,'两股偏差乘积':a*b})
    v1=(a*a).sum()/(n-1);v2=(b*b).sum()/(n-1);cov=(a*b).sum()/(n-1)
    s1=math.sqrt(v1);s2=math.sqrt(v2);rho=cov/(s1*s2) if s1>0 and s2>0 else np.nan
    stats=dict(n=n,m1=m1,m2=m2,v1=v1,v2=v2,cov=cov,s1=s1,s2=s2,rho=rho)
    return joined,detail,stats

def portfolio(st,w,A):
    w=np.asarray(w,dtype=float);mean=w*st['m1']+(1-w)*st['m2']
    var=w*w*st['v1']+(1-w)**2*st['v2']+2*w*(1-w)*st['cov']
    if np.any(var < -1e-12):raise ValueError('组合方差为明显负数，请检查方差/协方差口径')
    var=np.maximum(var,0)
    utility=mean-.5*A*var
    if not np.all(np.isfinite(utility)):raise ValueError('A与当前数据产生数值溢出，请减小A')
    return mean,var,np.sqrt(var),utility

def effective_set(risk,mean,tol=1e-12):
    # 按风险排序并建立最大收益前缀；避免10001点的平方级比较。
    order=np.argsort(risk,kind='stable');x=np.asarray(risk)[order];y=np.asarray(mean)[order];prefix=np.maximum.accumulate(y)
    efficient=np.ones(len(x),dtype=bool)
    for i,(r,m) in enumerate(zip(x,y)):
        near=np.searchsorted(x,r+tol,side='right')-1
        lower=np.searchsorted(x,r-tol,side='left')-1
        if prefix[near]>m+tol or (lower>=0 and prefix[lower]>=m-tol):efficient[i]=False
    result=np.zeros(len(x),dtype=bool);result[order]=efficient;return result

def sweep(st,divisions,A):
    w=np.arange(divisions+1,dtype=float)/divisions
    mean,var,sd,u=portfolio(st,w,A)
    df=pd.DataFrame({'股票1权重':w,'股票2权重':1-w,'组合预期收益率':mean,'组合方差':var,'组合标准差':sd,'效用U':u})
    df['有效集']=effective_set(sd,mean)
    tol=1e-12
    df['最低方差并列']=np.abs(var-var.min())<=tol
    df['最高效用并列']=np.abs(u-u.max())<=tol
    return df

def representative(grid,flag):
    # 并列全部保留，图中代表点选其中风险最低者，不宣称效用严格更高。
    return grid[grid[flag]].sort_values(['组合方差','股票1权重'],kind='stable').iloc[0]

def formula_table(st,detail,w,A):
    mean,var,sd,u=portfolio(st,w,A);rows=[]
    def add(step,formula,sub,value):rows.append({'步骤':step,'公式':formula,'数值代入':sub,'结果（小数）':float(value)})
    for i in [1,2]:
        add(f'1 股票{i}预期收益','平均收益率=Σr/n',f'{detail[f"股票{i}收益率"].sum():.12g}/{st["n"]}',st[f'm{i}'])
        add(f'1 股票{i}方差','s²=Σ(r−平均收益率)²/(n−1)',f'{detail[f"股票{i}偏差平方"].sum():.12g}/({st["n"]}−1)',st[f'v{i}'])
        add(f'1 股票{i}标准差','s=√s²',f'√({st[f"v{i}"]:.12g})',st[f's{i}'])
    add('2 协方差','Cov=Σ[(r1−均值1)(r2−均值2)]/(n−1)',f'{detail["两股偏差乘积"].sum():.12g}/({st["n"]}−1)',st['cov'])
    add('2 相关系数ρ','ρ=Cov/(s1×s2)',f'{st["cov"]:.12g}/({st["s1"]:.12g}×{st["s2"]:.12g})；零标准差时无法计算',st['rho'])
    add('3 组合预期收益','E(rp)=w×均值1+(1−w)×均值2',f'{w:.8g}×{st["m1"]:.12g}+{1-w:.8g}×{st["m2"]:.12g}',mean)
    add('4 组合方差','σp²=w²s1²+(1−w)²s2²+2w(1−w)Cov',f'{w:.8g}²×{st["v1"]:.12g}+{1-w:.8g}²×{st["v2"]:.12g}+2×{w:.8g}×{1-w:.8g}×{st["cov"]:.12g}',var)
    add('4 组合标准差','σp=√σp²',f'√({float(var):.12g})',sd)
    add('5 效用','U=E(rp)−0.5Aσp²',f'{float(mean):.12g}−0.5×{A:.12g}×{float(var):.12g}',u)
    return pd.DataFrame(rows)

def configure_fonts():
    names={f.name for f in font_manager.fontManager.ttflist};fonts=[n for n in ['Microsoft YaHei','PingFang SC','Heiti TC','Arial Unicode MS','SimHei','Noto Sans CJK SC'] if n in names]
    if not fonts:raise ValueError('请安装中文字体，例如Noto Sans CJK SC')
    rcParams['font.family']=fonts[0];rcParams['axes.unicode_minus']=False

def make_figure(grid,st,w,A,context,names,mode='收益—风险主图'):
    configure_fonts();fig=Figure(figsize=(14,8.5),dpi=100,facecolor='white')
    ink='#172B3A';blue='#006E90';orange='#BE542C'
    main=fig.add_axes([.08,.34,.59,.45])
    stock_names=[re.sub(r'\s+\d{6}$','',name).strip() for name in names]
    precision=max(2,-Decimal(str(context.get('权重步长百分数',.01))).normalize().as_tuple().exponent)
    freq=context['收益率频率'];risk_mode=mode=='收益—风险主图'
    col={'权重—收益':'组合预期收益率','权重—标准差':'组合标准差','权重—效用':'效用U'}.get(mode,'组合预期收益率')
    x=grid['组合标准差' if risk_mode else '股票1权重'].to_numpy();y=grid[col].to_numpy()
    mask=grid['有效集'].to_numpy()
    if risk_mode:
        main.plot(x,y,color=orange,lw=2.5,ls='--',label='非有效部分')
        main.plot(np.where(mask,x,np.nan),np.where(mask,y,np.nan),color=blue,lw=3,label='有效集')
    else:main.plot(x,y,color=blue,lw=3,label=mode)
    low=representative(grid,'最低方差并列');best=representative(grid,'最高效用并列')
    points=[(grid.iloc[-1],f'100%{stock_names[0]}',orange,'o'),(grid.iloc[0],f'100%{stock_names[1]}',blue,'o'),(low,'最小方差','#169B79','D'),(best,'最高效用','#CD4777','*')]
    grouped=[]
    for row,label,color,marker in points:
        xx=row['组合标准差' if risk_mode else '股票1权重'];yy=row[col]
        found=next((p for p in grouped if abs(p[0]-xx)<=1e-12 and abs(p[1]-yy)<=1e-12),None)
        if found:found[2].append(label)
        else:grouped.append([xx,yy,[label],color,marker])
    for xx,yy,labels,color,marker in grouped:
        main.scatter(xx,yy,s=155 if marker=='*' else 65,marker=marker,color=color,edgecolor='white',linewidth=1.3,zorder=5)
        right=xx>(float(np.min(x))+float(np.max(x)))/2
        below='最高效用' in labels or labels==[f'100%{stock_names[0]}']
        main.annotate(' / '.join(labels),(xx,yy),xytext=(-7 if right else 7,-7 if below else 7),textcoords='offset points',ha='right' if right else 'left',va='top' if below else 'bottom',fontsize=9,color=color,zorder=7,
                      bbox=dict(facecolor='white',edgecolor='none',alpha=.85,pad=1),arrowprops=dict(arrowstyle='-',color=color,lw=.8,shrinkA=1,shrinkB=4))
    main.set_xlabel(f'风险：收益标准差（{freq}口径）' if risk_mode else f'{stock_names[0]}权重（{stock_names[1]}为其余比例）',color=ink,labelpad=8)
    main.set_ylabel(f'收益：平均收益率（{freq}口径）' if risk_mode else f'{col}（{freq}口径）',color=ink,labelpad=8)
    main.xaxis.set_major_formatter(PercentFormatter(1))
    if col!='效用U':main.yaxis.set_major_formatter(PercentFormatter(1))
    main.spines[['top','right']].set_visible(False)
    for side in ['left','bottom']:main.spines[side].set_color('#DDE3E7')
    main.tick_params(colors=ink,length=0,pad=8);main.grid(axis='y',color='#DDE3E7',lw=.7);main.set_axisbelow(True)
    main.legend(loc='upper left',frameon=False,ncol=2,fontsize=10);main.margins(.16)
    # 只放大实际非有效分支附近，范围随数据与展示频率变化。
    if risk_mode and np.any(~mask) and np.any(mask):
        zx=x[~mask];zy=y[~mask];lx=float(low['组合标准差']);ly=float(low['组合预期收益率'])
        xmin=min(float(zx.min()),lx);xmax=max(float(zx.max()),lx);ymin=min(float(zy.min()),ly);ymax=max(float(zy.max()),ly)
        dx=max(xmax-xmin,float(np.ptp(x))*.02,1e-8);dy=max(ymax-ymin,float(np.ptp(y))*.02,1e-8)
        zoom=fig.add_axes([.43,.42,.19,.14],facecolor='white')
        zoom.plot(x,y,color=orange,lw=2,ls='--');zoom.plot(np.where(mask,x,np.nan),np.where(mask,y,np.nan),color=blue,lw=2)
        zoom.scatter(lx,ly,color='#169B79',s=25,zorder=5)
        zoom.set_xlim(max(0,xmin-.2*dx),xmax+.8*dx);zoom.set_ylim(ymin-.2*dy,ymax+.8*dy)
        zoom.set_title('非有效分支附近 · 局部放大',loc='left',fontsize=9,color=ink)
        zoom.xaxis.set_major_formatter(PercentFormatter(1));zoom.yaxis.set_major_formatter(PercentFormatter(1));zoom.tick_params(labelsize=8,length=0,colors=ink)
        for spine in zoom.spines.values():spine.set_color('#DDE3E7')
    fig.text(.065,.925,'投资组合的收益与风险' if risk_mode else mode,fontsize=24,weight='bold',color=ink)
    fig.text(.065,.878,f'{names[0]} × {names[1]}',fontsize=11,color=ink)
    fig.text(.065,.838,f'{context["研究开始"]}—{context["研究结束"]} · {freq}口径 · {context["复权口径"]} · A={A:g}',fontsize=10,color=ink,alpha=.7)
    for yy,title,row,color,flag in [(.73,'01   最低风险组合',low,'#169B79','最低方差并列'),(.49,'02   当前 A 下最高效用',best,'#CD4777','最高效用并列')]:
        fig.text(.72,yy,title,fontsize=12,weight='bold',color=color)
        fig.text(.72,yy-.055,f'{row["股票1权重"]:.{precision}%} / {row["股票2权重"]:.{precision}%}',fontsize=19,weight='bold',color=ink)
        fig.text(.72,yy-.09,f'{stock_names[0]} / {stock_names[1]}',fontsize=10,color=ink,alpha=.7)
        fig.text(.72,yy-.135,f'收益 {row["组合预期收益率"]:.3%}   风险 {row["组合标准差"]:.3%}',fontsize=10,color=ink)
        fig.text(.72,yy-.17,f'效用 {row["效用U"]:.6g} · 并列 {int(grid[flag].sum())} 组',fontsize=9,color=ink,alpha=.7)
    footer=(f'共同样本 {st["n"]:,} 个 · 权重 {len(grid):,} 组 · 图示为并列中风险最低的代表\n'
            +context.get('计算日期口径','')+'\n'
            +'历史算术平均估计不代表未来收益；固定权重单期统计，不等同多年不调仓持有。\n'
            +'有效集基于当前步长；完整数据见“全部权重结果”。')
    fig.text(.08,.025,footer,fontsize=8,color=ink,alpha=.75,va='bottom',linespacing=1.5)
    tip=main.annotate('',xy=(0,0),xytext=(12,12),textcoords='offset points',bbox={'boxstyle':'round','fc':'white','alpha':.95},fontsize=9);tip.set_visible(False)
    def hover(event):
        if event.inaxes is not main:
            if tip.get_visible():tip.set_visible(False);fig.canvas.draw_idle()
            return
        pts=main.transData.transform(np.column_stack([x,y]));dist=np.hypot(pts[:,0]-event.x,pts[:,1]-event.y);i=int(np.argmin(dist))
        if dist[i]>15:tip.set_visible(False);fig.canvas.draw_idle();return
        r=grid.iloc[i];tip.xy=(x[i],y[i]);tip.set_position((-220 if event.x>main.bbox.x0+main.bbox.width*.6 else 12,-110 if event.y>main.bbox.y0+main.bbox.height*.6 else 12))
        tip.set_text(f'{names[0]}：{r["股票1权重"]:.2%}\n{names[1]}：{r["股票2权重"]:.2%}\n预期收益率 {r["组合预期收益率"]:.4%}\n方差 {r["组合方差"]:.8e}\n标准差 {r["组合标准差"]:.4%}\n效用U {r["效用U"]:.10g}')
        tip.set_visible(True);fig.canvas.draw_idle()
    fig.canvas.mpl_connect('motion_notify_event',hover);fig.canvas.mpl_connect('button_press_event',hover)
    fig._portfolio_tip=tip
    return fig

# 以下为增量升级：日历、状态核实及年化。原价格从不由API覆盖。
from decimal import Decimal, InvalidOperation
from datetime import datetime, timezone
import hashlib, subprocess, sys, copy

CACHE_VERSION=1

def weight_divisions(value):
    if str(value).strip() in STEPS:return STEPS[str(value).strip()]
    text=str(value).strip().replace('个百分点','').removesuffix('%')
    try:
        step=Decimal(text)
        if not step.is_finite() or step<=0 or step>100:raise ValueError()
        # 使用整数比而不是受Decimal上下文舍入影响的循环加法。
        numerator,denominator=step.as_integer_ratio()
        total=100*denominator
        if total % numerator:raise ValueError()
        return total//numerator
    except (InvalidOperation,ValueError,OverflowError):
        raise ValueError('步长须为0至100之间的正数，并能整除100。例如0.5、0.01；0.3不能完整等分。')

def cache_home():
    return Path(os.environ.get('PORTFOLIO_CACHE_DIR',str(Path.home()/'.a_share_portfolio_cache')))

def validate_calendar(rows,start,end):
    df=pd.DataFrame(rows)
    if not {'calendar_date','is_trading_day'}.issubset(df):raise ValueError('日历字段缺失，应为calendar_date、is_trading_day')
    dates=pd.to_datetime(df['calendar_date'],errors='coerce').dt.normalize()
    flags=df['is_trading_day'].astype(str)
    if dates.isna().any() or dates.duplicated().any() or not flags.isin(['0','1']).all():raise ValueError('日历存在无效日期、重复或未知开市标志')
    df=pd.DataFrame({'日期':dates,'是否交易日':flags.astype(int)}).sort_values('日期')
    expected=pd.date_range(start,end,freq='D')
    if not pd.DatetimeIndex(df['日期']).equals(expected):raise ValueError('日历未逐日完整覆盖请求区间，禁止计算K')
    if ((df['日期'].dt.dayofweek>=5)&df['是否交易日'].eq(1)).any():raise ValueError('日历将周末标为A股交易日，请核查来源')
    if not df['是否交易日'].any():raise ValueError('日历区间没有交易日，不能核实')
    return df

def _api_result(rs):
    if rs.error_code!='0':raise RuntimeError(f'BaoStock {rs.error_code}: {rs.error_msg}')
    rows=[]
    while rs.next():
        if rs.error_code!='0':raise RuntimeError(f'BaoStock {rs.error_code}: {rs.error_msg}')
        rows.append(rs.get_row_data())
    if rs.error_code!='0':raise RuntimeError(f'BaoStock {rs.error_code}: {rs.error_msg}')
    return pd.DataFrame(rows,columns=rs.fields)

def _network_fetch(kind,start,end,code=''):
    import baostock as bs
    login=bs.login()
    if login.error_code!='0':raise RuntimeError(f'BaoStock登录失败：{login.error_msg}')
    try:
        if kind=='calendar':
            return {'rows':_api_result(bs.query_trade_dates(start_date=start,end_date=end)).to_dict('records'),'source':'BaoStock query_trade_dates'}
        market=('sh.' if code.startswith('6') else 'sz.' if code.startswith(('0','3')) else '')
        if not market:raise ValueError('当前状态核实仅支持沪深A股；其他市场需导入明确交易状态')
        full=market+code
        status=_api_result(bs.query_history_k_data_plus(full,'date,code,tradestatus',start_date=start,end_date=end,frequency='d',adjustflag='3'))
        basic_warning=''
        try:basic=_api_result(bs.query_stock_basic(code=full))
        except Exception as e:basic=pd.DataFrame();basic_warning='上市退市资料未核实：'+str(e)
        if status.empty:raise ValueError('交易状态接口返回空数据')
        if not {'date','code','tradestatus'}.issubset(status):raise ValueError('交易状态返回字段不符')
        return {'rows':status.to_dict('records'),'basic':basic.to_dict('records'),'source':'BaoStock tradestatus / query_stock_basic','basic_warning':basic_warning}
    finally:bs.logout()

def fetch_api(kind,start,end,code=''):
    # 在独立进程设总时限，避免网络库内部阻塞导致Tk界面无法响应。
    request={'kind':kind,'start':start,'end':end,'code':code}
    error=''
    for attempt in range(2):
        try:
            result=subprocess.run([sys.executable,str(Path(__file__).resolve()),'--api-worker',json.dumps(request)],capture_output=True,text=True,timeout=45)
            lines=[line for line in result.stdout.splitlines() if line.startswith('PORTFOLIO_JSON:')]
            if not lines:raise RuntimeError('API子进程未返回结果：'+result.stderr[-400:])
            payload=json.loads(lines[-1].split(':',1)[1])
            if payload.get('error'):raise RuntimeError(payload['error'])
            return payload
        except Exception as e:error=str(e)
    raise RuntimeError('API下载失败（已有限重试）：'+error)

def cached_api(kind,start,end,code='',directory=None,refresh=False,fetcher=None):
    folder=Path(directory) if directory else cache_home();folder.mkdir(parents=True,exist_ok=True)
    path=folder/f'{kind}_{code}_{start}_{end}.json'
    def validate(payload):
        if kind=='calendar':validate_calendar(payload['rows'],start,end)
        else:
            df=pd.DataFrame(payload['rows'])
            if df.empty or not {'date','code','tradestatus'}.issubset(df):raise ValueError('状态缓存为空或字段缺失')
            d=pd.to_datetime(df['date'],errors='coerce')
            if d.isna().any() or d.duplicated().any() or not df['tradestatus'].astype(str).isin(['0','1']).all():raise ValueError('状态日期/标志无效')
            if (d<pd.Timestamp(start)).any() or (d>pd.Timestamp(end)).any():raise ValueError('状态数据越过请求期间')
            if not df['code'].astype(str).str.endswith('.'+code).all():raise ValueError('状态证券代码不符')
    cached=None
    if path.exists():
        try:
            env=json.loads(path.read_text(encoding='utf-8'));body=env['payload']
            if env['version']!=CACHE_VERSION or env['request']!=[kind,start,end,code]:raise ValueError('缓存请求不符')
            digest=hashlib.sha256(json.dumps(body,ensure_ascii=False,sort_keys=True).encode()).hexdigest()
            if digest!=env['sha256']:raise ValueError('缓存校验和不符')
            validate(body);cached=body
        except Exception:cached=None
    if cached is not None and not refresh:return dict(cached,cache_used=True,cache_path=str(path))
    try:
        payload=(fetcher or fetch_api)(kind,start,end,code);validate(payload)
        payload['downloaded_at']=datetime.now(timezone.utc).isoformat();payload['coverage']=[start,end]
        env={'version':CACHE_VERSION,'request':[kind,start,end,code],'payload':payload,'sha256':hashlib.sha256(json.dumps(payload,ensure_ascii=False,sort_keys=True).encode()).hexdigest()}
        fd,tmp=tempfile.mkstemp(dir=folder,suffix='.json');os.close(fd)
        try:Path(tmp).write_text(json.dumps(env,ensure_ascii=False),encoding='utf-8');os.replace(tmp,path)
        finally:
            if os.path.exists(tmp):os.unlink(tmp)
        return dict(payload,cache_used=False,cache_path=str(path))
    except Exception as e:
        if cached is not None:return dict(cached,cache_used=True,download_warning=str(e),cache_path=str(path))
        raise

def calendar_years(calendar,start,end):
    rows=[]
    for year in range(start.year,end.year+1):
        a=pd.Timestamp(year,1,1);b=pd.Timestamp(year,12,31)
        section=calendar[calendar['日期'].between(a,b)]
        full=start<=a and end>=b and len(section)==len(pd.date_range(a,b))
        rows.append({'年份':year,'市场实际交易日数':int(section['是否交易日'].sum()),'是否完整年度':full,'纳入K':full})
    df=pd.DataFrame(rows);valid=df.loc[df['纳入K'],'市场实际交易日数']
    return df,float(valid.mean()) if len(valid) else np.nan

def aligned_status(stock,calendar,external=None):
    data=stock['data'];ext=pd.DataFrame((external or {}).get('rows',[]))
    lookup=ext.set_index('date')['tradestatus'].astype(str).to_dict() if len(ext) else {}
    basic=(external or {}).get('basic',[]);basic=basic[0] if basic else {}
    ipo=parse_date(basic.get('ipoDate',stock.get('meta',{}).get('上市日期')))
    out=parse_date(basic.get('outDate',stock.get('meta',{}).get('退市日期')))
    rows=[];last_status=None;last_valid=np.nan;last_date=pd.NaT;halt_start=None;run=0;events=[]
    for day,market in calendar[['日期','是否交易日']].itertuples(index=False,name=None):
        raw=data.loc[day] if day in data.index else None
        price=float(raw['价格']) if raw is not None and pd.notna(raw['价格']) else np.nan
        local=str(raw.get('交易状态','')) if raw is not None else ''
        if raw is not None and bool(raw['停牌']):local='0'
        api=lookup.get(str(day.date()),'');source=(external or {}).get('source','') if api else '本地明确交易状态' if local in ['0','1'] else '未提供交易状态'
        state='原因未知';conflict=local in ['0','1'] and api in ['0','1'] and local!=api
        if not market:state='市场休市'
        elif pd.notna(ipo) and day<ipo:state='上市前无行情'
        elif pd.notna(out) and day>=out:state='退市后无行情'
        elif conflict:state='状态来源冲突'
        elif (api or local)=='0':state='股票停牌'
        elif (api or local)=='1':state='股票复牌' if last_status=='股票停牌' else '正常交易'
        reason=''
        if state in ['正常交易','股票复牌'] and pd.isna(price):reason='正常交易但原文件价格缺失/冲突'
        elif state=='原因未知':reason='无交易状态，不能凭价格、成交量或缺记录推断'
        elif state=='状态来源冲突':reason='本地与API交易状态冲突，排除'
        valid=state in ['正常交易','股票复牌'] and pd.notna(price)
        mark=price if valid else np.nan;carried=False
        if market:
            if state=='股票停牌':
                if halt_start is None:halt_start=day;run=0
                run+=1
                # 仅复权价格、且从明确正常交易连续进入已确认停牌，才可沿用。
                if stock['basis'] in ['后复权','前复权'] and pd.notna(last_valid):mark=last_valid;carried=True
            else:
                if halt_start is not None:
                    events.append({'代码':stock['code'],'停牌日期':halt_start,'恢复交易日期':day if state=='股票复牌' else pd.NaT,'连续停牌天数':run,'上次有效价格日期':last_date,'上次有效价格':last_valid,'复牌价格':price if state=='股票复牌' else np.nan,'复牌期间累计变化':price/last_valid-1 if state=='股票复牌' and pd.notna(price) and pd.notna(last_valid) else np.nan,'是否纳入主要风险分析':False,'排除原因':'复牌跨期变化不作普通单日收益' if state=='股票复牌' else '后续状态未确认，不能确定复牌','数据来源':source})
                    halt_start=None;run=0
                if valid:last_valid=price;last_date=day
                else:last_valid=np.nan;last_date=pd.NaT
            last_status=state
        rows.append({'日期':day,'代码':stock['code'],'名称':stock['name'],'市场交易日':bool(market),'休市类型':('周末' if day.dayofweek>=5 else '法定或其他休市（API未细分名称）') if not market else '', '状态':state,'原始价格':price,'严格有效价格':price if valid else np.nan,'临时估值价格':mark if market else np.nan,'是否沿用值':carried,'可执行交易':valid,'排除说明':reason,'状态来源':source})
    if halt_start is not None:events.append({'代码':stock['code'],'停牌日期':halt_start,'恢复交易日期':pd.NaT,'连续停牌天数':run,'上次有效价格日期':last_date,'上次有效价格':last_valid,'复牌价格':np.nan,'复牌期间累计变化':np.nan,'是否纳入主要风险分析':False,'排除原因':'截至区间末未确认复牌','数据来源':(external or {}).get('source','本地')})
    frame=pd.DataFrame(rows).set_index('日期')
    days=frame.loc[frame['市场交易日']].copy()
    for column,outcol in [('严格有效价格','严格日收益率'),('临时估值价格','估值日收益率')]:
        days[outcol]=days[column]/days[column].shift(1)-1
    frame=frame.join(days[['严格日收益率','估值日收益率']])
    return frame,pd.DataFrame(events)

def daily_period_table(frame,start,end,valuation=False):
    days=frame[frame['市场交易日']];col='临时估值价格' if valuation else '严格有效价格';ret='估值日收益率' if valuation else '严格日收益率'
    rows=[]
    for i,(day,row) in enumerate(days.iterrows()):
        if not start<=day<=end:continue
        prev=days.iloc[i-1] if i else None;left=days.index[i-1] if i else pd.NaT
        reason=''
        if pd.isna(row[ret]):
            reason='缺少上个市场交易日价格' if prev is None else '；'.join([str(prev['状态']),str(prev['排除说明']),str(row['状态']),str(row['排除说明']),'两端有效价格不足'])
        rows.append({'期间':str(day.date()),'所需期初日期':left,'所需期末日期':day,'实际期初价格日期':left if prev is not None and pd.notna(prev[col]) else pd.NaT,'实际期末价格日期':day if pd.notna(row[col]) else pd.NaT,'期初价格':prev[col] if prev is not None else np.nan,'期末价格':row[col],'收益率':row[ret],'排除原因':reason,'内部提示':'临时账面估值，非成交价，不可假设停牌期间交易' if valuation else '严格同步；排除停牌可能存在样本选择偏差'})
    return pd.DataFrame(rows)

def build_analysis(stocks,start,end,freq,calendar,calendar_meta=None,status_payloads=None):
    if end>=pd.Timestamp('2024-01-01'):raise ValueError('投资决策时点为2024-01-01，研究结束日期必须早于该日；不得使用决策后的行情优化')
    if start>end:raise ValueError('开始日期晚于结束日期')
    if stocks[0]['code']==stocks[1]['code']:raise ValueError('请选择不同股票')
    if stocks[0]['basis']!=stocks[1]['basis'] or stocks[0]['basis']=='未知':raise ValueError('需确认两股一致复权口径')
    bases=[s.get('base','未注明') for s in stocks]
    if all(b not in ['未注明','未知',''] for b in bases) and bases[0]!=bases[1]:raise ValueError('两股调整基准明确不一致')
    calendar=calendar[calendar['日期']<=end].copy()
    if calendar.empty:raise ValueError('交易日历为空')
    expected=pd.date_range(calendar['日期'].min(),end)
    if not pd.DatetimeIndex(calendar['日期']).equals(expected):raise ValueError('日历不完整，不能进行严格计算')
    trade=pd.DatetimeIndex(calendar.loc[calendar['是否交易日'].eq(1),'日期']);coverage=(calendar['日期'].min(),end)
    frames=[];events=[];ps=[];valuation=[];diagnostics=[];yearly=[]
    for s in stocks:
        frame,event=aligned_status(s,calendar,(status_payloads or {}).get(s['code']));frames.append(frame)
        if not event.empty:
            event=event[(event['停牌日期']<=end)&(event['恢复交易日期'].isna()|(event['恢复交易日期']>=start))].copy()
        events.append(event)
        if freq=='日':p=daily_period_table(frame,start,end)
        else:
            strict=copy.copy(s);strict['data']=s['data'].reindex(trade).copy()
            strict['data']['价格']=frame['严格有效价格'].reindex(trade);strict['data']['停牌']=frame['状态'].reindex(trade).eq('股票停牌')
            strict['data']['原因']=frame['状态'].reindex(trade).fillna('原因未知')
            p=periods(strict,trade,start,end,freq,coverage)
        ps.append(p);valuation.append(daily_period_table(frame,start,end,True))
        part=frame.loc[start:end];market=part[part['市场交易日']]
        halt=market['状态'].eq('股票停牌');runs=halt.groupby((~halt).cumsum()).sum()
        diagnostics.append({'代码':s['code'],'名称':s['name'],'市场总交易日数':len(market),'有效正常报价日数':int(market['严格有效价格'].notna().sum()),'停牌总天数':int(halt.sum()),'最长连续停牌天数':int(runs.max()) if len(runs) else 0,'复牌次数':int(market['状态'].eq('股票复牌').sum()),'未知或冲突状态天数':int(market['状态'].isin(['原因未知','状态来源冲突']).sum()),'正常交易缺价天数':int((market['状态'].isin(['正常交易','股票复牌'])&market['原始价格'].isna()).sum())})
    joined,detail,st=estimate(*ps)
    strict_daily=estimate(*[daily_period_table(f,start,end) for f in frames]);valued=estimate(*valuation)
    years,K=calendar_years(calendar,start,end)
    factor=K if freq=='日' else 12. if freq=='月' else 1.
    for year in range(start.year,end.year+1):
        market=calendar[(calendar['日期'].dt.year==year)&calendar['日期'].between(start,end)&calendar['是否交易日'].eq(1)]
        daily_join=strict_daily[0];valid=daily_join.loc[daily_join.index.str.startswith(str(year)),'共同纳入'].sum()
        yearly.append({'年份':year,'市场交易日数':len(market),'共同有效日收益样本数':int(valid),'日样本覆盖率':float(valid/len(market)) if len(market) else np.nan})
    ctx={'研究开始':str(start.date()),'研究结束':str(end.date()),'投资决策时点':'2024-01-01','收益率频率':freq,'计算日期口径':'严格市场日历；交易状态未核实者排除','复权口径':stocks[0]['basis'],'共同样本数':st['n'] if st else 0,'排除期间数':len(joined)-len(detail),'数据来源':'；'.join(s['source'] for s in stocks),'日历来源':(calendar_meta or {}).get('source','本地已确认日历'),'日历下载时间':(calendar_meta or {}).get('downloaded_at','本地导入'),'日历覆盖':str(coverage[0].date())+' 至 '+str(coverage[1].date()),'市场年均交易日数K':K,'当前频率年化乘数':factor,'年化说明':'均值/方差/协方差乘以年化乘数，标准差乘其平方根；忽略跨期自相关的近似，非复合年增长率','停牌说明':'剔除停牌可能产生样本选择偏差，不能完整捕捉长期停牌风险','调整基准说明':'保留文件元数据；未注明者未获独立认证；不跨来源拼接','日历缓存提示':(calendar_meta or {}).get('download_warning',''),'资料时点':'行情截断在决策前；文件及API为事后下载，不声称历史时点可得性已认证'}
    return dict(stocks=stocks,joined=joined,detail=detail,stats=st,context=ctx,calendar=calendar,calendar_meta=calendar_meta or {},frames=frames,status_payloads=status_payloads or {},events=pd.concat(events,ignore_index=True),diagnostics=pd.DataFrame(diagnostics),yearly=pd.DataFrame(yearly),years=years,K=K,factor=factor,strict_daily=strict_daily,valuation=valued)

def annual_metrics(st,grid,factor):
    rows=[]
    for i in [1,2]:rows.append({'对象':f'股票{i}','原频率算术平均收益率':st[f'm{i}'],'原频率方差':st[f'v{i}'],'原频率标准差':st[f's{i}']})
    for flag,label in [('最低方差并列','最低方差组合（代表）'),('最高效用并列','最高效用组合（代表）')]:
        r=representative(grid,flag);rows.append({'对象':label,'股票1权重':r['股票1权重'],'股票2权重':r['股票2权重'],'原频率算术平均收益率':r['组合预期收益率'],'原频率方差':r['组合方差'],'原频率标准差':r['组合标准差']})
    for r in rows:
        r.update({'年化乘数':factor,'简单年化预期收益率':r['原频率算术平均收益率']*factor,'年化方差':r['原频率方差']*factor,'年化标准差':r['原频率标准差']*math.sqrt(factor),'年化协方差':st['cov']*factor,'相关系数ρ':st['rho']})
        r['达到10%简单年化预期收益目标']=bool(r['简单年化预期收益率']>=.1) if math.isfinite(factor) else '无法判断：缺少完整年度日历'
    return pd.DataFrame(rows)

def analytic_checks(st,grid,A):
    d=st['v1']+st['v2']-2*st['cov'];delta=st['m1']-st['m2'];rows=[]
    for flag,label in [('最低方差并列','最小方差'),('最高效用并列','最高效用')]:
        flat=False
        if flag=='最低方差并列':
            optimum=float(np.clip((st['v2']-st['cov'])/d,0,1)) if d>1e-18 else 0.;flat=d<=1e-18
        elif A>0 and d>1e-18:optimum=float(np.clip((delta/A+st['v2']-st['cov'])/d,0,1))
        else:optimum=1. if delta>0 else 0.;flat=abs(delta)<=1e-18
        r=representative(grid,flag);m,v,sd,u=portfolio(st,optimum,A)
        rows.append({'核验项':label,'解析股票1权重':np.nan if flat else optimum,'网格代表股票1权重':r['股票1权重'],'网格权重误差':0. if flat else abs(r['股票1权重']-optimum),'解析目标值':float(v if flag=='最低方差并列' else u),'网格目标值':r['组合方差'] if flag=='最低方差并列' else r['效用U'],'目标值误差':abs(float(v if flag=='最低方差并列' else u)-(r['组合方差'] if flag=='最低方差并列' else r['效用U'])),'说明':'目标恒定，全部权重并列' if flat else '限制0—100%；解析解仅独立核对，不替代网格结果'})
    return pd.DataFrame(rows)

def sensitivity(state,div,A):
    rows=[]
    for label,result in [('严格同步日收益',state['strict_daily']),('持仓临时估值日收益',state['valuation'])]:
        st=result[2]
        if st is None:rows.append({'口径':label,'样本数':int(result[0]['共同纳入'].sum()),'说明':'不足两个共同样本，不计算优化'});continue
        g=sweep(st,div,A);low=representative(g,'最低方差并列');best=representative(g,'最高效用并列')
        rows.append({'口径':label,'样本数':st['n'],'相关系数ρ':st['rho'],'协方差':st['cov'],'股票1标准差':st['s1'],'股票2标准差':st['s2'],'最低方差股票1权重':low['股票1权重'],'最高效用股票1权重':best['股票1权重'],'最低方差并列数量':int(g['最低方差并列'].sum()),'最高效用并列数量':int(g['最高效用并列'].sum()),'说明':'日口径；A不年化；估值非成交，不执行调仓'})
    df=pd.DataFrame(rows);warning=''
    if len(rows)==2 and all('最低方差股票1权重' in r for r in rows):
        changes=[abs(rows[0][c]-rows[1][c]) for c in ['相关系数ρ','最低方差股票1权重','最高效用股票1权重']]
        if any(v>=.1 for v in changes):warning='组合估计对停牌处理方式较敏感，建议谨慎解释历史最优权重。'
    df['提示']=warning;df['差异提示规则']='相关系数差≥0.1，或任一代表权重差≥10个百分点；只是提示阈值'
    return df

def build_reference_analysis(stocks,start,end,freq):
    """保留旧版明确选择的本地探索流程；绝不当作严格同步结果，不推算市场K。"""
    if end>=pd.Timestamp('2024-01-01') or start>end:raise ValueError('研究区间必须早于2024-01-01且起止顺序正确')
    if stocks[0]['basis']!=stocks[1]['basis'] or stocks[0]['basis']=='未知':raise ValueError('需明确且一致的复权口径')
    reference=pd.DatetimeIndex(sorted(set(stocks[0]['data'].index)|set(stocks[1]['data'].index)))
    reference=reference[reference<=end]
    ps=[periods(s,reference,start,end,freq,exploratory=freq!='日') for s in stocks]
    joined,detail,st=estimate(*ps)
    daily=estimate(*[periods(s,reference,start,end,'日') for s in stocks])
    unavailable=pd.DataFrame({'共同纳入':False},index=daily[0].index)
    frames=[];diagnostics=[]
    for s in stocks:
        frame=s['data'].reindex(reference).copy();frame['状态']='未核实，不能认定正常交易或停牌';frame['代码']=s['code'];frames.append(frame)
        diagnostics.append({'代码':s['code'],'名称':s['name'],'本地原始记录数':len(s['data']),'区间有效价格数':int(s['data'].loc[start:end,'价格'].notna().sum()),'相对日期并集缺失数':int(s['data'].reindex(reference[(reference>=start)&(reference<=end)])['价格'].isna().sum()),'严格风险分析状态':'日历和交易状态未核实，不能计算严格结果'})
    ctx={'研究开始':str(start.date()),'研究结束':str(end.date()),'投资决策时点':'2024-01-01','收益率频率':freq,'计算日期口径':'未核实参考：股票日期并集，不能识别共同缺口或确认期末','结果性质':'旧口径本地参考，非严格同步、非停牌核实后结果','复权口径':stocks[0]['basis'],'共同样本数':st['n'] if st else 0,'排除期间数':len(joined)-len(detail),'数据来源':'；'.join(s['source'] for s in stocks),'市场年均交易日数K':np.nan,'年化说明':'未取得可靠市场日历；日频年化留空。月年边界未核实，年化仅参考，不作为严格结果','停牌说明':'缺记录不等于停牌；尚不能给出真实停牌处理前后差异','资料时点':'本地文件事后下载，未认证为2024-01-01时点可得版本'}
    return dict(stocks=stocks,joined=joined,detail=detail,stats=st,context=ctx,calendar=pd.DataFrame({'日期':reference,'说明':'本地日期并集，不是核实后的市场日历'}),calendar_meta={'source':'本地日期并集，未核实'},frames=frames,events=pd.DataFrame(columns=['代码','停牌日期','恢复交易日期','说明']),diagnostics=pd.DataFrame(diagnostics),yearly=pd.DataFrame(),years=pd.DataFrame(columns=['年份','市场实际交易日数','纳入K']),K=np.nan,factor=np.nan if freq=='日' else 12. if freq=='月' else 1.,strict_daily=(unavailable,pd.DataFrame(),None),valuation=(unavailable,pd.DataFrame(),None))

class App:
    def __init__(self,root):
        self.root=root;root.title('投资组合分析器');root.geometry('1350x950');self.stocks=[];self.calendar=None;self.coverage=None;self.calendar_path='';self.calendar_frame=None;self.calendar_meta={};self.local_status={};self.status_paths=set();self.approved_divisions=set();self.generation=0;self.state=None;self.scenarios=[];self.pending=None;self.busy=False;self.q=queue.Queue()
        top=ttk.Frame(root,padding=8);top.pack(fill='x')
        for label,fn in [('导入行情或筛选Excel',self.load),('导入本地日历备用',self.load_calendar),('导入本地交易状态备用',self.load_status),('查看数据来源（无需填写）',self.confirm_sources),('仅本地参考（未核实）',self.reference)]:ttk.Button(top,text=label,command=fn).pack(side='left',padx=4)
        self.sel1=tk.StringVar();self.sel2=tk.StringVar();self.boxes=[]
        row=ttk.Frame(root);row.pack(fill='x',padx=8)
        for label,var in [('股票1',self.sel1),('股票2',self.sel2)]:
            ttk.Label(row,text=label).pack(side='left');box=ttk.Combobox(row,textvariable=var,state='readonly',width=43);box.pack(side='left',padx=4);self.boxes.append(box)
        self.start=tk.StringVar(value='2014-01-01');self.end=tk.StringVar(value='2023-12-31');self.freq=tk.StringVar(value='年');self.step=tk.StringVar(value='10001组（步长0.01%）');self.A=tk.StringVar(value='2')
        row=ttk.Frame(root,padding=8);row.pack(fill='x')
        for label,var in [('开始',self.start),('结束',self.end)]:ttk.Label(row,text=label).pack(side='left');ttk.Entry(row,textvariable=var,width=12).pack(side='left')
        ttk.Label(row,text='收益率频率').pack(side='left');ttk.Combobox(row,textvariable=self.freq,values=['日','月','年'],width=5,state='readonly').pack(side='left')
        ttk.Button(row,text='自动检查并开始计算',command=self.check).pack(side='left',padx=8)
        row=ttk.Frame(root,padding=8);row.pack(fill='x')
        ttk.Label(row,text='软件按步长遍历股票1权重0%—100%，股票2为其余比例').pack(side='left',padx=8)
        ttk.Label(row,text='组合组数（也可输入步长%）').pack(side='left');ttk.Combobox(row,textvariable=self.step,values=list(STEPS),width=24,state='normal').pack(side='left')
        ttk.Label(row,text='风险厌恶指数A').pack(side='left');ttk.Entry(row,textvariable=self.A,width=9).pack(side='left')
        ttk.Button(row,text='保存当前A方案',command=self.save_A).pack(side='left',padx=5)
        row=ttk.Frame(root);row.pack(fill='x',padx=8)
        for title,fn in [('导出Excel',self.export),('导出当前图PNG/SVG',self.image)]:ttk.Button(row,text=title,command=fn).pack(side='left',padx=4)
        row=ttk.Frame(root,padding=(8,4));row.pack(fill='x')
        self.chart_mode=tk.StringVar(value='收益—风险主图');self.display_basis=tk.StringVar(value='原频率');self.compare_halts=tk.BooleanVar(value=False)
        ttk.Label(row,text='查看图表：').pack(side='left',padx=8)
        chart_box=ttk.Combobox(row,textvariable=self.chart_mode,values=['收益—风险主图','权重—收益','权重—标准差','权重—效用'],state='readonly',width=17);chart_box.pack(side='left');chart_box.bind('<<ComboboxSelected>>',lambda _:self.draw_chart())
        display_box=ttk.Combobox(row,textvariable=self.display_basis,values=['原频率','年化展示'],state='readonly',width=9);display_box.pack(side='left');display_box.bind('<<ComboboxSelected>>',lambda _:self.draw_chart())
        ttk.Checkbutton(row,text='比较停牌临时估值（日）',variable=self.compare_halts,command=self.schedule).pack(side='left')
        self.status=tk.StringVar(value='导入文件 → 选择两股和日期 → 自动检查并开始计算。软件自动获取日历与状态；失败可导入已核实本地资料。');ttk.Label(root,textvariable=self.status,wraplength=1300).pack(anchor='w',padx=8)
        self.tabs=ttk.Notebook(root);self.tabs.pack(fill='both',expand=True)
        self.plot=ttk.Frame(self.tabs);self.tabs.add(self.plot,text='可行集与有效集')
        def text_panel(title):
            frame=ttk.Frame(self.tabs);self.tabs.add(frame,text=title)
            text=tk.Text(frame,wrap='none');text.grid(row=0,column=0,sticky='nsew')
            sx=ttk.Scrollbar(frame,orient='horizontal',command=text.xview);sx.grid(row=1,column=0,sticky='ew')
            sy=ttk.Scrollbar(frame,command=text.yview);sy.grid(row=0,column=1,sticky='ns')
            text.configure(xscrollcommand=sx.set,yscrollcommand=sy.set);frame.columnconfigure(0,weight=1);frame.rowconfigure(0,weight=1)
            return text
        self.quality=text_panel('数据质量与停牌');self.results=text_panel('全部比例比较');self.text=text_panel('分步公式与代入');self.detail=text_panel('期间检查、明细与A方案')
        for var in [self.sel1,self.sel2,self.start,self.end,self.freq]:var.trace_add('write',lambda *_:self.invalidate())
        for var in [self.A,self.step]:var.trace_add('write',lambda *_:self.schedule())
        root.after(100,self.poll)
    def invalidate(self):
        self.generation+=1
        self.state=None;self.scenarios=[];self.status.set('研究条件已变更，请重新检查数据；旧图不再可导出。')
    def labels(self):return [f'{i+1}. {s["name"]} {s["code"]} | {s["basis"]} | {Path(s["path"]).name}' for i,s in enumerate(self.stocks)]
    def pair(self):
        if not self.stocks:raise ValueError('请先点击导入行情或筛选Excel')
        ids=[int(v.get().split('.',1)[0])-1 for v in [self.sel1,self.sel2]];ss=[self.stocks[i] for i in ids]
        if ss[0]['code']==ss[1]['code']:raise ValueError('请选择两只不同股票')
        return ss
    def load(self):
        if self.busy:return
        paths=filedialog.askopenfilenames(filetypes=[('Excel','*.xlsx')])
        if not paths:return
        self.invalidate();self.busy=True;self.status.set('正在读取本地价格……')
        def worker():
            try:self.q.put(([s for p in paths for s in import_workbook(p)],None))
            except Exception as e:self.q.put((None,str(e)))
        threading.Thread(target=worker,daemon=True).start()
    def poll(self):
        try:
            item=self.q.get_nowait()
            if isinstance(item,dict):
                self.handle_job(item);return
            items,error=item;self.busy=False
            if error:messagebox.showerror('读取失败',error)
            else:
                self.stocks=items
                for box in self.boxes:box['values']=self.labels()
                if items:self.sel1.set(self.labels()[0])
                if len(items)>1:self.sel2.set(self.labels()[1])
                self.status.set('文件已读取。选择两只股票及文件实际覆盖的研究日期，再点击自动检查并开始计算；无需手工清洗。')
        except queue.Empty:pass
        self.root.after(100,self.poll)
    def handle_job(self,item):
        self.root.after(100,self.poll)
        if item['kind']=='progress':self.status.set(item['text']);return
        self.busy=False
        if item.get('generation',self.generation)!=self.generation:
            self.status.set('研究条件已变化，本次旧条件结果已丢弃，请重新检查');return
        if item.get('error'):
            self.status.set(item['error']);messagebox.showerror('无法完成核实',item['error']);return
        self.state=item['state'];self.scenarios=[]
        self.detail.delete('1.0','end');self.detail.insert('end',self.state['joined'].to_string())
        self.quality.delete('1.0','end');self.quality.insert('end',self.state['diagnostics'].to_string(index=False)+'\n\n'+self.state['yearly'].to_string(index=False)+'\n\n'+self.state['events'].to_string(index=False))
        if self.state['stats'] is None:
            self.tabs.select(self.quality.master);self.status.set('已完成数据检查，但不足两个已核实共同样本。可导出诊断Excel。')
            messagebox.showwarning('有效样本不足','请查看数据质量与期间检查页。未核实状态不会当作正常交易；缺少上期末价格也会排除。可先导出诊断，再补充交易状态或调整收益频率。');return
        self.update();self.tabs.select(self.plot)
        messagebox.showinfo('检查完成',f'共同样本{self.state["stats"]["n"]}个。已按步长遍历全部比例；年化与停牌诊断可导出。')
    def reference(self):
        if self.busy:return
        try:
            if not messagebox.askyesno('本地参考限制','保留旧版试算，但没有交易日历/交易状态时不能声称严格同步，日频不能计算K或年化风险。仅用于核对已有价格与旧结果，继续？'):return
            ss=self.pair()
            for stock in ss:
                if stock['basis']=='未知':self.resolve_unknown(stock)
            self.controls();result=build_reference_analysis(copy.deepcopy(ss),pd.Timestamp(self.start.get()),pd.Timestamp(self.end.get()),self.freq.get())
            self.state=result;self.scenarios=[]
            self.quality.delete('1.0','end');self.quality.insert('end',result['diagnostics'].to_string(index=False))
            self.detail.delete('1.0','end');self.detail.insert('end',result['joined'].to_string())
            if result['stats'] is None:self.status.set('本地参考共同样本不足，可导出诊断');return
            self.update();self.tabs.select(self.plot)
        except Exception as e:messagebox.showerror('本地参考失败',str(e))
    def load_status(self):
        p=filedialog.askopenfilename(filetypes=[('状态数据','*.csv *.xlsx')])
        if not p:return
        try:
            df=pd.read_csv(p,dtype=str) if p.lower().endswith('.csv') else pd.read_excel(p,dtype=str)
            if not {'date','code','tradestatus'}.issubset(df):raise ValueError('需date、code、tradestatus列；0=已确认停牌，1=正常交易。不得凭价格自行推断。')
            if not df['tradestatus'].isin(['0','1']).all():raise ValueError('交易状态必须为0或1')
            for code,g in df.groupby('code'):
                short=str(code).split('.')[-1].zfill(6)
                if pd.to_datetime(g['date'],errors='coerce').isna().any() or g['date'].duplicated().any():raise ValueError('交易状态日期无效或重复')
                self.local_status[short]={'rows':g.to_dict('records'),'source':'用户导入状态：'+str(Path(p).resolve())}
            self.status_paths.add(str(Path(p).resolve()))
            self.invalidate();self.status.set('已导入明确交易状态，重新检查即可')
        except Exception as e:messagebox.showerror('状态导入失败',str(e))
    def confirm_sources(self):
        try:
            ss=self.pair()
            text='\n\n'.join(f"{x['name']} {x['code']}\n复权：{x['basis']}\n来源：{x['source']}\n文件：{Path(x['path']).name}\n调整基准：{x['base']}" for x in ss)
            messagebox.showinfo('数据来源（自动读取）',text+'\n\n已知口径直接使用，无需填写专业说明。未注明的调整基准如实记录，本程序不拼接不同来源的价格序列。')
        except Exception as e:messagebox.showerror('选择提示',str(e))
    def resolve_unknown(self,stock):
        win=tk.Toplevel(self.root);win.title('文件未注明复权方式');win.transient(self.root);win.grab_set()
        ttk.Label(win,text=f"{stock['name']}的文件没有写明复权方式。\n请根据下载时的选项选择；不知道就取消，不要猜测。",padding=15).pack()
        var=tk.StringVar(value='请选择');ttk.Combobox(win,textvariable=var,values=['后复权','前复权','不复权'],state='readonly').pack(padx=15,pady=8)
        answer=[]
        def choose():
            if var.get() in ['后复权','前复权','不复权']:answer.append(var.get());win.destroy()
        ttk.Button(win,text='确定',command=choose).pack(pady=8);ttk.Button(win,text='取消',command=win.destroy).pack(pady=8)
        self.root.wait_window(win)
        if not answer:raise ValueError('复权方式未确认；请查原下载记录或补充带下载说明的文件')
        stock['basis']=answer[0];stock['confirmation']='用户根据原下载选项选择：'+answer[0]
    def load_calendar(self):
        p=filedialog.askopenfilename(filetypes=[('交易日历','*.xlsx *.csv')])
        if not p:return
        try:
            df=pd.read_csv(p,dtype=str) if p.lower().endswith('.csv') else pd.read_excel(p,dtype=str)
            df=df.rename(columns={'日期':'calendar_date','是否交易日':'is_trading_day'})
            if not {'calendar_date','is_trading_day'}.issubset(df):raise ValueError('新版备用日历需含每天（包括休市日）的日期与是否交易日0/1，不能仅提供股票日期并集')
            df['is_trading_day']=df['is_trading_day'].str.replace('.0','',regex=False)
            a=pd.to_datetime(df['calendar_date']).min();b=pd.to_datetime(df['calendar_date']).max()
            frame=validate_calendar(df.to_dict('records'),a,b)
            if not messagebox.askyesno('确认本地日历来源','请确认这是完整市场日历，含休市日期，来自可靠来源；不是单股价格日期。继续？'):return
            self.calendar_frame=frame;self.calendar_meta={'source':'用户确认完整本地日历','path':str(Path(p).resolve()),'coverage':[str(a.date()),str(b.date())]}
            self.calendar_path=str(Path(p).resolve());self.invalidate();self.status.set('本地完整日历已导入')
        except Exception as e:messagebox.showerror('日历错误',str(e))
    def check(self):
        if self.busy:self.status.set('正在处理，请稍候');return
        self.state=None;self.last_controls=None
        try:
            ss=self.pair();start=pd.Timestamp(self.start.get());end=pd.Timestamp(self.end.get());freq=self.freq.get();self.controls()
            if pd.isna(start) or pd.isna(end) or start>end:raise ValueError('研究日期无效')
            if end>=pd.Timestamp('2024-01-01'):raise ValueError('研究结束必须早于投资决策时点2024-01-01')
            for stock in ss:
                if stock['basis']=='未知':self.resolve_unknown(stock)
                if stock['period']=='未知' and not messagebox.askyesno('缺少频率信息',stock['name']+'确认是每日收盘价格吗？'):return
            if ss[0]['basis']!=ss[1]['basis']:raise ValueError('复权口径不一致')
            if ss[0]['basis']=='不复权' and not messagebox.askyesno('口径提示','不复权仅用于价格变化，不含股息；禁止沿用停牌价进行分红调整的估值。继续严格模式？'):return
            local=self.calendar_frame.copy() if self.calendar_frame is not None else None
            localmeta=dict(self.calendar_meta);localstatus=copy.deepcopy(self.local_status);stocks=copy.deepcopy(ss);generation=self.generation
            self.busy=True;self.status.set('正在自动获取市场日历与交易状态（价格仍使用本地文件）……')
            def worker():
                try:
                    a=f'{start.year-1}-01-01';b=str(end.date())
                    if local is not None:
                        cal=local[local['日期'].between(pd.Timestamp(a),end)];meta=localmeta
                        if cal.empty or cal['日期'].min()>pd.Timestamp(a) or cal['日期'].max()<end:raise ValueError('本地日历需从研究开始前一年1月1日完整覆盖至研究结束，以核实前期末和K')
                    else:
                        meta=cached_api('calendar',a,b);cal=validate_calendar(meta['rows'],a,b)
                    statuses={};warnings=[]
                    for stock in stocks:
                        self.q.put({'kind':'progress','text':'正在核实 '+stock['name']+' 的交易状态……'})
                        if stock['code'] in localstatus:statuses[stock['code']]=localstatus[stock['code']];continue
                        try:statuses[stock['code']]=cached_api('status',a,b,stock['code'])
                        except Exception as e:warnings.append(stock['name']+'状态获取失败：'+str(e))
                    result=build_analysis(stocks,start,end,freq,cal,meta,statuses)
                    result['context']['API提示']='；'.join(warnings)
                    self.q.put({'kind':'analysis','state':result,'generation':generation})
                except Exception as e:self.q.put({'kind':'analysis','generation':generation,'error':str(e)+'\n可导入已核实的本地完整日历/交易状态后重试。不会使用252或将未知状态当作正常交易。'})
            threading.Thread(target=worker,daemon=True).start()
        except Exception as e:self.status.set(str(e));messagebox.showerror('检查提示',str(e))
    def schedule(self):
        if self.pending:self.root.after_cancel(self.pending)
        self.pending=self.root.after(220,self.update)
    def controls(self):
        A=float(self.A.get())
        if not math.isfinite(A) or A<0:raise ValueError('A必须为有限非负数')
        div=weight_divisions(self.step.get())
        if div+1>20000 and div not in self.approved_divisions:
            if not messagebox.askyesno('计算规模提示',f'将完整生成{div+1:,}组比例及公式，可能占用较多内存与时间。继续？'):raise ValueError('已取消，请选择其他步长')
            self.approved_divisions.add(div)
        return A,div
    def update(self):
        self.pending=None
        if self.state is None or self.state['stats'] is None:return
        try:
            A,div=self.controls();st=self.state['stats'];grid=sweep(st,div,A);self.grid=grid
            w=float(representative(grid,'最高效用并列')['股票1权重'])
            self.formulas=formula_table(st,self.state['detail'],w,A)
            # 每次交互用真实组合序列交叉核验，保持同一共同样本。
            direct=w*self.state['detail']['股票1收益率']+(1-w)*self.state['detail']['股票2收益率'];mean,var,sd,u=portfolio(st,w,A)
            if not np.isclose(direct.mean(),mean,atol=1e-12) or not np.isclose(direct.var(ddof=1),var,atol=1e-12):raise ValueError('组合序列交叉校验失败')
            self.text.delete('1.0','end');self.text.insert('end','组合公式代入采用当前步长下最高效用结果（如并列，展示其中风险最低者；完整并列名单在下方）。\n\n'+self.formulas.to_string(index=False)+'\n\n各列求和：\n'+self.state['detail'].sum().to_string())
            for flag,label in [('最低方差并列','当前步长下最低方差'),('最高效用并列','当前A、当前步长下最高效用')]:
                tied=grid[grid[flag]]
                self.text.insert('end',f'\n\n{label}：共{len(tied)}个；股票1权重范围{tied["股票1权重"].min():.4%}—{tied["股票1权重"].max():.4%}（范围不意味着所有中间权重均并列，完整名单如下）\n'+tied.to_string(index=False))
            self.detail.delete('1.0','end');self.detail.insert('end',self.state['joined'].to_string()+'\n\n偏差计算明细：\n'+self.state['detail'].to_string()+'\n\n已保存A方案：\n'+self.scenario_table().to_string(index=False))
            display=grid.copy()
            for c in ['股票1权重','股票2权重','组合预期收益率','组合标准差']:display[c]=display[c].map(lambda v:f'{v:.4%}')
            for c in ['组合方差','效用U']:display[c]=display[c].map(lambda v:f'{v:.10g}')
            for c in ['有效集','最低方差并列','最高效用并列']:display[c]=display[c].map({True:'是',False:''})
            self.results.delete('1.0','end')
            self.results.insert('end',f'股票1：{self.state["stocks"][0]["name"]}；股票2：{self.state["stocks"][1]["name"]}\n步长：{100/div:g}%；共 {len(grid)} 种比例，含0%和100%，以下按股票1权重升序排列。\n最低方差＝当前步长下风险最低；最高效用＝当前A下效用最高，两者可能不同。\n所有指标采用同一组共同样本；完整数值可导出Excel“全部权重结果”。\n\n'+display.to_string(index=False))
            if self.compare_halts.get():
                comparison=sensitivity(self.state,div,A);self.quality.delete('1.0','end');self.quality.insert('end',self.state['diagnostics'].to_string(index=False)+'\n\n'+comparison.to_string(index=False)+'\n\n'+self.state['events'].to_string(index=False))
            self._draw_chart()
            self.last_controls=(A,div)
            small='年度样本少于10个，估计可能不稳定。' if self.freq.get()=='年' and st['n']<10 else ''
            self.status.set(self.state['context'].get('计算日期口径','')+'；'+f'共同样本{st["n"]}个；{self.freq.get()}口径。'+small+('A=0只比较预期收益。' if A==0 else '')+'不同频率不能直接视为同一年度目标；并列结果全部保留。')
        except Exception as e:self.status.set(str(e));self.last_controls=None
    def draw_chart(self):
        try:self._draw_chart()
        except Exception as e:self.status.set('绘图未完成：'+str(e));messagebox.showerror('绘图提示',str(e))
    def _draw_chart(self):
        if self.state is None or self.state['stats'] is None:return
        A,_=self.controls();st=self.state['stats'];w=float(representative(self.grid,'最高效用并列')['股票1权重'])
        grid=self.grid;ctx=dict(self.state['context'],权重步长百分数=100/self.controls()[1])
        if self.display_basis.get()=='年化展示':
            k=self.state['factor']
            if not math.isfinite(k):
                self.display_basis.set('原频率');messagebox.showerror('无法年化','没有已核实完整年度日历，无法计算市场K；已恢复原频率');return
            grid=grid.copy();grid['组合预期收益率']*=k;grid['组合方差']*=k;grid['组合标准差']*=math.sqrt(k);grid['效用U']*=k
            ctx['收益率频率']='简单年化';ctx['计算日期口径']+='；忽略跨期自相关近似，非复合增长'
        for child in self.plot.winfo_children():child.destroy()
        self.fig=make_figure(grid,st,w,A,ctx,[s['name']+' '+s['code'] for s in self.state['stocks']],self.chart_mode.get())
        canvas=FigureCanvasTkAgg(self.fig,master=self.plot);NavigationToolbar2Tk(canvas,self.plot);canvas.get_tk_widget().pack(fill='both',expand=True);canvas.draw()
    def save_A(self):
        if self.state is None:return
        try:
            A,_=self.controls()
            if A not in self.scenarios:self.scenarios.append(A)
            self.update()
        except Exception as e:messagebox.showerror('参数错误',str(e))
    def scenario_table(self):
        rows=[]
        if self.state is None or self.state['stats'] is None:return pd.DataFrame()
        _,div=self.controls()
        for A in self.scenarios:
            g=sweep(self.state['stats'],div,A)
            for _,r in g[g['最高效用并列']].iterrows():rows.append(dict(A=A,**r.to_dict()))
        return pd.DataFrame(rows)
    def target(self,image=False):
        if self.state is None:raise ValueError('请先完成数据检查')
        if self.state['stats'] is None:
            if image:raise ValueError('无可计算样本，不能导出优化图')
        elif self.controls()!=getattr(self,'last_controls',None):self.update()
        if self.state['stats'] is not None and getattr(self,'last_controls',None) is None:raise ValueError('参数尚未有效，请修正')
        p=filedialog.asksaveasfilename(initialfile='投资组合分析'+('.png' if image else '.xlsx'),defaultextension='.png' if image else '.xlsx',filetypes=[('PNG','*.png'),('SVG','*.svg')] if image else [('Excel','*.xlsx')],confirmoverwrite=True)
        protected={Path(s['path']).resolve() for s in self.stocks}
        if self.calendar_path:protected.add(Path(self.calendar_path).resolve())
        protected.update(Path(p).resolve() for p in self.status_paths)
        if p and Path(p).resolve() in protected:raise ValueError('不能覆盖原始输入文件')
        return p
    def image(self):
        try:
            p=self.target(True)
            if p:
                if Path(p).suffix.lower() not in ['.png','.svg']:raise ValueError('只支持PNG/SVG')
                self.fig._portfolio_tip.set_visible(False);self.fig.savefig(p,dpi=300,facecolor=self.fig.get_facecolor());self.status.set('图片已导出：'+p)
        except Exception as e:messagebox.showerror('导出失败',str(e))
    def export(self):
        try:
            p=self.target()
            if not p:return
            A,div=self.controls();st=self.state['stats'];g=self.grid if st else None
            tables=analysis_tables(self.state,g,A,div,self.scenario_table() if st else None,self.compare_halts.get())
            self.status.set('正在写入完整Excel公式并逐行校验……');self.root.update_idletasks()
            verify=(g,self.state['detail'],st,A,self.state['factor'],div) if st else None
            write_excel(tables,p,verify);self.status.set('Excel已导出并核验：'+p)
        except Exception as e:messagebox.showerror('导出失败',str(e))

def analysis_tables(state,grid,A,div,scenarios=None,include_sensitivity=True):
    st=state['stats'];ctx=dict(state['context'],风险厌恶指数A=A,比较步长百分数=100/div,计算组合总数=div+1,并列绝对容差=1e-12)
    tables={'分析总览':pd.DataFrame(ctx.items(),columns=['项目','数值或说明'])}
    if st:
        annual=annual_metrics(st,grid,state['factor']);w=float(representative(grid,'最高效用并列')['股票1权重'])
        # 总览以Python数值保存，不依赖Excel首次打开重算。
        records=[{'项目':f'股票{i}名称及代码','数值或说明':f'{s["name"]} {s["code"]}'} for i,s in enumerate(state['stocks'],1)]
        for _,row in annual.iterrows():
            for key,val in row.items():
                if key!='对象':records.append({'项目':str(row['对象'])+'：'+key,'数值或说明':val})
        records.extend([{'项目':'原频率协方差','数值或说明':st['cov']},{'项目':'相关系数ρ','数值或说明':st['rho']}])
        tables['分析总览']=pd.concat([tables['分析总览'],pd.DataFrame(records)],ignore_index=True)
        tables.update({'年化统计':annual,'市场年度交易日数':state['years'],'解析解精度核对':analytic_checks(st,grid,A),'全部权重结果':grid,'最低方差及最高效用':grid[grid['最低方差并列']|grid['最高效用并列']],'公式代入过程':formula_table(st,state['detail'],w,A)})
        summary=pd.DataFrame([{'样本数':st['n'],'股票1算术平均收益率':st['m1'],'股票2算术平均收益率':st['m2'],'股票1方差':st['v1'],'股票2方差':st['v2'],'股票1标准差':st['s1'],'股票2标准差':st['s2'],'协方差':st['cov'],'相关系数ρ':st['rho']}])
        tables['个股风险及协方差']=summary
    else:tables['分析总览'].loc[len(tables['分析总览'])]=['计算状态','不足两个已核实共同样本；未生成优化结果或虚假公式数值']
    tables['分析设置与方法']=pd.DataFrame(ctx.items(),columns=['项目','说明'])
    tables['期间检查及排除']=state['joined'].reset_index();tables['期间收益与偏差明细']=state['detail'].reset_index()
    tables['偏差列求和']=state['detail'].sum().rename('合计').reset_index()
    tables['多个A的比较']=scenarios if scenarios is not None else pd.DataFrame(columns=['A','说明'])
    tables['数据质量与停牌记录']=state['events'];tables['样本诊断']=state['diagnostics'];tables['每年样本覆盖率']=state['yearly']
    tables['日历原始记录']=state['calendar'];tables['日历来源']=pd.DataFrame([(k,json.dumps(v,ensure_ascii=False,default=str) if isinstance(v,(list,dict)) else v) for k,v in state['calendar_meta'].items() if k!='rows'],columns=['项目','说明'])
    if include_sensitivity:
        tables['停牌敏感性比较']=sensitivity(state,div,A)
        tables['临时估值共同收益']=state['valuation'][0].reset_index()
    for i,(s,f) in enumerate(zip(state['stocks'],state['frames']),1):
        tables[f'股票{i}逐日状态与估值']=f.reset_index()
        payload=state.get('status_payloads',{}).get(s['code'],{})
        tables[f'股票{i}API原始状态']=pd.DataFrame(payload.get('rows',[]))
        tables[f'股票{i}状态核实来源']=pd.DataFrame([(k,json.dumps(v,ensure_ascii=False,default=str) if isinstance(v,(list,dict)) else v) for k,v in payload.items() if k!='rows'],columns=['项目','说明'])
        tables[f'股票{i}清洗日志']=pd.DataFrame(s['logs']);tables[f'股票{i}价格明细']=s['data'].loc[:pd.Timestamp(ctx['研究结束'])].reset_index()
        tables[f'股票{i}来源']=pd.DataFrame([{'代码':s['code'],'名称':s['name'],'文件':s['path'],'工作表':s['sheet'],'复权口径':s['basis'],'调整基准':s['base'],'原元数据':json.dumps(s['meta'],ensure_ascii=False,default=str),'原清洗说明':json.dumps(s['audit'],ensure_ascii=False,default=str)}])
    return tables

def independent_values(detail,w,A,K):
    # 与组合方差展开式独立：直接构造组合收益率序列，再计算样本方差。
    r1=detail['股票1收益率'].to_numpy(float);r2=detail['股票2收益率'].to_numpy(float)
    rp=w*r1+(1-w)*r2;m=float(np.mean(rp));v=float(np.var(rp,ddof=1));sd=math.sqrt(v)
    return [w,1-w,m,v,sd,m-.5*A*v,m*K if math.isfinite(K) else '',sd*math.sqrt(K) if math.isfinite(K) else '']

def weight_formulas(row):
    return [f'=(ROW()-10)/$I$5',f'=1-A{row}',f'=A{row}*$A$5+B{row}*$B$5',f'=A{row}^2*$C$5^2+B{row}^2*$D$5^2+2*A{row}*B{row}*$E$5',f'=IF(D{row}>=-1E-12,SQRT(MAX(D{row},0)),SQRT(D{row}))',f'=C{row}-0.5*$F$5*D{row}',f'=IF(ISNUMBER($G$5),C{row}*$G$5,"")',f'=IF(ISNUMBER($G$5),E{row}*SQRT($G$5),"")']

def evaluate_weight_formula(formula,cells,row):
    # 仅用于本程序输出的八列算术公式，按实际保存的引用求值，不执行任意代码。
    import ast
    expression=formula[1:].replace('^','**')
    refs={}
    def ref(match):
        key='ref'+str(len(refs));refs[key]=cells.get(match.group().replace('$',''));return key
    expression=re.sub(r'\$?[A-Z]+\$?\d+',ref,expression)
    tree=ast.parse(expression,mode='eval')
    def calc(n):
        if isinstance(n,ast.Constant):return n.value
        if isinstance(n,ast.Name):return refs[n.id]
        if isinstance(n,ast.UnaryOp) and isinstance(n.op,ast.USub):return -calc(n.operand)
        if isinstance(n,ast.Compare) and len(n.ops)==1 and isinstance(n.ops[0],ast.GtE):return calc(n.left)>=calc(n.comparators[0])
        if isinstance(n,ast.BinOp):
            a,b=calc(n.left),calc(n.right)
            if isinstance(n.op,ast.Add):return a+b
            if isinstance(n.op,ast.Sub):return a-b
            if isinstance(n.op,ast.Mult):return a*b
            if isinstance(n.op,ast.Div):return a/b
            if isinstance(n.op,ast.Pow):return a**b
        if isinstance(n,ast.Call) and isinstance(n.func,ast.Name):
            name=n.func.id
            if name=='IF':return calc(n.args[1] if calc(n.args[0]) else n.args[2])
            args=[calc(a) for a in n.args]
            if name=='ISNUMBER':return isinstance(args[0],(int,float)) and not isinstance(args[0],bool)
            if name=='ROW':return row
            if name=='SQRT':return math.sqrt(args[0])
            if name=='MAX':return max(args)
        raise ValueError('公式超出核验器支持范围')
    return calc(tree.body)

def verify_export(path,grid=None,detail=None):
    import openpyxl
    with __import__('zipfile').ZipFile(path) as z:
        if z.testzip() is not None:raise ValueError('Excel压缩包完整性错误')
    wf=openpyxl.load_workbook(path,data_only=False,read_only=True);wv=openpyxl.load_workbook(path,data_only=True,read_only=True)
    try:
        if grid is None:return {'状态':'诊断工作簿，无优化公式'}
        sf=wf['权重计算验证'];sv=wv['权重计算验证'];inputs=list(sv.iter_rows(min_row=5,max_row=5,max_col=9,values_only=True))[0]
        cells={f'{chr(65+i)}5':v for i,v in enumerate(inputs)};A=inputs[5];K=inputs[6] if inputs[6] is not None else np.nan
        max_error=0.;count=0
        for row,(fs,vs) in enumerate(zip(sf.iter_rows(min_row=10,max_col=8,values_only=True),sv.iter_rows(min_row=10,max_col=8,values_only=True)),10):
            if row-10>=len(grid):break
            expected=independent_values(detail,float(grid.iloc[row-10]['股票1权重']),A,K)
            if list(fs)!=weight_formulas(row):raise ValueError(f'第{row}行公式与模板不一致')
            for j,(f,v,ind) in enumerate(zip(fs,vs,expected)):
                actual=evaluate_weight_formula(f,cells,row);cells[f'{chr(65+j)}{row}']=actual
                if ind=='':
                    if actual!='' or v not in ['',None]:raise ValueError('缺少年化系数时不应生成伪造年化值')
                else:
                    error=max(abs(actual-ind),abs(float(v)-ind));max_error=max(max_error,error)
                    if error>1e-10*max(1.,abs(ind)):raise ValueError(f'Excel公式、缓存与直接收益序列不一致：第{row}行第{j+1}列')
            count+=1
        if count!=len(grid) or sf.max_row!=len(grid)+9:raise ValueError('Excel公式行数与用户选择不一致')
        return {'核验行数':count,'最大误差':max_error,'核验方式':'保存文件公式视图+数值视图；独立算术解释器；非桌面Excel执行证明'}
    finally:wf.close();wv.close()

def write_excel(tables,path,verification=None):
    import xlsxwriter
    hidden_sheets={'市场年度交易日数','每年样本覆盖率','股票1来源','股票2来源'}
    if verification:
        tables=dict(tables)
        detail_table=tables['期间收益与偏差明细'].copy()
        prices=tables['期间检查及排除'].set_index('期间',verify_integrity=True)
        aligned=prices.loc[detail_table['期间']]
        for source,target in [('期初价格1','股票1期初价格'),('期末价格1','股票1期末价格'),('期初价格2','股票2期初价格'),('期末价格2','股票2期末价格')]:
            values=aligned[source].to_numpy(dtype=float)
            if not np.all(np.isfinite(values)) or np.any(values<=0):raise ValueError('共同样本存在缺失或非正价格，不能生成收益率公式')
            detail_table[target]=values
        for stock in [1,2]:
            calculated=detail_table[f'股票{stock}期末价格']/detail_table[f'股票{stock}期初价格']-1
            if not np.allclose(calculated,detail_table[f'股票{stock}收益率'],rtol=1e-10,atol=1e-12):raise ValueError('期初期末价格与共同收益样本不一致')
        tables['期间收益与偏差明细']=detail_table
        report_sheets=['分析总览','最低方差及最高效用','年化统计','个股风险及协方差','全部权重结果','公式代入过程','期间收益与偏差明细','样本诊断','数据质量与停牌记录','市场年度交易日数','每年样本覆盖率','股票1来源','股票2来源']
        tables={name:tables[name] for name in report_sheets if name in tables}
    if verification and len(verification[0])+9>1048576:raise ValueError('所选组合数超出单张Excel工作表行数上限；未减少组数，请调整步长')
    path=Path(path);fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=path.parent);os.close(fd)
    try:
        with xlsxwriter.Workbook(tmp,{'strings_to_formulas':False,'strings_to_urls':False}) as book:
            book.set_calc_mode('auto')
            header=book.add_format({'bold':True,'font_color':'white','bg_color':'#24476B','text_wrap':True,'valign':'vcenter'})
            textfmt=book.add_format({'font_name':'Arial Unicode MS','font_size':10,'valign':'top'})
            percent=book.add_format({'num_format':'0.0000%','font_size':10});precise=book.add_format({'num_format':'0.0000000000E+00','font_size':10})
            dates=book.add_format({'num_format':'yyyy-mm-dd','font_size':10});wrap=book.add_format({'text_wrap':True,'valign':'top','font_size':10})
            def put(ws,r,c,value,fmt=None):
                if value is None or (not isinstance(value,(list,dict)) and pd.isna(value)):return
                if isinstance(value,(pd.Timestamp,datetime)):ws.write_datetime(r,c,value.to_pydatetime() if isinstance(value,pd.Timestamp) else value,dates)
                elif isinstance(value,(bool,np.bool_)):ws.write_boolean(r,c,bool(value),fmt)
                elif isinstance(value,(float,int,np.number)):ws.write_number(r,c,float(value),fmt)
                else:ws.write_string(r,c,str(value),fmt or textfmt)
            for name,df in tables.items():
                ws=book.add_worksheet(name);ws.hide_gridlines(2);ws.set_default_row(20);ws.set_tab_color('#24476B' if name=='分析总览' else '#B6C7D9')
                if verification and name in hidden_sheets:ws.hide()
                if len(df.columns)==0:ws.write(0,0,'无记录');continue
                for j,col in enumerate(df.columns):
                    ws.write(0,j,col,header)
                    wide=any(k in str(col) for k in ['说明','代入','公式','元数据','文件','来源位置'])
                    ws.set_column(j,j,55 if wide else 24)
                if name in ['分析总览','分析设置与方法','日历来源'] or name.endswith('状态核实来源'):
                    ws.set_column(0,0,44);ws.set_column(1,1,100)
                ws.set_row(0,34);ws.freeze_panes(1,0);ws.autofilter(0,0,max(1,len(df)),len(df.columns)-1)
                for i,row in enumerate(df.itertuples(index=False,name=None),1):
                    for j,v in enumerate(row):
                        namecol=str(df.columns[j]);fmt=percent if any(k in namecol for k in ['权重','收益率','标准差','覆盖率']) and '数量' not in namecol else precise if any(k in namecol for k in ['方差','平方','乘积','效用','误差']) and '数量' not in namecol else textfmt
                        if name=='分析总览' and j==1:
                            label=str(row[0]);fmt=percent if any(k in label for k in ['收益率','标准差','权重']) else precise if label.endswith(('方差','协方差','效用U')) else textfmt
                        if isinstance(v,str) and len(v)>60:fmt=wrap;ws.set_row(i,min(150,20*(1+len(v)//50)))
                        put(ws,i,j,v,fmt)
            if verification:
                grid,detail,st,A,K,div=verification
                from xlsxwriter.utility import xl_col_to_name
                detail_df=tables['期间收益与偏差明细']
                end=len(detail_df)+1
                detail_ws=book.get_worksheet_by_name('期间收益与偏差明细')
                for row,values in enumerate(detail_df.itertuples(index=False,name=None),2):
                    formulas=[f'=J{row}/I{row}-1',f'=B{row}-AVERAGE($B$2:$B${end})',f'=(B{row}-AVERAGE($B$2:$B${end}))^2',f'=L{row}/K{row}-1',f'=E{row}-AVERAGE($E$2:$E${end})',f'=(E{row}-AVERAGE($E$2:$E${end}))^2',f'=(B{row}-AVERAGE($B$2:$B${end}))*(E{row}-AVERAGE($E$2:$E${end}))']
                    for col,formula in enumerate(formulas,1):detail_ws.write_formula(row-1,col,formula,percent if col in [1,4] else precise,float(values[col]))
                ranges=[f"'期间收益与偏差明细'!${xl_col_to_name(detail_df.columns.get_loc(name))}$2:${xl_col_to_name(detail_df.columns.get_loc(name))}${end}" for name in ['股票1收益率','股票2收益率']]
                r1,r2=ranges
                stats_formulas=[f'=COUNT({r1})',f'=AVERAGE({r1})',f'=AVERAGE({r2})',f'=VAR({r1})',f'=VAR({r2})',f'=STDEV({r1})',f'=STDEV({r2})',f'=COVAR({r1},{r2})*COUNT({r1})/(COUNT({r1})-1)',f'=CORREL({r1},{r2})']
                stats_ws=book.get_worksheet_by_name('个股风险及协方差')
                for j,formula in enumerate(stats_formulas):
                    stats_ws.write_formula(1,j,formula,precise,float(tables['个股风险及协方差'].iloc[0,j]))
                result_ws=book.get_worksheet_by_name('全部权重结果')
                last_result=len(grid)+1
                for i,row in enumerate(grid.itertuples(index=False,name=None),2):
                    mean=f'A{i}*AVERAGE({r1})+B{i}*AVERAGE({r2})'
                    variance=f'A{i}^2*VAR({r1})+B{i}^2*VAR({r2})+2*A{i}*B{i}*COVAR({r1},{r2})*COUNT({r1})/(COUNT({r1})-1)'
                    formulas=["=(ROW()-2)/'权重计算验证'!$I$5",f'=1-A{i}',f'={mean}',f'={variance}',f'=IF(({variance})>=-1E-12,SQRT(MAX(({variance}),0)),SQRT({variance}))',f"=({mean})-0.5*'权重计算验证'!$F$5*({variance})"]
                    for j,formula in enumerate(formulas):result_ws.write_formula(i-1,j,formula,precise if j in [3,5] else percent,float(row[j]))
                    for j,formula in [(7,f'=ABS(D{i}-MIN($D$2:$D${last_result}))<=1E-12'),(8,f'=ABS(F{i}-MAX($F$2:$F${last_result}))<=1E-12')]:result_ws.write_formula(i-1,j,formula,None,bool(row[j]))
                result_ws.write(0,6,'有效集（导出时快照）',header)
                book.get_worksheet_by_name(next(iter(tables))).activate()
                result_ws.set_selection('C2')
                ws=book.add_worksheet('权重计算验证');ws.hide();ws.hide_gridlines(2);ws.set_column('A:H',23);ws.set_column('I:M',24)
                title=book.add_format({'font_size':14,'bold':True,'font_color':'#19334D'})
                inp=book.add_format({'font_color':'#0000FF','bg_color':'#FFF1CD','num_format':'0.0000000000E+00'})
                ws.write('A1','权重计算验证',title)
                ws.write('A2','均值、标准差及协方差引用收益明细计算；修改A或收益明细可重算权重结果。有效集和其他汇总为导出时快照，需重新导出更新。')
                for j,label in enumerate(['μ1','μ2','σ1','σ2','Cov12','A','年化乘数K']):ws.write(3,j,label,header)
                values=[st['m1'],st['m2'],st['s1'],st['s2'],st['cov'],A,K]
                for j,v in enumerate(values):
                    if j<5:
                        source=['B','C','F','G','H'][j]
                        ws.write_formula(4,j,f"='个股风险及协方差'!{source}2",precise,float(v))
                    else:put(ws,4,j,v,inp)
                ws.data_validation('F5',{'validate':'decimal','criteria':'>=','value':0})
                # 标准差由共同样本公式计算。
                ws.data_validation('G5',{'validate':'decimal','criteria':'>','value':0})
                # 协方差与标准差引用同一组收益样本。
                ws.write('I4','等分数量（改变后需重新导出）',header);ws.write_number('I5',div)
                ws.write('A7','K：日频用完整年度市场交易日均值，月频12，年频1。空白表示日历不足，年化结果留空。')
                ws.write('I7','并列时INDEX/MATCH显示首个精确极值；全部容差并列见Python结果表。')
                for j,h in enumerate(['股票1权重','股票2权重','组合平均收益率','组合方差','组合标准差','效用U','简单年化预期收益率','年化标准差']):ws.write(8,j,h,header)
                ws.set_row(8,34);ws.freeze_panes(9,2);ws.autofilter(8,0,len(grid)+8,7)
                cached_values=[]
                for i,w in enumerate(grid['股票1权重']):
                    r=i+10;v=independent_values(detail,float(w),A,K);cached_values.append(v)
                    for j,(formula,value) in enumerate(zip(weight_formulas(r),v)):ws.write_formula(r-1,j,formula,precise if j in [3,5] else percent,value)
                last=len(grid)+9
                for label,cell,formula,value in [('最低方差','J5',f'=MIN(D10:D{last})',min(v[3] for v in cached_values)),('最高效用','K5',f'=MAX(F10:F{last})',max(v[5] for v in cached_values))]:
                    ws.write(3,9 if cell=='J5' else 10,label,header);ws.write_formula(cell,formula,precise,value)
                ws.write('I6','对应股票1权重（精确极值首个）')
                ws.set_row(3,36)
                for col,metric,target in [('J','D','J5'),('K','F','K5')]:
                    pos=int(np.argmin([v[3] for v in cached_values]) if col=='J' else np.argmax([v[5] for v in cached_values]))
                    ws.write_formula(f'{col}6',f'=INDEX($A$10:$A${last},MATCH({target},${metric}$10:${metric}${last},0))',percent,float(grid.iloc[pos]['股票1权重']))
        verify_export(tmp,verification[0] if verification else None,verification[1] if verification else None)
        os.replace(tmp,path)
    finally:
        if os.path.exists(tmp):os.unlink(tmp)

if __name__=='__main__':
    if len(sys.argv)>1 and sys.argv[1]=='--api-worker':
        try:
            request=json.loads(sys.argv[2]);payload=_network_fetch(**request)
        except Exception as e:payload={'error':str(e)}
        print('PORTFOLIO_JSON:'+json.dumps(payload,ensure_ascii=False))
    else:
        root=tk.Tk();App(root);root.mainloop()
