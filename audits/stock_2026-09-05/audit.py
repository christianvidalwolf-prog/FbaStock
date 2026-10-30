"""Read-only audit of supplier snapshots and Amazon exports; production downloads disabled."""
import sys, pathlib, json, csv, collections, hashlib, contextlib, io
ROOT=pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
import pandas as pd
import auto_update_stock_fast as m
OUT=pathlib.Path(__file__).resolve().parent
paths={'dcasa':str(ROOT/'DataWeb20260904.csv'),'mina':str(ROOT/'minerales_feed.xml'),'signes':str(ROOT/'signes_stock.csv')}
catalog=json.loads((ROOT/'catalog_robust.json').read_text())
prices=json.loads((ROOT/'excel_prices_final.json').read_text())
by_sku={}
for e in catalog: by_sku.setdefault(e['sku'],e)
feeds={}
dc=pd.read_csv(paths['dcasa'],sep=';',encoding='latin1');dc.columns=dc.columns.str.strip()
for _,r in dc.iterrows(): feeds.setdefault('DC',{})[str(r['CODIGO']).replace('.0','')]=(m.to_num(r['STOCK_DISPONIBLE']),m.to_num(r['Tarifa A']))
sg=pd.read_csv(paths['signes'],sep=';',encoding='latin1')
for _,r in sg.iterrows(): feeds.setdefault('SG',{})[str(r.iloc[0]).replace('SG-','').strip()]=(m.to_num(r.iloc[2]),m.to_num(r.iloc[3]))
vc=pd.read_csv(paths['mina'],sep='|',header=None)
for _,r in vc.iterrows(): feeds.setdefault('VC',{})[str(r.iloc[0]).replace('.0','')]=(m.to_num(r.iloc[7]),0)
feeds['MD']={k:(v,0) for k,v in m.load_trediser_stocks(str(ROOT/'STOCK TREDISER.xls')).items()}
def read_export(p):
 rows=list(csv.reader(p.open(encoding='utf-8-sig'),delimiter='\t'))
 ix=2 if any('fulfillment_availability#1.quantity' in r for r in rows[:7]) else 4
 return {r[0]:int(r[ix]) for i,r in enumerate(rows) if len(r)>ix and r[ix].isdigit() and not (i<6 and r[0]=='ABC123')}
actual=read_export(ROOT/'STOCK AMZ.txt')
# Replay exact production generator without network calls or production file writes.
m.download_files=lambda: paths
m.OUTPUT_FILE=str(OUT/'replayed_STOCK_AMZ.txt')
m.shutil.copy2=lambda *a,**kw: None
with contextlib.redirect_stdout(io.StringIO()) as log: m.run_fast_update()
(OUT/'replay.log').write_text(log.getvalue())
replay=read_export(pathlib.Path(m.OUTPUT_FILE))
assert replay==actual, 'Replay differs from current export'
records=[]
forced={str(s).upper().strip() for s in m.SKUS_FORZAR_CERO}
for sku,qty in actual.items():
 e=by_sku.get(sku,{'provider':'SG' if sku.endswith('SGRI') else 'Unknown','id':sku[:-4] if sku.endswith('SGRI') else sku})
 provider=e['provider'];ref=e['id'].replace('.0',''); source=feeds.get(provider,{})
 raw,cost=source.get(ref,(0,0));div=prices.get({'DC':'DcasaWeb','SG':'Signes','VC':'INICIOVC','MD':'Madelcar'}.get(provider,''),{}).get(e['id'],{})
 div=div.get('divisor',1) if isinstance(div,dict) else 1
 reason='exportado_con_stock' if qty>0 else 'stock_origen_cero'
 if qty==0:
  if provider=='Unknown':reason='proveedor_desconocido'
  elif ref not in source:reason='referencia_ausente_del_fichero'
  elif raw>0:
   if any(sku.upper().strip().startswith(s) for s in forced) or (provider=='VC' and ref in {'4085','5324','2598'}) or (provider=='SG' and '83627' in sku):reason='exclusion_explicita'
   elif provider=='VC' and raw<5:reason='umbral_minerales_5'
   elif provider in {'SG','DC'} and raw<=3:reason='umbral_stock_4'
   elif provider in {'SG','DC'} and raw<20 and cost*(div if provider=='DC' else 1)<20:reason='umbral_20_unidades_o_20_euros'
   else:reason='otra_causa_revisar'
 records.append(dict(sku=sku,proveedor=provider,referencia=ref,stock_origen=raw,coste_origen=cost,divisor=div,stock_amazon=qty,motivo=reason))
with (OUT/'detalle_skus.csv').open('w',newline='',encoding='utf-8-sig') as f:
 w=csv.DictWriter(f,fieldnames=list(records[0]));w.writeheader();w.writerows(records)
summary={}
for provider,feed in feeds.items():
 rec=[r for r in records if r['proveedor']==provider];mapped={r['referencia'] for r in rec}
 summary[provider]={'referencias_origen':len(feed),'referencias_origen_con_stock':sum(v[0]>0 for v in feed.values()),'sku_exportados':len(rec),'sku_con_stock':sum(r['stock_amazon']>0 for r in rec),'motivos':dict(collections.Counter(r['motivo'] for r in rec)),'referencias_con_stock_sin_sku_exportado':sum(v[0]>0 and k not in mapped for k,v in feed.items())}
history=[]
for p in sorted(pathlib.Path('/Users/christianvidalwolf/Library/CloudStorage/Dropbox/Christian/Ficheros Stock').glob('STOCK AMZ*2026090*')):
 d=read_export(p);count=collections.Counter(by_sku.get(s,{'provider':'SG' if s.endswith('SGRI') else 'Unknown'})['provider'] for s,q in d.items() if q>0)
 history.append({'file':p.name,'total':sum(count.values()),'providers':dict(count),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
result={'replay_exact_match':True,'total_skus':len(actual),'positive_skus':sum(q>0 for q in actual.values()),'providers':summary,'history':history,'source_files':{str(p.relative_to(ROOT)):{'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in [ROOT/'STOCK AMZ.txt',*[pathlib.Path(p) for p in paths.values()],ROOT/'STOCK TREDISER.xls']}}
(OUT/'summary.json').write_text(json.dumps(result,indent=2,ensure_ascii=False));print(json.dumps(result,indent=2,ensure_ascii=False))
