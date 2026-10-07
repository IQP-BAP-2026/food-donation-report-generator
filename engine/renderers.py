from pathlib import Path
import tempfile,asyncio,os,base64,mimetypes,json,sys,io
from jinja2 import Environment,FileSystemLoader
try:from playwright.async_api import async_playwright
except Exception:async_playwright=None
try:from weasyprint import HTML as WeasyHTML
except Exception:WeasyHTML=None
BASE=Path(__file__).resolve().parent.parent
APP_DIR=Path(sys.executable).resolve().parent if getattr(sys,'frozen',False) else BASE
EXTERNAL_MONEY_TEMPLATES=APP_DIR/'templates'/'money';BUNDLED_MONEY_TEMPLATES=BASE/'templates'/'money'
def money_templates_root():return EXTERNAL_MONEY_TEMPLATES if EXTERNAL_MONEY_TEMPLATES.exists() else BUNDLED_MONEY_TEMPLATES
def discover_money_templates():
 root=money_templates_root();out=[]
 if not root.exists():return out
 for folder in sorted((x for x in root.iterdir() if x.is_dir()),key=lambda x:x.name.lower()):
  try:
   cfg=json.loads((folder/'template.json').read_text(encoding='utf-8'));entry=cfg.get('entry','report.html');css=cfg.get('stylesheet','style.css')
   if not (folder/entry).exists() or not (folder/css).exists():continue
   out.append({'id':folder.name,'name':cfg.get('name',folder.name),'description':cfg.get('description',''),'path':str(folder)})
  except Exception:continue
 return out
def validate_money_template(tid):
 folder=money_templates_root()/tid
 if not (folder/'template.json').exists():raise ValueError(f'No se encontró la plantilla monetaria {tid}.')
 cfg=json.loads((folder/'template.json').read_text(encoding='utf-8'))
 for x in (cfg.get('entry','report.html'),cfg.get('stylesheet','style.css')):
  if not (folder/x).exists():raise ValueError(f'La plantilla monetaria no contiene {x}.')
 return folder,cfg
def _len(v):return len(str(v or '').strip())
def _range(label,v,lo,hi):
 n=_len(v)
 if n<lo or n>hi:raise ValueError(f'{label}: use entre {lo} y {hi} caracteres (actualmente {n}).')
 # Very long unbroken tokens are unsafe in fixed Canva text boxes.
 for token in str(v or '').split():
  if len(token)>55:raise ValueError(f'{label}: contiene una palabra o secuencia demasiado larga. Agregue espacios o acórtela.')
# ---- V4: cierres, indicadores e iconos ----
CLOSING_STYLES={'closing_a':('Ondas y corazones',5),'closing_b':('Marco verde y franjas',5),'closing_c':('Hexágonos',3)}
KPI_LIMITS={'results_photos':(2,5),'results_numbers':(1,4)}
ICON_EXTS=('.png','.jpg','.jpeg','.webp','.svg')
ICON_DIR_DEFAULT=BASE/'assets'/'icons'   # iconos incluidos; la aplicación mantiene además una carpeta editable (data['icon_dir'])
def resolve_icon(value,icon_dir=None):
 """Nombre de archivo (en la carpeta de iconos) o ruta completa -> Path existente, o None."""
 v=str(value or '').strip()
 if not v:return None
 p=Path(v)
 if p.is_absolute():return p if p.exists() else None
 for base in ([Path(icon_dir)] if icon_dir else [])+[ICON_DIR_DEFAULT]:
  q=base/v
  if q.exists():return q
 return None
# Rangos permitidos (pt) para tamaños de fuente elegidos manualmente. Fuera de rango se rechaza antes de generar el PDF.
MANUAL_RANGES={'title':(8,48),'text':(6,32),'kpi_value':(8,60),'kpi_label':(5,24),'cover_title':(10,60),'donor':(7,36),'tagline':(6,24),'closing':(10,48)}
def _check_manual(label,mf):
 for k,v in (mf or {}).items():
  if k not in MANUAL_RANGES:continue
  lo,hi=MANUAL_RANGES[k]
  try:x=float(v)
  except Exception:raise ValueError(f'{label}: el tamaño de fuente manual "{v}" no es un número.')
  if not lo<=x<=hi:raise ValueError(f'{label}: el tamaño de fuente manual debe estar entre {lo} y {hi} pt (actualmente {x:g}).')
def validate_money_data(d,page_only=False):
 if not page_only:
  _range('Nombre del donante',d.get('donor'),2,80);_range('Título del informe',d.get('report_title'),3,70);_range('Frase / slogan',d.get('tagline'),5,100);_range('Mensaje de cierre',d.get('closing_title'),5,120)
  if len(d.get('cover_photos',[]))!=4:raise ValueError('La portada requiere exactamente 4 fotografías.')
  style=d.get('closing_style') or 'closing_a'
  if style not in CLOSING_STYLES:raise ValueError('Diseño de página de cierre no reconocido.')
  need=CLOSING_STYLES[style][1]
  if len(d.get('closing_photos',[]))!=need:raise ValueError(f'El cierre "{CLOSING_STYLES[style][0]}" requiere exactamente {need} fotografías (hay {len(d.get("closing_photos",[]))}).')
  if not d.get('pages'):raise ValueError('Agregue al menos una página de contenido al informe monetario.')
 if not page_only:_check_manual('Portada / cierre',d.get('manual_font'))
 for idx,p in enumerate(d.get('pages',[]),2):
  pre=f'Página {idx}' if not page_only else 'Página';_check_manual(pre,p.get('manual_font'));_range(f'{pre} - título',p.get('title'),3,70);typ=p.get('type')
  if p.get('program_logo') and not Path(str(p['program_logo'])).exists():raise ValueError(f'{pre}: no se encontró el archivo del logo del programa ({p["program_logo"]}).')
  rules={'text_photos':(40,500,3),'mixed_photos':(40,450,3),'results_photos':(30,350,3),'results_numbers':(20,300,0)}
  if typ in rules:
   lo,hi,photos=rules[typ];_range(f'{pre} - texto principal',p.get('text'),lo,hi)
   if len(p.get('photos',[]))!=photos:raise ValueError(f'{pre}: este diseño requiere exactamente {photos} fotografía(s).')
  elif typ=='gallery':
   cnt=int(p.get('gallery_count',0));
   if cnt not in (6,9,12):raise ValueError(f'{pre}: la galería debe ser de 6, 9 o 12 fotos.')
   if len(p.get('photos',[]))!=cnt:raise ValueError(f'{pre}: seleccionó galería de {cnt}; debe cargar exactamente {cnt} fotos.')
  else:raise ValueError(f'{pre}: tipo de página no compatible con la plantilla aprobada.')
  if typ in ('results_photos','results_numbers'):
   ks=p.get('kpis',[]);lo_k,hi_k=KPI_LIMITS[typ]
   if not lo_k<=len(ks)<=hi_k:raise ValueError(f'{pre}: este diseño requiere entre {lo_k} y {hi_k} indicadores (hay {len(ks)}).')
   for n,k in enumerate(ks,1):
    _range(f'{pre} - valor del indicador {n}',k[0],1,25);_range(f'{pre} - descripción del indicador {n}',k[1],2,60)
    if not page_only and len(k)>2 and str(k[2] or '').strip():
     ic=resolve_icon(k[2],d.get('icon_dir'))
     if ic is None:raise ValueError(f'{pre} - indicador {n}: no se encontró el icono "{k[2]}".')
     if ic.suffix.lower() not in ICON_EXTS:raise ValueError(f'{pre} - indicador {n}: formato de icono no compatible ({ic.suffix}). Use PNG, JPG, WEBP o SVG.')
 return True
async def _playwright_pdf(html,out):
 if async_playwright is None:raise RuntimeError('Playwright no está disponible')
 tmp=Path(tempfile.gettempdir())/'bap_preview.html';tmp.write_text(html,encoding='utf-8')
 async with async_playwright() as p:
  kw={}
  if os.name!='nt' and Path('/usr/bin/chromium').exists():kw['executable_path']='/usr/bin/chromium'
  b=await p.chromium.launch(**kw);page=await b.new_page(viewport={'width':900,'height':1800});await page.goto(tmp.as_uri(),wait_until='networkidle');await page.emulate_media(media='print');await page.evaluate('document.fonts.ready.then(()=>true)');await page.pdf(path=str(out),print_background=True,prefer_css_page_size=True,margin={'top':'0','right':'0','bottom':'0','left':'0'});await b.close()
def render_html_pdf(html,out,base):
 try:asyncio.run(_playwright_pdf(html,out));return
 except Exception as first:
  if WeasyHTML is not None:WeasyHTML(string=html,base_url=str(base)).write_pdf(str(out));return
  raise RuntimeError(f'No se pudo renderizar el PDF: {first}')
def file_uri(path):
 if not path:return ''
 p=Path(path)
 if not p.exists():return ''
 mime=mimetypes.guess_type(str(p))[0] or 'application/octet-stream';return f'data:{mime};base64,'+base64.b64encode(p.read_bytes()).decode('ascii')
def _font_css(folder):
 """Incrusta (base64) las fuentes Avenir que el usuario coloque en templates/money/<plantilla>/fonts/ (o en ./fonts junto al EXE).
 Familia 'BAP Avenir'. El peso se deduce del nombre del archivo. Sin archivos: no agrega nada y se usa Avenir instalado en el sistema o una alternativa."""
 wmap=[('extralight',200),('ultralight',200),('light',300),('semibold',600),('demi',600),('medium',500),('heavy',800),('extrabold',800),('black',900),('bold',700),('roman',400),('book',400),('regular',400),('normal',400)]
 fmts={'.otf':('font/otf','opentype'),'.ttf':('font/ttf','truetype'),'.woff':('font/woff','woff'),'.woff2':('font/woff2','woff2')}
 seen={};rules=[]
 for d in (Path(folder)/'fonts',APP_DIR/'fonts',BASE/'fonts'):
  if not d.exists():continue
  for f in sorted(d.iterdir()):
   ext=f.suffix.lower()
   if ext not in fmts or 'serif' in f.stem.lower():continue
   n=f.stem.lower().replace('-','').replace('_','').replace(' ','')
   w=next((v for k,v in wmap if k in n),400);st='italic' if ('oblique' in n or 'italic' in n) else 'normal'
   if (w,st) in seen:continue
   seen[(w,st)]=1;mime,fmt=fmts[ext]
   try:b64=base64.b64encode(f.read_bytes()).decode('ascii')
   except Exception:continue
   rules.append("@font-face{font-family:'BAP Avenir';src:url(data:%s;base64,%s) format('%s');font-weight:%d;font-style:%s}"%(mime,b64,fmt,w,st))
 # Serif opcional (cierre de hexágonos): cualquier archivo cuyo nombre contenga "serif" (p. ej. DMSerifDisplay-Regular.ttf) -> familia 'BAP Serif'
 for d in (Path(folder)/'fonts',APP_DIR/'fonts',BASE/'fonts'):
  if not d.exists():continue
  for f in sorted(d.iterdir()):
   ext=f.suffix.lower()
   if ext in fmts and 'serif' in f.stem.lower():
    try:b64=base64.b64encode(f.read_bytes()).decode('ascii')
    except Exception:continue
    mime,fmt=fmts[ext];rules.append("@font-face{font-family:'BAP Serif';src:url(data:%s;base64,%s) format('%s');font-weight:400;font-style:normal}"%(mime,b64,fmt));break
  else:continue
  break
 return '\n'.join(rules)
def _trim_margins(im):
 """Recorta los márgenes vacíos de un logo (transparentes o blancos) para que el dibujo ocupe todo su espacio."""
 from PIL import Image,ImageChops
 try:
  if im.mode in ('RGBA','LA') or (im.mode=='P' and 'transparency' in im.info):
   a=im.convert('RGBA').split()[3];box=a.point(lambda v:255 if v>8 else 0).getbbox()
  else:
   rgb=im.convert('RGB');box=ImageChops.difference(rgb,Image.new('RGB',rgb.size,(255,255,255))).convert('L').point(lambda v:255 if v>14 else 0).getbbox()
  if not box:return im
  w,h=im.size;pad=max(2,int(max(box[2]-box[0],box[3]-box[1])*0.02));box=(max(0,box[0]-pad),max(0,box[1]-pad),min(w,box[2]+pad),min(h,box[3]+pad))
  if (box[2]-box[0])*(box[3]-box[1])>=0.97*w*h:return im   # casi sin margen: no tocar
  return im.crop(box)
 except Exception:return im
def logo_prepared(path,max_side=1600):
 """(data-URI, ancho en px) de un logo ya recortado. Para logos del donante y del programa."""
 if not path or not Path(str(path)).exists():return '',0
 try:
  from PIL import Image,ImageOps
  im=_trim_margins(ImageOps.exif_transpose(Image.open(path)));w=im.size[0]
  alpha=im.mode in ('RGBA','LA') or (im.mode=='P' and 'transparency' in im.info)
  if max(im.size)>max_side:im.thumbnail((max_side,max_side),Image.LANCZOS)
  buf=io.BytesIO()
  if alpha:im.convert('RGBA').save(buf,'PNG',optimize=True);mime='image/png'
  else:im.convert('RGB').save(buf,'JPEG',quality=92,optimize=True);mime='image/jpeg'
  return f'data:{mime};base64,'+base64.b64encode(buf.getvalue()).decode('ascii'),w
 except Exception:return image_uri(path,max_side),_natural_width_px(path)
def image_uri(path,max_side=2000,quality=90):
 """Photo/logo -> data URI. Applies EXIF rotation and downsizes very large images (phone photos are often 4000+ px),
 which keeps the PDF small and fast to render without visible quality loss at A4 print size."""
 if not path:return ''
 p=Path(path)
 if not p.exists():return ''
 try:
  from PIL import Image,ImageOps
  im=Image.open(p);im=ImageOps.exif_transpose(im)
  alpha=im.mode in ('RGBA','LA') or (im.mode=='P' and 'transparency' in im.info)
  if max(im.size)>max_side:im.thumbnail((max_side,max_side),Image.LANCZOS)
  buf=io.BytesIO()
  if alpha:im.convert('RGBA').save(buf,'PNG',optimize=True);mime='image/png'
  else:im.convert('RGB').save(buf,'JPEG',quality=quality,optimize=True);mime='image/jpeg'
  return f'data:{mime};base64,'+base64.b64encode(buf.getvalue()).decode('ascii')
 except Exception:
  return file_uri(path)
def _natural_width_px(path):
 try:
  from PIL import Image
  return Image.open(path).size[0]
 except Exception:return 600
def _kpi_boxes(cfg,typ,n):
 """Exact card rectangles [x,y,w,h] in mm for this page type and indicator count (from template.json)."""
 return (cfg.get('kpi_boxes',{}).get(typ,{}) or {}).get(str(n),[])
def icon_uri(path):
 p=Path(path)
 if p.suffix.lower()=='.svg':return 'data:image/svg+xml;base64,'+base64.b64encode(p.read_bytes()).decode('ascii')
 return image_uri(p,384) or ''
# Colores de las tarjetas por cantidad de indicadores (tal como en la guía de Canva)
PALETTE={5:['green','orange','blue','blue','green'],4:['green','orange','blue','green'],3:['green','orange','blue'],2:['green','orange'],1:['green']}
# Tamaños máximos / mínimos (pt) del autoajuste de valor y descripción
KPI_FONT={('results_photos',5):(16,7.5),('results_photos',4):(19,8.8),('results_photos',3):(21,9.5),('results_photos',2):(21,9.5),('results_numbers',1):(36,14),('results_numbers',2):(30,12),('results_numbers',3):(26,11),('results_numbers',4):(26,11)}
def generate_money(data,tid,out):
 validate_money_data(data);folder,cfg=validate_money_template(tid);env=Environment(loader=FileSystemLoader(str(folder)),autoescape=True);tpl=env.get_template(cfg.get('entry','report.html'))
 d=dict(data);d['css']=(folder/cfg.get('stylesheet','style.css')).read_text(encoding='utf-8');d['font_css']=_font_css(folder);d['manual']={k:v for k,v in (data.get('manual_font') or {}).items() if k in MANUAL_RANGES}
 logos=BASE/'assets'/'logos'
 d['logos']={'bap':file_uri(logos/'logo-bap.png'),'global':file_uri(logos/'logo-global-foodbanking.png'),'camara':file_uri(logos/'logo-camara.png'),'sumarse':file_uri(logos/'logo-sumarse.png')}
 raw_logo=d.get('donor_logo')
 # never show the donor logo larger than ~150 dpi allows: small logos stay crisp instead of being stretched
 d['donor_logo'],_dw=logo_prepared(raw_logo,1600)
 d['donor_logo_max_mm']=round(max(34,min(85,_dw/110*25.4)),1) if _dw else 60   # ~110 ppp mínimo efectivo
 logo_cache={}
 def logo_uri(path):
  if not path:return ''
  if path not in logo_cache:logo_cache[path]=logo_prepared(path,1200)[0]
  return logo_cache[path]
 style=d.get('closing_style') or 'closing_a';d['closing_style']=style
 d['cover_photos']=[image_uri(x) for x in d['cover_photos']];d['closing_photos']=[image_uri(x) for x in d['closing_photos']]
 pages=[];needed={(d.get('cover_variant') or 'cover_1')+'.png',style+'.png'}
 for p in d['pages']:
  q=dict(p);q['photos']=[image_uri(x) for x in p.get('photos',[])];q['photo_count']=len(q['photos']);q['indicator_count']=len(q.get('kpis',[]))
  n=_len(q.get('text'));tn=_len(q.get('title'))
  q['text_class']='body-xl' if n<=110 else 'body-lg' if n<=190 else 'body-md' if n<=300 else 'body-sm' if n<=400 else 'body-xs'
  q['title_class']='title-xl' if tn<=24 else 'title-lg' if tn<=42 else 'title-md' if tn<=58 else 'title-sm'
  q['manual']={k:v for k,v in (p.get('manual_font') or {}).items() if k in MANUAL_RANGES}
  t=q.get('type');lp=p.get('program_logo') or (d.get('program_logo') if p.get('show_program_logo',True) else '')
  q['program_logo_uri']=logo_uri(lp) if t!='gallery' else '';q['show_logo']=bool(q['program_logo_uri'])
  q['icons_white']=bool(p.get('icons_white',True))
  if t in ('results_photos','results_numbers'):
   ks=q.get('kpis',[]);cnt=len(ks);boxes=_kpi_boxes(cfg,t,cnt);pal=PALETTE.get(cnt,['green']*cnt);items=[]
   q['kpi_vmax'],q['kpi_lmax']=KPI_FONT.get((t,cnt),(20,9))
   for i,k in enumerate(ks):
    ic=resolve_icon(k[2],d.get('icon_dir')) if len(k)>2 and k[2] else None
    it={'value':k[0],'label':k[1],'color':pal[i],'icon':icon_uri(ic) if ic else ''}
    if i<len(boxes):
     x,y,w,h=boxes[i];dd=round(0.44*h,1);it.update(x=x,y=y,w=w,h=h,d=dd,pad=round(.48*dd,1),gap=round(.37*dd,1),vh=round(.46*h,1),lh=round(.34*h,1))
    items.append(it)
   q['kpi_items']=items
  if t=='gallery':needed.add(f"gallery_{q['photo_count']}.png")
  elif t in ('results_photos','results_numbers'):needed.add(f"{t}_{q['indicator_count']}.png")
  else:needed.add(f'{t}.png')
  pages.append(q)
 d['pages']=pages;rt=_len(d.get('report_title'));ct=_len(d.get('closing_title'))
 d['cover_title_class']='cover-title-xl' if rt<=24 else 'cover-title-lg' if rt<=45 else 'cover-title-sm'
 d['closing_title_class']='closing-title-xl' if ct<=28 else 'closing-title-lg' if ct<=65 else 'closing-title-sm'
 d['cover_variant']=d.get('cover_variant') or 'cover_1'
 # only the backgrounds this report actually uses are loaded and embedded
 d['backgrounds']={x.name:file_uri(x) for x in (folder/'backgrounds').glob('*.png') if x.name in needed}
 html=tpl.render(**d);render_html_pdf(html,out,folder);return Path(out)
