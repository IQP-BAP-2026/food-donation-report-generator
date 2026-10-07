import importlib,json,logging,os,queue,re,shutil,subprocess,sys,threading,traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
BASE=Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parent)); os.environ.setdefault('PLAYWRIGHT_BROWSERS_PATH',str(BASE/'ms-playwright'))
WORK=Path(os.environ.get('LOCALAPPDATA',Path.home()))/'BAP Report Generator V2.8'
# ---- FOOD: rutas y entorno (misma lógica que la aplicación original de alimentos) ----
FOOD_DATA=WORK/'alimentos_datos'   # copia editable de plantilla, CSS, Excel y assets de alimentos
FOOD_OUTPUT=Path.home()/'Documents'/'BAP Donor Traceability Reports'
if not (Path.home()/'Documents').exists():FOOD_OUTPUT=FOOD_DATA/'output'
FOOD_RESOURCE=BASE/'food'          # recursos incluidos en la aplicación
FOOD_FILES=['food-traceability-template.html','food-traceability.css','MASTER-SHEET.xlsx']
def configure_food_environment():
 """Indica al generador de alimentos dónde están los datos/salida y las DLL nativas de WeasyPrint."""
 os.environ.setdefault('BAP_REPORT_DATA_DIR',str(FOOD_DATA));os.environ.setdefault('BAP_REPORT_OUTPUT_DIR',str(FOOD_OUTPUT))
 bundled=BASE/'weasy_dlls'
 if bundled.exists():os.environ['WEASYPRINT_DLL_DIRECTORIES']=str(bundled)
 else:
  msys=Path(r'C:\msys64\ucrt64\bin')
  if msys.exists():os.environ.setdefault('WEASYPRINT_DLL_DIRECTORIES',str(msys))
def run_food_worker(donor):
 """Proceso hijo: genera UN reporte de alimentos con food_generator.py (sin modificar) e imprime la ruta del PDF."""
 os.environ['BAP_REPORT_DATA_DIR']=os.environ.get('BAP_REPORT_DATA_DIR',str(FOOD_DATA));os.environ['BAP_REPORT_OUTPUT_DIR']=os.environ.get('BAP_REPORT_OUTPUT_DIR',str(FOOD_OUTPUT))
 configure_food_environment();gen=importlib.import_module('food_generator');pdf=gen.generate_report_for_donor(donor);print(str(pdf),flush=True);return 0
if __name__=='__main__' and len(sys.argv)>=3 and sys.argv[1]=='--generate':
 raise SystemExit(run_food_worker(sys.argv[2]))
configure_food_environment()
for p_ in ['borradores/alimentos','borradores/monetarios','reportes','logs','datos','alimentos_datos','iconos','logos_programas']:(WORK/p_).mkdir(parents=True,exist_ok=True)
logging.basicConfig(filename=WORK/'logs'/'bap_v28.log',level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
from engine.renderers import generate_money,discover_money_templates,money_templates_root,validate_money_data,CLOSING_STYLES,KPI_LIMITS,ICON_EXTS
# ---- Biblioteca de iconos para los indicadores monetarios (carpeta editable) ----
ICON_DIR=WORK/'iconos'
def seed_icons():
 """Copia los iconos incluidos a la carpeta editable la primera vez (no sobrescribe nada)."""
 src=BASE/'assets'/'icons'
 if src.exists():
  for f in src.iterdir():
   if f.is_file() and not (ICON_DIR/f.name).exists():shutil.copy2(f,ICON_DIR/f.name)
seed_icons()
def list_icon_names():
 return sorted((f.name for f in ICON_DIR.iterdir() if f.is_file() and f.suffix.lower() in ICON_EXTS),key=str.casefold) if ICON_DIR.exists() else []
def import_icon_file(path):
 """Copia un icono elegido por el usuario a la biblioteca y devuelve su nombre de archivo."""
 src=Path(path)
 if src.suffix.lower() not in ICON_EXTS:raise ValueError('Formato de icono no compatible. Use PNG, JPG, WEBP o SVG.')
 ICON_DIR.mkdir(parents=True,exist_ok=True);dest=ICON_DIR/src.name
 if dest.exists() and dest.read_bytes()!=src.read_bytes():
  k=2
  while (ICON_DIR/f'{src.stem}_{k}{src.suffix}').exists():k+=1
  dest=ICON_DIR/f'{src.stem}_{k}{src.suffix}'
 if not dest.exists():shutil.copy2(src,dest)
 return dest.name
# ---- Biblioteca de logos de programas (carpeta editable; PNG con fondo transparente) ----
LOGO_DIR=WORK/'logos_programas'
LOGO_EXTS=('.png','.jpg','.jpeg','.webp')
def seed_program_logos():
 """Copia los logos preparados que trae el programa a la carpeta editable (no sobrescribe ni borra nada)."""
 src=BASE/'assets'/'program_logos'
 if src.exists():
  for f in src.iterdir():
   if f.is_file() and not (LOGO_DIR/f.name).exists():shutil.copy2(f,LOGO_DIR/f.name)
seed_program_logos()
def list_program_logos():
 return sorted((f.name for f in LOGO_DIR.iterdir() if f.is_file() and f.suffix.lower() in LOGO_EXTS),key=str.casefold) if LOGO_DIR.exists() else []
CLOSING_CHOICES=[('closing_a','Ondas y corazones (5 fotos)'),('closing_b','Marco verde y franjas (5 fotos)'),('closing_c','Hexágonos (3 fotos)')]
CLOSING_NEED={k:v[1] for k,v in CLOSING_STYLES.items()}
VERSION='4.3'; DARK='#29483A'; GREEN='#4F8554'; PINK='#EF88A2'; LG='#DDE8DC'; LP='#FBE9EE'
def slug(s):return re.sub(r'[^A-Za-z0-9_-]+','_',str(s)).strip('_') or 'reporte'
def open_path(p):
 if os.name=='nt':os.startfile(str(p))
 elif sys.platform=='darwin':subprocess.Popen(['open',str(p)])
 else:subprocess.Popen(['xdg-open',str(p)])
# ---- FOOD: ayudas de Excel y fotos (no tocan food_generator.py) ----
FOOD_PHOTO_SLOTS=[f'closing_photo{i}.jpg' for i in range(1,5)]+[f'custom_photo{i}.jpg' for i in range(1,4)]
FOOD_MIN_PHOTO_BYTES=10_000   # food_generator.py ignora las fotos personalizadas de 10 KB o menos
def _food_slug(v):
 import unicodedata
 t='' if v is None else str(v).strip().lower();t=''.join(c for c in unicodedata.normalize('NFD',t) if unicodedata.category(c)!='Mn');return ''.join(c for c in t if c.isalnum())
def food_detect_donor_column(columns):
 preferred={'donor','donante','empresa','company'};exact=[];partial=[]
 for col in columns:
  raw=str(col);key=_food_slug(raw)
  if key in preferred:exact.append(raw)
  elif any(t in key for t in preferred):partial.append(raw)
 return (exact or partial or [None])[0]
def food_workbook_info(path):
 """Hoja MASTER-{AÑO} más reciente (misma regla que food_generator.find_master_year_sheet) y lista de donantes."""
 import pandas as pd
 x=pd.ExcelFile(path);found=[]
 for name in x.sheet_names:
  m=re.match(r'^\s*MASTER-(\d{4})',str(name),re.IGNORECASE)
  if m:found.append((int(m.group(1)),name))
 if not found:raise ValueError('No se encontró una hoja MASTER-{AÑO}. Hojas disponibles: '+', '.join(map(str,x.sheet_names)))
 year,sheet=max(found,key=lambda i:i[0]);df=pd.read_excel(path,sheet_name=sheet);col=food_detect_donor_column(list(df.columns))
 if col is None:raise ValueError('No se encontró una columna DONOR / DONANTE / EMPRESA / COMPANY.')
 bad={'nan','none','donor','donante','empresa','company'}
 donors=sorted({str(v).strip() for v in df[col].dropna().tolist() if str(v).strip() and str(v).strip().lower() not in bad},key=str.casefold)
 return {'sheet':sheet,'year':year,'donors':donors}
def prepare_food_photos(paths):
 """Coloca exactamente 7 fotos en las 7 posiciones que usa food_generator.py:
 fotos 1-4 -> closing_photo1-4 (franja inferior, página 2); fotos 5-7 -> custom_photo1-3 (zona sin mapa regional).
 Los JPG se copian tal cual; PNG/WEBP se convierten a JPG real."""
 if len(paths)!=7:raise ValueError('Se requieren exactamente 7 fotos.')
 dest=FOOD_DATA/'assets'/'photos';dest.mkdir(parents=True,exist_ok=True)
 for n,(src,name) in enumerate(zip(paths,FOOD_PHOTO_SLOTS),1):
  src=Path(src);target=dest/name
  if not src.exists():raise FileNotFoundError(f'No se encontró la foto {n}: {src}')
  if src.suffix.lower() in ('.jpg','.jpeg'):shutil.copy2(src,target)
  else:
   from PIL import Image,ImageOps
   im=ImageOps.exif_transpose(Image.open(src))
   if im.mode in ('RGBA','LA','P'):
    rgba=im.convert('RGBA');bg=Image.new('RGB',rgba.size,'white');bg.paste(rgba,mask=rgba.split()[3]);im=bg
   else:im=im.convert('RGB')
   im.save(target,'JPEG',quality=92)
  if target.stat().st_size<=FOOD_MIN_PHOTO_BYTES:raise ValueError(f'La foto {n} ({src.name}) es demasiado pequeña (10 KB o menos). Use una fotografía de mayor tamaño.')

class App(tk.Tk):
 def __init__(self):
  super().__init__();self.title(f'Generador de Reportes BAP v{VERSION}');self.geometry('1220x820');self.minsize(1000,700);self.style=ttk.Style(self)
  try:self.style.theme_use('vista')
  except:pass
  self.style.configure('H2.TLabel',font=('Segoe UI',16,'bold'),foreground=GREEN);self.style.configure('Primary.TButton',font=('Segoe UI',10,'bold'),padding=(14,9))
  self.vars={};self.pages=[];self.report_type='';self.donor_logo='';self.cover_photos=[];self.closing_photos=[];self.food_photos=[];self.food_donor_photos={};self.excel_path='';self.food_busy=False;self._menu();self.home()
 def _menu(self):
  m=tk.Menu(self);a=tk.Menu(m,tearoff=0);a.add_command(label='Inicio',command=self.home);a.add_command(label='Abrir carpeta de reportes',command=lambda:open_path(WORK/'reportes'));m.add_cascade(label='Archivo',menu=a);h=tk.Menu(m,tearoff=0);h.add_command(label='Manual de usuario',command=self.open_manual);h.add_command(label='Registros',command=lambda:open_path(WORK/'logs'));h.add_command(label='Diagnóstico',command=self.diagnostics);m.add_cascade(label='Ayuda',menu=h);self.config(menu=m)
 def clear(self):
  for w in self.winfo_children():w.destroy()
  self.vars={}
 def open_manual(self):
  for p in [Path(sys.executable).resolve().parent/'Guia_Generador_Informes_Trazabilidad_BAP_Espanol.pdf',BASE/'Guia_Generador_Informes_Trazabilidad_BAP_Espanol.pdf']:
   if p.exists():open_path(p);return
  messagebox.showerror('Manual no encontrado','No se encontró el manual de usuario.')
 def open_templates(self):p=money_templates_root();p.mkdir(parents=True,exist_ok=True);open_path(p)
 def report_card(self,parent,text,sub,fill,command,width=330,height=150):
  c=tk.Canvas(parent,width=width,height=height,bg='white',highlightthickness=0,cursor='hand2')
  # Rounded rectangle card. The home screen is navigation only; report content stays in the workflows.
  r=26;x1,y1,x2,y2=5,5,width-5,height-5
  c.create_arc(x1,y1,x1+2*r,y1+2*r,start=90,extent=90,fill=fill,outline=fill)
  c.create_arc(x2-2*r,y1,x2,y1+2*r,start=0,extent=90,fill=fill,outline=fill)
  c.create_arc(x1,y2-2*r,x1+2*r,y2,start=180,extent=90,fill=fill,outline=fill)
  c.create_arc(x2-2*r,y2-2*r,x2,y2,start=270,extent=90,fill=fill,outline=fill)
  c.create_rectangle(x1+r,y1,x2-r,y2,fill=fill,outline=fill);c.create_rectangle(x1,y1+r,x2,y2-r,fill=fill,outline=fill)
  c.create_text(width/2,height/2-17,text=text,fill='white',font=('Segoe UI',19,'bold'),width=width-55,justify='center')
  c.create_text(width/2,height/2+28,text=sub,fill='white',font=('Segoe UI',10),width=width-65,justify='center')
  c.bind('<Button-1>',lambda e:command())
  return c
 def home(self):
  self.clear();self.configure(bg='white')
  top=tk.Frame(self,bg=GREEN,height=64);top.pack(fill='x');top.pack_propagate(False)
  tk.Label(top,text='BANCO DE ALIMENTOS PANAMÁ',bg=GREEN,fg='white',font=('Segoe UI',17,'bold')).pack(side='left',padx=30,pady=17)
  tk.Label(top,text=f'Generador de Reportes  •  v{VERSION}  •  Offline',bg=GREEN,fg='white',font=('Segoe UI',10)).pack(side='right',padx=30,pady=20)
  main=tk.Frame(self,bg='white');main.pack(fill='both',expand=True)
  center=tk.Frame(main,bg='white');center.place(relx=.5,rely=.48,anchor='center')
  tk.Label(center,text='¿QUÉ REPORTE DESEA CREAR?',bg='white',fg=DARK,font=('Segoe UI',25,'bold')).pack(pady=(0,32))
  cards=tk.Frame(center,bg='white');cards.pack()
  self.report_card(cards,'ALIMENTOS','Crear reporte de donación de alimentos',GREEN,self.food_workflow).pack(side='left',padx=22)
  self.report_card(cards,'MONETARIO','Crear reporte de donación monetaria',PINK,self.money_workflow).pack(side='left',padx=22)
  tk.Button(center,text='Abrir borrador guardado',command=self.open_draft,bg='white',fg=DARK,relief='flat',font=('Segoe UI',12,'bold'),highlightbackground=GREEN,highlightthickness=2,padx=24,pady=9).pack(pady=(34,0))
  footer=tk.Frame(self,bg=PINK,height=76);footer.pack(fill='x',side='bottom');footer.pack_propagate(False)
  inner=tk.Frame(footer,bg=PINK);inner.place(relx=.5,rely=.5,anchor='center')
  for label,cmd in [('Manual de usuario',self.open_manual),('Reportes guardados',lambda:open_path(WORK/'reportes')),('Plantillas monetarias',self.open_templates),('Configuración',self.diagnostics),('Salir',self.destroy)]:
   tk.Button(inner,text=label,command=cmd,bg=PINK,fg='white',activebackground=PINK,activeforeground='white',relief='flat',font=('Segoe UI',10,'bold'),padx=14).pack(side='left',padx=6)
 def shell(self,title):
  self.clear();top=ttk.Frame(self,padding=(22,14));top.pack(fill='x');ttk.Button(top,text='← Inicio',command=self.home).pack(side='left');ttk.Label(top,text=title,style='H2.TLabel').pack(side='left',padx=18);ttk.Button(top,text='Manual de usuario',command=self.open_manual).pack(side='right');body=ttk.Frame(self,padding=(22,0));body.pack(fill='both',expand=True);self.nb=ttk.Notebook(body);self.nb.pack(fill='both',expand=True);foot=ttk.Frame(self,padding=(22,12));foot.pack(fill='x');
  ttk.Button(foot,text='Guardar borrador',command=self.save_draft).pack(side='left');ttk.Button(foot,text='Abrir carpeta de reportes',command=lambda:open_path(WORK/'reportes')).pack(side='left',padx=8);ttk.Button(foot,text='Vista previa',command=self.preview).pack(side='right',padx=8);ttk.Button(foot,text='GENERAR PDF',command=self.generate,style='Primary.TButton').pack(side='right')
 def field(self,p,label,key,row,default=''):
  ttk.Label(p,text=label).grid(row=row,column=0,sticky='w',padx=7,pady=7);v=tk.StringVar(value=default);self.vars[key]=v;ttk.Entry(p,textvariable=v).grid(row=row,column=1,sticky='ew',padx=7,pady=7);p.columnconfigure(1,weight=1)
 def val(self,k):return self.vars.get(k,tk.StringVar()).get().strip()
 # FOOD  (interfaz de V3.0; el PDF lo produce food_generator.py del compañero, sin modificar)
 def food_copy_initial_resources(self):
  """Instala plantilla, CSS, Excel y assets incluidos en una carpeta editable. No sobrescribe lo que ya exista."""
  FOOD_DATA.mkdir(parents=True,exist_ok=True);(FOOD_DATA/'assets').mkdir(parents=True,exist_ok=True)
  # La plantilla y el CSS son código del generador: si el programa trae una versión nueva, se actualizan (el Excel y las fotos del usuario NO se tocan).
  import hashlib
  hh=hashlib.sha256()
  for name in ('food-traceability-template.html','food-traceability.css'):
   if (FOOD_RESOURCE/name).exists():hh.update((FOOD_RESOURCE/name).read_bytes())
  stamp=FOOD_DATA/'.codigo_alimentos.sha'
  if not stamp.exists() or stamp.read_text(encoding='utf-8').strip()!=hh.hexdigest():
   for name in ('food-traceability-template.html','food-traceability.css'):
    if (FOOD_RESOURCE/name).exists():shutil.copy2(FOOD_RESOURCE/name,FOOD_DATA/name)
   stamp.write_text(hh.hexdigest(),encoding='utf-8')
  for name in FOOD_FILES:
   src=FOOD_RESOURCE/name;dst=FOOD_DATA/name
   if src.exists() and not dst.exists():shutil.copy2(src,dst)
  src_assets=FOOD_RESOURCE/'assets'
  if src_assets.exists():
   for src in src_assets.rglob('*'):
    if not src.is_file():continue
    dst=FOOD_DATA/'assets'/src.relative_to(src_assets);dst.parent.mkdir(parents=True,exist_ok=True)
    if not dst.exists():shutil.copy2(src,dst)
  FOOD_OUTPUT.mkdir(parents=True,exist_ok=True)
 def food_workflow(self,data=None):
  d=data or {};self.report_type='food';self.food_copy_initial_resources();self.food_busy=False;self.excel_path=d.get('excel_path','');self.food_photos=d.get('shared_photos',[]);self.food_donor_photos=d.get('donor_photos',{});self.shell('Reporte de donación de alimentos')
  f=ttk.Frame(self.nb,padding=22);self.nb.add(f,text='Reporte de alimentos');ttk.Label(f,text='1. Archivo Excel',style='H2.TLabel').grid(row=0,column=0,columnspan=3,sticky='w');self.excel_label=tk.StringVar(value=self.excel_path or 'Ningún archivo seleccionado');ttk.Label(f,textvariable=self.excel_label,wraplength=850).grid(row=1,column=0,columnspan=2,sticky='w',pady=8);ttk.Button(f,text='Cargar Excel',command=self.choose_excel).grid(row=1,column=2,padx=8)
  ttk.Label(f,text='2. Donantes',style='H2.TLabel').grid(row=2,column=0,columnspan=3,sticky='w',pady=(18,4));ttk.Label(f,text='Seleccione uno o varios donantes. Se generará un PDF separado por donante.').grid(row=3,column=0,columnspan=3,sticky='w');self.donor_list=tk.Listbox(f,selectmode='extended',height=12,exportselection=False);self.donor_list.grid(row=4,column=0,columnspan=3,sticky='nsew',pady=8);self.donor_list.bind('<<ListboxSelect>>',lambda e:self.show_donor_photo_status());self.food_status=tk.StringVar();ttk.Label(f,textvariable=self.food_status).grid(row=5,column=0,columnspan=3,sticky='w')
  ttk.Label(f,text='3. Fotografías - exactamente 7 por reporte',style='H2.TLabel').grid(row=6,column=0,columnspan=3,sticky='w',pady=(18,4));self.photo_mode=tk.StringVar(value=d.get('photo_mode','shared'));ttk.Radiobutton(f,text='Usar las mismas 7 fotos para todos los donantes',variable=self.photo_mode,value='shared').grid(row=7,column=0,columnspan=3,sticky='w');ttk.Radiobutton(f,text='Usar 7 fotos diferentes para cada donante',variable=self.photo_mode,value='per_donor').grid(row=8,column=0,columnspan=3,sticky='w');ttk.Button(f,text='Seleccionar 7 fotos compartidas',command=self.choose_shared_food_photos).grid(row=9,column=0,sticky='w',pady=8);self.food_photo_status=tk.StringVar(value=f'{len(self.food_photos)}/7 seleccionadas');ttk.Label(f,textvariable=self.food_photo_status).grid(row=9,column=1,sticky='w');ttk.Button(f,text='Asignar 7 fotos al donante seleccionado',command=self.assign_donor_food_photos).grid(row=10,column=0,sticky='w');self.food_specific_status=tk.StringVar(value='');ttk.Label(f,textvariable=self.food_specific_status).grid(row=10,column=1,columnspan=2,sticky='w')
  ttk.Label(f,text='Orden de las fotos: las fotos 1 a 4 forman la franja inferior de la página 2; las fotos 5 a 7 se muestran en el centro de la página 2 cuando el donante no tiene mapa regional. Cada foto debe pesar más de 10 KB.',wraplength=850,foreground='#5d6964').grid(row=11,column=0,columnspan=3,sticky='w',pady=(6,0));f.rowconfigure(4,weight=1);f.columnconfigure(1,weight=1);self.refresh_donors(d.get('donors',[]))
 def choose_excel(self):
  p=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx *.xlsm *.xls')])
  if not p:return
  dest=WORK/'datos'/Path(p).name
  try:shutil.copy2(p,dest)
  except shutil.SameFileError:pass
  self.excel_path=str(dest);self.excel_label.set(self.excel_path);self.refresh_donors([])
 def refresh_donors(self,selected):
  self.donor_list.delete(0,'end')
  if not self.excel_path or not Path(self.excel_path).exists():self.food_status.set('Seleccione un Excel que contenga una hoja MASTER-{AÑO}.');return
  try:
   info=food_workbook_info(self.excel_path);self.food_year=str(info['year'])
   for x in info['donors']:self.donor_list.insert('end',x)
   for i,x in enumerate(info['donors']):
    if x in selected:self.donor_list.selection_set(i)
   self.food_status.set(f"Hoja: {info['sheet']} • Año: {info['year']} • {len(info['donors'])} donantes")
  except Exception as e:self.err(e)
 def selected_donors(self):return [self.donor_list.get(i) for i in self.donor_list.curselection()]
 def show_donor_photo_status(self):
  ds=self.selected_donors()
  if len(ds)==1:
   n=len(self.food_donor_photos.get(ds[0],[]));self.food_specific_status.set(f'{ds[0]}: {n}/7 fotos' if n else '')
 def pick_seven_photos(self,title):
  ps=filedialog.askopenfilenames(title=title,filetypes=[('Imágenes','*.png *.jpg *.jpeg *.webp')])
  if not ps:return None
  if len(ps)!=7:messagebox.showwarning('Se requieren exactamente 7 fotos',f'Seleccionó {len(ps)} foto(s). Cada reporte necesita exactamente 7, ni más ni menos.');return None
  return list(ps)
 def choose_shared_food_photos(self):
  ps=self.pick_seven_photos('Seleccione exactamente 7 fotos')
  if ps:self.food_photos=ps;self.food_photo_status.set('7/7 seleccionadas')
 def assign_donor_food_photos(self):
  ds=self.selected_donors()
  if len(ds)!=1:return messagebox.showwarning('Seleccione un donante','Seleccione exactamente un donante para asignarle sus 7 fotos.')
  ps=self.pick_seven_photos(f'Seleccione exactamente 7 fotos para {ds[0]}')
  if ps:self.food_donor_photos[ds[0]]=ps;self.food_specific_status.set(f'{ds[0]}: 7/7 fotos')
 def food_generate_one(self,donor,photos,excel_path,out_dir):
  """Genera UN PDF ejecutando food_generator.py (sin modificar) en un proceso aparte, igual que la aplicación del compañero."""
  try:shutil.copy2(excel_path,FOOD_DATA/'MASTER-SHEET.xlsx')
  except shutil.SameFileError:pass
  prepare_food_photos(photos);Path(out_dir).mkdir(parents=True,exist_ok=True)
  env=os.environ.copy();env['BAP_REPORT_DATA_DIR']=str(FOOD_DATA);env['BAP_REPORT_OUTPUT_DIR']=str(out_dir);env['PYTHONIOENCODING']='utf-8';env['PYTHONUTF8']='1'
  command=[sys.executable,'--generate',donor] if getattr(sys,'frozen',False) else [sys.executable,str(Path(__file__).resolve()),'--generate',donor]
  r=subprocess.run(command,env=env,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=120,cwd=str(BASE),creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
  if r.returncode!=0:
   full=(r.stderr or r.stdout or 'El generador terminó con un error.').strip();logging.error('food_generator falló para %s:\n%s',donor,full);lines=[x for x in full.splitlines() if x.strip()];raise RuntimeError(lines[-1][:400] if lines else full[-400:])
  for line in reversed(r.stdout.splitlines()):
   c=Path(line.strip())
   if c.suffix.lower()=='.pdf' and c.exists():return c
  raise RuntimeError('El PDF no fue creado.\n\n'+(r.stdout or '')[-800:])
 def food_run(self,donors,out_dir,d,preview=False):
  """Genera uno o varios PDFs en segundo plano con una ventana de progreso (la aplicación no se congela)."""
  if getattr(self,'food_busy',False):return messagebox.showinfo('Generando','Ya hay un reporte en proceso. Espere a que termine.')
  self.food_busy=True;win=tk.Toplevel(self);win.title('Generando reporte de alimentos');win.geometry('540x170');win.transient(self);win.grab_set();win.protocol('WM_DELETE_WINDOW',lambda:None)
  msg=tk.StringVar(value='Preparando...');ttk.Label(win,textvariable=msg,wraplength=500).pack(padx=20,pady=(26,12));bar=ttk.Progressbar(win,mode='indeterminate');bar.pack(fill='x',padx=20);bar.start(12)
  q=queue.Queue()
  def worker():
   ok=[];fails=[]
   for i,donor in enumerate(donors,1):
    q.put(('msg',f'Generando {i} de {len(donors)}: {donor}'))
    try:photos=d['shared_photos'] if d['photo_mode']=='shared' else d['donor_photos'][donor];ok.append(self.food_generate_one(donor,photos,d['excel_path'],out_dir))
    except subprocess.TimeoutExpired:fails.append((donor,'La generación tardó más de 2 minutos y fue detenida. Revise la configuración de PDF/WeasyPrint del equipo.'))
    except Exception as e:logging.error('food %s: %s\n%s',donor,e,traceback.format_exc());fails.append((donor,str(e)))
   q.put(('done',ok,fails))
  threading.Thread(target=worker,daemon=True).start()
  def poll():
   try:
    while True:
     item=q.get_nowait()
     if item[0]=='msg':msg.set(item[1])
     else:bar.stop();win.destroy();self.food_busy=False;self.food_finished(item[1],item[2],out_dir,preview);return
   except queue.Empty:pass
   self.after(200,poll)
  poll()
 def food_finished(self,ok,fails,out_dir,preview):
  detail='\n'.join(f'• {n}: {m[:300]}' for n,m in fails)
  if preview:
   if ok:
    p=WORK/'reportes'/'_VISTA_PREVIA_ALIMENTOS.pdf';shutil.copy2(ok[0],p);shutil.rmtree(out_dir,ignore_errors=True);open_path(p)
   else:messagebox.showerror('No se pudo crear la vista previa',detail or 'Error desconocido.')
   return
  if ok and not fails:messagebox.showinfo('Reportes generados',f'Se generaron {len(ok)} PDF(s) en:\n{out_dir}');open_path(out_dir)
  elif ok:messagebox.showwarning('Generación parcial',f'Se generaron {len(ok)} PDF(s) en:\n{out_dir}\n\nNo se pudieron generar:\n{detail}');open_path(out_dir)
  else:messagebox.showerror('No se pudo generar',detail or 'Error desconocido.')
 # MONEY
 def money_workflow(self,data=None):
  d=data or {};self.report_type='money';self.pages=d.get('pages',[])
  if d.get('program_logo'):  # borradores V4.0: el logo era único para todo el informe -> ahora vive en cada página de contenido
   for q in self.pages:
    if q.get('type')!='gallery' and not q.get('program_logo') and q.get('show_program_logo',True):q['program_logo']=d['program_logo']
  self.cover_variant=tk.StringVar(value=d.get('cover_variant','cover_1'));self.donor_logo=d.get('donor_logo','');self.cover_photos=d.get('cover_photos',[]);self.closing_photos=d.get('closing_photos',[]);self.closing_style=tk.StringVar(value=next((l for k,l in CLOSING_CHOICES if k==d.get('closing_style','closing_a')),CLOSING_CHOICES[0][1]));self.shell('Reporte de donación monetaria')
  t1=ttk.Frame(self.nb,padding=22);self.nb.add(t1,text='1. Portada');self.field(t1,'Donante / organización','donor',0,d.get('donor',''));self.field(t1,'Año','year',1,d.get('year',str(datetime.now().year)));self.field(t1,'Título del informe','report_title',2,d.get('report_title','Informe de Impacto y Resultados'));self.field(t1,'Frase / slogan','tagline',3,d.get('tagline','Alimentamos a más panameños cada día'));ttk.Label(t1,text='Diseño de portada').grid(row=4,column=0,sticky='w',padx=7,pady=7);ttk.Combobox(t1,textvariable=self.cover_variant,state='readonly',values=['cover_1','cover_2'],width=18).grid(row=4,column=1,sticky='w',padx=7,pady=7);ttk.Button(t1,text='Seleccionar logo del donante',command=self.choose_logo).grid(row=5,column=0,sticky='w',pady=8);self.logo_status=tk.StringVar(value=self.donor_logo or 'No seleccionado');ttk.Label(t1,textvariable=self.logo_status).grid(row=5,column=1,sticky='w');ttk.Button(t1,text='Seleccionar 4 fotos de portada',command=self.choose_cover_photos).grid(row=6,column=0,sticky='w');self.cover_status=tk.StringVar(value=f'{len(self.cover_photos)}/4 seleccionadas');ttk.Label(t1,textvariable=self.cover_status).grid(row=6,column=1,sticky='w');ttk.Label(t1,text='La portada requiere exactamente 4 fotografías.',foreground='#7A4E00',wraplength=850).grid(row=7,column=0,columnspan=2,sticky='w',pady=12)
  self.cover_font=self.font_override_box(t1,[('cover_title','Título de portada',10,60,27),('donor','Nombre del donante',7,36,16),('tagline','Frase / slogan',6,24,10.5)],d.get('manual_font',{}),8)
  t2=ttk.Frame(self.nb,padding=22);self.nb.add(t2,text='2. Páginas');head=ttk.Frame(t2);head.pack(fill='x');ttk.Label(head,text='Páginas del informe',style='H2.TLabel').pack(side='left');ttk.Button(head,text='+ Agregar página',command=self.add_page_dialog,style='Primary.TButton').pack(side='right');ttk.Label(t2,text='Cada tipo de página aplica automáticamente sus límites de texto, indicadores y fotografías.',wraplength=900).pack(anchor='w',pady=6);scrollwrap=ttk.Frame(t2);scrollwrap.pack(fill='both',expand=True);self.page_canvas=tk.Canvas(scrollwrap,highlightthickness=0);sb=ttk.Scrollbar(scrollwrap,orient='vertical',command=self.page_canvas.yview);self.page_canvas.configure(yscrollcommand=sb.set);sb.pack(side='right',fill='y');self.page_canvas.pack(side='left',fill='both',expand=True);self.page_list=ttk.Frame(self.page_canvas);self.page_window=self.page_canvas.create_window((0,0),window=self.page_list,anchor='nw');self.page_list.bind('<Configure>',lambda e:self.page_canvas.configure(scrollregion=self.page_canvas.bbox('all')));self.page_canvas.bind('<Configure>',lambda e:self.page_canvas.itemconfigure(self.page_window,width=e.width));self.page_canvas.bind_all('<MouseWheel>',lambda e:self.page_canvas.yview_scroll(int(-1*(e.delta/120)),'units') if self.report_type=='money' else None);self.refresh_page_list()
  tc=ttk.Frame(self.nb,padding=22);self.nb.add(tc,text='3. Cierre');ttk.Label(tc,text='Página de cierre',style='H2.TLabel').grid(row=0,column=0,columnspan=2,sticky='w',pady=(0,12));self.field(tc,'Mensaje de cierre','closing_title',1,d.get('closing_title','¡Gracias por su apoyo!'));ttk.Label(tc,text='El mensaje debe tener entre 5 y 120 caracteres. El año del informe se muestra automáticamente en el cierre.').grid(row=2,column=0,columnspan=2,sticky='w',pady=(0,12))
  ttk.Label(tc,text='Diseño de la página de cierre').grid(row=3,column=0,sticky='w',padx=7,pady=7);self.closing_combo=ttk.Combobox(tc,textvariable=self.closing_style,state='readonly',values=[l for k,l in CLOSING_CHOICES],width=34);self.closing_combo.grid(row=3,column=1,sticky='w',padx=7,pady=7);self.closing_combo.bind('<<ComboboxSelected>>',lambda e:self.update_closing_info())
  self.close_btn=ttk.Button(tc,text='Seleccionar fotos de cierre',command=self.choose_closing_photos);self.close_btn.grid(row=4,column=0,sticky='w',pady=8);self.close_status=tk.StringVar();ttk.Label(tc,textvariable=self.close_status).grid(row=4,column=1,sticky='w');self.close_note=tk.StringVar();ttk.Label(tc,textvariable=self.close_note,foreground='#7A4E00',wraplength=800).grid(row=5,column=0,columnspan=2,sticky='w',pady=12);tc.columnconfigure(1,weight=1);self.update_closing_info()
  self.closing_font=self.font_override_box(tc,[('closing','Mensaje de cierre',10,48,25)],d.get('manual_font',{}),6)
  t3=ttk.Frame(self.nb,padding=22);self.nb.add(t3,text='4. Diseño');self.money_templates=discover_money_templates();names=[x['name'] for x in self.money_templates];self.template_name_var=tk.StringVar(value=names[0] if names else '');ttk.Label(t3,text='Diseño monetario',style='H2.TLabel').pack(anchor='w');ttk.Label(t3,text='Esta plantilla se usa únicamente para los reportes monetarios. Los reportes de alimentos conservan el diseño de Cam.',wraplength=900).pack(anchor='w',pady=8);ttk.Combobox(t3,textvariable=self.template_name_var,state='readonly',values=names,width=70).pack(anchor='w');ttk.Button(t3,text='Abrir carpeta de plantillas',command=self.open_templates).pack(anchor='w',pady=10)

 def font_override_box(self,parent,fields,initial,row,win=None):
  """Tamaño de fuente manual (opcional). Apagado = ajuste automático. Al activarlo se pide confirmar una advertencia."""
  initial=initial or {}
  box=ttk.LabelFrame(parent,text='Tamaño de fuente (opcional)',padding=8);box.grid(row=row,column=0,columnspan=2,sticky='ew',pady=8)
  on=tk.BooleanVar(value=any(f[0] in initial for f in fields));spins=[];vars_={}
  def apply_state():
   for sp in spins:sp.configure(state='normal' if on.get() else 'disabled')
  def toggled():
   if on.get() and not messagebox.askyesno('¿Cambiar el tamaño de la fuente?','¿Está seguro de que desea cambiar el tamaño de la fuente?\n\nPor defecto el programa ajusta el tamaño automáticamente para que todo quepa en el diseño. Si lo cambia manualmente, el texto podría salirse de su zona, cortarse o verse desalineado en el PDF.\n\nSi eso ocurre, vuelva a desactivar esta opción para recuperar el ajuste automático.',icon='warning',default='no',parent=win or self):on.set(False)
   apply_state()
  ttk.Checkbutton(box,text='Ajustar el tamaño manualmente (desactiva el ajuste automático)',variable=on,command=toggled).grid(row=0,column=0,columnspan=3,sticky='w')
  ttk.Label(box,text='Recomendado: dejar desactivado. Tamaños en puntos (pt).',foreground='#7A4E00').grid(row=1,column=0,columnspan=3,sticky='w',pady=(2,6))
  for i,(key,label,lo,hi,default) in enumerate(fields):
   v=tk.StringVar(value=str(initial.get(key,default)));vars_[key]=(v,label,lo,hi)
   ttk.Label(box,text=f'{label} ({lo}–{hi} pt)').grid(row=2+i,column=0,sticky='w',padx=(0,10),pady=2)
   sp=ttk.Spinbox(box,from_=lo,to=hi,increment=0.5,textvariable=v,width=8);sp.grid(row=2+i,column=1,sticky='w',pady=2);spins.append(sp)
  apply_state()
  def get():
   if not on.get():return {}
   out={}
   for key,(v,label,lo,hi) in vars_.items():
    try:x=float(v.get().replace(',','.'))
    except Exception:raise ValueError(f'{label}: escriba un número (pt).')
    if not lo<=x<=hi:raise ValueError(f'{label}: use un tamaño entre {lo} y {hi} pt.')
    out[key]=round(x,1)
   return out
  return type('FontBox',(),{'get':staticmethod(get)})()

 def logo_warning(self,parent):
  """Aviso antes de elegir un logo que no es de la biblioteca. Devuelve True si el usuario decide continuar."""
  res={'go':False};w=tk.Toplevel(parent);w.title('Antes de usar otro logo');w.transient(parent);w.grab_set();w.resizable(False,False);f=ttk.Frame(w,padding=20);f.pack()
  ttk.Label(f,text='Use un logo preparado',style='H2.TLabel').pack(anchor='w')
  ttk.Label(f,text='Un logo que no esté preparado puede verse pequeño, con un recuadro blanco o borroso, y dar errores en el informe.\n\nAntes de continuar, lea en el manual de usuario la sección «Cómo preparar un logo del programa» (fondo transparente, recortado al dibujo, PNG de buena resolución).\n\nLos logos de la lista ya están preparados.',wraplength=470,justify='left').pack(anchor='w',pady=12)
  bar=ttk.Frame(f);bar.pack(fill='x',pady=(8,0))
  def go():res['go']=True;w.destroy()
  ttk.Button(bar,text='Abrir el manual',command=self.open_manual).pack(side='left');ttk.Button(bar,text='Cancelar',command=w.destroy).pack(side='right');ttk.Button(bar,text='Ya preparé mi logo, continuar',command=go,style='Primary.TButton').pack(side='right',padx=8)
  w.wait_window();return res['go']
 def choose_logo(self):
  p=filedialog.askopenfilename(filetypes=[('Imágenes','*.png *.jpg *.jpeg *.webp')]);
  if p:self.donor_logo=p;self.logo_status.set(p)
 def choose_cover_photos(self):
  ps=filedialog.askopenfilenames(filetypes=[('Imágenes','*.png *.jpg *.jpeg *.webp')]);
  if ps:self.cover_photos=list(ps);self.cover_status.set(f'{len(ps)}/4 seleccionadas')
 def closing_key(self):return next((k for k,l in CLOSING_CHOICES if l==self.closing_style.get()),'closing_a')
 def update_closing_info(self):
  k=self.closing_key();n=CLOSING_NEED[k];have=len(self.closing_photos);state='✓' if have==n else '⚠ debe seleccionar exactamente %d'%n
  self.close_status.set(f'{have}/{n} seleccionadas  {state}')
  self.close_btn.configure(text=f'Seleccionar / reemplazar las {n} fotos de cierre')
  self.close_note.set('Este diseño usa exactamente 5 fotografías. Orden: foto 1 = vertical (izquierda), foto 2 = cuadrada, foto 3 = horizontal (abajo a la derecha), fotos 4 y 5 = horizontales (arriba).' if n==5 else 'Este diseño usa exactamente 3 fotografías en hexágonos: foto 1 = hexágono grande (izquierda), fotos 2 y 3 = hexágonos de la derecha (arriba y abajo).')
 def choose_closing_photos(self):
  n=CLOSING_NEED[self.closing_key()];ps=filedialog.askopenfilenames(title=f'Seleccione exactamente {n} fotos',filetypes=[('Imágenes','*.png *.jpg *.jpeg *.webp')])
  if not ps:return
  if len(ps)!=n:return messagebox.showwarning(f'Se requieren exactamente {n} fotos',f'Seleccionó {len(ps)} foto(s). Este diseño de cierre necesita exactamente {n}, ni más ni menos.')
  self.closing_photos=list(ps);self.update_closing_info()
 def default_page(self,typ,label):return {'type':typ,'title':label,'text':'','kpis':[],'photos':[],'gallery_count':6,'program_logo':'','icons_white':True}
 def page_status(self,p):
  typ=p.get('type');problems=[]
  title=str(p.get('title','')).strip()
  if not 3<=len(title)<=70:problems.append('título')
  rules={'text_photos':(40,500,3),'mixed_photos':(40,450,3),'results_photos':(30,350,3),'results_numbers':(20,300,0)}
  if typ in rules:
   lo,hi,n=rules[typ];L=len(str(p.get('text','')).strip())
   if not lo<=L<=hi:problems.append(f'texto {L}/{hi}')
   if len(p.get('photos',[]))!=n:problems.append(f'fotos {len(p.get("photos",[]))}/{n}')
  if typ=='gallery':
   n=int(p.get('gallery_count',6));
   if len(p.get('photos',[]))!=n:problems.append(f'fotos {len(p.get("photos",[]))}/{n}')
  if typ in KPI_LIMITS and not KPI_LIMITS[typ][0]<=len(p.get('kpis',[]))<=KPI_LIMITS[typ][1]:problems.append('%d-%d indicadores'%KPI_LIMITS[typ])
  return ('✓ Completa' if not problems else '⚠ Incompleta - '+', '.join(problems))
 def refresh_page_list(self):
  for w in self.page_list.winfo_children():w.destroy()
  labels={'text_photos':'Texto + 3 fotografías','mixed_photos':'Texto + fotos de distintos tamaños','results_photos':'Resultados + fotografías (2-5 indicadores)','results_numbers':'Resultados numéricos grandes (1-4 indicadores)','gallery':'Galería de fotografías'}
  for i,p in enumerate(self.pages):
   card=ttk.LabelFrame(self.page_list,text=f'Página {i+2} · {labels.get(p.get("type"),p.get("type"))}',padding=10);card.pack(fill='x',pady=5);ttk.Label(card,text=f'{p.get("title","Sin título")}  •  {len(p.get("kpis",[]))} indicadores  •  {len(p.get("photos",[]))} fotos  •  {self.page_status(p)}').pack(side='left');
   for txt,cmd in [('Editar',lambda n=i:self.edit_page(n)),('Duplicar',lambda n=i:self.duplicate_page(n)),('↑',lambda n=i:self.move_page(n,-1)),('↓',lambda n=i:self.move_page(n,1)),('Eliminar',lambda n=i:self.delete_page(n))]:ttk.Button(card,text=txt,command=cmd).pack(side='right',padx=2)
 def add_page_dialog(self):
  w=tk.Toplevel(self);w.title('Agregar página');w.geometry('560x360');w.transient(self);w.grab_set();f=ttk.Frame(w,padding=20);f.pack(fill='both',expand=True);ttk.Label(f,text='¿Qué tipo de página desea agregar?',style='H2.TLabel').pack(anchor='w');opts=[('text_photos','Texto + 3 fotografías'),('mixed_photos','Texto + fotografías de distintos tamaños'),('results_photos','Resultados + fotografías (2 a 5 indicadores con icono)'),('results_numbers','Resultados numéricos grandes (1 a 4 indicadores con icono)'),('gallery','Galería de fotografías')];cb=ttk.Combobox(f,state='readonly',values=[x[1] for x in opts],width=55);cb.current(0);cb.pack(fill='x',pady=14);ttk.Label(f,text='La página mostrará solamente los campos compatibles con el diseño seleccionado. Los requisitos de fotos, texto e indicadores se validan antes de guardar y antes de generar el PDF.',wraplength=500).pack(anchor='w')
  def add():typ,label=next(x for x in opts if x[1]==cb.get());self.pages.append(self.default_page(typ,label));w.destroy();self.edit_page(len(self.pages)-1)
  ttk.Button(f,text='Agregar y editar',command=add,style='Primary.TButton').pack(anchor='e',pady=18)
 def duplicate_page(self,i):self.pages.insert(i+1,json.loads(json.dumps(self.pages[i])));self.refresh_page_list()
 def move_page(self,i,d):
  j=i+d
  if 0<=j<len(self.pages):self.pages[i],self.pages[j]=self.pages[j],self.pages[i];self.refresh_page_list()
 def delete_page(self,i):
  if messagebox.askyesno('Eliminar página','¿Eliminar esta página?'):self.pages.pop(i);self.refresh_page_list()
 def edit_page(self,i):
  p=self.pages[i];typ=p['type'];w=tk.Toplevel(self);w.title('Editar página');w.geometry('980x860');w.transient(self);outer=ttk.Frame(w);outer.pack(fill='both',expand=True);cv=tk.Canvas(outer,highlightthickness=0);vsb=ttk.Scrollbar(outer,orient='vertical',command=cv.yview);cv.configure(yscrollcommand=vsb.set);vsb.pack(side='right',fill='y');cv.pack(side='left',fill='both',expand=True);frm=ttk.Frame(cv,padding=18);wid=cv.create_window((0,0),window=frm,anchor='nw');frm.bind('<Configure>',lambda e:cv.configure(scrollregion=cv.bbox('all')));cv.bind('<Configure>',lambda e:cv.itemconfigure(wid,width=e.width));title=tk.StringVar(value=p.get('title',''));ttk.Label(frm,text='Título (3–70 caracteres)').grid(row=0,column=0,sticky='w');ttk.Entry(frm,textvariable=title).grid(row=0,column=1,sticky='ew',pady=5);text=None
  if typ!='gallery':
   limits={'text_photos':(40,500),'mixed_photos':(40,450),'results_photos':(30,350),'results_numbers':(20,300)}[typ];ttk.Label(frm,text=f'Texto principal ({limits[0]}–{limits[1]} caracteres)').grid(row=1,column=0,sticky='nw');text=tk.Text(frm,height=8,wrap='word');text.insert('1.0',p.get('text',''));text.grid(row=1,column=1,sticky='nsew',pady=5);counter=tk.StringVar();ttk.Label(frm,textvariable=counter).grid(row=2,column=1,sticky='e');text.bind('<KeyRelease>',lambda e:counter.set(f'{len(text.get("1.0","end-1c"))}/{limits[1]} caracteres'));counter.set(f'{len(p.get("text",""))}/{limits[1]} caracteres')
  kvars=[];row=3;prog=tk.StringVar(value=p.get('program_logo',''));icons_white=tk.BooleanVar(value=p.get('icons_white',True))
  if typ!='gallery':
   lf=ttk.LabelFrame(frm,text='Logo del programa junto al título (opcional)',padding=8);lf.grid(row=row,column=0,columnspan=2,sticky='ew',pady=(4,6));prog_status=tk.StringVar(value=(Path(prog.get()).name if prog.get() and Path(prog.get()).parent==LOGO_DIR else prog.get()) or 'Sin logo: el título se centra solo.');lib_var=tk.StringVar(value=('(sin logo)' if not prog.get() else (Path(prog.get()).name if Path(prog.get()).parent==LOGO_DIR else '(otro archivo)')))
   def set_logo():
    if not self.logo_warning(w):return
    f=filedialog.askopenfilename(parent=w,title='Seleccionar logo del programa',filetypes=[('Imágenes','*.png *.jpg *.jpeg *.webp')])
    if f:prog.set(f);prog_status.set(f);lib_var.set('(otro archivo)')
   def pick_library(*_):
    v=lib_var.get()
    if v in ('(sin logo)','(otro archivo)'):
     if v=='(sin logo)':clear_logo()
     return
    prog.set(str(LOGO_DIR/v));prog_status.set(v)
   def clear_logo():prog.set('');prog_status.set('Sin logo: el título se centra solo.');lib_var.set('(sin logo)')
   def apply_all():
    if not prog.get():return messagebox.showinfo('Logo del programa','Primero seleccione un logo para esta página.',parent=w)
    if messagebox.askyesno('Aplicar a todas las páginas','¿Usar este logo en todas las páginas de contenido (no incluye galerías)? Reemplaza el logo que tengan ahora.',parent=w):
     for q in self.pages:
      if q.get('type')!='gallery':q['program_logo']=prog.get()
     messagebox.showinfo('Logo del programa','Listo. Se aplicó a todas las páginas de contenido.',parent=w)
   def lib_values():return ['(sin logo)']+list_program_logos()+['(otro archivo)']
   ttk.Label(lf,text='Logos de programas (ya preparados):').grid(row=0,column=0,sticky='w');lib_cb=ttk.Combobox(lf,textvariable=lib_var,state='readonly',values=lib_values(),width=44);lib_cb.configure(postcommand=lambda:lib_cb.configure(values=lib_values()));lib_cb.grid(row=0,column=1,columnspan=2,sticky='w',padx=6);lib_cb.bind('<<ComboboxSelected>>',pick_library)
   ttk.Button(lf,text='Abrir carpeta de logos',command=lambda:(LOGO_DIR.mkdir(parents=True,exist_ok=True),open_path(LOGO_DIR))).grid(row=0,column=3,padx=6)
   ttk.Button(lf,text='Usar otro archivo…',command=set_logo).grid(row=1,column=0,sticky='w',pady=(8,0));ttk.Button(lf,text='Quitar',command=clear_logo).grid(row=1,column=1,sticky='w',padx=6,pady=(8,0));ttk.Button(lf,text='Usar en todas las páginas de contenido',command=apply_all).grid(row=1,column=2,columnspan=2,sticky='w',padx=6,pady=(8,0))
   ttk.Label(lf,textvariable=prog_status,wraplength=800,foreground='#5d6964').grid(row=2,column=0,columnspan=4,sticky='w',pady=(6,0));ttk.Label(lf,text='Con logo, título y logo quedan juntos y centrados. Sin logo, el título se centra solo. Las galerías no usan logo. Para agregar un logo a la lista, copie su PNG (preparado) a la carpeta de logos.',foreground='#5d6964',wraplength=800).grid(row=3,column=0,columnspan=4,sticky='w');row+=1
  if typ in KPI_LIMITS:
   lo_k,hi_k=KPI_LIMITS[typ];box=ttk.LabelFrame(frm,text=f'Indicadores - mínimo {lo_k}, máximo {hi_k}. Cada uno puede llevar un icono.',padding=8);box.grid(row=row,column=0,columnspan=2,sticky='ew',pady=8);old=p.get('kpis',[]);NONE='(sin icono)'
   ttk.Label(box,text='Valor (máx. 25)').grid(row=0,column=1,sticky='w',padx=4);ttk.Label(box,text='Descripción (máx. 60)').grid(row=0,column=2,sticky='w',padx=4);ttk.Label(box,text='Icono').grid(row=0,column=3,sticky='w',padx=4)
   def refresh_values():return [NONE]+list_icon_names()
   for n in range(hi_k):
    val,desc,ic=(tuple(old[n])+('','',''))[:3] if n<len(old) else ('','','');v=tk.StringVar(value=val);dsc=tk.StringVar(value=desc);iv=tk.StringVar(value=ic or NONE);kvars.append((v,dsc,iv));ttk.Label(box,text=f'{n+1}.').grid(row=n+1,column=0,sticky='w');ttk.Entry(box,textvariable=v,width=16).grid(row=n+1,column=1,padx=4,pady=3);ttk.Entry(box,textvariable=dsc,width=34).grid(row=n+1,column=2,padx=4)
    cb=ttk.Combobox(box,textvariable=iv,state='readonly',values=refresh_values(),width=24);cb.configure(postcommand=lambda c=cb:c.configure(values=refresh_values()));cb.grid(row=n+1,column=3,padx=4)
    def upload(ivar=iv,c=cb):
     f=filedialog.askopenfilename(parent=w,title='Seleccionar icono',filetypes=[('Iconos','*.png *.jpg *.jpeg *.webp *.svg')])
     if not f:return
     try:name=import_icon_file(f)
     except Exception as e:return messagebox.showerror('Icono no válido',str(e),parent=w)
     c.configure(values=refresh_values());ivar.set(name)
    ttk.Button(box,text='Subir…',command=upload).grid(row=n+1,column=4,padx=4)
   ttk.Checkbutton(box,text='Mostrar los iconos en blanco dentro del círculo de color (recomendado para iconos de un solo color)',variable=icons_white).grid(row=hi_k+1,column=0,columnspan=5,sticky='w',pady=(8,0))
   ttk.Button(box,text='Abrir carpeta de iconos',command=lambda:(ICON_DIR.mkdir(parents=True,exist_ok=True),open_path(ICON_DIR))).grid(row=hi_k+2,column=0,columnspan=3,sticky='w',pady=(6,0));ttk.Label(box,text='Guarde aquí sus iconos (PNG, JPG, WEBP o SVG) y aparecerán en la lista.',foreground='#5d6964').grid(row=hi_k+2,column=3,columnspan=2,sticky='w');row+=1
  gallery_count=tk.IntVar(value=p.get('gallery_count',6))
  if typ=='gallery':ttk.Label(frm,text='Cantidad obligatoria de fotos').grid(row=row,column=0,sticky='w');ttk.Combobox(frm,textvariable=gallery_count,state='readonly',values=[6,9,12],width=10).grid(row=row,column=1,sticky='w');row+=1
  requirements={'text_photos':'Exactamente 3 fotografías.','mixed_photos':'Exactamente 3: 1 vertical, 1 cuadrada y 1 horizontal.','results_photos':'Exactamente 3 fotografías y entre 2 y 5 indicadores.','results_numbers':'No usa fotografías; requiere entre 1 y 4 indicadores (tarjetas grandes).','gallery':'Exactamente 6, 9 o 12 fotografías según la opción elegida.'}[typ];ttk.Label(frm,text=requirements,foreground='#7A4E00').grid(row=row,column=0,columnspan=2,sticky='w',pady=8);row+=1;photo_status=tk.StringVar(value=f'{len(p.get("photos",[]))} fotos seleccionadas');ttk.Label(frm,textvariable=photo_status).grid(row=row,column=1,sticky='w')
  if typ!='results_numbers':
   def photos():
    ps=filedialog.askopenfilenames(parent=w,filetypes=[('Imágenes','*.png *.jpg *.jpeg *.webp')]);
    if ps:p['photos']=list(ps);photo_status.set(f'{len(ps)} fotos seleccionadas')
   ttk.Button(frm,text='Seleccionar fotografías',command=photos).grid(row=row,column=0,sticky='w');row+=1
  fbox=self.font_override_box(frm,{'text_photos':[('title','Título',8,48,24),('text','Texto principal',6,32,14)],'mixed_photos':[('title','Título',8,48,24),('text','Texto principal',6,32,14)],'results_photos':[('title','Título',8,48,24),('text','Texto principal',6,32,14),('kpi_value','Valor de los indicadores',8,60,26),('kpi_label','Descripción de los indicadores',5,24,11)],'results_numbers':[('title','Título',8,48,24),('text','Texto principal',6,32,14),('kpi_value','Valor de los indicadores',8,60,28),('kpi_label','Descripción de los indicadores',5,24,11)],'gallery':[('title','Título',8,48,24)]}[typ],p.get('manual_font',{}),row+1,w);row+=2
  def save():
   q=dict(p);q['title']=title.get().strip();q['text']=text.get('1.0','end-1c').strip() if text else '';q['gallery_count']=gallery_count.get();q['kpis']=[(a.get().strip(),b.get().strip(),'' if c.get()=='(sin icono)' else c.get()) for a,b,c in kvars if a.get().strip() or b.get().strip()];q['program_logo']=prog.get();q.pop('show_program_logo',None);q['icons_white']=bool(icons_white.get())
   try:q['manual_font']=fbox.get()
   except Exception as e:return messagebox.showerror('Revise esta página',str(e),parent=w)
   try:validate_money_data({'donor':'XX','year':'2026','report_title':'XXX','tagline':'XXXXX','donor_logo':'x','cover_photos':['x']*4,'closing_title':'XXXXX','closing_photos':['x']*5,'pages':[q]},page_only=True)
   except Exception as e:return messagebox.showerror('Revise esta página',str(e),parent=w)
   self.pages[i]=q;self.refresh_page_list();w.destroy()
  ttk.Button(frm,text='Guardar página',command=save,style='Primary.TButton').grid(row=row,column=1,sticky='e',pady=12);frm.columnconfigure(1,weight=1);frm.rowconfigure(1,weight=1)
 def collect(self):
  if self.report_type=='food':return {'excel_path':self.excel_path,'year':getattr(self,'food_year',''),'donors':self.selected_donors(),'photo_mode':self.photo_mode.get(),'shared_photos':self.food_photos,'donor_photos':self.food_donor_photos}
  tid=next((x['id'] for x in self.money_templates if x['name']==self.template_name_var.get()),'Plantilla_BAP_Oficial');return {'donor':self.val('donor'),'year':self.val('year'),'report_title':self.val('report_title'),'tagline':self.val('tagline'),'donor_logo':self.donor_logo,'cover_photos':self.cover_photos,'cover_variant':self.cover_variant.get(),'pages':self.pages,'template_id':tid,'closing_title':self.val('closing_title'),'closing_style':self.closing_key(),'closing_photos':self.closing_photos,'icon_dir':str(ICON_DIR),'manual_font':{**self.cover_font.get(),**self.closing_font.get()}}
 def validate_food(self,d):
  if not d['excel_path'] or not Path(d['excel_path']).exists():raise ValueError('Seleccione el archivo Excel.')
  if not d['donors']:raise ValueError('Seleccione al menos un donante.')
  if d['photo_mode']=='shared':sets={'las fotos compartidas':d['shared_photos']}
  else:
   missing=[x for x in d['donors'] if len(d['donor_photos'].get(x,[]))!=7]
   if missing:raise ValueError('Cada donante necesita exactamente 7 fotos. Falta completar: '+', '.join(missing))
   sets={x:d['donor_photos'][x] for x in d['donors']}
  for label,ps in sets.items():
   if len(ps)!=7:raise ValueError(f'Se requieren exactamente 7 fotos ({label}: {len(ps)}).')
   for n,ph in enumerate(ps,1):
    if not Path(ph).exists():raise ValueError(f'No se encontró la foto {n} de {label}: {ph}')
    if Path(ph).stat().st_size<=FOOD_MIN_PHOTO_BYTES:raise ValueError(f'La foto {n} de {label} ({Path(ph).name}) es demasiado pequeña (10 KB o menos).')
 def save_draft(self):
  try:d=self.collect()
  except Exception as e:return self.err(e)
  folder='alimentos' if self.report_type=='food' else 'monetarios';name=slug((d.get('donors') or [d.get('donor') or 'Sin_nombre'])[0]);p=WORK/'borradores'/folder/f'{name}_{slug(d.get("year",datetime.now().year))}.json';p.write_text(json.dumps({'type':self.report_type,'data':d},ensure_ascii=False,indent=2),encoding='utf-8');messagebox.showinfo('Borrador guardado',f'Guardado en:\n{p}')
 def open_draft(self):
  p=filedialog.askopenfilename(initialdir=WORK/'borradores',filetypes=[('Borrador BAP','*.json')]);
  if not p:return
  try:o=json.loads(Path(p).read_text(encoding='utf-8'));(self.food_workflow if o.get('type')=='food' else self.money_workflow)(o.get('data',{}))
  except Exception as e:self.err(e)
 def preview(self):
  try:
   d=self.collect()
   if self.report_type=='food':
    self.validate_food(d);tmp=WORK/'reportes'/'_vista_previa_tmp';shutil.rmtree(tmp,ignore_errors=True);tmp.mkdir(parents=True,exist_ok=True);self.food_run(d['donors'][:1],tmp,d,preview=True);return
   validate_money_data(d);p=WORK/'reportes'/'_VISTA_PREVIA_MONETARIO.pdf';generate_money(d,d['template_id'],p);open_path(p)
  except Exception as e:self.err(e)
 def generate(self):
  try:
   d=self.collect()
   if self.report_type=='food':
    self.validate_food(d);folder=filedialog.askdirectory(initialdir=WORK/'reportes',title='Seleccione la carpeta para los PDFs')
    if not folder:return
    self.food_run(d['donors'],Path(folder),d);return
   validate_money_data(d);name=f'Informe_Monetario_{slug(d["donor"])}_{slug(d["year"])}.pdf';p=filedialog.asksaveasfilename(initialdir=WORK/'reportes',initialfile=name,defaultextension='.pdf',filetypes=[('PDF','*.pdf')])
   if not p:return
   generate_money(d,d['template_id'],Path(p))
   if messagebox.askyesno('Reporte generado',f'PDF creado:\n{p}\n\n¿Abrir ahora?'):open_path(p)
  except Exception as e:self.err(e)
 def err(self,e):logging.error('%s\n%s',e,traceback.format_exc());messagebox.showerror('No se pudo completar la acción',f'{e}\n\nLos detalles se guardaron en Registros.')
 def diagnostics(self):messagebox.showinfo('Diagnóstico',f'Versión: {VERSION}\nDatos: {WORK}\nAlimentos (datos): {FOOD_DATA}\nAlimentos (reportes): {FOOD_OUTPUT}\nPlantillas monetarias: {money_templates_root()}\nLogos de programas: {LOGO_DIR}\nReportes: {WORK/"reportes"}\nRegistros: {WORK/"logs"}')
if __name__=='__main__':App().mainloop()
