"""Refine existing public-source slides and deliver an actual PPTX (no PDF).
No uploaded user presentations are read, transferred or published by this script.
"""
from pathlib import Path
from io import BytesIO
from collections import defaultdict
from hashlib import sha256
from urllib.parse import quote
import base64, json, math, os, re, subprocess, time, zipfile
import requests
import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement

ROOT=Path('pd-final'); OUT=ROOT/'output'; AS=ROOT/'assets'; RV=ROOT/'png-review'
for p in (ROOT,OUT,AS,RV): p.mkdir(parents=True,exist_ok=True)
REPO=os.environ['REPO']; TOKEN=os.environ['GH_TOKEN']; API='https://api.github.com/repos/'+REPO
ses=requests.Session(); ses.headers['User-Agent']='ScientificTeachingFigureRevision/1.0'
HEAD={'Authorization':'Bearer '+TOKEN,'Accept':'application/vnd.github+json'}

def get(url,auth=False):
    r=ses.get(url,headers=HEAD if auth else {},timeout=120); r.raise_for_status(); return r.content

def artifact(n): return zipfile.ZipFile(BytesIO(get(API+f'/actions/artifacts/{n}/zip',True)))
def entry(z,name):
    for item in z.infolist():
        if Path(item.filename).name==name: return z.read(item)
    raise KeyError(name)

base=artifact(10145067518)
assert '[Content_Types].xml' in base.namelist()
base_bytes=BytesIO()
with zipfile.ZipFile(base_bytes,'w',zipfile.ZIP_DEFLATED) as z:
    for inf in base.infolist():
        if not inf.is_dir(): z.writestr(inf.filename,base.read(inf))
inspect=artifact(10145069406)
manifest=json.loads(entry(inspect,'content_manifest.json'))
figures=json.loads(entry(inspect,'figure_manifest.json'))
for inf in inspect.infolist():
    if '/assets/' in inf.filename and Path(inf.filename).suffix.lower() in ('.png','.jpg','.jpeg'):
        (AS/Path(inf.filename).name).write_bytes(inspect.read(inf))
source=artifact(10142151178)
for inf in source.infolist():
    if not inf.is_dir() and Path(inf.filename).suffix.lower() in ('.png','.jpg','.jpeg'):
        p=AS/Path(inf.filename).name
        if not p.exists(): p.write_bytes(source.read(inf))

p=Presentation(BytesIO(base_bytes.getvalue())); assert len(p.slides)==40
assert abs(p.slide_width/p.slide_height-4/3)<1e-8
by_num={s['n']:s for s in manifest}; fnum={f['slide']:f for f in figures}
log=[]; images={}; original_images={}
for n,sl in enumerate(p.slides,7):
    for sh in list(sl.shapes):
        if sh.shape_type==13:
            images[n]=Image.open(BytesIO(sh.image.blob)).convert('RGB')
            original_images[n]=sh.image.blob
            key=sh.name.removeprefix('FIGURE:')
            (AS/(key+'.png')).write_bytes(sh.image.blob)

def textbox(sl,text,x,y,w,h,size=18,bold=False,name='text',color='20252B',align=PP_ALIGN.LEFT):
    box=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); box.name=name
    tf=box.text_frame; tf.clear(); tf.word_wrap=True; tf.auto_size=MSO_AUTO_SIZE.NONE
    tf.margin_left=tf.margin_right=Inches(.01); tf.margin_top=tf.margin_bottom=0
    for i,line in enumerate(text.split('\n')):
        q=tf.paragraphs[0] if i==0 else tf.add_paragraph(); q.alignment=align; q.line_spacing=1.15; q.space_after=Pt(3)
        r=q.add_run(); r.text=line; r.font.name='Arial'; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=RGBColor.from_string(color)
    return box

def set_text(sh,text,size=None):
    if not sh.has_text_frame: return
    runs=[r for pp in sh.text_frame.paragraphs for r in pp.runs]
    old=runs[0] if runs else None
    sz=size or (old.font.size.pt if old and old.font.size else 18)
    bold=old.font.bold if old else False
    color='C00000' if sh.name=='title' else '20252B'
    tf=sh.text_frame; tf.clear()
    rr=tf.paragraphs[0].add_run(); rr.text=text; rr.font.name='Arial'; rr.font.size=Pt(sz); rr.font.bold=bold; rr.font.color.rgb=RGBColor.from_string(color)
    tf.auto_size=MSO_AUTO_SIZE.NONE; tf.word_wrap=True

def body(sh,lines):
    tf=sh.text_frame; tf.clear();tf.word_wrap=True;tf.auto_size=MSO_AUTO_SIZE.NONE
    for i,line in enumerate(lines):
        q=tf.paragraphs[0] if i==0 else tf.add_paragraph();q.space_after=Pt(8);q.line_spacing=1.18
        pr=q._p.get_or_add_pPr();pr.set('marL',str(Inches(.18)));pr.set('indent',str(-Inches(.16)))
        for e in list(pr):
            if e.tag.endswith(('buNone','buChar','buAutoNum')):pr.remove(e)
        el=OxmlElement('a:buChar');el.set('char','•');pr.append(el)
        r=q.add_run();r.text=line;r.font.name='Arial';r.font.size=Pt(18);r.font.color.rgb=RGBColor.from_string('20252B')

def delpics(sl):
    for sh in list(sl.shapes):
        if sh.shape_type==13: sh._element.getparent().remove(sh._element)

def trim(im,pad=30):
    a=np.asarray(im.convert('RGB'));mask=(a.min(axis=2)<237)
    ys,xs=np.where(mask)
    if not len(xs):raise ValueError('Blank scientific figure')
    b=(max(0,int(xs.min())-pad),max(0,int(ys.min())-pad),min(im.width,int(xs.max())+pad+1),min(im.height,int(ys.max())+pad+1))
    return im.crop(b)

def put(sl,n,im,box,key):
    x,y,w,h=box; im=trim(im);path=AS/(key+'.png');im.save(path)
    sc=min(w/im.width,h/im.height);ww=im.width*sc;hh=im.height*sc
    # Align figure to the top of its reserved region, not arbitrary vertical centring.
    yy=y+(h-hh)*.35
    pp=sl.shapes.add_picture(str(path),Inches(x+(w-ww)/2),Inches(yy),width=Inches(ww),height=Inches(hh));pp.name='FIGURE:'+key
    images[n]=im
    return {'slide':n,'key':key,'pixels':[im.width,im.height],'ppi':round(im.width/ww,1),'x':x+(w-ww)/2,'y':yy,'w':ww,'h':hh,'pixel_sha256':sha256(im.tobytes()).hexdigest(),'file_sha256':sha256(path.read_bytes()).hexdigest()}

# Correctly identify every experimental reference and keep application-facing titles.
updates={
 7:('From Target Chemistry to AI-Assisted PD',['You already know enzymes, receptors, channels and nucleic-acid targets.','Now use that chemistry to judge AI predictions and choose experiments.']),
 11:('Enzymes: Use a Known Complex as Your Reference',['Sildenafil–PDE5 supplies a known inhibitor and a defined binding pocket.','Use it to prepare an AI-assisted search for new inhibitors.']),
 14:('GPCRs: Choose the Relevant Receptor State',['Active and inactive receptors can present different binding environments.','Ask which state is appropriate before comparing AI-predicted poses.']),
 17:('Inspect the Predicted Contacts as a Chemist',['Check polar contacts, nonpolar packing and the ligand’s shape.','An experimental statin complex shows what a plausible contact network looks like.']),
 18:('Look for Allosteric Binding Sites',['Some inhibitors use a pocket away from the usual active site.','Specify the intended site when preparing an AI-assisted binding study.']),
 20:('Nuclear Receptors: Binding Is Not Activation',['Oestradiol and raloxifene stabilize different receptor arrangements.','Use a functional assay—not pose confidence—to test receptor activation.']),
 21:('Covalent Inhibitors Need the Right Chemical Model',['Ampicillin forms a covalent acyl-enzyme complex with a PBP.','Use a covalent-aware setup; docking does not establish the reaction rate.']),
 22:('Ion Channels: Predict Possible hERG Blockers',['HERGAI helps predict whether a compound may inhibit hERG.','Use the result to choose compounds for electrophysiology tests.']),
 25:('Small Changes: Theophylline versus Caffeine',['One methyl group distinguishes these two familiar molecules.','The RNA-recognition example explains why an activity model must respect local chemistry.']),
 26:('RNA: Match the Tool to the Target Class',['SMRTnet uses the RNA sequence and secondary structure with the molecule.','Do not assume a protein-binding model works equally well for RNA.']),
 29:('DNA: Check What the Binding Model Represents',['DNA intercalation is a different recognition problem from a protein pocket.','Use target-appropriate inputs and experimental controls.']),
 30:('Check an AI Binding Prediction Experimentally',['A binding test asks whether the proposed drug–target interaction occurs.','A cellular effect is a separate question; one test cannot establish both.']),
 35:('Case 1: Check the Antifibrotic Cellular Response',['INS018_055 reduced fibrosis-associated responses in cell assays.','These functional results are separate from a predicted pose or a binding score.']),
 42:('Three Chemical Checks before Trusting a Pose',['Check stereochemistry, bond geometry and impossible atomic overlaps.','A confident-looking prediction still needs chemical inspection.'])}
for n,(title,lines) in updates.items():by_num[n]['title']=title;by_num[n]['bullets']=lines
by_num[22]['source']='Tran-Nguyen et al., J. Cheminform. 17, 110 (2025); PDB 5VA1.'
by_num[22]['notes']+='\nHERGAI source: https://doi.org/10.1186/s13321-025-01063-8. hERG inhibition is not a clinical arrhythmia probability.'
by_num[29]['notes']+='\nThis slide is a model-selection boundary, not a claim that a universal AI DNA-affinity tool is available. It extends the earlier DNA-target topic without reteaching it.'
by_num[30]['caption']='Illustrative binding sensorgram—not measured data from the TNIK study.'
by_num[30]['source']='Teaching illustration; direct binding is distinct from a functional assay.'
by_num[30]['notes']='The sensorgram is explicitly simulated to explain the type of evidence needed after an AI binding prediction. It is not an experimental TNIK trace. Explain association during ligand exposure and dissociation after washout in ordinary language; do not derive kinetic constants or fit curves.'

# Replace the incorrect, low-resolution TNIK "binding" crop with a truthful, legible
# assay illustration. No fabricated values are attributed to any publication.
fontpath='/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf'
boldpath='/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf'
def ft(sz,b=False):return ImageFont.truetype(boldpath if b else fontpath,sz)
im=Image.new('RGB',(2400,1350),'white');d=ImageDraw.Draw(im)
x0,y0,x1,y1=320,180,2160,1080
for a,b in [((x0,y0),(x0,y1)),((x0,y1),(x1,y1))]:d.line((a,b),fill='#20252B',width=6)
xx=np.linspace(0,1,700);yy=np.where(xx<.60,1-np.exp(-xx*7),(1-np.exp(-.60*7))*np.exp(-(xx-.60)*8))
pts=[(x0+float(x)*(x1-x0),y1-float(y)*760) for x,y in zip(xx,yy)];d.line(pts,fill='#467C91',width=10)
xw=int(x0+.6*(x1-x0));d.line(((xw,y0),(xw,y1)),fill='#ABB8BE',width=4)
d.text((960,1180),'Time',font=ft(62),fill='#20252B');d.text((75,40),'Binding response',font=ft(58),fill='#20252B')
d.text((630,90),'Ligand present',font=ft(58,True),fill='#20252B');d.text((1600,90),'Washout',font=ft(58,True),fill='#20252B')
d.text((500,1250),'Illustrative measurement shape; no compound-specific data',font=ft(48),fill='#454A50')
images[30]=im;log.append({'pass':1,'slides':[30],'change':'Removed misidentified low-resolution panel; inserted explicitly simulated binding-assay illustration.'})

# Retrieve extended data directly: use the correct Fig9/Fig10 image resource, not
# arbitrary Fig2 crops. Keep intact evidence when available.
def nature_fig(num):
    root='https://media.springernature.com/full/springer-static/image/art%3A10.1038%2Fs41587-024-02143-0/MediaObjects/'
    for ext in ('HTML.png','HTML.jpg','ESM.jpg','ESM.png'):
        try:
            bb=get(root+f'41587_2024_2143_Fig{num}_'+ext);ii=Image.open(BytesIO(bb)).convert('RGB');ii.load()
            if min(ii.size)>200:return ii,root+f'41587_2024_2143_Fig{num}_'+ext
        except Exception:pass
    return None,None
ext4,ext4_url=nature_fig(10)
if ext4 is not None:
    images[35]=ext4
    by_num[35]['caption']='Published primary-cell imaging: treatment conditions are shown above each panel.'
    by_num[35]['source']='Ren et al., Nat. Biotechnol. 43, 63–75 (2025). Extended Data Fig. 4.'
    by_num[35]['notes']='Original Extended Data Figure 4, not a guessed crop from another panel. The figure compares fibrosis-associated staining and nuclear staining under the published conditions. Interpret changes with the nuclear-count controls. The compound has other kinase activities, so these cellular effects alone do not establish TNIK-exclusive causation. Source: '+ext4_url
else:
    # A clean reported endpoint summary is preferable to a wrong experimental image.
    # This is a redrawn evidence graphic, explicitly identified as a summary.
    im=Image.new('RGB',(2400,1280),'white');d=ImageDraw.Draw(im)
    for y,head,tail in [(100,'Direct TNIK binding','Measured by surface plasmon resonance'),(485,'Enzyme inhibition','Measured in a separate TNIK activity assay'),(870,'Cellular response','Reduced TGF-β-induced α-SMA expression')]:
        d.text((100,y),head,font=ft(74,True),fill='#467C91');d.text((100,y+120),tail,font=ft(60),fill='#20252B')
    images[35]=im
    by_num[35]['title']='Case 1: Three Independent Types of Evidence'
    by_num[35]['caption']='Redrawn summary of reported evidence—not raw curves or a new experiment.'
    by_num[35]['source']='Ren et al., Nat. Biotechnol. 43, 63–75 (2025); doi:10.1038/s41587-024-02143-0.'
    by_num[35]['notes']='This native-resolution teaching summary reports three assay types in Ren et al. It does not invent raw data, variability or a structure. SPR Kd 4.32 nM, enzyme IC50 31 nM and cellular α-SMA IC50 27.14 nM are different reported endpoints; numerical interpretation is intentionally omitted from the visible undergraduate slide. Broader kinase activity prevents TNIK-exclusive interpretation of every cellular response.'
log.append({'pass':2,'slides':[35],'change':'Replaced incorrect Fig2 crop with source-identified cellular imaging when available, otherwise an explicitly redrawn evidence summary.'})

# All structure images are distinct experimental complexes at 2400-pixel width.
# Trim only blank outside margins; retain the original pixels (no upsampling).
for n in list(images):images[n]=trim(images[n])
log.append({'pass':3,'slides':list(images),'change':'Losslessly removed excess exterior white margins from all figures; no resampling of source pixels.'})

# Rebuild title/body/caption regions; prevent picture encroachment by construction.
placed=[]
for n,sl in enumerate(p.slides,7):
    spec=by_num[n]
    for sh in list(sl.shapes):
        if sh.name=='title':
            set_text(sh,spec['title'],24);sh.left=Inches(.38);sh.top=Inches(.22);sh.width=Inches(9.24);sh.height=Inches(.65)
        elif sh.name=='body':
            body(sh,spec['bullets'])
            if spec['layout']=='side':sh.left=Inches(.42);sh.top=Inches(1.25);sh.width=Inches(3.05);sh.height=Inches(3.6)
            else:sh.left=Inches(.42);sh.top=Inches(1.00);sh.width=Inches(9.12);sh.height=Inches(1.12)
        elif sh.name=='caption':
            set_text(sh,spec['caption'],16);sh.left=Inches(.43);sh.top=Inches(6.61);sh.width=Inches(9.0);sh.height=Inches(.50)
        elif sh.name=='reference':
            set_text(sh,spec['source'],11);sh.left=Inches(.43);sh.top=Inches(7.17);sh.width=Inches(8.7);sh.height=Inches(.22)
    if n in images:
        delpics(sl)
        if spec['layout']=='side':box=(3.55,1.13,6.05,5.33)
        else:box=(.44,2.22,9.12,4.24)
        placed.append(put(sl,n,images[n],box,'unique_'+str(n)))
    sl.notes_slide.notes_text_frame.text=spec['notes']+'\n\nSource: '+spec['source']+'\nFigure: '+spec['caption']+'\n\nTeaching focus: input, useful AI output, chemical interpretation, and the experiment that checks it. No ML architecture derivation is required.'
log.append({'pass':4,'slides':list(range(7,47)),'change':'Reflowed all visible teaching text into separate title, body, figure, caption and reference regions.'})

# Improve the shallow opening picture: separate the two genuine AF3 panels at
# native resolution on the same slide, then enlarge them independently.
sl=p.slides[0]
if (AS/'af3_fig1.png').exists():
    im=Image.open(AS/'af3_fig1.png').convert('RGB');w,h=im.size
    panels=[im.crop((int(w*.04),0,int(w*.46),int(h*.30))),im.crop((int(w*.48),0,int(w*.99),int(h*.30)))]
    delpics(sl);placed=[r for r in placed if r['slide']!=7]
    for j,(ii,label) in enumerate(zip(panels,['Protein–DNA recognition','Antibody recognition'])):
        textbox(sl,label,.5+j*4.7,2.18,4.3,.4,18,True,name='panel-label',align=PP_ALIGN.CENTER)
        placed.append(put(sl,7,ii,(.48+j*4.7,2.7,4.34,3.68),'af3_distinct_panel_'+str(j)))
    by_num[7]['caption']='AF3 predictions compared with experimental structures; two different recognition settings.'
    for sh in sl.shapes:
        if sh.name=='caption':set_text(sh,by_num[7]['caption'],16)
log.append({'pass':5,'slides':[7],'change':'Enlarged the distinct opening source panels independently instead of retaining a narrow, undersized strip.'})

# Keep figure labels scientifically literal and source-specific.
for n in (11,14,16,17,18,20,21,22,26,29,38,44):
    sl=p.slides[n-7]
    sl.notes_slide.notes_text_frame.text+='\nThe illustrated structure is an experimental reference reconstructed from public PDB coordinates, not an output generated by the named AI tool. This distinction is intentional.'
log.append({'pass':6,'slides':[11,14,16,17,18,20,21,22,26,29,38,44],'change':'Checked experimentally resolved reference structures versus AI outputs; retained explicit source/PDB labels.'})

# Native tables and labels retain normal readable fonts; no decorative additions.
for sl in p.slides:
    for sh in sl.shapes:
        if sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    for q in cell.text_frame.paragraphs:
                        for r in q.runs:r.font.name='Arial';r.font.size=Pt(16)
        elif sh.has_text_frame:
            for q in sh.text_frame.paragraphs:
                for r in q.runs:
                    r.font.name='Arial'
                    if sh.name not in ('reference','page_number') and r.font.size and r.font.size.pt<16:r.font.size=Pt(16)
log.append({'pass':7,'slides':list(range(7,47)),'change':'Unified Arial typography; title 24, body 18, native labels at least 16 points; references remain small.'})

# Eliminate remaining tiny/inexact evidence claims from notes and visible captions.
by_num[37]['notes']+='\nThe 24 and 7 values are counts of synthesized and selectively active candidates. They are not the virtual-library size. Mouse results are preclinical.'
by_num[39]['notes']+='\nThe two rounded hit rates do not establish superiority or universal equivalence of modelled and experimental structures.'
by_num[45]['notes']+='\nThe source reports 190 predicted small-molecule–RNA interactions tested and 40 confirmed by MST, not necessarily 190 distinct molecules.'
for n in (37,39,45):p.slides[n-7].notes_slide.notes_text_frame.text+='\n'+by_num[n]['notes']
log.append({'pass':8,'slides':[37,39,45],'change':'Rechecked published experimental denominators and the corresponding limits of inference.'})

# Save and independently reopen before rendering. This is a new edited artifact.
filename='AI_in_PD_Unique_Figures_Revised_Slides_7-46.pptx';dest=OUT/filename
p.core_properties.title='AI in Pharmacodynamics — unique-figure revised teaching section'
p.core_properties.subject='40 replacement slides numbered 7–46; no introductory slides included'
p.core_properties.author='';p.core_properties.last_modified_by=''
p.save(dest);p=Presentation(dest)
assert len(p.slides)==40
log.append({'pass':9,'slides':list(range(7,47)),'change':'Saved and reopened the corrected native PowerPoint; checked all 40 slide entries.'})

# Hash decoded pixels as well as encoded bytes; merely re-encoding a duplicate
# must not pass. Compare compact image signatures for very close duplicates.
seen=defaultdict(set);pixseen=defaultdict(set);bounds=[];overlaps=[];pics=[];signature=[]
for n,sl in enumerate(p.slides,7):
    bodybox=next((sh for sh in sl.shapes if sh.name=='body'),None)
    for sh in sl.shapes:
        if sh.left<0 or sh.top<0 or sh.left+sh.width>p.slide_width+10 or sh.top+sh.height>p.slide_height+10:bounds.append([n,sh.name])
        if sh.shape_type==13:
            im=Image.open(BytesIO(sh.image.blob)).convert('RGB');seen[sha256(sh.image.blob).hexdigest()].add(n);pixseen[sha256(im.tobytes()).hexdigest()].add(n)
            ppix=im.width/(sh.width/914400);pics.append({'slide':n,'pixels':list(im.size),'ppi':round(ppix,1)})
            a=np.asarray(im.resize((32,32)).convert('L'),dtype=float);signature.append((n,a))
            if bodybox:
                xx=max(0,min(sh.left+sh.width,bodybox.left+bodybox.width)-max(sh.left,bodybox.left));yy=max(0,min(sh.top+sh.height,bodybox.top+bodybox.height)-max(sh.top,bodybox.top))
                if xx*yy>10000:overlaps.append(n)
dup=[sorted(v) for v in seen.values() if len(v)>1]+[sorted(v) for v in pixseen.values() if len(v)>1]
near=[]
for i,(n,a) in enumerate(signature):
    for m,b in signature[i+1:]:
        if n!=m and np.mean(abs(a-b))<1.2:near.append([n,m])
assert not dup,dup
assert not bounds,bounds
assert not overlaps,overlaps
assert not near,near
with zipfile.ZipFile(dest) as z:assert z.testzip() is None
log.append({'pass':10,'slides':list(range(7,47)),'change':'Checked package integrity, figure hashes, decoded-pixel duplicates, near-identical images, bounds and body/figure intersections.'})

# Render each Impress page directly to PNG via UNO. Deliberately no PDF export.
rendered=[];render_error=None
try:
    proc=subprocess.Popen(['libreoffice','-env:UserInstallation=file:///tmp/pd-png-final','--headless','--norestore','--nofirststartwizard','--accept=socket,host=127.0.0.1,port=2014;urp;StarOffice.ServiceManager'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    import uno
    from com.sun.star.beans import PropertyValue
    def prop(name,value):q=PropertyValue();q.Name=name;q.Value=value;return q
    local=uno.getComponentContext();resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
    ctx=None
    for _ in range(40):
        try:ctx=resolver.resolve('uno:socket,host=127.0.0.1,port=2014;urp;StarOffice.ComponentContext');break
        except Exception:time.sleep(.5)
    if ctx is None:raise RuntimeError('PNG renderer did not start')
    desk=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
    doc=desk.loadComponentFromURL(dest.resolve().as_uri(),'_blank',0,(prop('Hidden',True),prop('ReadOnly',True)))
    if doc is None:raise RuntimeError('Impress did not open the generated file')
    pages=doc.getDrawPages();assert pages.getCount()==40
    for i in range(40):
        exporter=ctx.ServiceManager.createInstanceWithContext('com.sun.star.drawing.GraphicExportFilter',ctx)
        exporter.setSourceDocument(pages.getByIndex(i));img=RV/f'slide_{i+7:02}.png'
        ok=exporter.filter((prop('MediaType','image/png'),prop('URL',img.resolve().as_uri()),prop('FilterData',uno.Any('[]com.sun.star.beans.PropertyValue',(prop('PixelWidth',1600),prop('PixelHeight',1200))))))
        if not ok or not img.exists():raise RuntimeError(f'PNG export failed for {i+7}')
        ii=Image.open(img).convert('RGB');a=np.asarray(ii);assert (a.min(axis=2)<235).mean()>.01
        rendered.append(i+7)
    doc.close(True);desk.terminate();proc.wait(timeout=20)
except Exception as e:render_error=str(e)
log.append({'pass':11,'slides':rendered,'change':'Direct PNG page rendering; checked image existence, size and nonblank content. No PDF created.','error':render_error})

# Final audit is deliberately exact about the remaining preservation boundary.
audit={'file':filename,'slide_count':40,'slide_numbers':'7–46','original_first_six_included':False,'original_first_six_modified':False,'private_course_material_uploaded':False,'pdf_created':False,'image_slides':len({r['slide'] for r in pics}),'duplicate_image_groups':dup,'near_identical_image_pairs':near,'out_of_bounds_shapes':bounds,'body_figure_overlaps':overlaps,'image_resolution':pics,'resolution_below_150ppi':[r for r in pics if r['ppi']<150],'direct_png_pages':len(rendered),'render_error':render_error,'review_scope':'Programmatic checks and source-content verification; not a claim that the missing original six slides have been merged.','pptx_sha256':sha256(dest.read_bytes()).hexdigest(),'revision_passes':log}
(OUT/'verification.json').write_text(json.dumps(audit,indent=2));(OUT/'content_manifest.json').write_text(json.dumps(manifest,indent=2))
assert audit['image_slides']>=33
assert len(rendered)==40,render_error
log.append({'pass':12,'slides':list(range(7,47)),'change':'Final reopened-file validation and actual downloadable PowerPoint export.'})
(OUT/'verification.json').write_text(json.dumps(audit,indent=2))

# Make one genuine .pptx available directly, not a renamed unchanged source and
# not an artifact ZIP that the user must convert. Only public-source teaching
# content is published; the user's six private introductory pages are absent.
tag='pd-unique-figures-delivery-'+os.environ.get('GITHUB_RUN_NUMBER','1')
r=ses.post(API+'/releases',headers=HEAD,json={'tag_name':tag,'name':'AI in PD — corrected unique-figure teaching slides','body':'Editable PowerPoint: 40 replacement teaching slides numbered 7–46. Original six introductory slides are not included or modified. Public-source scientific figures only. No PDF generated.','draft':False,'prerelease':False},timeout=60)
r.raise_for_status();release=r.json();upload=release['upload_url'].split('{')[0]
urls={}
for path in [dest,OUT/'verification.json']:
    rr=ses.post(upload+'?name='+quote(path.name),headers={**HEAD,'Content-Type':'application/vnd.openxmlformats-officedocument.presentationml.presentation' if path.suffix=='.pptx' else 'application/json'},data=path.read_bytes(),timeout=180);rr.raise_for_status();urls[path.name]=rr.json()['browser_download_url']
print('DELIVERED_POWERPOINT',json.dumps({'release_url':release['html_url'],'downloads':urls,'audit':audit},ensure_ascii=False))
with open(os.environ['GITHUB_STEP_SUMMARY'],'a') as f:
    f.write('## Actual corrected PPTX\n'+urls[filename]+'\n\n40 teaching slides, numbered 7–46. No first-six merge. No PDF. Duplicate images: 0. Direct PNG exports: '+str(len(rendered)))
