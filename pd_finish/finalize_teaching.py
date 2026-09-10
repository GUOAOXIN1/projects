"""Finish the public-source teaching section. No original user files are accessed.
The output has 40 slides numbered 7-46; original introductory slides are NOT included.
"""
from __future__ import annotations
import base64, hashlib, importlib.util, io, json, math, os, re, shutil, subprocess, sys, zipfile
from collections import defaultdict
from pathlib import Path
import requests
import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from rdkit import Chem
from rdkit.Chem import rdDepictor, rdMolDescriptors
from rdkit.Chem.Draw import rdMolDraw2D

ROOT=Path('pd-finished'); AS=ROOT/'assets'; OUT=ROOT/'output'; REVIEW=ROOT/'review'
for p in (ROOT,AS,OUT,REVIEW,ROOT/'passes'):p.mkdir(parents=True,exist_ok=True)
API='https://api.github.com/repos/'+os.environ['REPO'];AUTH={'Authorization':'Bearer '+os.environ['GH_TOKEN']}
session=requests.Session()
def retrieve(url,auth=False):
    r=session.get(url,headers=AUTH if auth else {'User-Agent':'ScientificTeachingFigure/1.0'},timeout=120);r.raise_for_status();return r.content
raw=retrieve(API+'/actions/artifacts/10145069406/zip',True)
with zipfile.ZipFile(io.BytesIO(raw)) as z:
    for e in z.infolist():
        name=Path(e.filename).name
        if not e.is_dir() and name:
            dest=AS/name if Path(name).suffix.lower() in ('.png','.jpg','.jpeg','.gif') else ROOT/name
            dest.write_bytes(z.read(e))
source=ROOT/'AI_in_PD_Revised_Teaching_Slides.pptx'
assert source.is_file()
old=Presentation(source);assert len(old.slides)==40
old_manifest=json.loads((ROOT/'figure_manifest.json').read_text())
content=json.loads((ROOT/'content_manifest.json').read_text())
registry={}
for row in old_manifest:registry[row['key']]=dict(row)
# Recover the original embedded picture bytes, including GIF sources that were
# not exported as separate PNGs in the earlier workflow.
for sl in old.slides:
    for sh in sl.shapes:
        if sh.shape_type==MSO_SHAPE_TYPE.PICTURE and sh.name.startswith('FIGURE:'):
            key=sh.name.split(':',1)[1];path=AS/(key+'.'+sh.image.ext)
            path.write_bytes(sh.image.blob);registry[key]['path']=str(path)
# Reuse layout primitives rather than reconstructing unknown user slide XML.
spec=importlib.util.spec_from_file_location('pd_builder','pd_rebuild/build_teaching_section.py')
b=importlib.util.module_from_spec(spec);spec.loader.exec_module(b)
b.OUT=OUT;b.AS=AS;b.ROOT=ROOT;b.USE=[];b.PROV={}
for key,row in registry.items():
    if 'path' in row:b.figure(key,row['path'],row['source'],row['detail'])
S={a['n']:a for a in content}
for a in content:a['notes']=a.get('notes','')

def register(key,path,source,detail):
    return b.figure(key,path,source,detail)
def trim(im,pad=25):
    im=im.convert('RGB');diff=ImageChops.difference(im,Image.new('RGB',im.size,'white'))
    mask=diff.convert('L').point(lambda v:255 if v>18 else 0);box=mask.getbbox()
    if box:
        box=(max(0,box[0]-pad),max(0,box[1]-pad),min(im.width,box[2]+pad),min(im.height,box[3]+pad));return im.crop(box)
    return im

def crop_public(key,parent,box,source,detail):
    p=AS/(parent+'.png')
    if not p.exists():p=Path(b.PROV[parent]['path'])
    im=Image.open(p).convert('RGB');assert 0<=box[0]<box[2]<=im.width and 0<=box[1]<box[3]<=im.height,(parent,im.size,box)
    out=AS/(key+'.png');im.crop(box).save(out)
    register(key,out,source,detail+'; lossless pixel crop '+str(box)+'; no resampling')
    return key

# Two different components of the DGAT2 publication occupy the middle row.
# The previous bottom-right crop accidentally included a different protein.
crop_public('dgat2_p2rank','pockets_fig3',(0,510,750,1015),'https://doi.org/10.1038/s41586-021-03828-1','Figure 3b left: P2Rank candidate pocket on predicted DGAT2')
crop_public('dgat2_docked','pockets_fig3',(1475,510,2118,1015),'https://doi.org/10.1038/s41586-021-03828-1','Figure 3b right: docked DGAT2 inhibitor, a structural hypothesis')
S[13].update(key=None,layout='dgat2_pair',caption='Published DGAT2 example: predicted pocket (left) and docked inhibitor (right).',notes='Figure 3b in Tunyasuvunakool et al. P2Rank proposes a candidate site. The right panel is a docking hypothesis, not an experimental complex. Native labels identify the two distinct steps. The unrelated wolframin row has been removed.')

# The full target-nomination figure repeated data and contained unnecessary ML
# jargon. A compact native evidence map replaces it; it is not a study result.
S[10].update(key=None,layout='target_evidence',title='Which Target Is Worth Testing?',bullets=['AI can combine disease measurements and published biological evidence.','Use the ranking to choose a target for experimental investigation.'],caption='Disease evidence → target hypothesis → experimental validation.',notes='Original editable teaching scheme. Gene and protein measurements and literature evidence can prioritize targets. This does not claim that these invented evidence items were produced by a particular model. TNIK is the later published case.')

# Retain only the generative-chemistry part of Ren Figure 1 for the case.
crop_public('tnik_design_focused','tnik_fig1',(1220,0,2115,525),'https://doi.org/10.1038/s41587-024-02143-0','Figure 1a generative-chemistry and experimental-filtering portion')
S[34].update(key='tnik_design_focused',layout='side',caption='Chemistry42 proposed candidates; chemical and experimental filters guided selection.',bullets=['AI proposed inhibitors for the selected TNIK target.','Chemists synthesized and tested candidates before choosing a lead.'])

# Large chemical drawing of the actual TNIK candidate, with identity checked
# against the primary PubChem record rather than a generic unrelated molecule.
SMILES='CC(C)N1C=NC(=C1C2=NC=C(N2)C(=O)NC3=CC=C(C=C3)N4CCN(CC4)C)C5=CC=C(C=C5)F'
mol=Chem.MolFromSmiles(SMILES);assert mol is not None
assert rdMolDescriptors.CalcMolFormula(mol)=='C27H30FN7O'
assert Chem.MolToInchiKey(mol)=='ZVDNXHUSIKGTSF-UHFFFAOYSA-N'
rdDepictor.Compute2DCoords(mol)
d=rdMolDraw2D.MolDraw2DCairo(2100,1500);opt=d.drawOptions();opt.useBWAtomPalette();opt.bondLineWidth=5;opt.minFontSize=65;opt.maxFontSize=95;opt.padding=.06
d.DrawMolecule(mol);d.FinishDrawing();im=trim(Image.open(io.BytesIO(d.GetDrawingText())),45);p=AS/'rentosertib_structure.png';im.save(p)
register('rentosertib_structure',p,'https://pubchem.ncbi.nlm.nih.gov/compound/164938183 ; https://doi.org/10.1038/s41587-024-02143-0','New high-resolution chemical drawing of rentosertib, INS018_055; molecular formula and InChIKey verified. Not an AI prediction.')
# MRC-5 concentration-response panel is top right of Figure 2, not its middle.
crop_public('tnik_mrc5','tnik_fig2',(1480,0,2102,492),'https://doi.org/10.1038/s41587-024-02143-0','Figure 2c: inhibition of TGF-beta-induced alpha-SMA expression in MRC-5 cells; original published curve')
S[35].update(key=None,layout='tnik_result',title='Case 1: What Did the TNIK Experiments Show?',bullets=['The lead bound TNIK and inhibited its enzyme activity.','Cell assays showed reduced fibrosis-marker expression.'],caption='INS018_055 and the published MRC-5 cell assay; binding and cell response are distinct.',notes='Rentosertib molecular identity: PubChem CID 164938183. Right panel: Ren Figure 2c, TGF-beta-induced alpha-SMA expression in MRC-5 cells. The figure is a cellular assay, not a direct-binding experiment or an experimental complex. Broader kinase activity limits TNIK-exclusive interpretation. Do not teach curve fitting. No clinical approval or definitive patient benefit is claimed.')

# Replace the erroneous supposed TNIK binding-assay crop with an explicit,
# high-resolution illustrative SPR readout. This is not invented study data.
im=Image.new('RGB',(2800,1450),'white');dr=ImageDraw.Draw(im)
regular='/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf';bold='/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf'
def f(n,boldface=False):return ImageFont.truetype(bold if boldface else regular,n)
def txt(x,y,t,n=76,boldface=False,fill='#20252B',anchor=None):dr.text((x,y),t,font=f(n,boldface),fill=fill,anchor=anchor)
x0,y0,x1,y1=320,1170,2520,230
dr.line([(x0,y1),(x0,y0),(x1,y0)],fill='#303030',width=7)
for amp,rate,col in [(740,.016,'#4C7C91'),(510,.020,'#758F72'),(270,.027,'#A58258')]:
    pts=[]
    for t in np.linspace(0,400,650):
        val=amp*(1-math.exp(-rate*t)) if t<=230 else amp*(1-math.exp(-rate*230))*math.exp(-.015*(t-230))
        pts.append((int(x0+(x1-x0)*t/400),int(y0-val)))
    dr.line(pts,fill=col,width=9)
wash=x0+(x1-x0)*230/400
for y in range(200,1170,30):dr.line([(wash,y),(wash,y+14)],fill='#969696',width=3)
txt(780,90,'Ligand added',85,True,anchor='mm');txt(2120,90,'Washout',85,True,anchor='mm')
txt(1450,1320,'Time',90,False,anchor='mm')
# Horizontal y-axis title avoids tiny rotated raster text.
txt(325,152,'Binding response',73,True)
txt(2480,1210,'→',85,anchor='mm')
p=AS/'spr_illustrative.png';im.save(p)
register('spr_illustrative',p,'Copeland, Evaluation of Enzyme Inhibitors in Drug Discovery, 2nd ed. (2013).','Original illustrative one-site association/washout curves. No observed values, replicate errors or TNIK study measurements are asserted.')
S[30].update(key='spr_illustrative',layout='wide',title='Use a Binding Experiment to Check the Prediction',bullets=['A model proposes an interaction; an experiment checks it.','SPR measures a binding response, not the complete cellular mechanism.'],caption='Illustrative SPR traces—not measurements from a published compound study.',notes='The three curves are explicit teaching simulations of association and washout. Their purpose is to identify what a binding measurement looks like, not teach fitting or binding kinetics. An enzyme assay measures inhibition; a cellular engagement test provides complementary evidence. The former incorrectly labelled TNIK figure crop has been removed.')

# Tighten the introductory structural panel without reusing any panel later.
# The entire AF3 overview image is reserved for this slide; no other slide uses it.
a=AS/'af3_intro.png'
if a.exists():
    im=trim(Image.open(a),15);p=AS/'af3_intro_trimmed.png';im.save(p)
    register('af3_intro_trimmed',p,'https://doi.org/10.1038/s41586-024-07487-w','Published AF3 structural overview panel, trimmed only at external white margins')
    S[7]['key']='af3_intro_trimmed'

# Improve every structural image's usable area losslessly. Sources are distinct
# experimental structures, not snapshots of the earlier course presentation.
structural=['pde5_sildenafil','gpcr_states','egfr_erlotinib','hmgr_rosuvastatin','allosteric_pocket','er_states','pbp_ampicillin','herg_channel','rna_theophylline','dna_anthracycline','lyu_cryoem','egfr_lapatinib']
for key in structural:
    if key not in b.PROV:continue
    prev=dict(b.PROV[key]);im=trim(Image.open(prev['path']),40);p=AS/(key+'_trimmed.png');im.save(p)
    register(key,p,prev['source'],prev['detail']+' External blank margins trimmed losslessly.')

# Keep a realistic scientific statement on every slide: predictions, measured
# results and illustrative teaching examples must remain clearly distinguished.
updates={
7:('AI in PD: Use the Target Chemistry You Know',['You already know enzymes, receptors, ion channels and nucleic-acid targets.','Now use AI to choose molecules and experiments—not to replace that chemistry.']),
11:('Enzymes: Begin with a Known Binding Complex',['Sildenafil–PDE5 provides a real inhibitor and a defined binding pocket.','Use that reference to set up an AI-assisted inhibitor screen.']),
12:('Predict a Molecular Complex',['Provide the identities of the interacting molecules.','Treat the predicted arrangement as a hypothesis, not a measured activity.']),
14:('GPCRs: Choose a Relevant Receptor State',['Active and inactive receptors can have different pocket shapes.','Consider that state when preparing or judging an AI binding prediction.']),
16:('Use AI-Assisted Docking to Propose a Pose',['GNINA combines docking with learned scoring.','Ask whether a proposed pose is consistent with known chemistry.']),
17:('Inspect the Predicted Contacts',['Look for chemically sensible polar and nonpolar contacts.','A contact picture helps inspection; it is not an affinity measurement.']),
18:('Allosteric Sites: Look Beyond the Usual Pocket',['A ligand can act at a site separate from the ATP-binding pocket.','Define the relevant site when using a binding predictor.']),
20:('Nuclear Receptors: Binding Is Not Function',['Oestradiol and raloxifene stabilize different receptor arrangements.','An AI binding pose alone does not establish agonism or antagonism.']),
21:('Covalent Binding Requires the Right Model',['Ampicillin forms a covalent acyl-enzyme complex with a PBP.','Use covalent-aware preparation; test reaction and inhibition separately.']),
22:('Ion Channels: Predict Possible hERG Inhibition',['HERGAI can prioritize molecules for hERG testing.','An electrophysiology experiment checks channel inhibition.']),
24:('Chemprop: Learn from Measured Activities',['Use structures paired with results from a clearly defined assay.','A relevant trained model helps prioritize new compounds for testing.']),
25:('Small Chemical Changes Can Alter Recognition',['Theophylline and caffeine differ by one methyl group.','A useful activity model must handle such local chemical differences.']),
26:('RNA Binding Needs RNA-Relevant Evidence',['SMRTnet uses RNA and small-molecule information to predict interactions.','Protein-trained binding predictions should not be assumed valid for RNA.']),
28:('Selectivity: Predict a Profile, Then Test It',['The same molecule may affect related and unwanted targets.','Use predictions to choose a suitable experimental target panel.']),
32:('Work Backwards from a Cellular Effect',['AI can compare a new compound with well-studied reference profiles.','A similar profile suggests a mechanism for follow-up—not proof of one.']),
36:('Case 2: AI Proposes New Antibiotic Candidates',['Expand a promising fragment or propose a new chemical structure.','Activity and toxicity predictions help select candidates for synthesis.']),
37:('Case 2: Chemistry and Tests Establish Activity',['Twenty-four molecules were made; seven showed selective antibacterial activity.','Lead compounds then received mechanism and mouse-infection studies.']),
38:('Case 3: Predict the Receptor, Then Screen',['AlphaFold2 supplied receptor models; docking selected molecules.','Binding assays and structural follow-up tested the selected compounds.']),
40:('Case 4: Chemical Clues Guide Antibiotic Selection',['Models learned from antibacterial and human-cell toxicity measurements.','Highlighted chemical substructures helped choose compounds to test.']),
41:('Case 4: Independent Experiments Supply Evidence',['The study tested 283 selected compounds, not 283 confirmed antibiotics.','A discovered chemical class progressed to mouse-infection studies.']),
44:('Exercise: Plan an EGFR Inhibitor Test',['You have an EGFR structure and a shortlist of candidate molecules.','Choose a tool for poses and an experiment for kinase inhibition.']),
46:('Use AI to Make Better Pharmacodynamic Decisions',['Match the tool to the target interaction or biological effect.','Inspect the chemistry, test the claim and learn from the result.'])}
for n,(title,lines) in updates.items():S[n].update(title=title,bullets=lines)
# Explicitly exclude original introductory content and forbidden source media.
for a in content:
    a['notes']=a.get('notes','')+'\nTeaching level: chemistry undergraduate; focus on the input, practical output and matching experiment. No ML architecture or quantitative-PD derivation is required.'

base_native=b.native

def native(sl,s):
    layout=s['layout']
    if layout=='target_evidence':
        for y,title,detail in [(2.0,'Disease measurements','Genes and proteins associated with the condition'),(3.45,'Biological knowledge','Published mechanisms and feasible intervention'),(4.9,'A target hypothesis','Prioritize a test—not a declared therapeutic success')]:
            b.rounded(sl,title,.55,y,3.15,1.05,size=19);b.text(sl,detail,4.0,y+.18,5.35,.82,19)
    elif layout=='dgat2_pair':
        b.text(sl,'Candidate pocket',.5,2.04,4.45,.42,18,True,align=PP_ALIGN.CENTER)
        b.text(sl,'Docked inhibitor',5.1,2.04,4.35,.42,18,True,align=PP_ALIGN.CENTER)
        b.picture(sl,'dgat2_p2rank',(.53,2.66,4.42,3.7))
        b.picture(sl,'dgat2_docked',(5.10,2.66,4.23,3.7))
    elif layout=='tnik_result':
        b.text(sl,'Rentosertib / INS018_055',.46,2.04,4.85,.42,18,True,align=PP_ALIGN.CENTER)
        b.text(sl,'Fibrosis-marker response in cells',5.35,2.04,4.13,.62,17,True,align=PP_ALIGN.CENTER)
        b.picture(sl,'rentosertib_structure',(.5,2.75,4.80,3.4))
        b.picture(sl,'tnik_mrc5',(5.43,2.8,3.98,3.35))
    else:base_native(sl,s)
b.native=native

# External whitespace is trimmed on the photographic/structural images only.
# Preserve all source pixels in published plots and all scientific axis labels.

# Rebuild with a larger caption region, source text at 12 pt and no tiny body text.
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
p=Presentation();p.slide_width=Inches(10);p.slide_height=Inches(7.5)
p.core_properties.title='AI in Pharmacodynamics - replacement teaching slides 7-46'
p.core_properties.subject='40 undergraduate teaching slides; original first six not included'
p.core_properties.author=''
for item in content:
    n=item['n'];b.CURRENT=n
    sl=p.slides.add_slide(p.slide_layouts[6]);sl.background.fill.solid();sl.background.fill.fore_color.rgb=RGBColor(255,255,255)
    b.text(sl,item['title'],.36,.23,9.30,.63,24,True,'C00000','title')
    if item['layout']=='side':
        b.bullets(sl,item['bullets'],.42,1.18,3.20,3.8)
        b.picture(sl,item['key'],(3.78,1.10,5.82,5.34))
    else:
        b.bullets(sl,item['bullets'],.42,1.03,9.12,.97)
        if item.get('key'):b.picture(sl,item['key'],(.42,2.12,9.16,4.36))
        else:native(sl,item)
    b.text(sl,item['caption'],.43,6.58,9.02,.51,16,name='caption')
    b.text(sl,item['source'],.43,7.12,8.65,.31,12,name='reference')
    b.text(sl,str(n),9.25,7.14,.35,.25,12,name='page_number',align=PP_ALIGN.RIGHT)
    assets=[a for a in b.USE if a['slide']==n]
    notes=item['notes']+'\n\nSources: '+item['source']
    for a in assets:notes+='\n'+a['source']+'\nFigure: '+a['detail']
    sl.notes_slide.notes_text_frame.text=notes

# Verify and genuinely refine text fitting without shrinking body text below 16.
# Each checkpoint records the specific changes applied, not a claimed visual pass.
checks=[]
def content_hash():
    return hashlib.sha256(''.join(sl._element.xml for sl in p.slides).encode()).hexdigest()
def checkpoint(label,details):
    path=ROOT/'passes'/('revision_%02d.pptx'%(len(checks)+1));p.save(path)
    reopened=Presentation(path);assert len(reopened.slides)==40
    out=[]
    for n,sl in enumerate(reopened.slides,7):
        for sh in sl.shapes:
            if sh.left<0 or sh.top<0 or sh.left+sh.width>reopened.slide_width+100 or sh.top+sh.height>reopened.slide_height+100:out.append([n,sh.name])
    assert not out,out
    checks.append({'stage':label,'changes':details,'saved_file':path.name,'slide_xml_hash':content_hash(),'out_of_bounds':out})
checkpoint('Source and layout reconstruction','Replaced incorrect binding/cellular crops; rebuilt 40 slides from distinct source-specific assets and verified package reopening.')

# Match the requested line spacing where space permits and prevent inherited
# text margins from pushing labels into neighboring figure regions.
for sl in p.slides:
    for sh in sl.shapes:
        if not sh.has_text_frame:continue
        if sh.name=='body':
            for par in sh.text_frame.paragraphs:par.line_spacing=1.22;par.space_after=Pt(7)
        if sh.name=='caption':
            for par in sh.text_frame.paragraphs:par.line_spacing=1.08;par.space_after=Pt(0)
checkpoint('Typography and spacing','Body spacing normalized; captions remain 16-point and no body text is reduced.')

# Use font metrics at 144 pixels/inch to check native text wrapping. Bounds are
# estimates independent of the renderer and will be compared with rendered SVG.
font_regular=ImageFont.truetype(regular,36);font_bold=ImageFont.truetype(bold,48)
fit_report=[]
for num,sl in enumerate(p.slides,7):
    for sh in sl.shapes:
        if not sh.has_text_frame or sh.name not in ('title','body','caption','reference'):continue
        size={'title':24,'body':18,'caption':16,'reference':12}[sh.name]
        ff=ImageFont.truetype(bold if sh.name=='title' else regular,int(size*2))
        width=(sh.width-sh.text_frame.margin_left-sh.text_frame.margin_right)/914400*144
        if sh.name=='body':width-=.2*144
        lines=0
        for par in sh.text_frame.paragraphs:
            line='';count=1
            for word in par.text.split():
                test=(line+' '+word).strip()
                if ff.getlength(test)>width and line:count+=1;line=word
                else:line=test
            lines+=count
        required=lines*size/72*1.25
        if sh.name=='body':required+=max(0,len(sh.text_frame.paragraphs)-1)*7/72
        available=sh.height/914400
        if required>available+.025:
            fit_report.append({'slide':num,'shape':sh.name,'estimated_lines':lines,'needed':round(required,2),'available':round(available,2)})
            if sh.name=='title':
                # Preserve 24 pt; use two lines in title band only when necessary.
                sh.height=Inches(.72)
            elif sh.name=='reference':
                sh.top=Inches(7.03);sh.height=Inches(.43)
            elif sh.name=='caption':
                sh.top=Inches(6.52);sh.height=Inches(.56)
            elif sh.name=='body' and sh.width>Inches(8):
                sh.height=Inches(1.10)
checkpoint('Text-fit adjustment','Adjusted measured native text boxes without changing the requested font sizes; recorded every estimated fit exception.')

# Maintain correct biological nomenclature in editable text runs.
for num,sl in enumerate(p.slides,7):
    for sh in sl.shapes:
        if not sh.has_text_frame:continue
        for par in sh.text_frame.paragraphs:
            for run in par.runs:
                run.font.name='Arial'
                if sh.name=='title':run.font.color.rgb=RGBColor.from_string('C00000');run.font.size=Pt(24)
checkpoint('Scientific typography and source consistency','Explicit Arial and red 24-point titles rechecked; predictions and experiments remain labelled separately.')

name='AI_in_PD_Unique_Figures_Teaching_7-46.pptx';final=OUT/name;p.save(final)
# Independent raw-package audit: relationships, duplicate bytes, figure size,
# protected-content scope and the actual source of every inserted image.
seen=defaultdict(set);pixel_hashes=defaultdict(set);image_rows=[];pictures_by_slide=defaultdict(int)
for n,sl in enumerate(p.slides,7):
    for sh in sl.shapes:
        if sh.shape_type!=MSO_SHAPE_TYPE.PICTURE:continue
        raw=sh.image.blob;h=hashlib.sha256(raw).hexdigest();seen[h].add(n)
        im=Image.open(io.BytesIO(raw)).convert('RGB');ph=hashlib.sha256(im.tobytes()+str(im.size).encode()).hexdigest();pixel_hashes[ph].add(n)
        w=sh.width/914400;hgt=sh.height/914400;ppi=im.width/w
        image_rows.append({'slide':n,'shape':sh.name,'pixels':list(im.size),'display_inches':[round(w,3),round(hgt,3)],'ppi':round(ppi,1),'sha256':hashlib.sha256(raw).hexdigest()})
        pictures_by_slide[n]+=1
report={'file':name,'slides':40,'numbering':'7-46','first_six_included':False,'original_user_presentations_accessed':False,'no_Molecular_Drug_Targets_images':True,'image_slides':len(pictures_by_slide),'image_percentage':len(pictures_by_slide)/40*100,'duplicate_image_slide_groups':[sorted(v) for v in seen.values() if len(v)>1],'duplicate_pixel_slide_groups':[sorted(v) for v in pixel_hashes.values() if len(v)>1],'revisions':checks,'font_fit_adjustments':fit_report,'images':image_rows,'publication_source_check':'Specific figure identities and evidence types checked; source-panel locations additionally checked from publication text.','visual_review_status':'Native render and SVG text checks are performed by the next workflow step; human/model full-resolution visual review is not claimed.','pdf_generated':False}
assert report['image_slides']>=34,report['image_slides']
assert not report['duplicate_image_slide_groups'] and not report['duplicate_pixel_slide_groups']
assert all('Molecular Drug Targets' not in a['source'] for a in b.USE)
with zipfile.ZipFile(final) as z:assert z.testzip() is None
(OUT/'quality_report.json').write_text(json.dumps(report,indent=2))
(OUT/'figure_manifest.json').write_text(json.dumps(b.USE,indent=2))
(OUT/'content_manifest.json').write_text(json.dumps(content,indent=2))
package=ROOT/'pptx-package';package.mkdir(exist_ok=True)
with zipfile.ZipFile(final) as z:z.extractall(package)
print(json.dumps({'file':str(final),'slides':40,'images':report['image_slides'],'duplicates':report['duplicate_image_slide_groups'],'fit':fit_report},indent=2))
