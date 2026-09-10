"""Build public-source teaching slides; never reads or uploads a user's presentation.
An optional LOCAL merge utility preserves an existing presentation's first six slides.
"""
from __future__ import annotations
import argparse, hashlib, io, json, math, os, re, subprocess, textwrap, time, zipfile
from pathlib import Path
from collections import Counter, defaultdict
from urllib.parse import quote
import requests
from PIL import Image, ImageChops, ImageDraw, ImageFont
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from rdkit import Chem
from rdkit.Chem import rdDepictor
from rdkit.Chem.Draw import rdMolDraw2D

ROOT=Path('pd-build'); AS=ROOT/'assets'; OUT=ROOT/'output'
for d in (ROOT,AS,OUT):d.mkdir(parents=True,exist_ok=True)
SESSION=requests.Session()
SESSION.headers['User-Agent']='EducationalScientificFigureBuilder/1.0'
BLUE='#467C91'; GREEN='#688C72'; AMBER='#A67F49'; PURPLE='#84759C'; INK='#20252B'
FONT='/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf'
BOLD='/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf'
PROV={}; USE=[]

def fetch(url,path,tries=3):
    path=Path(path)
    if path.exists() and path.stat().st_size>100:return path
    for i in range(tries):
        try:
            r=SESSION.get(url,timeout=70);r.raise_for_status();path.write_bytes(r.content);return path
        except Exception:
            if i==tries-1:raise
            time.sleep(2)

def figure(key,path,source,detail):
    p=Path(path); im=Image.open(p);im.load()
    PROV[key]={'path':str(p),'source':source,'detail':detail,'pixels':list(im.size),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
    return key

def public_assets():
    token=os.environ.get('GH_TOKEN')
    if token:
        r=SESSION.get('https://api.github.com/repos/GUOAOXIN1/projects/actions/artifacts/10142151178/zip',headers={'Authorization':'Bearer '+token},timeout=120)
        r.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            for info in z.infolist():
                p=Path(info.filename)
                if p.is_absolute() or '..' in p.parts:raise ValueError('Unsafe archive entry')
                if not info.is_dir():
                    (AS/p.name).write_bytes(z.read(info))
    rows=json.loads((AS/'provenance.json').read_text())
    for row in rows:
        p=AS/row.get('file','')
        if p.is_file() and p.suffix.lower() in ('.png','.jpg','.gif'):
            try:figure(p.stem,p,row.get('source',''),row.get('caption','Public-source original figure; source pixels retained.'))
            except Exception:pass
    for name,doi,stem,num in [
        ('reinvent_fig4','10.1186/s13321-024-00812-5','13321_2024_812',4),
        ('af3_fig5','10.1038/s41586-024-07487-w','41586_2024_7487',5)]:
        url=f'https://media.springernature.com/full/springer-static/image/art%3A{quote(doi,safe="")}/MediaObjects/{stem}_Fig{num}_HTML.png'
        try:figure(name,fetch(url,AS/(name+'.png')),'https://doi.org/'+doi,'Original published figure')
        except Exception as exc:print('OPTIONAL_ASSET_FAILED',name,str(exc))

# Cropping is lossless and cannot overlap a panel reused on another slide.
CROPS=defaultdict(list)
def crop(key,parent,box,detail):
    src=PROV[parent]; im=Image.open(src['path']).convert('RGB');w,h=im.size
    b=tuple(int(v*s) for v,s in zip(box,(w,h,w,h)))
    for prior in CROPS[parent]:
        x=max(0,min(b[2],prior[2])-max(b[0],prior[0]));y=max(0,min(b[3],prior[3])-max(b[1],prior[1]))
        if x*y>0:raise ValueError('Repeated/overlapping source panel: '+parent)
    CROPS[parent].append(b)
    p=AS/(key+'.png');im.crop(b).save(p)
    figure(key,p,src['source'],detail+'; lossless crop '+str(b)+' from '+parent)

def font(size,bold=False):return ImageFont.truetype(BOLD if bold else FONT,size)
def canvas():return Image.new('RGB',(2400,1280),'white')
def label(draw,xy,text,size=54,bold=False,fill=INK,anchor=None):
    draw.text(xy,text,font=font(size,bold),fill=fill,anchor=anchor,spacing=12)
def arrow(draw,a,b,fill=BLUE,width=7):
    draw.line([a,b],fill=fill,width=width)
    ang=math.atan2(b[1]-a[1],b[0]-a[0]);r=23
    draw.polygon([b,(b[0]-r*math.cos(ang-.48),b[1]-r*math.sin(ang-.48)),(b[0]-r*math.cos(ang+.48),b[1]-r*math.sin(ang+.48))],fill=fill)
def save_art(key,im,source,detail):
    p=AS/(key+'.png');im.save(p);figure(key,p,source,detail)

def mol_image(smiles,w=700,h=520,highlight=None):
    mol=Chem.MolFromSmiles(smiles)
    if mol is None:raise ValueError('Invalid chemical structure: '+smiles)
    rdDepictor.Compute2DCoords(mol)
    d=rdMolDraw2D.MolDraw2DCairo(w,h)
    opts=d.drawOptions();opts.bondLineWidth=3.4;opts.minFontSize=32;opts.maxFontSize=45;opts.padding=.08
    opts.useBWAtomPalette()
    d.DrawMolecule(mol,highlightAtoms=highlight or []);d.FinishDrawing()
    return Image.open(io.BytesIO(d.GetDrawingText())).convert('RGB'),mol

def chemical_figures():
    im=canvas();d=ImageDraw.Draw(im)
    sm='C[NH2+]C[C@H](O)c1ccc(O)c(O)c1';m=Chem.MolFromSmiles(sm);Chem.AssignStereochemistry(m,cleanIt=True,force=True)
    centers=Chem.FindMolChiralCenters(m,includeUnassigned=False)
    if centers[0][1]!='R':sm=sm.replace('[C@H]','[C@@H]')
    mi,m=mol_image(sm,1350,1000);im.paste(mi,(120,120))
    label(d,(1580,190),'Adrenaline input',58,True)
    for y,t in [(370,'Correct connectivity'),(565,'Defined R stereocentre'),(760,'Protonated secondary amine')]:label(d,(1530,y),t,48)
    save_art('adrenaline_input',im,'Chemical teaching example; R-adrenaline connectivity and stereochemistry checked with RDKit.','Original high-resolution chemical illustration. No activity prediction is assigned.')
    im=canvas();d=ImageDraw.Draw(im)
    smiles=['Cn1c(=O)c2[nH]cnc2n(C)c1=O','Cn1c(=O)c2c(ncn2C)n(C)c1=O']
    for i,(s,t) in enumerate(zip(smiles,['Theophylline','Caffeine'])):
        mi,m=mol_image(s,1000,760);im.paste(mi,(90+i*1200,170));label(d,(590+i*1200,1060),t,64,True,anchor='mm')
    label(d,(1200,75),'A single methyl group changes the recognition problem',55,True,anchor='mm')
    save_art('xanthine_pair',im,'Zimmermann et al., Nat. Struct. Biol. 4, 644–649 (1997). doi:10.1038/nsb0897-644','Theophylline and caffeine chemical structures; a known RNA aptamer discriminates between them. No new AI prediction.')
    im=canvas();d=ImageDraw.Draw(im)
    candidates=['CC(=O)Nc1cccc(Cl)c1','O=C(Nc1ccncc1)c1ccccc1','CCOc1ccc(C(N)=O)cc1','CCN1CCC(c2ccccc2)CC1','COc1ccc2[nH]c(C)nc2c1','O=C(O)c1ccc(Cl)cc1','CCOc1ccc(NC(N)=O)cc1','CC(=O)N1CCN(c2ccccc2)CC1','COc1cc(C(N)=O)ccc1F']
    for k,s in enumerate(candidates):
        mi,m=mol_image(s,520,280);im.paste(mi,(45+(k%3)*565,130+(k//3)*320))
    arrow(d,(1740,595),(1940,595));label(d,(2175,360),'AI-assisted',58,True,anchor='mm');label(d,(2175,440),'ranking',58,True,anchor='mm')
    for y,t in [(650,'Inspect'),(780,'Select'),(910,'Test')]:label(d,(2175,y),t,56,anchor='mm')
    save_art('screening_collection',im,'McNutt et al., J. Cheminform. 17, 28 (2025). doi:10.1186/s13321-025-00973-x','Original teaching illustration: distinct plausible chemical candidates, with no assigned activity or claimed model output.')
    im=canvas();d=ImageDraw.Draw(im)
    examples=[('COc1ccccc1C(N)=O','Measured result'),('O=C(NCc1ccccc1)c1ccncc1','Measured result'),('CCOc1ccccc1C(=O)N','Predict, then test')]
    for i,(s,t) in enumerate(examples):
        mi,m=mol_image(s,690,540);im.paste(mi,(70+790*i,200));label(d,(415+790*i,875),t,55,True,anchor='mm')
    label(d,(1200,70),'Use one clearly defined assay',60,True,anchor='mm')
    label(d,(800,1130),'Known examples',55,anchor='mm');arrow(d,(1400,1100),(1610,1100));label(d,(1940,1130),'New candidate',55,anchor='mm')
    save_art('assay_examples',im,'Heid et al., J. Chem. Inf. Model. 64, 9–17 (2024). doi:10.1021/acs.jcim.3c01250','Original illustration of assay-trained prediction, not a measured dataset or actual Chemprop prediction.')
    # Qualitative activity panel: labels state explicitly that data are hypothetical.
    im=canvas();d=ImageDraw.Draw(im)
    cols=['Target A','Target B','Target C','hERG'];rows=['Candidate 1','Candidate 2','Candidate 3'];vals=[[.92,.11,.08,.06],[.74,.53,.12,.09],[.79,.71,.52,.68]]
    for j,t in enumerate(cols):label(d,(800+j*390,140),t,55,True,anchor='mm')
    for i,t in enumerate(rows):
        label(d,(100,330+i*270),t,53,True)
        for j,v in enumerate(vals[i]):
            c=tuple(int(244+(x-244)*v) for x in (63,112,132));x=630+j*390;y=270+i*270
            d.rounded_rectangle((x,y,x+325,y+190),radius=15,fill=c)
    label(d,(1200,1160),'Darker cells = stronger inhibition in this invented example',50,anchor='mm')
    save_art('selectivity_panel',im,'Hopkins, Nat. Chem. Biol. 4, 682–690 (2008). doi:10.1038/nchembio.118','Original hypothetical target-panel illustration; not measured activity or a safety rating.')
    im=canvas();d=ImageDraw.Draw(im);rng=np.random.default_rng(7)
    profiles=np.array([[.8,.7,-.3,-.5,.4,.9,-.7,-.5],[.85,.62,-.25,-.44,.35,.8,-.65,-.6],[-.7,.4,.9,.2,-.5,-.4,.5,.8],[.1,-.7,.4,.85,.5,-.6,.2,-.1]])
    names=['New compound','Reference A','Reference B','Reference C']
    for i,t in enumerate(names):
        label(d,(55,240+i*230),t,50,True)
        for j,v in enumerate(profiles[i]):
            target=(88,129,153) if v>0 else (186,135,101);a=abs(v);c=tuple(int(245+(x-245)*a) for x in target)
            x=605+j*210;y=180+i*230;d.rectangle((x,y,x+190,y+180),fill=c)
    label(d,(1490,90),'Different measured features',55,True,anchor='mm')
    label(d,(1200,1175),'Similar profiles suggest a testable mechanism—not proof of one',48,anchor='mm')
    save_art('profile_comparison',im,'Chandrasekaran et al., Nat. Methods 21, 1114–1121 (2024). doi:10.1038/s41592-024-02241-6','Original simulated perturbation profiles, explicitly distinct from published microscopy and measured study results.')
    im=canvas();d=ImageDraw.Draw(im)
    for side,(title,hits) in enumerate([('AI-selected set',7),('Comparison set',2)]):
        x0=180+side*1180;label(d,(x0+460,100),title,60,True,anchor='mm')
        for k in range(20):
            x=x0+(k%5)*185;y=245+(k//5)*195;c=GREEN if k<hits else '#E4E7E8'
            d.ellipse((x,y,x+125,y+125),fill=c,outline=INK,width=2)
        label(d,(x0+460,1150),f'{hits} useful hits / 20 tested',55,True,anchor='mm')
    save_art('hit_comparison',im,'Hypothetical teaching comparison; not results for a named AI model.','Original dot display of two equal-size test sets. Not a statistical demonstration of improvement.')
    im=canvas();d=ImageDraw.Draw(im)
    for i,(t,p,c) in enumerate([('AF2 receptor model',54,BLUE),('Experimental structure',51,GREEN)]):
        y=260+i*380;label(d,(80,y+65),t,52,True);d.rectangle((790,y,790+p*24,y+180),fill=c);label(d,(830+p*24,y+60),f'≈{p}%',60,True)
    label(d,(1200,1060),'Published σ₂ screening campaigns',56,True,anchor='mm')
    label(d,(1200,1160),'Similar hit rates in this comparison; not universal equivalence',46,anchor='mm')
    save_art('lyu_hit_rates',im,'Lyu et al., Science 384, eadn6354 (2024). doi:10.1126/science.adn6354','Redrawn published rounded hit rates: AF2 64/119 (~54%); experimental-structure campaign ~51%. Rates not significantly different.')
    im=canvas();d=ImageDraw.Draw(im)
    label(d,(1200,85),'RNA-binding predictions were checked experimentally',58,True,anchor='mm')
    for i,(t,n,c) in enumerate([('Selected predictions tested',190,'#DCE5E9'),('Confirmed interactions',40,GREEN)]):
        y=280+350*i;label(d,(90,y+70),t,51,True);d.rectangle((880,y,880+n*6.2,y+185),fill=c);label(d,(930+n*6.2,y+60),str(n),66,True)
    label(d,(1200,1120),'Ten disease-associated RNA targets; MST binding tests',50,anchor='mm')
    save_art('rna_validation',im,'Fei et al., Nat. Biotechnol. (2026). doi:10.1038/s41587-025-02942-z','Original chart of published counts: 190 predicted small molecule–RNA interactions tested; 40 confirmed. These are interactions, not necessarily 190 unique molecules.')

# Every structural rendering has an explicit experimental identity.
def structural_figures():
    import pymol
    pymol.finish_launching(['pymol','-cq'])
    from pymol import cmd
    def reset():
        cmd.reinitialize();cmd.bg_color('white');cmd.set('orthoscopic',1);cmd.set('ray_opaque_background',1);cmd.set('ray_shadows',0);cmd.set('antialias',2);cmd.set('ambient',.5);cmd.set('specular',.16);cmd.set('cartoon_fancy_helices',1);cmd.set('cartoon_smooth_loops',1);cmd.set('stick_radius',.20);cmd.set('sphere_scale',.28)
    def getpdb(p):return fetch(f'https://files.rcsb.org/download/{p}.pdb',AS/(p+'.pdb'))
    def export(key,source,detail,selection='all',buffer=2):
        cmd.zoom(selection,buffer);path=AS/(key+'.png');cmd.png(str(path),width=2400,height=1500,dpi=300,ray=1);figure(key,path,source,detail)
    def pocket(p,lig,key,chain='A',buffer=8):
        reset();cmd.load(str(getpdb(p)),'mol');cmd.remove('solvent')
        if chain:cmd.remove('not chain '+chain)
        cmd.hide('everything');cmd.show('cartoon','polymer.protein');cmd.color('lightblue','polymer.protein')
        cmd.select('ligand','resn '+lig)
        if cmd.count_atoms('ligand')<4:raise ValueError('Ligand missing '+p+' '+lig)
        cmd.show('sticks','ligand');cmd.color('orange','ligand and elem C');cmd.color('red','ligand and elem O');cmd.color('blue','ligand and elem N');cmd.color('yellow','ligand and elem S')
        cmd.show('sticks','byres (polymer.protein within 4.0 of ligand)');cmd.orient('ligand')
        export(key,'https://doi.org/10.2210/pdb'+p+'/pdb','New ray-traced rendering of experimental coordinates '+p+'; ligand '+lig+'. Not an AI prediction.','ligand',buffer)
    for p,l,k,c,b in [('1UDT','VIA','pde5_sildenafil','A',8),('1M17','AQ4','egfr_erlotinib','A',8),('1HWL','FBI','hmgr_rosuvastatin','',9),('2EX6','AIX','pbp_ampicillin','A',7),('1XKK','FMM','egfr_lapatinib','A',8)]:
        pocket(p,l,k,c,b)
    reset()
    for p in ['1ERE','1ERR']:cmd.load(str(getpdb(p)),p)
    cmd.remove('not chain A');cmd.remove('solvent');cmd.align('1ERR and polymer.protein','1ERE and polymer.protein');cmd.hide('everything');cmd.show('cartoon','polymer.protein');cmd.color('lightblue','1ERE');cmd.color('wheat','1ERR');cmd.color('marine','1ERE and resi 538-552');cmd.color('orange','1ERR and resi 538-552');cmd.show('sticks','resn EST+RAL');cmd.color('marine','1ERE and resn EST');cmd.color('orange','1ERR and resn RAL');cmd.orient('1ERE and polymer.protein');cmd.turn('x',20)
    export('er_states','https://doi.org/10.2210/pdb1ERE/pdb ; https://doi.org/10.2210/pdb1ERR/pdb','Experimental oestrogen receptor comparison: estradiol-bound blue and raloxifene-bound wheat/orange; helix 12 highlighted. No AI calculation.',buffer=1)
    reset();cmd.load(str(getpdb('3SN6')),'active');cmd.load(str(getpdb('2RH1')),'inactive');cmd.remove('(active and not chain R) or (inactive and not chain A)');cmd.remove('solvent');cmd.align('active and polymer.protein and resi 35-342','inactive and polymer.protein and resi 35-342');cmd.hide('everything');cmd.show('cartoon','polymer.protein');cmd.color('lightblue','inactive');cmd.color('wheat','active');cmd.color('marine','inactive and resi 267-298');cmd.color('orange','active and resi 267-298');cmd.orient('inactive and polymer.protein and resi 35-342')
    export('gpcr_states','https://doi.org/10.2210/pdb3SN6/pdb ; https://doi.org/10.2210/pdb2RH1/pdb','Experimental beta-2 adrenergic receptor states, inactive blue and active wheat; TM6 highlighted. Experimental comparison, not an AI trajectory.','polymer.protein and resi 35-342',2)
    reset();cmd.load(str(getpdb('1EHT')),'rna',state=1);cmd.frame(1);cmd.remove('solvent');cmd.hide('everything');cmd.show('cartoon','polymer.nucleic');cmd.color('lightblue','polymer.nucleic');cmd.show('sticks','resn TEP or (byres (polymer.nucleic within 4.0 of resn TEP))');cmd.color('orange','resn TEP');cmd.orient('all')
    export('rna_theophylline','https://doi.org/10.2210/pdb1EHT/pdb','Theophylline-bound RNA aptamer. First deposited NMR model, not the entire ensemble or an AI-generated result.',buffer=2)
    reset();cmd.load(str(getpdb('5VA1')),'herg');cmd.remove('solvent');cmd.hide('everything');cmd.show('cartoon','polymer.protein')
    for ch,col in zip(cmd.get_chains('polymer.protein'),['lightblue','wheat','palegreen','lightpink']):cmd.color(col,'chain '+ch)
    cmd.orient('polymer.protein')
    export('herg_channel','https://doi.org/10.2210/pdb5VA1/pdb','Experimental hERG channel structure; no blocker pose or new HERGAI output is shown.',buffer=2)
    reset();cmd.load(str(getpdb('1D12')),'dna');cmd.remove('solvent');cmd.hide('everything');cmd.show('cartoon','polymer.nucleic');cmd.show('sticks','organic');cmd.color('lightblue','polymer.nucleic');cmd.color('orange','organic');cmd.orient('all')
    export('dna_complex','https://doi.org/10.2210/pdb1D12/pdb','Experimental anthracycline–DNA complex from entry 1D12; no sequence-selective AI drug prediction is asserted.',buffer=2)
    # A target's allosteric ligand demonstrates why the binding site must be specified.
    reset();cmd.load(str(getpdb('7QIE')),'allosteric');cmd.remove('solvent');cmd.hide('everything');cmd.show('cartoon','polymer.protein');cmd.color('lightblue','polymer.protein');cmd.show('sticks','organic');cmd.color('orange','organic');cmd.orient('all')
    export('allosteric_pocket','https://doi.org/10.2210/pdb7QIE/pdb','Experimental PI5P4K-gamma allosteric inhibitor complex 7QIE; NIH-12848 analogue. Not an AI affinity calculation.',buffer=2)
    reset();cmd.load(str(getpdb('8UWL')),'complex');cmd.remove('solvent');cmd.hide('everything');cmd.show('cartoon','polymer.protein');cmd.color('lightblue','polymer.protein');cmd.show('sticks','organic');cmd.color('orange','organic');cmd.orient('all')
    export('lyu_cryoem','https://doi.org/10.2210/pdb8UWL/pdb ; https://doi.org/10.1126/science.adn6354','Experimental 5-HT2A complex associated with the Lyu et al. prospective screening study; labelled as experimental follow-up, not an AF2 model.',buffer=2)

# Plain-English statements; technical qualifications and source provenance stay in notes.
SPECS=[]
def S(n,title,bullets,key,caption,source,notes='',layout='wide'):
    SPECS.append(dict(n=n,title=title,bullets=bullets,key=key,caption=caption,source=source,notes=notes,layout=layout))

def curriculum():
    A='Abramson et al., Nature 630, 493–500 (2024).'
    G='McNutt et al., J. Cheminform. 17, 28 (2025).'
    R='Ren et al., Nat. Biotechnol. 43, 63–75 (2025; online 2024).'
    W='Wong et al., Nature 626, 177–185 (2024).'
    L='Lyu et al., Science 384, eadn6354 (2024).'
    K='Krishnan et al., Cell 188, 5962–5979.e22 (2025).'
    RE='Loeffler et al., J. Cheminform. 16, 20 (2024).'
    SM='Fei et al., Nat. Biotechnol. (2026), doi:10.1038/s41587-025-02942-z.'
    C='Chandrasekaran et al., Nat. Methods 21, 1114–1121 (2024).'
    S(7,'From Target Chemistry to AI-Assisted PD',['Use what you know about target structure and chemical recognition.','Ask how AI helps choose molecules and the experiments that test them.'],'af3_intro','Published AF3 complexes: different molecules require different recognition models.',A,'Connect to prior target classes without reteaching the pharmacology chapters. These are published predictions, not calculations made for this lesson.')
    S(8,'Follow a Drug–Target Research Question',['Define the interaction or effect before choosing a tool.','Move from a target to a testable chemical decision.'],None,'Target → structure → candidates → experiments → next decision.','Teaching workflow.',layout='workflow')
    S(9,'AI Tools for Pharmacodynamic Research',['Choose by the question, the available data and the meaning of the output.'],None,'Examples of useful research tools—not a popularity or performance ranking.','Official documentation and original tool papers; details in notes.',layout='tools')
    S(10,'Use Disease Evidence to Prioritize Targets',['Combine disease-related measurements with established biological knowledge.','AI ranking proposes a target; it does not validate the target.'],'tnik_evidence','TNIK study: biological evidence used in target prioritization.',R,'Input-evidence panel from the published workflow. Explain genes, proteins and disease associations, not the model architecture.')
    S(11,'Enzymes: Start from a Real Binding Complex',['Sildenafil–PDE5 provides a known inhibitor and a defined pocket.','Use that reference to prepare an AI-assisted screening question.'],'pde5_sildenafil','Experimental sildenafil–PDE5 structure, PDB 1UDT.','RCSB PDB 1UDT.','Sildenafil was not discovered by the AI methods taught here. This experimental reference is specific to the stated enzyme project.','side')
    S(12,'Predict a Complex When a Structure Is Missing',['Provide the identities of the interacting molecules.','Use predicted geometry as a starting hypothesis—not measured activity.'],'af3_fig3','AF3 examples include protein–RNA, antibody and small-molecule complexes.',A,'Published Figure 3. Blue/green: predicted proteins; orange: predicted ligands/glycans; purple: predicted RNA; grey: experimental comparison. Explain selected examples, not numerical structure metrics.')
    S(13,'Find Candidate Pockets with P2Rank',['A pocket predictor highlights regions that may bind a ligand.','Known chemistry and experiments help decide which site matters.'],'dgat2_pocket','Published DGAT2 pocket illustration; a candidate region is not a confirmed drug site.','Tunyasuvunakool et al., Nature 596, 590–596 (2021); P2Rank.','The panel concerns DGAT2 and predicted pocket features. P2Rank/PrankWeb proposes binding regions, not affinity or cellular efficacy.','side')
    S(14,'GPCRs: The Receptor State Matters',['Active and inactive receptors need not present the same geometry.','Choose a relevant state before interpreting a ligand prediction.'],'gpcr_states','β₂-adrenoceptor: inactive and active experimental structures; TM6 highlighted.','RCSB PDB 2RH1 and 3SN6; Rasmussen et al., Nature 477, 549–555 (2011).','Inactive blue; active wheat/orange. The overlay is a comparison of experimental states, not a computed transition or thermodynamic ensemble.','side')
    S(15,'Give the Tool the Correct Chemical Structure',['Keep connectivity, charge and stereochemistry explicit.','Adrenaline illustrates why those chemical details cannot be ignored.'],'adrenaline_input','R-adrenaline shown with a protonated secondary amine.','Original chemical illustration; molecular identity checked computationally.','The input form is a teaching example. No numerical affinity or activity is assigned.')
    S(16,'Docking: Propose How a Ligand Fits',['GNINA combines docking with learned pose scoring.','Compare the proposal with known contacts and experimental structures.'],'egfr_erlotinib','Experimental erlotinib–EGFR reference, PDB 1M17; not a new GNINA result.',G+' RCSB PDB 1M17.','This defined kinase–inhibitor complex explains what a pose is and what redocking would try to recover. Conventional sampling and AI scoring are different operations.','side')
    S(17,'Inspect Contacts with the Chemistry You Know',['Check whether polar and nonpolar groups face suitable environments.','A contact picture suggests chemistry; contact counting does not measure affinity.'],'hmgr_rosuvastatin','Rosuvastatin in the experimental HMG-CoA reductase complex, PDB 1HWL.','Istvan & Deisenhofer, Science 292, 1160–1164 (2001); PDB 1HWL.','The experimental statin complex is used as a chemically specific reference for inspecting model predictions. The drug is not described as AI-discovered.','side')
    S(18,'Search Beyond the Usual Active Site',['An allosteric ligand can act at a different pocket.','Specify the intended site when predicting or scoring binding.'],'allosteric_pocket','Experimental allosteric inhibitor complex of PI5P4Kγ, PDB 7QIE.',A+' RCSB PDB 7QIE.','The AF3 paper includes this allosteric-site example. This new rendering uses experimental coordinates; it is not the published AF3 prediction.','side')
    S(19,'Boltz-2: Read the Right Output',['Structural confidence, binding likelihood and activity estimates are different.','Validate the particular claim represented by each output.'],None,'A confident-looking structure is not an experimental binding measurement.','Boltz-2 original paper and official prediction documentation (2025).','Do not teach logarithmic score conversion. Do not rename predicted activity as experimental Kd. Model setup and its documented applicability matter.',layout='outputs')
    S(20,'Nuclear Receptors: Binding Is Not Function',['Oestradiol and raloxifene bind the receptor but stabilize different arrangements.','A pose prediction alone cannot determine agonism or antagonism.'],'er_states','Oestrogen receptor comparison: oestradiol blue; raloxifene wheat/orange.','Brzozowski et al., Nature 389, 753–758 (1997); PDB 1ERE/1ERR.','Helix 12 is highlighted. Raloxifene is a selective oestrogen receptor modulator with tissue-dependent effects, not a universal antagonist.','side')
    S(21,'Covalent Inhibitors Need a Covalent-Aware Model',['Ampicillin forms a covalent acyl-enzyme complex with a PBP.','Check the reaction and the bound chemical form—not just the fit.'],'pbp_ampicillin','Experimental covalent ampicillin–PBP4 complex, PDB 2EX6.','Kishida et al., Biochemistry 45, 783–792 (2006). '+G,'PBP4 is a cell-wall-remodelling DD-peptidase, not the bifunctional PBP1b. GNINA covalent docking requires the bound form; it does not predict the reaction. Its CNN was not trained on covalent complexes.','side')
    S(22,'Ion Channels: Flag Possible hERG Blockers',['HERGAI predicts hERG inhibition from molecule–channel information.','Use the prediction to prioritize electrophysiology tests.'],'herg_channel','Experimental hERG channel structure, PDB 5VA1; no blocker pose is asserted.','Tran-Nguyen et al., J. Cheminform. 17, 110 (2025); PDB 5VA1.','hERG is a potassium channel. A blocker prediction is not a probability of clinical arrhythmia. No architecture or threshold calculations are needed.','side')
    S(23,'Screen Broadly; Test a Manageable Shortlist',['AI-assisted ranking can reduce the number of compounds sent to the laboratory.','Keep chemically sensible and sufficiently diverse candidates.'],'screening_collection','Illustrative candidate collection; no activities or model scores are assigned.',G,'The structures are distinct synthetic teaching candidates, not compounds from a claimed experiment.')
    S(24,'Predict Activity from Measured Examples',['Chemprop can learn a defined activity endpoint from appropriate assay data.','Test predictions on compounds not used to build the model.'],'assay_examples','A structure and an assay result form a learning example.','Heid et al., J. Chem. Inf. Model. 64, 9–17 (2024).','Only an input–output explanation is needed. Chemprop needs relevant data or a suitable existing model; it is not universally pretrained for every target.')
    S(25,'Small Chemical Changes Can Matter',['Theophylline and caffeine differ by one methyl group.','A model must capture the recognition problem—not just overall similarity.'],'xanthine_pair','A theophylline-binding RNA aptamer discriminates against caffeine.','Zimmermann et al., Nat. Struct. Biol. 4, 644–649 (1997).','This is a documented molecular-recognition example, not a new AI prediction or a general rule about methylation.')
    S(26,'RNA Targets Need RNA-Relevant Models',['SMRTnet combines RNA sequence/secondary structure with the candidate molecule.','Protein-trained binding predictions should not be transferred to RNA without testing.'],'rna_theophylline','Experimental theophylline–RNA aptamer, PDB 1EHT; one NMR model.',SM+' PDB 1EHT.','The specific aptamer illustrates RNA recognition; SMRTnet studied theophylline-related discrimination. This rendering is experimental reference geometry, not a model output.','side')
    S(27,'Ask AI to Propose New Chemical Ideas',['REINVENT 4 can decorate a scaffold, link fragments or propose new molecules.','Design goals guide proposals; synthesis and assays determine their value.'],'reinvent_fig4','Four molecular-design tasks in REINVENT 4.',RE,'Original Figure 4. Explain the chemical changes and fixed starting information; omit neural architectures and optimization equations.')
    S(28,'Predict a Target Profile, Not Just One Score',['Related proteins and unwanted targets may also interact with the molecule.','Follow promising predictions with an appropriate target panel.'],'selectivity_panel','Hypothetical matched-assay inhibition patterns—not measured candidates.','Hopkins, Nat. Chem. Biol. 4, 682–690 (2008).','Some multi-target effects are useful. Selectivity is not sufficient to establish clinical safety.')
    S(29,'DNA Recognition Is a Different Binding Problem',['Intercalation and groove recognition require the appropriate molecular geometry.','Use DNA-relevant data and validation, not untested transfer from protein models.'],'dna_complex','Experimental anthracycline–DNA complex, PDB 1D12.','RCSB PDB 1D12; drug–nucleic-acid recognition context.','This extends the previous nucleic-acid course to model choice. No universal AI DNA-affinity tool or sequence selectivity is claimed.','side')
    S(30,'Check the Predicted Interaction Directly',['A binding experiment addresses whether the proposed interaction occurs.','It does not by itself establish the entire cellular mechanism.'],'tnik_binding','Published TNIK binding evidence; this is a measurement, not a docking score.',R,'Explain the measurement type without kinetic fitting or quantitative PD derivations.')
    S(31,'Use Cell Images to Measure the Response',['Images capture changes in cell shape and internal organization.','AI-assisted image analysis needs controls and a clearly defined biological question.'],'cellpaint_fig1','Published Cell Painting example; images are experimental inputs.',C,'Do not describe a fluorescence image as a direct binding assay or infer an unstated cell identity. Cell Painting and image segmentation are not equivalent to target identification.')
    S(32,'Work Backwards from an Unexplained Cell Effect',['Compare the compound’s profile with well-characterized references.','A similar pattern suggests a mechanism to test, not a proven shared target.'],'profile_comparison','Simulated profiles show how reference matching can generate a hypothesis.',C,'Transcriptomic and morphological profiles are complementary sources, not mandatory sequential stages. Distinguish shared stress from shared target action.')
    S(33,'Let Experiments Change the Next Design',['Use both successful and unsuccessful compounds to refine the next choice.','The best next experiment may test a new chemical idea, not the highest score.'],None,'Design → make → test → learn: AI supports the experimental cycle.',RE,'Do not introduce acquisition functions or claim that synthesis and biological validation have been automated in this teaching example.',layout='cycle')
    S(34,'Case 1: AI Helped Design TNIK Inhibitors',['After target prioritization, Chemistry42 proposed candidate structures.','Synthesis and experimental filtering guided candidate selection.'],'tnik_design','Published structure-based design inputs and filters—not a fully automated discovery.',R,'The approximately eighteen-month milestone was preclinical candidate nomination, not first-in-human testing. It is not a controlled estimate of the time saved by AI.')
    S(35,'Case 1: Separate Prediction from Measured PD',['The candidate bound TNIK and inhibited the enzyme.','Cell assays showed fibrosis-related effects; those need mechanistic interpretation.'],'tnik_cell','Published cellular assay evidence for INS018_055.',R,'Broader kinase activity prevents attributing every cellular effect exclusively to TNIK. A predicted pose is not an experimental co-crystal. No clinical approval is claimed.')
    S(36,'Case 2: Generate Antibiotic Candidates',['Expand a selected fragment or propose compounds without a fixed fragment.','Predicted activity and toxicity guide which molecules are made.'],'antibiotics_workflow','Published generative-antibiotic strategy; candidate generation precedes validation.',K,'Do not teach VAE, genetic-algorithm or graph-network details. Large computational candidate pools are not experimentally tested compounds.')
    S(37,'Case 2: From Candidates to Antibacterial Evidence',['Twenty-four compounds were synthesized; seven had selective antibacterial activity.','Lead compounds received mechanism and mouse-infection follow-up.'],'antibiotics_evidence','Published antibacterial study: chemistry and experiments, not human efficacy.',K,'NG1 was investigated against Neisseria gonorrhoeae and DN1 against Staphylococcus aureus. Mechanistic follow-up is not direct proof of a resolved drug–target complex.')
    S(38,'Case 3: Predict the Receptor, Then Screen Ligands',['AlphaFold2 supplied receptor models; docking selected candidate molecules.','Experiments checked binding and followed up selected complexes.'],'lyu_cryoem','Experimental 5-HT₂A complex from the study’s structural follow-up, PDB 8UWL.',L,'The pictured structure is explicitly experimental follow-up. AI predicted the receptor; the screened molecules were not generated by AlphaFold2.','side')
    S(39,'Case 3: Did Model-Based Screening Help?',['Both predicted-structure and experimental-structure campaigns found binders.','Similar results in this comparison do not guarantee success on every target.'],'lyu_hit_rates','Rounded published hit rates refer to compounds actually tested.',L,'The model-based sigma-2 campaign found 64 hits among 119 tested compounds. The roughly 54% versus 51% difference was not significant. Do not claim superiority.')
    S(40,'Case 4: Use Chemical Clues to Select Antibiotics',['Models learned from antibacterial activity and human-cell toxicity.','Highlighted substructures helped guide prospective compound selection.'],'wong_fig3','Published chemical rationales and candidate scaffolds.',W,'The source’s real chemical substructures replace unrelated teaching molecules. Explain a rationale as a model-associated chemical clue, not a proven biological mechanism.')
    S(41,'Case 4: Laboratory Evidence Goes Beyond a Score',['The study experimentally tested 283 selected compounds.','A discovered chemical class received mechanistic and mouse-infection studies.'],'wong_fig5','Published preclinical efficacy evidence; not a human clinical result.',W,'No claim that all 283 were active. Selected subgroup comparisons were not randomized tests of the value of explainability. Explain the named animal readout simply.')
    S(42,'Inspect Chemical Validity Before Trusting a Pose',['Check stereochemistry, bond geometry and impossible atomic overlaps.','High confidence does not replace those chemical checks.'],'chemical_checks','Illustrative invalid-versus-valid input checks; not results from a named model.','Buttenschoen et al., Chem. Sci. 15, 3130–3139 (2024).','The illustration is explicitly hypothetical. Do not invent measured benchmark results or attribute a synthetic geometry error to a specific software package.')
    S(43,'Measure Whether AI Improved the Decision',['Count useful hits among compounds actually tested.','Compare with a sensible alternative under comparable conditions.'],'hit_comparison','Hypothetical example: twenty experimentally tested compounds per group.','Teaching comparison; no claimed performance for a named model.','The 7/20 and 2/20 counts are invented. They illustrate the denominator and do not establish statistical significance.')
    S(44,'Exercise: Choose the Next EGFR Experiment',['You have an EGFR structure and a shortlist of candidate inhibitors.','Which tool addresses the pose—and which assay addresses inhibition?'],'egfr_lapatinib','Experimental lapatinib–EGFR complex, PDB 1XKK, as the exercise reference.','Wood et al., Cancer Res. 64, 6652–6659 (2004); PDB 1XKK.','A suitable docking/scoring workflow can propose poses; a kinase assay tests inhibition. Discuss receptor conformation and target selectivity. Lapatinib is not presented as AI-discovered.','side')
    S(45,'Exercise: What Did an RNA Binding Screen Prove?',['SMRTnet nominated interactions across ten disease-associated RNA targets.','What additional evidence is needed to connect binding to a cellular effect?'],'rna_validation','Published counts: 190 predicted interactions tested; 40 confirmed by MST.',SM,'Count interactions, not necessarily unique compounds. Binding validation does not alone establish cellular specificity, causality or therapeutic benefit. Ask for a functional experiment and a specificity control.')
    S(46,'From Target Knowledge to Better PD Decisions',['Define the target interaction or biological effect.','Choose a suitable AI tool, inspect the chemistry and test the claim.'],None,'Chemical knowledge makes AI predictions useful—not automatically true.','Summary of the cited research examples.','Close the link to enzymes, GPCRs, nuclear receptors, ion channels and nucleic acids. Students should name the question, model output and matching experiment.',layout='summary')
    assert len(SPECS)==40 and [s['n'] for s in SPECS]==list(range(7,47))

def additional_panels():
    # Separate, non-overlapping panels; full original assets are never repeated on slides.
    crop('af3_intro','af3_fig1',(0,0,1,.30),'AF3 Figure 1a–b, structural examples only; architecture and benchmark panels omitted')
    crop('tnik_evidence','tnik_fig1',(0,0,.48,.68),'Target-prioritization portion of Ren Figure 1')
    crop('tnik_design','tnik_fig1',(.50,0,1,.68),'Molecule-design portion of Ren Figure 1')
    crop('tnik_binding','tnik_fig2',(0,.27,.46,.52),'Binding-assay portion of Ren Figure 2; verify source panel before final use')
    crop('tnik_cell','tnik_fig2',(.48,.27,1,.54),'Cellular-assay portion of Ren Figure 2; verify source panel before final use')
    crop('dgat2_pocket','pockets_fig3',(.49,.50,1,1),'DGAT2-region crop from AlphaFold human-proteome Figure 3; verify selected panel')
    # Retain entire figures until their precise panel captions and boundaries are inspected.
    for new,old in [('antibiotics_workflow','antibiotics_fig1'),('antibiotics_evidence','antibiotics_fig2')]:
        PROV[new]=dict(PROV[old])
    im=canvas();d=ImageDraw.Draw(im)
    items=[('Defined stereochemistry','Use an explicit stereochemical form.'),('Plausible bond geometry','Inspect distorted rings and bonds.'),('No impossible overlaps','Check clashes with the target.')]
    for i,(a,b) in enumerate(items):
        y=150+i*350;d.rounded_rectangle((100,y,520,y+210),radius=35,outline=BLUE,width=8)
        if i==0:
            d.line((190,y+145,315,y+45,420,y+155),fill=INK,width=8);d.polygon([(315,y+45),(330,y+160),(355,y+145)],fill=INK)
        elif i==1:
            pts=[(310+120*math.cos(k*math.pi/3),y+105+90*math.sin(k*math.pi/3)) for k in range(6)];d.line(pts+[pts[0]],fill=INK,width=8)
        else:
            d.ellipse((160,y+50,310,y+200),outline=AMBER,width=8);d.ellipse((325,y+15,475,y+165),outline=BLUE,width=8)
        label(d,(650,y+15),a,64,True);label(d,(650,y+125),b,54)
    save_art('chemical_checks',im,'Buttenschoen et al., Chem. Sci. 15, 3130–3139 (2024). doi:10.1039/D3SC04185A','Original schematic checklist; not model output or experimental data.')

# Native PowerPoint text, with actual bullet properties.
def text(slide,txt,x,y,w,h,size=18,bold=False,color='20252B',name='text',align=PP_ALIGN.LEFT):
    b=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));b.name=name
    tf=b.text_frame;tf.clear();tf.word_wrap=True;tf.auto_size=MSO_AUTO_SIZE.NONE
    tf.margin_left=tf.margin_right=Inches(.015);tf.margin_top=tf.margin_bottom=0
    for i,line in enumerate(txt.split('\n')):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.alignment=align;p.space_after=Pt(3);p.line_spacing=1.12
        r=p.add_run();r.text=line;r.font.name='Arial';r.font.size=Pt(size);r.font.bold=bold;r.font.color.rgb=RGBColor.from_string(color)
    return b

def bullets(slide,lines,x,y,w,h):
    box=text(slide,'',x,y,w,h,name='body');tf=box.text_frame
    for i,line in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=line;p.line_spacing=1.18;p.space_after=Pt(8)
        pp=p._p.get_or_add_pPr();pp.set('marL',str(Inches(.19)));pp.set('indent',str(-Inches(.17)))
        for tag in ('a:buNone','a:buChar','a:buAutoNum'):
            for e in list(pp.findall(tag,pp.nsmap)):pp.remove(e)
        el=OxmlElement('a:buChar');el.set('char','•');pp.append(el)
        for r in p.runs:r.font.name='Arial';r.font.size=Pt(18);r.font.color.rgb=RGBColor.from_string('20252B')
    return box

def picture(slide,key,box):
    rec=PROV[key];im=Image.open(rec['path']);w,h=im.size;x,y,bw,bh=box
    scale=min(bw/w,bh/h);pw=w*scale;ph=h*scale
    p=slide.shapes.add_picture(rec['path'],Inches(x+(bw-pw)/2),Inches(y+(bh-ph)/2),width=Inches(pw),height=Inches(ph));p.name='FIGURE:'+key
    USE.append({'slide':CURRENT,'key':key,'sha256':rec['sha256'],'source':rec['source'],'detail':rec['detail'],'pixels':rec['pixels'],'x':p.left/914400,'y':p.top/914400,'w':pw,'h':ph,'ppi':w/pw})

def rounded(slide,txt,x,y,w,h,color='E5EFF2',size=18):
    b=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h));b.fill.solid();b.fill.fore_color.rgb=RGBColor.from_string(color);b.line.fill.background()
    text(slide,txt,x+.08,y+.1,w-.16,h-.16,size=size,align=PP_ALIGN.CENTER)

def native(slide,s):
    lay=s['layout']
    if lay=='tools':
        rows=[['Research question','Example tool','Typical result'],['What could the complex look like?','AlphaFold 3','Predicted 3D arrangement'],['Where might a ligand bind?','P2Rank / PrankWeb','Candidate binding pockets'],['How might a ligand fit?','GNINA','Docked poses and scores'],['Which pairs look promising?','Boltz-2','Structure and binding outputs'],['Which analogue may be active?','Chemprop','Assay-trained predictions'],['What could we make next?','REINVENT 4','Candidate chemical structures']]
        table=slide.shapes.add_table(7,3,Inches(.4),Inches(1.68),Inches(9.2),Inches(4.65)).table
        for col,width in zip(table.columns,[3.68,2.40,3.12]):col.width=Inches(width)
        for i,row in enumerate(rows):
            for j,value in enumerate(row):
                cell=table.cell(i,j);cell.text=value;cell.margin_left=Inches(.12);cell.margin_right=Inches(.08);cell.margin_top=Inches(.10);cell.margin_bottom=Inches(.04);cell.vertical_anchor=MSO_ANCHOR.MIDDLE
                cell.fill.solid();cell.fill.fore_color.rgb=RGBColor.from_string('E0EBEF' if i==0 else ('F3F6F7' if i%2==0 else 'FFFFFF'))
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:r.font.name='Arial';r.font.size=Pt(16);r.font.bold=i==0;r.font.color.rgb=RGBColor.from_string('20252B')
        s['notes']+='\nSources: https://www.nature.com/articles/s41586-024-07487-w ; https://prankweb.cz/ ; https://github.com/gnina/gnina ; https://github.com/jwohlwend/boltz/blob/main/docs/prediction.md ; https://chemprop.readthedocs.io/ ; https://doi.org/10.1186/s13321-024-00812-5 . Tools differ in training needs, licensing and setup.'
    elif lay=='workflow':
        steps=[('Target','What biological question?'),('Structure / pocket','Where could it bind?'),('Candidate molecules','What should be tested?'),('Experiments','Binding and function'),('Next choice','Use the new evidence')]
        for i,(a,b) in enumerate(steps):
            y=1.8+i*.84;rounded(slide,a,.65,y,2.5,.64);text(slide,b,3.65,y+.13,5.5,.5,18)
            if i<4:text(slide,'↓',1.78,y+.63,.5,.3,18,align=PP_ALIGN.CENTER)
    elif lay=='outputs':
        for y,a,b in [(1.9,'Structure','Is this molecular arrangement plausible?'),(3.35,'Binding-related prediction','What precisely was the model trained to predict?'),(4.8,'Experimental result','Which binding or functional assay checks it?')]:
            rounded(slide,a,.65,y,3.0,1.10);text(slide,b,4.0,y+.22,5.25,.85,20)
    elif lay=='cycle':
        for (x,y),t in zip([(1.1,2.2),(6.0,2.2),(6.0,4.6),(1.1,4.6)],['Design','Make','Test','Learn']):rounded(slide,t,x,y,2.7,.9,size=22)
        for x,y,t in [(4.3,2.36,'→'),(7.12,3.6,'↓'),(4.3,4.8,'←'),(2.2,3.6,'↑')]:text(slide,t,x,y,.65,.5,26,align=PP_ALIGN.CENTER)
        text(slide,'Measurements change the next chemical decision.',2.4,3.72,5.2,.7,18,align=PP_ALIGN.CENTER)
    else:
        for y,a,b in [(1.9,'A specific PD question','Target • pocket • binding • function'),(3.4,'An appropriate AI output','Structure • ranked candidates • chemical proposals'),(4.9,'A matching experiment','Binding test • functional assay • causal follow-up')]:
            rounded(slide,a,.7,y,3.3,1.1);text(slide,b,4.35,y+.19,4.9,.87,19)

def build():
    global CURRENT
    p=Presentation();p.slide_width=Inches(10);p.slide_height=Inches(7.5)
    p.core_properties.title='AI in Pharmacodynamics — revised teaching slides 7–46'
    p.core_properties.subject='Undergraduate chemistry; application-led; unique scientific figures'
    p.core_properties.author=''
    for s in SPECS:
        CURRENT=s['n'];sl=p.slides.add_slide(p.slide_layouts[6]);sl.background.fill.solid();sl.background.fill.fore_color.rgb=RGBColor(255,255,255)
        text(sl,s['title'],.38,.24,9.24,.56,24,True,'C00000','title')
        if s['layout']=='side':
            bullets(sl,s['bullets'],.42,1.24,3.25,2.65)
            picture(sl,s['key'],(3.75,1.16,5.85,5.22))
        else:
            bullets(sl,s['bullets'],.42,1.0,9.12,.92)
            if s['key']:picture(sl,s['key'],(.43,2.00,9.14,4.46))
            else:native(sl,s)
        text(sl,s['caption'],.43,6.62,9.04,.49,16,name='caption')
        text(sl,s['source'],.43,7.14,8.72,.22,11,name='reference')
        text(sl,str(s['n']),9.25,7.12,.35,.23,12,name='page_number',align=PP_ALIGN.RIGHT)
        prov=PROV[s['key']] if s['key'] else {}
        notes=s['notes']+'\n\nSources: '+s['source']+'\n'+prov.get('source','')+'\n\nFigure provenance: '+prov.get('detail','Native editable teaching diagram/table.')
        sl.notes_slide.notes_text_frame.text=notes
    path=OUT/'AI_in_PD_Revised_Teaching_Slides.pptx';p.save(path)
    (OUT/'figure_manifest.json').write_text(json.dumps(USE,indent=2))
    (OUT/'content_manifest.json').write_text(json.dumps(SPECS,indent=2))
    audit(p,path)
    return path

def audit(p,path):
    report={'slides':len(p.slides),'numbering':'7–46','preserved_intro_in_this_file':False,'aspect_ratio':'4:3','image_slides':len(set(a['slide'] for a in USE)),'duplicate_images':[],'warnings':[],'out_of_bounds':[],'body_fonts':[],'title_fonts':[],'forbidden_source_images':[]}
    seen=defaultdict(list)
    for a in USE:
        seen[a['sha256']].append(a['slide'])
        if a['ppi']<150:report['warnings'].append({'slide':a['slide'],'issue':'effective resolution below 150 ppi','ppi':round(a['ppi'],1)})
        if 'Molecular Drug Targets' in a['source']:report['forbidden_source_images'].append(a['slide'])
    report['duplicate_images']=[v for v in seen.values() if len(v)>1]
    for n,sl in enumerate(p.slides,7):
        for sh in sl.shapes:
            if sh.left<0 or sh.top<0 or sh.left+sh.width>p.slide_width+100 or sh.top+sh.height>p.slide_height+100:report['out_of_bounds'].append((n,sh.name))
            if sh.has_text_frame and sh.name in ('body','title'):
                for par in sh.text_frame.paragraphs:
                    for r in par.runs:
                        report[sh.name+'_fonts'].append({'slide':n,'font':r.font.name,'size':r.font.size.pt if r.font.size else None})
    with zipfile.ZipFile(path) as z:assert z.testzip() is None
    (OUT/'build_audit.json').write_text(json.dumps(report,indent=2))
    assert not report['duplicate_images'],report['duplicate_images']
    assert not report['out_of_bounds'],report['out_of_bounds']
    assert report['image_slides']>=34
    assert all(x['font']=='Arial' and x['size']>=16 for x in report['body_fonts'])
    assert all(x['font']=='Arial' and x['size']==24 for x in report['title_fonts'])
    print('ACTUAL_BUILD_SAVED',path,'IMAGES',report['image_slides'],'DUPLICATES',report['duplicate_images'])

if __name__=='__main__':
    public_assets();chemical_figures();structural_figures();additional_panels();curriculum();build()
