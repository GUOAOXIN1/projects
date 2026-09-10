"""Run with the system Python that provides LibreOffice UNO."""
import json, time, subprocess, re
from pathlib import Path
import xml.etree.ElementTree as ET
import uno
from com.sun.star.beans import PropertyValue
root=Path('pd-finished').resolve()
src=root/'output/AI_in_PD_Unique_Figures_Teaching_7-46.pptx'
out=root/'review';out.mkdir(exist_ok=True)
def pv(name,value):
    p=PropertyValue();p.Name=name;p.Value=value;return p
proc=subprocess.Popen(['soffice','-env:UserInstallation=file:///tmp/pd-direct-render','--headless','--invisible','--nodefault','--norestore','--nofirststartwizard','--accept=socket,host=127.0.0.1,port=2013;urp;StarOffice.ComponentContext'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
local=uno.getComponentContext();resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
context=None
for k in range(45):
    try:
        context=resolver.resolve('uno:socket,host=127.0.0.1,port=2013;urp;StarOffice.ComponentContext');break
    except Exception:time.sleep(1)
if context is None:raise RuntimeError('LibreOffice socket unavailable')
sm=context.ServiceManager;desktop=sm.createInstanceWithContext('com.sun.star.frame.Desktop',context)
doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(src)),'_blank',0,(pv('Hidden',True),pv('ReadOnly',True),pv('UpdateDocMode',0)))
if doc is None:raise RuntimeError('PPTX failed to import into LibreOffice')
pages=doc.getDrawPages();assert pages.getCount()==40
specs=json.loads((root/'output/content_manifest.json').read_text())
report=json.loads((root/'output/quality_report.json').read_text())
missing=[];rendered=[]
norm=lambda t:re.sub(r'[^a-z0-9]','',t.lower())
for i in range(pages.getCount()):
    page=pages.getByIndex(i);n=i+7
    exp=sm.createInstanceWithContext('com.sun.star.drawing.GraphicExportFilter',context)
    exp.setSourceDocument(page)
    png=out/f'slide_{n:02}.png'
    filters=(pv('PixelWidth',1800),pv('PixelHeight',1350),pv('Translucent',False))
    args=(pv('URL',uno.systemPathToFileUrl(str(png))),pv('MediaType','image/png'),pv('FilterData',filters))
    if not exp.filter(args) or not png.is_file():raise RuntimeError('PNG export failed '+str(n))
    svg=out/f'slide_{n:02}.svg'
    args=(pv('URL',uno.systemPathToFileUrl(str(svg))),pv('MediaType','image/svg+xml'),pv('FilterData',(pv('ExportTextAsPath',False),)))
    if not exp.filter(args) or not svg.is_file():raise RuntimeError('SVG export failed '+str(n))
    rendered_text=' '.join(ET.parse(svg).getroot().itertext())
    for role,values in [('title',[specs[i]['title']]),('body',specs[i]['bullets']),('caption',[specs[i]['caption']])]:
        for text in values:
            if norm(text) not in norm(rendered_text):missing.append({'slide':n,'part':role,'text':text})
    rendered.append(n)
doc.close(True);desktop.terminate();proc.wait(timeout=30)
report['rendered_slides']=rendered
report['direct_png_svg_render']=True
report['missing_rendered_title_body_caption']=missing
report['visual_review_status']='All slides rendered to PNG and SVG and their title/body/caption text checked. Full-resolution human/model visual inspection is not claimed.'
report['pdf_generated']=False
(root/'output/quality_report.json').write_text(json.dumps(report,indent=2))
assert not missing,missing
assert not list(root.rglob('*.pdf'))
print('DIRECT_PNG_SVG_RENDER_COMPLETE',len(rendered),'NO_PDF')
