"""Flagship case 006: blue biotech editorial, figure-led but information-dense."""
from pathlib import Path
from PIL import Image
from pptx_designer import Presentation
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.layout import page_number
from pptx_designer.tools.shapes import arrow, oval, rect, rrect
from pptx_designer.tools.text import multiline, text

ROOT = Path(__file__).resolve().parent
ASSET = ROOT / "assets"
OUT = ROOT / "output" / "car_t_single_cell_atlas_blue_editorial.pptx"
CROP_DIR = ASSET / "crops"
C = {"bg":"#F3F7FA","panel":"#FFFFFF","panel2":"#E5F0F5","ink":"#12314A","muted":"#627B8C","line":"#C5D8E2","navy":"#123B5D","blue":"#1F6FAE","cyan":"#08A6B5","pink":"#D48728","lime":"#658E42","orange":"#C57A2C","white":"#FFFFFF"}

def tx(s,x,y,w,h,v,size=12,color="ink",bold=False,name="Aptos",align=None):
    kw=dict(font_size=size,color=color,bold=bold,font_name=name,C=C)
    if align is not None: kw["align"]=align
    return text(s,x,y,w,h,v,**kw)
def ml(s,x,y,w,h,lines,size=11.5,color="muted",spacing=1.08):
    return multiline(s,x,y,w,h,lines,font_size=size,color=color,C=C,line_spacing=spacing)
def line(s,x,y,w,h=.014,color="line"): rect(s,x,y,w,h,fill=C[color],C=C)
def box(s,x,y,w,h,fill="panel",accent=None):
    rect(s,x,y,w,h,fill=C[fill],line=C[accent or fill] if accent else C["line"],C=C)
def tag(s,x,y,w,v,color="blue"):
    rect(s,x,y,w,.25,fill=C["panel"],line=C[color],C=C); tx(s,x,y+.055,w,.12,v,8,color,True,"Consolas","center")
def image(s,n,x,y,w,h,label=None):
    image_file(s,ASSET/f"41586_2024_7762_Fig{n}_HTML.png",x,y,w,h,label)

def image_file(s,path,x,y,w,h,label=None):
    box(s,x-.035,y-.035,w+.07,h+.07,"panel","line")
    cover_image(s,x,y,w,h,str(path))
    if label: tx(s,x,y+h+.055,w,.12,label,7.8,"muted",False,"Consolas")

def crop(n,key,rect):
    CROP_DIR.mkdir(exist_ok=True)
    out=CROP_DIR/f"fig{n}_{key}.png"
    src=ASSET/f"41586_2024_7762_Fig{n}_HTML.png"
    if not out.exists() or out.stat().st_mtime < src.stat().st_mtime:
        Image.open(src).crop(rect).save(out)
    return out

def crop_image(s,n,key,rect,x,y,w,h,label=None):
    image_file(s,crop(n,key,rect),x,y,w,h,label)
def base(prs,n,section,title,sub):
    s=prs.slides.add_slide(prs.slide_layouts[6]); rect(s,0,0,13.333,7.5,fill=C["bg"],C=C)
    for x in [.62,3.68,6.74,9.80,12.86]: line(s,x,.42,.012,6.42)
    tag(s,.72,.42,2.2,section,"navy"); tx(s,.74,.82,11.7,.42,title,22,"navy",True); tx(s,.76,1.28,11.7,.18,sub,9.5,"muted")
    page_number(s,n,12,C=C); tx(s,11.55,.5,1.0,.14,"ATLAS / 03",8,"muted",False,"Consolas","right")
    return s
def foot(s,v): tx(s,.74,7.08,12,.12,v,7.2,"muted",False,"Consolas")
def metric(s,x,y,w,num,label,color="blue"):
    box(s,x,y,w,.92,"panel",color); tx(s,x+.16,y+.12,w-.3,.34,num,26,color,True,"Aptos Display"); tx(s,x+.16,y+.57,w-.3,.14,label,8.5,"muted",True,"Consolas")
def note(s,x,y,w,h,head,body,color="blue"):
    box(s,x,y,w,h,"panel",color); tx(s,x+.15,y+.12,w-.3,.20,head,10.5,color,True,"Consolas"); ml(s,x+.15,y+.43,w-.3,h-.52,body if isinstance(body,list) else body.split("\n"),11,"ink",1.12)

def build():
    prs=Presentation()
    # 1 cover — conceptual hero visual, explicitly distinct from paper figures.
    s=prs.slides.add_slide(prs.slide_layouts[6]); cover_image(s,0,0,13.333,7.5,str(ASSET/"cart_hero_blue.png"))
    line(s,.82,.78,4.72,.035,"cyan"); tx(s,.84,1.08,4.8,.2,"PAPER CASE 03 / BIOTECH ATLAS",10,"cyan",True,"Consolas")
    tx(s,.84,1.55,5.25,1.25,"SINGLE-CELL\nCAR T ATLAS",30,"white",True)
    tx(s,.86,3.25,4.85,.58,"How a persistent clinical phenotype\nbecomes a cell-state hypothesis.",15,"cyan",True)
    ml(s,.88,4.25,4.8,1.0,["Bai et al. · Nature 634 · 702–711 · 2024","82 patients · 695,819 analyzed cells","CONCEPTUAL COVER VISUAL"],10.5,"white",1.25)
    tag(s,.86,6.2,2.75,"CLINICAL / CELLULAR / PRECLINICAL","cyan")
    tx(s,.86,6.82,5.3,.12,"DOI 10.1038/s41586-024-07762-w",7.2,"white",False,"Consolas")

    # 2 clinical persistence — data trace + scientific reasoning chain
    s=base(prs,2,"01 / CLINICAL PERSISTENCE","The clinical signal is unusually long-lived","Before asking what the cell is, establish how long it remains clinically visible.")
    crop_image(s,4,"top",(0,0,2124,600),.74,1.78,8.05,3.35,"Source figure / Fig. 4a–b · longitudinal serum measurements")
    tx(s,9.28,1.88,3.2,.18,"PERSISTENCE READOUT",10,"cyan",True,"Consolas"); line(s,9.28,2.22,3.2,.025,"cyan")
    tx(s,9.28,2.55,2.0,.55,"8.4",42,"orange",True,"Aptos Display"); tx(s,10.92,2.78,1.2,.2,"years",15,"orange",True)
    tx(s,9.28,3.30,3.0,.18,"median BCA-L duration",11,"navy",True)
    ml(s,9.28,3.72,3.15,.72,["The phenotype is not a single", "measurement—it is a longitudinal trace."],12,"muted",1.12)
    line(s,.76,5.75,11.72,.025,"blue")
    tx(s,.78,6.02,1.4,.16,"OBSERVATION",8,"cyan",True,"Consolas"); tx(s,2.35,6.00,2.4,.2,"long B-cell aplasia",12,"navy",True)
    arrow(s,5.0,6.04,.58,0,C["cyan"],C=C); tx(s,5.82,6.02,1.0,.16,"QUESTION",8,"cyan",True,"Consolas"); tx(s,7.02,6.00,2.5,.2,"what sustains the state?",12,"navy",True)
    arrow(s,9.75,6.04,.58,0,C["cyan"],C=C); tx(s,10.48,6.02,1.2,.16,"NEXT",8,"orange",True,"Consolas"); tx(s,10.48,6.28,1.8,.18,"single-cell atlas",11,"orange",True)
    foot(s,"Evidence anchor: BCA-L group median BCA duration 8.4 years. Clinical persistence is the entry point, not the conclusion.")

    # 3 study system — five-stage pipeline, no text containers
    s=base(prs,3,"02 / STUDY DESIGN","One cohort becomes a multi-omic research system","The paper connects patient persistence to a measured cell-state atlas through five linked readout layers.")
    x0=1.0; y=2.45; gap=2.15; line(s,x0+.25,y+.44,10.4,.035,"cyan")
    nodes=[("01","PATIENT\nCOHORT","82 ALL patients\n+ 6 healthy donors","orange"),("02","CAR T\nPRODUCT","manufactured\nCAR T cells","blue"),("03","SINGLE-CELL\nMULTIOMICS","10x 3′ assay\n>10⁶ cells","cyan"),("04","FUNCTIONAL\nREADOUTS","flow · secretomics\nATAC","blue"),("05","CLINICAL\nCORRELATION","17 clusters\nBCA strata","orange")]
    for i,(num,head,body,col) in enumerate(nodes):
        x=x0+i*gap; oval(s,x,y,.52,.52,fill=C[col],C=C); tx(s,x,y+.16,.52,.12,num,8,"white",True,"Consolas","center")
        tx(s,x-.28,y+.78,1.45,.38,head,10,col,True,"Consolas","center"); multiline(s,x-.47,y+1.28,1.85,.48,body.split("\n"),font_size=10,color="navy",C=C,line_spacing=1.08,align="center")
        if i<4: arrow(s,x+.58,y+.25,1.34,0,C["cyan"],C=C)
    rect(s,.76,4.76,7.0,1.28,fill=C["panel"],line=C["line"],C=C); cover_image(s,.86,4.88,6.8,1.04,str(crop(1,"top",(0,0,2122,360)))); tx(s,.88,6.08,6.7,.12,"Source figure / Fig. 1a · cohort → measurement layers → data analysis",7.5,"muted",False,"Consolas")
    tx(s,8.42,4.82,3.8,.18,"THE SYSTEM OUTPUT",9,"cyan",True,"Consolas"); line(s,8.42,5.14,3.85,.025,"cyan"); tx(s,8.42,5.42,2.35,.28,"695,819",26,"blue",True,"Aptos Display"); tx(s,10.92,5.54,1.45,.18,"analyzed cells",11,"navy",True); tx(s,8.42,5.92,3.6,.18,"17 cell states  ·  5 BCA strata",11,"navy",True); tx(s,8.42,6.34,3.8,.18,"Clinical observation becomes measurable state space.",10,"muted")
    foot(s,"Source figure / Fig. 1a. The pipeline is editorially redrawn; cohort and assay facts remain tied to the paper.")

    # 4 atlas
    s=base(prs,4,"03 / CELL ATLAS","695,819 cells resolve into 17 states","The UMAP is the map from which the later mechanistic claims depart.")
    image(s,1,.76,1.72,5.8,3.8,"Fig. 1b / UMAP + cluster annotation")
    metric(s,6.86,1.72,2.45,"695,819","analyzed cells","blue"); metric(s,9.55,1.72,2.45,"17","cell states","lime")
    tx(s,6.86,3.12,5.15,.18,"ATLAS READOUT / FOUR STATE FAMILIES",10,"cyan",True,"Consolas"); line(s,6.86,3.48,5.15,.025,"cyan")
    tx(s,6.86,3.92,2.0,.22,"NAIVE",13,"navy",True); tx(s,9.35,3.92,2.0,.22,"CYTOTOXIC",13,"navy",True)
    tx(s,6.86,4.52,2.0,.22,"REGULATORY",13,"navy",True); tx(s,9.35,4.52,2.0,.22,"PROLIFERATIVE",13,"navy",True)
    box(s,.76,5.82,11.25,.62,"panel2","line"); tx(s,.98,6.02,10.8,.18,"Design principle: preserve the paper’s cluster colors; use blue annotations to explain the map.",10,"navy",True)
    foot(s,"Source figure / Fig. 1b. Cell labels and original color encoding are preserved.")

    # 5 persistence
    s=base(prs,5,"04 / PERSISTENCE","Persistence is a gradient, not a binary","BCA strata create the clinical axis reused across the paper.")
    image(s,1,.76,1.72,5.6,2.15,"Fig. 1c / demographics and follow-up")
    box(s,.76,4.2,5.6,1.6,"panel","blue"); tx(s,.96,4.43,5.1,.18,"PERSISTENCE GRADIENT",10,"blue",True,"Consolas"); line(s,1.0,5.08,4.95,.055,"cyan")
    for i,(g,d,col) in enumerate([("BCA-L","101","pink"),("BCA-O","61","blue"),("BCA3","18","cyan"),("BCA2","4","orange"),("BCA1","1","lime")]):
        x=1.0+i*.98; tx(s,x,5.23,.8,.14,g,7.5,col,True,"Consolas","center"); tx(s,x,4.73,.8,.2,d,14,col,True,"Aptos Display","center")
    tx(s,6.78,1.95,2.55,.18,"BCA-L / LONGEST",10,"orange",True,"Consolas"); line(s,6.78,2.3,2.55,.025,"orange"); ml(s,6.78,2.58,2.55,.65,["n = 5 · 101 months", "no relapse observed"],12,"navy",1.12)
    tx(s,9.58,1.95,2.55,.18,"BCA1 / SHORTEST",10,"orange",True,"Consolas"); line(s,9.58,2.3,2.55,.025,"orange"); ml(s,9.58,2.58,2.55,.65,["n = 17 · 1 month", "relapse observed"],12,"navy",1.12)
    tx(s,6.78,4.48,5.35,.18,"INTERPRETATION",10,"cyan",True,"Consolas"); line(s,6.78,4.82,5.35,.025,"cyan"); ml(s,6.78,5.1,5.35,.55,["Persistence is a continuum with a clinical gradient, not a yes/no label."],11.5,"navy",1.12)
    foot(s,"Source figure / Fig. 1c. Duration values shown are the paper’s reported values; units are months in the source table.")

    # 6 type2 four evidence modules
    s=base(prs,6,"05 / TYPE 2 STATE","The long-remission state reads as type 2","Four orthogonal readouts point to the same biological state.")
    crop_image(s,2,"top",(0,0,1600,780),.76,1.72,3.0,2.0,"Fig. 2a–d / type 2 score")
    crop_image(s,2,"mid",(0,650,1600,1450),4.02,1.72,3.0,2.0,"Fig. 2e–f / secretion")
    crop_image(s,2,"bottom",(0,1350,1600,1959),.76,4.1,3.0,1.75,"Fig. 2g–j / ATAC + GATA3")
    line(s,4.02,4.1,.06,1.75,"orange")
    tx(s,4.28,4.18,2.4,.18,"PERTURBATION",9.5,"orange",True,"Consolas")
    tx(s,4.28,4.62,2.4,.22,"GATA3 / STAT6",12,"navy",True)
    tx(s,4.28,5.02,2.4,.18,"repeat stimulation",10.5,"navy")
    tx(s,4.28,5.34,2.4,.18,"→ tumour-cell lysis",10.5,"blue",True)
    # Replace the large right-hand card with a paper-like signal annotation:
    # a vertical accent, typographic hierarchy, and a convergence rule.
    line(s,7.52,1.72,.06,2.0,"orange")
    tx(s,7.78,1.76,3.9,.18,"TYPE 2 SIGNATURE",10,"orange",True,"Consolas")
    tx(s,7.78,2.20,4.0,.28,"IL4 · IL5 · IL13",18,"navy",True,"Aptos Display")
    tx(s,7.78,2.58,4.0,.24,"GATA3",16,"orange",True,"Aptos Display")
    tx(s,7.78,3.04,4.0,.18,"enriched in the long-remission group",10.5,"muted")
    line(s,7.52,4.10,4.5,.025,"cyan")
    tx(s,7.52,4.32,4.2,.18,"EVIDENCE RULE / CONVERGENCE",9,"cyan",True,"Consolas")
    tx(s,7.52,4.76,1.25,.20,"TRANSCRIPTOME",9.5,"navy",True,"Consolas")
    arrow(s,8.88,4.84,.38,0,C["cyan"],C=C)
    tx(s,9.38,4.76,1.05,.20,"SECRETION",9.5,"navy",True,"Consolas")
    arrow(s,10.58,4.84,.38,0,C["cyan"],C=C)
    tx(s,11.08,4.76,1.0,.20,"CHROMATIN",9.5,"navy",True,"Consolas")
    line(s,7.52,5.30,4.5,.025,"cyan")
    tx(s,7.52,5.50,4.5,.20,"→ one coherent type 2 state",12,"blue",True)
    foot(s,"Source figure / Fig. 2. Direct paper panels used as evidence modules; interpretation is newly organized.")

    # 7 mechanism
    s=base(prs,7,"06 / MECHANISM","Cluster 2 is the pressure point","The paper moves from association to a candidate regulatory interaction network.")
    crop_image(s,3,"top",(0,0,1420,760),.76,1.72,3.45,2.1,"Fig. 3a–b / ligand–receptor network")
    crop_image(s,3,"right",(1180,0,2125,760),4.48,1.72,3.45,2.1,"Fig. 3c–d / cluster 2 and DEGs")
    metric(s,8.18,1.72,2.0,"13.9%","cluster 2 proportion","pink")
    note(s,10.42,1.72,1.6,2.1,"STATE",["type 2", "signal", "cluster 2", "proliferation"],"cyan")
    box(s,.76,4.25,7.17,1.55,"panel","cyan"); tx(s,.98,4.48,6.6,.2,"EDITABLE MECHANISM SUMMARY",9,"cyan",True,"Consolas"); tx(s,1.0,5.02,6.6,.28,"Type 2 function  →  cluster 2  →  dysfunctional subpopulation",13,"navy",True)
    note(s,8.18,4.25,3.84,1.55,"CLAIM BOUNDARY",["Candidate mechanism", "not a complete causal proof"],"orange")
    foot(s,"Source figure / Fig. 3. Network diagrams and cluster 2 readouts are reproduced; the causal wording is deliberately bounded.")

    # 8 secretome
    s=base(prs,8,"07 / SECRETOME","The serum trace keeps the biology longitudinal","A state is more credible when it survives time, not just a snapshot.")
    crop_image(s,4,"top",(0,0,2124,600),.76,1.72,5.7,2.0,"Fig. 4a–b / collection timeline + time series")
    metric(s,6.78,1.72,1.65,"345","measurements","cyan"); metric(s,8.63,1.72,1.65,"30","cytokines","blue"); metric(s,10.48,1.72,1.65,"33","patients","pink")
    crop_image(s,4,"bottom",(0,500,2124,1283),.76,4.02,5.7,1.82,"Fig. 4c–e / cytokine windows + heatmap")
    note(s,6.78,4.02,5.35,1.82,"TYPE 2 WINDOW",["IL-4 · IL-5 · IL-13", "serial evidence across days 1–63", "33 discovery + 8 validation patients"],"pink")
    foot(s,"Source figure / Fig. 4. Longitudinal panels are retained; the surrounding blue modules explain scale and timing.")

    # 9 in vivo
    s=base(prs,9,"08 / IN VIVO TEST","Type 2-high cells pass the recall test","The state is tested as function, not merely described as phenotype.")
    image(s,5,.76,1.72,4.25,2.25,"Fig. 5a–c / mouse sequence + expansion")
    image(s,5,.76,4.28,4.25,1.55,"Fig. 5d / survival after rechallenge")
    tx(s,5.42,1.95,6.6,.18,"IN VIVO SEQUENCE",10,"cyan",True,"Consolas"); line(s,5.42,2.32,6.6,.035,"cyan")
    for x,day,label in [(5.55,"−7","Nalm6"),(7.05,"0","CAR T"),(8.55,"17","rechallenge"),(10.25,"24–28","imaging")]:
        line(s,x,2.2,.02,.28,"cyan"); tx(s,x-.18,2.72,.7,.18,day,10,"navy",True,"Consolas","center"); tx(s,x-.58,3.02,1.35,.18,label,8.2,"navy",True,"Consolas","center")
    tx(s,5.42,3.62,6.6,.18,"READOUT",10,"lime",True,"Consolas"); line(s,5.42,3.98,6.6,.025,"lime"); ml(s,5.42,4.28,6.6,.55,["Type 2-high cells expand more and show stronger recall in the reported model; selected readouts n=5 mice."],11.5,"navy",1.12)
    tx(s,5.42,5.32,6.6,.18,"EVIDENCE BOUNDARY",10,"orange",True,"Consolas"); ml(s,5.42,5.62,6.6,.35,["Preclinical model supports function; it does not establish clinical efficacy."],10.5,"navy",1.12)
    foot(s,"Source figure / Fig. 5. Experimental evidence is shown with the preclinical boundary kept explicit.")

    # 10 engineering
    s=base(prs,10,"09 / ENGINEERING","IL-4 turns the state into a design lever","The final figure converts a cell-state hypothesis into a manufacturing strategy.")
    crop_image(s,6,"left",(0,0,1100,1000),.76,1.72,4.0,2.28,"Fig. 6a / priming and ET2-L/H strategies")
    crop_image(s,6,"right",(950,0,2132,950),5.0,1.72,3.0,2.28,"Fig. 6b / in vivo imaging")
    metric(s,9.3,1.72,1.3,"10","ng ml⁻¹","cyan"); metric(s,10.78,1.72,1.3,"50","ng ml⁻¹","pink")
    crop_image(s,6,"bottom",(0,850,1250,1523),.76,4.34,4.0,1.5,"Fig. 6c–d / expansion + survival")
    note(s,5.0,4.34,3.0,1.5,"DESIGN OPTIONS",["priming: 10 ng ml⁻¹ / 12 h", "ET2-L: 10 · ET2-H: 50", "controlled manufacturing exposure"],"blue")
    note(s,8.28,4.34,3.8,1.5,"INTERPRETATION",["Promising engineering lever", "still bounded by model and study design"],"orange")
    foot(s,"Source figure / Fig. 6. Concentrations and group labels remain tied to the paper; the design framing is editorialized.")

    # 11 evidence chain
    s=base(prs,11,"10 / EVIDENCE CHAIN","The argument is strongest when each layer stays distinct","Do not collapse clinical association, mechanism and intervention into one claim.")
    # Continuous evidence spine: the four layers are stages in one argument,
    # not four empty containers.
    spine=[(1.18,"CLINICAL","8.4-year persistence","Fig. 1 / 4","OBSERVED","orange"),
           (4.20,"CELLULAR","type 2 state","Fig. 2","RESOLVED","cyan"),
           (7.22,"MECHANISM","cluster 2 network","Fig. 3","PROPOSED","blue"),
           (10.24,"PRECLINICAL","recall + IL-4","Fig. 5 / 6","TESTED","lime")]
    line(s,1.48,2.72,9.36,.045,"cyan")
    for i,(x,head,body,source,kind,col) in enumerate(spine):
        oval(s,x,2.42,.60,.60,fill=C[col],C=C); tx(s,x,2.62,.60,.12,str(i+1).zfill(2),8,"white",True,"Consolas","center")
        tx(s,x-.18,3.30,1.0,.18,kind,8,col,True,"Consolas","center")
        tx(s,x-.48,3.72,1.55,.24,head,12,"navy",True,"Aptos Display","center")
        tx(s,x-.48,4.15,1.55,.42,body,11.5,col,True,"Aptos Display","center")
        tx(s,x-.48,4.78,1.55,.16,source,8,"muted",False,"Consolas","center")
        if i<3: arrow(s,x+.68,2.70,1.95,0,C["blue"],C=C)
    line(s,.76,5.54,11.55,.025,"line")
    tx(s,.78,5.78,1.55,.18,"READING RULE",9,"blue",True,"Consolas")
    tx(s,2.52,5.76,9.5,.22,"Clinical persistence anchors the story; cellular evidence resolves the state; mechanism remains a candidate; preclinical work tests the direction.",11.5,"navy",True)
    tx(s,.78,6.28,11.4,.18,"Evidence boundary: association → state definition → candidate explanation → functional test / engineering lever",9.5,"muted",False,"Consolas")
    foot(s,"A rigorous research presentation makes evidence type and evidence boundary visible on the same page.")

    # 12 close
    s=prs.slides.add_slide(prs.slide_layouts[6]); rect(s,0,0,13.333,7.5,fill=C["bg"],C=C)
    tx(s,.84,.86,5,.2,"THE TAKEAWAY",10,"pink",True,"Consolas"); tx(s,.84,1.35,5.2,1.1,"Type 2 function\nmay sustain fitness.",28,"navy",True)
    ml(s,.86,2.98,4.85,1.2,["A durable clinical phenotype", "becomes a recognizable cell state,", "a candidate circuit and an engineering lever."],14,"muted",1.2)
    line(s,.86,4.75,4.5,.03,"cyan"); tx(s,.86,5.08,4.9,.35,"A strong research story shows its figures—and its limits.",12,"blue",True)
    image(s,2,6.6,.78,2.65,2.1,"Fig. 2")
    image(s,3,9.5,.78,2.65,2.1,"Fig. 3")
    image(s,6,6.6,3.38,5.55,2.25,"Fig. 6 / engineering strategy")
    tx(s,.86,6.75,5.4,.12,"Bai et al. · Nature 634 · 702–711 · 2024",7.2,"muted",False,"Consolas")
    prs.save(OUT); print(OUT)

if __name__ == "__main__": build()
