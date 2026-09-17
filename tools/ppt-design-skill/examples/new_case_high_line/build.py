from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.layout import page_number
from pptx_designer.tools.shapes import arrow, oval, rect
from pptx_designer.tools.text import multiline, text

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output" / "high_line_public_realm.pptx"
ASSET = ROOT / "assets"
C = {"paper":"#F3F0EA","ink":"#20272B","muted":"#6F7775","rule":"#C9C2B8","brick":"#B85C42","green":"#597461","blue":"#1F6F8B","sand":"#E5DED2","white":"#FFFFFF"}

def tx(s,x,y,w,h,v,size=12,color="ink",bold=False,name="Aptos",align=None):
    kw=dict(font_size=size,color=color,bold=bold,font_name=name,C=C)
    if align is not None: kw["align"] = align
    return text(s,x,y,w,h,v,**kw)

def ml(s,x,y,w,h,lines,size=11,color="muted",spacing=1.12):
    return multiline(s,x,y,w,h,lines,font_size=size,color=color,C=C,line_spacing=spacing)

def line(s,x,y,w,h=.02,color="rule"): return rect(s,x,y,w,h,fill=C[color],C=C)

def base(prs,n,section,title,sub):
    s=prs.slides.add_slide(prs.slide_layouts[6])
    rect(s,0,0,13.333,7.5,fill=C["paper"],C=C)
    for x in [0.68,3.18,5.68,8.18,10.68,12.68]: line(s,x,.45,.012,6.25,"sand")
    line(s,.72,.44,1.95,.035,"blue")
    tx(s,.72,.58,3,.16,section,8.5,"blue",True,"Consolas")
    tx(s,.74,.92,11.8,.45,title,24,"ink",True)
    tx(s,.76,1.34,11.6,.18,sub,10,"muted")
    page_number(s,n,10,C=C)
    tx(s,11.35,.6,1.2,.14,"CITY / HL01",8,"muted",False,"Consolas","right")
    return s

def foot(s,v): tx(s,.76,7.05,11.5,.14,v,7.2,"muted",False,"Consolas")

def main():
    prs=Presentation(); OUT.parent.mkdir(parents=True,exist_ok=True)
    # 1 cover
    s=prs.slides.add_slide(prs.slide_layouts[6]); rect(s,0,0,13.333,7.5,fill=C["ink"],C=C)
    cover_image(s,6.25,0,7.083,7.5,str(ASSET / "high_line_hero_v1.png"))
    rect(s,6.25,0,0.06,7.5,fill=C["blue"],C=C)
    line(s,7.0,.72,5.35,.035,"blue")
    line(s,.82,.82,2.95,.045,"blue"); tx(s,.84,1.1,5.2,.18,"URBAN REUSE / CASE HL01",9,"blue",True,"Consolas")
    tx(s,.84,1.65,5.3,1.2,"THE HIGH LINE\nPUBLIC REALM",29,"white",True)
    tx(s,.86,3.25,5.0,.55,"From freight infrastructure\nto civic operating model.",14,"white",True)
    ml(s,.88,4.35,4.9,.9,["A real case of preservation, landscape,", "access, programming, and stewardship."],11,"sand",1.2)
    tx(s,.88,6.45,5.5,.14,"OFFICIAL FACTS + EDITORIAL ANALYSIS",7.3,"blue",True,"Consolas")
    tx(s,.88,6.86,5.5,.12,"FRIENDS OF THE HIGH LINE / 2026",7.2,"muted",False,"Consolas")

    # 2 origin
    s=base(prs,2,"01 / ORIGIN","The asset was already there. The operating model was not.","A linear freight structure became a public-realm question.")
    tx(s,.82,1.84,4,.16,"ELEVATED INFRASTRUCTURE / 1934 → 1980",8.5,"blue",True,"Consolas"); line(s,.82,2.1,5.2,.025,"blue")
    line(s,1.05,4.4,5.2,.08,"ink"); rect(s,1.25,3.0,4.75,.18,fill=C["brick"],C=C)
    for x in [1.45,2.3,3.15,4.0,4.85,5.7]: rect(s,x,3.18,.12,1.22,fill=C["green"],C=C)
    tx(s,1.05,4.72,1.5,.18,"1934",18,"brick",True,"Consolas"); tx(s,4.55,4.72,1.5,.18,"1980",18,"blue",True,"Consolas")
    tx(s,1.05,5.08,2.4,.42,"freight rail begins",11,"muted"); tx(s,4.55,5.08,2.4,.42,"last train",11,"muted")
    tx(s,7.55,1.84,4.2,.18,"THE PUBLIC QUESTION",9,"brick",True,"Consolas"); line(s,7.55,2.12,4.2,.03,"brick")
    tx(s,7.55,2.55,4.1,.72,"Demolish,\npreserve, or transform?",22,"ink",True)
    ml(s,7.57,3.65,3.8,1.0,["The structure carried food and agricultural goods", "above Manhattan streets. Its next use had to", "carry public life instead."],11.5,"muted",1.18)
    line(s,7.55,5.35,4.0,.025,"blue"); tx(s,7.55,5.62,4.4,.25,"The first design move was reframing.",12,"blue",True)
    foot(s,"FACTS: Friends of the High Line official fact sheet / 1934 freight structure; last train 1980.")

    # 3 advocacy timeline
    s=base(prs,3,"02 / PRESERVATION","Preservation became a civic design process","The project advanced through advocacy, public support, design, and phased delivery.")
    tx(s,.82,1.84,4,.16,"PUBLIC PROCESS / OFFICIAL MILESTONES",8.5,"blue",True,"Consolas"); line(s,.82,2.1,11.4,.025,"blue")
    line(s,1.15,4.2,10.55,.045,"ink")
    milestones=[("1999","Friends founded","brick"),("2002","City support","green"),("2004","Design team selected","blue"),("2009","Section 1 opens","brick"),("2011","Section 2 opens","green")]
    for i,(year,label,col) in enumerate(milestones):
        x=1.15+i*2.64; oval(s,x-.16,3.88,.36,.36,fill=C[col],line=C[col],C=C)
        tx(s,x-.4,3.25,.8,.16,year,10,col,True,"Consolas","center")
        tx(s,x-.65,4.58,1.45,.34,label,10,"ink",True,"Aptos","center")
    tx(s,1.12,5.8,10.8,.35,"The sequence matters: legitimacy was built before the landscape was built.",17,"brick",True,"Consolas","center")
    foot(s,"FACTS: Friends of the High Line official 2014 corporate membership brochure / milestone chronology.")

    # 4 section
    s=base(prs,4,"03 / DESIGN LOGIC","The track bed became the design vocabulary","The strongest move preserved memory while changing use.")
    tx(s,.82,1.84,4,.16,"EDITABLE ELEVATED SECTION / RAIL + PLANTING",8.5,"blue",True,"Consolas"); line(s,.82,2.1,11.4,.025,"blue")
    rect(s,1.2,3.1,6.15,.2,fill=C["ink"],C=C); rect(s,1.2,4.75,6.15,.2,fill=C["ink"],C=C)
    for x in [1.45,2.2,3.0,3.8,4.7,5.55,6.35]: rect(s,x,3.32,.12,1.42,fill=C["brick"],C=C)
    for x in [1.6,2.45,3.25,4.05,4.95,5.8,6.55]: rect(s,x,2.65,.22,.45,fill=C["green"],C=C)
    tx(s,1.25,5.35,5.8,.22,"rail memory + self-seeded landscape + public movement",10,"blue",True,"Consolas","center")
    tx(s,8.1,2.12,3.9,.18,"DESIGN PRINCIPLE",9,"brick",True,"Consolas"); line(s,8.1,2.42,3.8,.03,"brick")
    tx(s,8.1,2.82,3.8,.9,"Keep the trace.\nChange the role.",24,"ink",True)
    ml(s,8.12,4.05,3.65,1.0,["More than one-third of the original", "rail tracks are featured in the park's design."],11,"muted",1.2)
    foot(s,"FACT: Friends of the High Line self-guided visit / more than one-third of original rail tracks retained in design.")

    # 5 system map
    s=base(prs,5,"04 / OPERATING SYSTEM","A park is not only a path","The public realm works as a linked system of access, landscape, culture, and care.")
    tx(s,.82,1.84,4,.16,"SYSTEMS MAP / NATIVE DIAGRAM",8.5,"blue",True,"Consolas"); line(s,.82,2.1,11.4,.025,"blue")
    oval(s,5.25,3.0,2.7,1.25,fill=C["ink"],line=C["ink"],C=C); tx(s,5.62,3.42,1.95,.3,"PUBLIC\nREALM",16,"white",True,"Consolas","center")
    nodes=[(1.15,2.55,"PRESERVE","structure","brick"),(1.15,4.65,"PLANT","seasonal life","green"),(9.45,2.55,"PROGRAM","art + learning","blue"),(9.45,4.65,"OPERATE","stewardship","brick")]
    for x,y,a,b,col in nodes:
        oval(s,x,y,2.1,.7,fill=C["paper"],line=C[col],C=C); tx(s,x+.2,y+.18,1.7,.16,a,10,col,True,"Consolas","center"); tx(s,x+.15,y+.82,1.8,.16,b,10,"muted",False,"Aptos","center")
    for x1,y1,x2,y2,col in [(3.2,2.9,5.25,3.45,"brick"),(3.2,5.0,5.25,3.85,"green"),(7.95,3.45,9.45,2.9,"blue"),(7.95,3.85,9.45,5.0,"brick")]: arrow(s,x1,y1,x2-x1,.02,fill=C[col],C=C)
    tx(s,4.18,5.75,4.4,.3,"VALUE IS MAINTAINED BETWEEN THE NODES.",13,"blue",True,"Consolas","center")
    foot(s,"EDITORIAL ANALYSIS: the system map is a synthesis; node labels are not a quoted project taxonomy.")

    # 6 scale chart
    s=base(prs,6,"05 / SCALE","The park changed the value of an obsolete corridor","A single linear structure can carry many kinds of public use.")
    tx(s,.82,1.84,4,.16,"OFFICIAL SCALE / 1.45 MILES",8.5,"blue",True,"Consolas"); line(s,.82,2.1,7.0,.025,"blue")
    line(s,1.25,5.45,6.0,.03,"ink")
    bars=[("STRUCTURE",1.45,"brick"),("OPEN 2009",.62,"blue"),("OPEN 2011",1.0,"green"),("VISITOR ROUTE",1.0,"brick")]
    for i,(label,val,col) in enumerate(bars):
        y=2.65+i*.62; tx(s,1.25,y,.95,.16,label,8.5,"muted",True,"Consolas"); rect(s,2.45,y,4.55,.18,fill=C["sand"],C=C); rect(s,2.45,y,4.55*(val/1.45),.18,fill=C[col],C=C); tx(s,7.08,y-.02,.7,.16,str(val),9,col,True,"Consolas")
    tx(s,1.25,5.85,6.1,.25,"Scale is physical. Meaning is operational.",13,"blue",True)
    tx(s,8.5,2.1,3.5,.18,"READING",9,"brick",True,"Consolas"); line(s,8.5,2.4,3.4,.03,"brick"); tx(s,8.5,2.85,3.3,.8,"Reuse creates\nmore than access.",21,"ink",True); ml(s,8.52,4.05,3.0,1.1,["It creates a platform for", "landscape, culture, and", "neighbourhood identity."],11,"muted",1.18)
    foot(s,"FACTS: official High Line fact sheet / 1.45 miles; sections opened in 2009 and 2011. Bars are editorial comparison, not attendance data.")

    # 7 design team / partnership
    s=base(prs,7,"06 / COLLABORATION","The design team was a coalition, not a single author","Landscape, architecture, planting, public agencies, and civic advocates formed one delivery system.")
    tx(s,.82,1.84,4,.16,"COLLABORATION FIELD",8.5,"blue",True,"Consolas"); line(s,.82,2.1,11.4,.025,"blue")
    roles=[("FIELD OPERATIONS","landscape / project lead","brick"),("DS+R","architecture + public realm","blue"),("PIET OUDOLF","planting design","green"),("CITY + FHL","jurisdiction + stewardship","brick")]
    for i,(a,b,col) in enumerate(roles):
        x=1.05+i*2.9; line(s,x,3.05,2.15,.05,col); tx(s,x,3.35,2.1,.28,a,12,"ink",True,"Consolas"); ml(s,x,3.85,2.1,.6,[b,"shared accountability"],10,"muted",1.15)
        if i<3: arrow(s,x+2.25,3.08,.38,.02,fill=C["rule"],C=C)
    line(s,1.05,5.2,10.9,.03,"blue"); tx(s,1.05,5.55,10.8,.3,"A memorable place requires an operating coalition behind the image.",16,"blue",True,"Consolas","center")
    foot(s,"FACT: official High Line materials identify Field Operations, Diller Scofidio + Renfro, and Piet Oudolf as the design collaboration.")

    # 8 stewardship
    s=base(prs,8,"07 / STEWARDSHIP","The project did not end at opening day","Public space becomes durable when maintenance, programming, and funding are designed with it.")
    tx(s,.82,1.84,4,.16,"OPERATING LOOP / EDITORIAL MODEL",8.5,"blue",True,"Consolas"); line(s,.82,2.1,11.4,.025,"blue")
    loop=[("VISIT","use"),("PROGRAM","return"),("CARE","trust"),("FUND","continue")]
    for i,(a,b) in enumerate(loop):
        x=1.25+i*2.75; oval(s,x,3.0,1.55,1.0,fill=C["paper"],line=C[["brick","blue","green","brick"][i]],C=C); tx(s,x+.15,3.28,1.25,.18,a,11,"ink",True,"Consolas","center"); tx(s,x+.2,4.32,1.15,.16,b,10,"muted",False,"Aptos","center");
        if i<3: arrow(s,x+1.65,3.45,.75,.02,fill=C["rule"],C=C)
    line(s,1.25,5.05,8.85,.025,"blue"); arrow(s,1.25,5.05,.34,.02,fill=C["blue"],C=C)
    tx(s,4.05,5.65,5.1,.3,"OPERATE THE PUBLIC REALM AS A LIVING ASSET.",13,"brick",True,"Consolas","center")
    foot(s,"EDITORIAL ANALYSIS: loop is a strategy model; it does not quantify the High Line's operating budget or impact.")

    # 9 decision matrix
    s=base(prs,9,"08 / TRANSFER","What should another city actually copy?","The transferable lesson is the operating logic, not the visual icon.")
    tx(s,.82,1.84,4,.16,"DECISION MATRIX / TRANSFERABILITY",8.5,"blue",True,"Consolas"); line(s,.82,2.1,7.0,.025,"blue")
    line(s,1.35,5.65,5.8,.025,"ink"); line(s,1.35,2.6,.025,3.05,"ink")
    tx(s,3.0,5.9,3.5,.18,"LOCAL CAPACITY →",9,"blue",True,"Consolas","center"); tx(s,.8,3.6,.18,1.2,"PUBLIC\nVALUE",9,"brick",True,"Consolas","center")
    for x,y,label,col in [(2.15,4.8,"ICON", "brick"),(3.85,3.85,"ACCESS","blue"),(5.8,3.15,"SYSTEM","green")]: oval(s,x,y,.55,.55,fill=C[col],line=C[col],C=C); tx(s,x-.25,y+.7,1.1,.16,label,9,col,True,"Consolas","center")
    tx(s,8.35,2.45,3.4,.18,"TRANSFER RULE",9,"brick",True,"Consolas"); line(s,8.35,2.75,3.3,.03,"brick"); tx(s,8.35,3.2,3.3,.75,"Copy the\nconditions, not the image.",21,"ink",True); ml(s,8.37,4.35,3.0,.85,["preserve a trace", "build a coalition", "fund the afterlife"],11,"muted",1.2)
    foot(s,"EDITORIAL ANALYSIS: matrix positions are strategic hypotheses, not a quantified benchmark of public projects.")

    # 10 close
    s=base(prs,10,"09 / DECISION","The High Line is a lesson in changing the role of infrastructure","Preservation became powerful when it was connected to design, stewardship, and public use.")
    tx(s,.9,2.15,4.9,.9,"Keep the trace.\nChange the role.",27,"ink",True); line(s,.9,3.55,4.4,.04,"brick")
    for y,n,v,col in [(4.05,"01","Preserve what gives the place memory.","brick"),(4.62,"02","Design the system around public use.","blue"),(5.19,"03","Fund the operating life after opening.","green")]: tx(s,.92,y,.35,.15,n,8.5,col,True,"Consolas"); tx(s,1.55,y-.01,4.6,.18,v,10.5,"ink",True)
    oval(s,7.2,2.0,4.35,3.8,fill=C["ink"],line=C["ink"],C=C); tx(s,7.7,2.55,3.3,.18,"THE BOUNDARY",9,"blue",True,"Consolas"); ml(s,7.7,3.1,3.2,1.3,["The High Line is a real case.","The transfer model is analysis.","Neither is a universal formula."],16,"white",1.25); line(s,7.7,5.25,3.1,.035,"green")
    foot(s,"SOURCES: Friends of the High Line official fact sheet, self-guided visit, and project materials. Editorial synthesis is labeled.")
    prs.save(str(OUT)); print(OUT)

if __name__ == "__main__": main()
