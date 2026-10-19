#!/usr/bin/env python3
"""Local AI-friendly SOC roster generator. Uses only user-approved names/roles."""
import argparse
import json
from collections import Counter
from datetime import date
from pathlib import Path
from shiftguard import SHIFTS, month_days, coverage, build_workbook

ROLES = ("SOC L1", "SOC L2", "Team Lead", "Other")

def create(config):
    people = config["analysts"]
    if not isinstance(people,list) or not people:
        raise ValueError("analysts must be a nonempty list")
    names=set()
    for p in people:
        if not isinstance(p.get("name"),str) or not p["name"].strip():
            raise ValueError("Analyst name required")
        name=p["name"].strip().casefold()
        if name in names: raise ValueError("Duplicate analyst name")
        names.add(name)
        if p.get("role") not in ROLES: raise ValueError("Unsupported role")
        if p.get("priority") not in ("general","soc"): raise ValueError("Invalid priority")
    if not any(p["priority"]=="soc" for p in people):
        raise ValueError("At least one SOC rotation analyst required")
    days=month_days(int(config["year"]),int(config["month"]))
    req=config.get("minimum_per_shift",{"SOC L1":1,"SOC L2":1})
    if not isinstance(req,dict) or any(k not in ("SOC L1","SOC L2") or type(v)!=int or v<0 for k,v in req.items()):
        raise ValueError("Invalid staffing minimums")
    roster={p["name"]:{} for p in people}
    counts={p["name"]:Counter() for p in people}
    previous={p["name"]:None for p in people}
    night_runs={p["name"]:0 for p in people}
    work_runs={p["name"]:0 for p in people}
    warnings=[]
    for day in days:
        key=day.isoformat()
        for p in people:
            if p["priority"]=="general":
                roster[p["name"]][key]="General" if day.weekday()<5 else "Week Off"
        assigned=set()
        for shift in ("Night","Afternoon","Morning"):
            for role,minimum in req.items():
                for _ in range(minimum):
                    candidates=[p for p in people if p["priority"]=="soc" and p["role"]==role and p["name"] not in assigned]
                    if not candidates: break
                    def score(p):
                        n=p["name"]
                        return (100*(previous[n]=="Night" and shift!="Night")
                                +70*(shift=="Night" and night_runs[n]>=2)
                                +60*(work_runs[n]>=5)
                                +counts[n][shift]*8
                                +sum(counts[n][s] for s in SHIFTS),n.casefold())
                    chosen=min(candidates,key=score)
                    roster[chosen["name"]][key]=shift
                    assigned.add(chosen["name"])
        for p in people:
            if p["priority"]!="soc": continue
            n=p["name"]
            shift=roster[n].setdefault(key,"Week Off")
            counts[n][shift]+=1
            if shift in SHIFTS:
                work_runs[n]+=1
                night_runs[n]=night_runs[n]+1 if shift=="Night" else 0
                if previous[n]=="Night" and shift in ("Morning","Afternoon"):
                    warnings.append({"date":key,"name":n,"warning":"Night-to-day transition: rest hours not verified"})
                if night_runs[n]>2:
                    warnings.append({"date":key,"name":n,"warning":"More than two consecutive nights"})
                if work_runs[n]>5:
                    warnings.append({"date":key,"name":n,"warning":"More than five consecutive working days"})
            else:
                work_runs[n]=night_runs[n]=0
            previous[n]=shift
    gaps=coverage(people,days,roster,req)
    return people,days,roster,gaps,warnings

def role():
    choices={"1":"SOC L1","2":"SOC L2","3":"Team Lead","4":"Other"}
    while True:
        choice=input("Role (1=L1, 2=L2, 3=Lead, 4=Other): ").strip()
        if choice in choices:return choices[choice]

def wizard():
    people=[]
    for priority,title in (("general","General-shift PRIORITY staff"),("soc","SOC ROTATING shift analysts")):
        print("\n"+title+" — enter names and roles. Blank name finishes section.")
        while True:
            name=input("Analyst name: ").strip()
            if not name: break
            people.append({"name":name,"role":role(),"priority":priority})
    now=date.today()
    period=input("Roster month YYYY-MM [current month]: ").strip() or now.strftime("%Y-%m")
    year,month=map(int,period.split("-"))
    print("General shift assumed Monday-Friday; confirm suitability before use.")
    l1=int(input("Minimum L1 on EACH shift [1]: ").strip() or "1")
    l2=int(input("Minimum L2 on EACH shift [1]: ").strip() or "1")
    if input("Confirm these assumptions and generate? [y/N]: ").strip().lower() not in ("y","yes"):
        raise ValueError("Cancelled")
    return {"year":year,"month":month,"analysts":people,
            "minimum_per_shift":{"SOC L1":l1,"SOC L2":l2}}

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--config",type=Path,help="Local JSON; omit for guided intake")
    ap.add_argument("--output",type=Path,default=Path("ShiftGuard_AI_Roster.xlsx"))
    ap.add_argument("--summary",type=Path,default=Path("ShiftGuard_Roster_Review.json"))
    ap.add_argument("--fail-on-gaps",action="store_true")
    args=ap.parse_args()
    config=json.loads(args.config.read_text(encoding="utf-8")) if args.config else wizard()
    people,days,roster,gaps,warnings=create(config)
    team=[dict(p,pattern="weekday" if p["priority"]=="general" else "rotating") for p in people]
    build_workbook(team,days,roster,gaps,args.output)
    report={"year":days[0].year,"month":days[0].month,
            "general_shift":sum(p["priority"]=="general" for p in people),
            "soc_shift":sum(p["priority"]=="soc" for p in people),
            "coverage_gaps":gaps,"rotation_warnings":warnings,
            "limitations":["General priority assumes weekdays only",
            "Shift hours, rest duration, leave, holidays, skills and legal requirements are not validated",
            "Human review and approval required"]}
    args.summary.write_text(json.dumps(report,indent=2),encoding="utf-8")
    print(f"Generated {args.output} and {args.summary}; coverage gaps: {len(gaps)}; warnings: {len(warnings)}")
    if gaps:print("WARNING: Uncovered slots exist. See Coverage Gaps worksheet.")
    return 2 if gaps and args.fail_on_gaps else 0

if __name__=="__main__":
    raise SystemExit(main())
