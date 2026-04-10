#---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
#---------------------------------------------------------------------------------------------
from openstaadpy import os_analytical

try:
    staad = os_analytical.connect()
except Exception as e:
    print(f"Error connecting to STAAD.Pro: {e}")
    exit(1)

geo = staad.Geometry
prop = staad.Property
sup = staad.Support
load = staad.Load

# --- Units initial FEET KIP (Feet=1, Kip=0) ---
staad.SetInputUnits(1,0)
staad.SaveModel(True)
# --- Nodes ---
node_coords = {
    1:(0.0,0.0,0.0), 2:(30.0,0.0,0.0), 3:(0.0,20.0,0.0), 4:(10.0,20.0,0.0), 5:(20.0,20.0,0.0), 6:(30.0,20.0,0.0),
    7:(0.0,35.0,0.0), 8:(30.0,35.0,0.0), 9:(7.5,35.0,0.0), 10:(22.5,35.0,0.0), 11:(15.0,35.0,0.0), 12:(5.0,38.0,0.0),
    13:(25.0,38.0,0.0), 14:(10.0,41.0,0.0), 15:(20.0,41.0,0.0), 16:(15.0,44.0,0.0)
}
for nid,(x,y,z) in node_coords.items():
    geo.CreateNode(nid,x,y,z)

# --- Members ---
member_incidence = {
    1:(1,3),2:(3,7),3:(2,6),4:(6,8),5:(3,4),6:(4,5),7:(5,6),8:(7,12),9:(12,14),10:(14,16),11:(15,16),12:(13,15),13:(8,13),
    14:(9,12),15:(9,14),16:(11,14),17:(11,15),18:(10,15),19:(10,13),20:(7,9),21:(9,11),22:(10,11),23:(8,10)
}
for mid,(n1,n2) in member_incidence.items():
    geo.CreateBeam(mid,n1,n2)

# --- Properties --- (Country code 1 assumed American)
cc = 1
w14 = prop.CreateBeamPropertyFromTable(cc,"W14X90",0,0.0,0.0)
w10 = prop.CreateBeamPropertyFromTable(cc,"W10X49",0,0.0,0.0)
w21 = prop.CreateBeamPropertyFromTable(cc,"W21X50",0,0.0,0.0)
w18 = prop.CreateBeamPropertyFromTable(cc,"W18X35",0,0.0,0.0)
angle = prop.CreateAnglePropertyFromTable(cc,"L40404",0,0.0)
prop.AssignBeamProperty([1,3,4],w14)
prop.AssignBeamProperty([2],w10)
prop.AssignBeamProperty([5,6,7],w21)
prop.AssignBeamProperty(list(range(8,14)),w18)
prop.AssignBeamProperty(list(range(14,24)),angle)
prop.AssignMaterialToMember("STEEL",list(range(1,24)))

# Beta angle
prop.AssignBetaAngle([3,4],90.0)

# --- Releases ---
# Member 5 START MZ -> release list indices [FX,FY,FZ,MX,MY,MZ]; release MZ only.
start_mp = prop.CreateMemberReleaseSpec(0,[0,0,0,0,0,1],[0,0,0,0,0,0])
start_my_mz = prop.CreateMemberPartialReleaseSpec(0,[0,1,1],[0.0,0.99,0.99])
end_my_mz = prop.CreateMemberPartialReleaseSpec(1,[0,1,1],[0.0,0.99,0.99])
prop.AssignMemberSpecToBeam([5],start_mp)
prop.AssignMemberSpecToBeam(list(range(14,24)),start_my_mz)
prop.AssignMemberSpecToBeam(list(range(14,24)),end_my_mz)


# --- Material Definition Unit Switch ---
# Switch to INCHES KIP (length=0 force=0) define material then back to Feet.
staad.SetInputUnits(0,0)
# Material is likely auto-defined; if needed ensure existence via Property or Root API; skipping explicit creation if wrapper handles.
staad.SetInputUnits(1,0)

# --- Supports ---
fixed_id = sup.CreateSupportFixed()
pinned_id = sup.CreateSupportPinned()
sup.AssignSupportToNode([1],fixed_id)
sup.AssignSupportToNode([2],pinned_id)

# --- Load Cases ---
# Case 1
case1 = load.CreateNewPrimaryLoadEx2("DEAD AND LIVE LOAD",0,1)
load.SetLoadActive(case1)
# Selfweight Y -1 (direction code: 2=GlobalY per typical OpenSTAAD convention) using AddSelfWeightInXYZ

load.AddSelfWeightInXYZ(2,-1.0)
# Joint loads
load.AddNodalLoad([4], 0.0, -15.0, 0.0, 0.0, 0.0, 0.0)
load.AddNodalLoad([5], 0.0, -15.0, 0.0, 0.0, 0.0, 0.0)
load.AddNodalLoad([11], 0.0, -35.0, 0.0, 0.0, 0.0, 0.0)
# Member loads
load.AddMemberUniformForce(list(range(8,14)),2,-0.9,0.0,0.0,0.0)
load.AddMemberUniformForce([6],2,-1.2,0.0,0.0,0.0)

# Case 2 Wind
case2 = load.CreateNewPrimaryLoadEx2("WIND FROM LEFT",3,2)
load.SetLoadActive(case2)
load.AddMemberUniformForce([1,2],4,0.6,0.0,0.0,0.0)  # 4=GlobalX
load.AddMemberUniformForce(list(range(8,11)),2,-1.0,0.0,0.0,0.0)

# Load Combination 3 (75% of 1 & 2)
comb3 = load.CreateNewLoadCombination("75 PERCENT DL LL WL",3)
# Add factor pairs (case1,case2) with 0.75 each
load.AddLoadAndFactorToCombination(3,1,0.75)
load.AddLoadAndFactorToCombination(3,2,0.75)

# Save model
staad.SaveModel(True)

# --- Perform Analysis Command---
staad.Command.PerformAnalysis(0)  # printOption=0 minimal output

print("Model created.")
print(f"Nodes: {len(node_coords)}, Members: {len(member_incidence)}")