#---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
#---------------------------------------------------------------------------------------------
from openstaadpy import os_analytical

staad = os_analytical.connect()

# Basic info
geometry = staad.Geometry
beam_list = geometry.GetBeamList()
print('Beam List:', beam_list)
print('File:', staad.GetSTAADFile())
print('Input Force Unit:', staad.GetInputUnitForForce())

# Ensure load case 1 exists (analysis model should already have been built by another script)
# Perform analysis (minimal print option)
status = staad.AnalyzeEx(1,1,1)
if(status < 0):
    print("Analysis failed with status code:", status)
    exit(1)
while(staad.IsAnalyzing()):
    pass
output = staad.Output

load_case_no = 1  # First primary load case
local_flag = 0    # 0=Global forces, 1=Local (assuming wrapper doc pattern)

# Get end forces for first 5 beams (or fewer if model smaller)
first_beams = beam_list[:5]
beam_forces_summary = []
for b in first_beams:
	try:
		start_forces = output.GetMemberEndForces(b, 0, load_case_no, local_flag)  # start end
		end_forces = output.GetMemberEndForces(b, 1, load_case_no, local_flag)    # end
		beam_forces_summary.append({
			'beam': b,
			'start': {
				'FX': start_forces[0], 'FY': start_forces[1], 'FZ': start_forces[2],
				'MX': start_forces[3], 'MY': start_forces[4], 'MZ': start_forces[5]
			},
			'end': {
				'FX': end_forces[0], 'FY': end_forces[1], 'FZ': end_forces[2],
				'MX': end_forces[3], 'MY': end_forces[4], 'MZ': end_forces[5]
			}
		})
	except Exception as e:
		beam_forces_summary.append({'beam': b, 'error': str(e)})

print("\nBeam end forces (Load Case 1):")
for bf in beam_forces_summary:
	if 'error' in bf:
		print(f" Beam {bf['beam']}: ERROR {bf['error']}")
	else:
		print(f" Beam {bf['beam']} Start FX={bf['start']['FX']:.3f} FY={bf['start']['FY']:.3f} FZ={bf['start']['FZ']:.3f} MX={bf['start']['MX']:.3f} MY={bf['start']['MY']:.3f} MZ={bf['start']['MZ']:.3f} | End FX={bf['end']['FX']:.3f} FY={bf['end']['FY']:.3f} FZ={bf['end']['FZ']:.3f} MX={bf['end']['MX']:.3f} MY={bf['end']['MY']:.3f} MZ={bf['end']['MZ']:.3f}")

# Support reactions: collect for nodes that are supports (heuristic: node 1 and 2 if present)
support_nodes = [n for n in [1,2] if n in geometry.GetNodeList()]
print("\nSupport reactions (Load Case 1):")
for nid in support_nodes:
	try:
		reactions = output.GetSupportReactions(nid, load_case_no)
		# reactions order FX,FY,FZ,MX,MY,MZ per wrapper docs
		print(f" Node {nid} RX={reactions[0]:.3f} RY={reactions[1]:.3f} RZ={reactions[2]:.3f} MX={reactions[3]:.3f} MY={reactions[4]:.3f} MZ={reactions[5]:.3f}")
	except Exception as e:
		print(f" Node {nid} ERROR {e}")

print("\nExtraction complete.")

