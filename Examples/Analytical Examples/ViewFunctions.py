#---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
#---------------------------------------------------------------------------------------------
from openstaadpy import os_analytical
from openstaadpy.os_analytical.oserrors import OsGroupAlreadyExists
import time

staad = os_analytical.connect()
view = staad.View
geo = staad.Geometry

print("Starting view demonstration with group creation...")

# 1. Toggle several diagram/display modes (IDs are wrapper-defined; adjust if needed)
diagram_modes = [1, 2, 9, 10, 12]  # sample: bending, shear, full section, outline, shrink
for mode in diagram_modes:
	try:
		view.SetDiagramMode(mode, True, True)
		print(f"Diagram mode {mode} ON")
		time.sleep(1.0)
		view.SetDiagramMode(mode, False, True)
		print(f"Diagram mode {mode} OFF")
	except Exception as e:
		print(f"Diagram mode {mode} error: {e}")
	time.sleep(1.0)

# 2. Select members parallel to global X, create a group from that selection, then re-select via group
try:
	group_name = "_ParallelX"
	view.SelectMembersParallelTo("X")
	print("Selected members parallel to X axis")
	selected_beams = geo.GetSelectedBeams()
	print(f"Parallel X selection count: {len(selected_beams)}")
	if selected_beams:
		try:
			geo.CreateGroupEx(2, group_name, selected_beams)  # 2 = Members group type
			print(f"Created group '{group_name}' with {len(selected_beams)} members")
		except OsGroupAlreadyExists:
			print(f"Group '{group_name}' already exists; continuing")
		geo.ClearMemberSelection()
		print("Cleared member selection after group demo")
	else:
		print("No members found parallel to X; skipping group creation")
except Exception as e:
	print("Parallel X selection/group creation error:", e)
time.sleep(0.8)

geo.ClearMemberSelection()
print("Cleared member selection before selecting group")
res = view.SelectGroup(group_name)
print(f"SelectGroup('{group_name}') -> {res}")
time.sleep(1.0)
geo.ClearMemberSelection()

# 4. Create a new view from (currently empty) selection just to illustrate API
try:
	geo.SelectMultipleBeams([1,2,3])
	view.CreateNewViewForSelections()
	print("Created new view for current (empty) selection.")
except Exception as e:
	print("CreateNewViewForSelections error:", e)
time.sleep(1.0)

# 5. Position window
availableL, availableW = view.GetApplicationDesktopSize()
view.SetWindowPosition(0, 0, availableL//2, availableW//2)
print("Window positioned (half size, top-left).")
time.sleep(1.0)

print("View demo sequence complete.")
