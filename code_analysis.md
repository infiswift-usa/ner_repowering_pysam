D:\VS_CODE\Infiswift\metpv_11_automation\run_pvlib_metpv11.py

PVLIB SIMULATION - 9-PCS GRANULAR MODEL
USING METPV-11 CLEANED DATA
======================================================================

Debug - First day mid-day values:
Time: 2016-01-01 12:00:00+09:00
Zenith: 58.42 deg
GHI: 477.78 W/m2
DNI: 673.69 W/m2

Running pvlib ModelChain for 9 systems...


FINAL RESULTS (PVLIB WORKFLOW)
======================================================================
Annual AC Energy: 1,386,744 kWh

Monthly Production (kWh):
  Jan:   88,143 kWh
  Feb:  105,760 kWh
  Mar:  139,284 kWh
  Apr:  134,153 kWh
  May:  142,980 kWh
  Jun:  107,949 kWh
  Jul:  127,338 kWh
  Aug:  138,398 kWh
  Sep:  110,776 kWh
  Oct:  111,588 kWh
  Nov:   93,448 kWh
  Dec:   86,927 kWh

Results saved to metpv11_pvlib_results.csv
======================================================================

"D:\VS_CODE\Infiswift\metpv_11_automation\run_pvlib_metpv11_with_modes.py"

PVLIB SIMULATION - 9-PCS GRANULAR MODEL
USING METPV-11 CLEANED DATA
======================================================================

[Mode: LEVEL_1] Calculating Mono-facial POA (Standard)...
Debug Values at Mid-Day:
Total POA: 652.20 W/m2

Running pvlib ModelChain for 9 systems...

======================================================================

FINAL RESULTS (PVLIB WORKFLOW)
======================================================================
Annual AC Energy: 1,352,112 kWh

Monthly Production (kWh):
  Jan:   83,703 kWh
  Feb:  100,515 kWh
  Mar:  134,128 kWh
  Apr:  132,039 kWh
  May:  142,962 kWh
  Jun:  108,832 kWh
  Jul:  126,967 kWh
  Aug:  135,611 kWh
  Sep:  107,875 kWh
  Oct:  108,029 kWh
  Nov:   89,164 kWh
  Dec:   82,286 kWh

Results saved to metpv11_pvlib_results.csv

======================================================================

"D:\VS_CODE\Infiswift\metpv_11_automation\run_pvlib_metpv11_with_modes.py"   

PVLIB SIMULATION - 9-PCS GRANULAR MODEL
USING METPV-11 CLEANED DATA
======================================================================

[Mode: LEVEL_2] Calculating Bifacial POA (Infinite Sheds)...
Bifacial Gain Applied: 80% with Albedo 0.2
Debug Values at Mid-Day:
Total POA: 725.60 W/m2

Running pvlib ModelChain for 9 systems...

======================================================================

FINAL RESULTS (PVLIB WORKFLOW)
======================================================================
Annual AC Energy: 1,547,818 kWh

Monthly Production (kWh):
  Jan:   94,556 kWh
  Feb:  114,050 kWh
  Mar:  152,962 kWh
  Apr:  151,871 kWh
  May:  165,324 kWh
  Jun:  126,629 kWh
  Jul:  147,299 kWh
  Aug:  156,355 kWh
  Sep:  123,580 kWh
  Oct:  122,377 kWh
  Nov:  100,357 kWh
  Dec:   92,459 kWh

Results saved to metpv11_pvlib_results.csv
======================================================================
======================================================================

PYSAM SIMULATION (PVWATTS V8) - 9-PCS GRANULAR MODEL
METPV-11 CORRECTED DATA
======================================================================

Weather Data: 8760 hours
Location: 34.870N, 136.453E, 70.0m

[Calculating DNI using pysam...]
Debug: Zenith at 2016-01-01 12:00:00+09:00 is 58.44 deg

Simulating Group A (PCS 01-04): 4 units x 140.0kW DC / 100.0kW AC

Simulating Group B (PCS 05-09): 5 units x 130.0kW DC / 95.0kW AC

======================================================================

FINAL RESULTS
======================================================================
Annual AC Energy: 1,499,135 kWh
LATEST MAXIFIT Target: 1,535,110 kWh
Difference:            -35,975 kWh (-2.3%)

Monthly Production (kWh):
  Jan:   99,046 kWh
  Feb:  116,519 kWh
  Mar:  151,017 kWh
  Apr:  143,384 kWh
  May:  151,942 kWh
  Jun:  114,354 kWh
  Jul:  135,332 kWh
  Aug:  147,444 kWh
  Sep:  118,641 kWh
  Oct:  121,992 kWh
  Nov:  102,705 kWh
  Dec:   96,760 kWh

Results saved to metpv11_pysam_results.csv
======================================================================
1. The Discrepancy explained:
run_pvlib_metpv11 copy.py (1,352k kWh): This script uses the Mid-Hour Position (e.g., 12:30). This is the industry standard for hourly data like METPV because 12:30 represents the "average" sun angle for the entire hour of 12:00–13:00.
run_pvlib_metpv11.py (1,386k kWh): In your manual edits, you switched back to mc.run_model(weather). This function calculates the sun position at the Top-of-the-Hour (e.g., 12:00, 13:00).

2. Why the "Copy" is more:
METPV irradiation is an accumulation of energy over 60 minutes. If you calculate the sun's power at exactly 12:00 (when it's lower), it doesn't represent the energy received during the rest of the hour. By using the angle at 12:30, we get a much better average of the sun's intensity for that data block.

The Problem with Bifaciality
To calculate the bifacial "bonus," the physics engine needs to know how high the panels are (Hub Height) and how reflective the ground is (Albedo). Since these aren't in any of the config files, I used industry standard defaults (1.5m and 0.2).

Albedo and Hub height are engineering assumptions made to enable the Bifacial IR model because those specific site-design details aren't in the module/inverter datasheets.

Here is the breakdown of those values:

Albedo (0.2): This is the industry standard default for grass or standard soil.

If the site has white gravel, this should be 0.3 - 0.4 (which would increase yield). If it's dark soil, it might be 0.15.

Hub Height (1.5m): This is the height of the module center from the ground.

This value is needed for the "Infinite Sheds" model to calculate how much ground-reflected light reaches the back side.

======================================================================

**METPV-20 DATA**
PURE PYSAM SIMULATION (PVWATTS V8) - 9-PCS GRANULAR MODEL
USING HORIZONTAL METPV WEATHER DATA
======================================================================

Weather Data: 8760 hours
Location: 34.856N, 136.452E, 70.0m

[Calculating DNI using pvlib...]
Debug: Zenith at 2016-01-01 12:00:00+09:00 is 58.42 deg

Simulating Group A (PCS 01-04): 4 units x 140.0kW DC / 100.0kW AC

Simulating Group B (PCS 05-09): 5 units x 130.0kW DC / 95.0kW AC

======================================================================

FINAL RESULTS (9-PCS GRANULAR MODEL)
======================================================================
Annual AC Energy: 1,612,803 kWh

Monthly Production (kWh):
  Jan:  119,251 kWh
  Feb:  119,655 kWh
  Mar:  155,448 kWh
  Apr:  148,512 kWh
  May:  168,389 kWh
  Jun:  131,621 kWh
  Jul:  146,972 kWh
  Aug:  160,755 kWh
  Sep:  129,103 kWh
  Oct:  121,013 kWh
  Nov:  106,724 kWh
  Dec:  105,361 kWh

Results saved to pure_pysam_results_granular.csv
======================================================================
