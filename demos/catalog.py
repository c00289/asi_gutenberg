"""Demo catalog for asi_gutenberg.

Each builder returns a Document.  Content is themed around an autonomous
wheel loader and its functional-safety system (the same world as the
asi_escher catalog) so the documents read like real engineering work
products rather than lorem ipsum.  The figures are asi_escher's bundled
example diagrams, carried as ``diagram`` blocks and rendered by ``build()``
into ``figures/`` next to the generated Markdown.
"""
from __future__ import annotations

import random

from asi_escher.examples import EXAMPLES

from asi_gutenberg import Document

ESCHER = dict(EXAMPLES)          # name -> builder of an asi_escher example diagram

ORG = "ASI Robots"
PROJECT = "WA485 Autonomous Wheel Loader"
DATE = "2026-10-27"
AUTHOR = "Roy (Capgemini, FuSa)"
APPROVERS = [("Functional Safety Manager", "Max Stone"), ("Safety Mentor", "Ida Moss"),
             ("Autonomy Lead", "TBD")]


def _std(doc: Document, rev_desc: str = "Initial draft") -> Document:
    doc.revision("0.1", DATE, "Roy", rev_desc)
    return doc


# ------------------------------------------------------------ spec --------
def spec_safety_requirements() -> Document:
    d = Document("Functional Safety Requirements Specification", "spec",
                 subtitle="Safety goals, functional and technical safety requirements for the WA485 autonomy kit",
                 doc_id="ASI-FS-SRS-001", revision="0.3", status="In Review", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal",
                 approvers=APPROVERS, keywords=["ISO 25119", "ISO 13849", "requirements", "PLr"])
    d.revision("0.1", "2026-10-13", "Roy", "Initial draft from HARA outputs")
    d.revision("0.2", "2026-10-20", "Roy", "Added TSRs for drive inhibit and zone monitoring")
    d.revision("0.3", DATE, "Roy", "Review comments from safety mentor incorporated")
    d.h1("Introduction")
    d.h2("Purpose")
    d.p("This specification captures the **safety goals** derived from the hazard analysis and risk "
        "assessment (HARA), the **functional safety requirements** (FSR) that implement them and the "
        "**technical safety requirements** (TSR) allocated to hardware and software elements of the "
        "WA485 autonomy kit.")
    d.h2("Scope")
    d.p("The scope is the autonomy kit fitted to the production WA485 wheel loader: perception suite, "
        "planning computer, vehicle interface and the safety system (safety PLC, E-stop circuit, drive "
        "inhibit relay, safety lidar and safety HMI). The base machine's own protective functions are "
        "out of scope and treated as *assumed safe*.")
    d.h2("Normative references")
    d.table(["Reference", "Title", "Use"],
            [["ISO 25119:2018", "Tractors and machinery for agriculture and forestry - safety-related parts of control systems", "AgPL determination, lifecycle"],
             ["ISO 13849-1:2023", "Safety of machinery - safety-related parts of control systems", "PL / category evaluation"],
             ["ISO 17757:2019", "Earth-moving machinery and mining - autonomous and semi-autonomous machine system safety", "ASAM operating concept"],
             ["IEC 61496-3", "Electro-sensitive protective equipment - AOPDDR", "Safety lidar type"]],
            caption="Normative references")
    d.h2("Conventions")
    d.note("Requirements use **shall** for binding statements. Every requirement carries a unique identifier, "
           "a required performance level (`plr`), a status and its upstream trace (`traces`).", "Requirement language")
    d.h1("Safety Goals")
    d.p("Safety goals are the top-level outcomes of the HARA. Each goal lists the hazard it mitigates and "
        "the required performance level inherited from the worst-case hazard scenario.")
    d.requirement("SG-01", "Prevent unintended motion",
                  "The loader shall not produce drive, steering or implement motion while in STANDBY or SAFE_STOP.",
                  plr="d", status="Approved", hazard="H-03 unexpected start-up")
    d.requirement("SG-02", "Stop before an occupied exclusion zone",
                  "The loader shall come to a stop before any part of the machine enters an exclusion zone occupied by a person or vehicle.",
                  plr="d", status="Approved", hazard="H-07 collision with pedestrian")
    d.requirement("SG-03", "Limit speed in degraded perception",
                  "When any perception channel reports reduced confidence the loader shall limit ground speed to 2 km/h or stop.",
                  plr="c", status="Approved", hazard="H-09 undetected obstacle")
    d.requirement("SG-04", "Prevent uncommanded implement motion",
                  "Boom and bucket shall not move without a valid trajectory command and a closed safety loop.",
                  plr="c", status="In Review", hazard="H-12 implement strike")
    d.diagram(ESCHER["flowchart_fusa_lifecycle"](), "Safety lifecycle context for these requirements (asi_escher)")
    d.h1("Functional Safety Requirements")
    d.h2("Drive and motion")
    d.requirement("FSR-01", "Drive inhibit",
                  "A drive inhibit function shall remove hydraulic pilot pressure and open the traction enable path on entry to SAFE_STOP.",
                  plr="d", status="Approved", traces="SG-01")
    d.requirement("FSR-02", "Dual-channel E-stop evaluation",
                  "The E-stop circuit shall be evaluated as two independent channels; a discrepancy longer than 500 ms shall be treated as a trip.",
                  plr="d", status="Approved", traces="SG-01")
    d.requirement("FSR-03", "Speed limitation in DEGRADED",
                  "The safety PLC shall enforce a 2 km/h limit by capping the speed command sent to the drive controller whenever the DEGRADED flag is set.",
                  plr="c", status="Approved", traces="SG-03")
    d.h2("Zone monitoring")
    d.requirement("FSR-04", "Protective field monitoring",
                  "Protective fields around the machine shall be monitored at no less than 10 Hz with a detection capability of a 200 mm test piece at 0.3 reflectivity.",
                  plr="d", status="Approved", traces="SG-02")
    d.requirement("FSR-05", "Field switching by speed",
                  "The active protective field set shall be selected from the measured ground speed such that the stopping distance plus a 0.5 m margin is always contained.",
                  plr="d", status="In Review", traces="SG-02")
    d.h2("Implement")
    d.requirement("FSR-06", "Implement enable interlock",
                  "Implement valves shall be enabled only while a trajectory heartbeat is received at 10 Hz and the safety loop is closed.",
                  plr="c", status="Draft", traces="SG-04")
    d.h1("Technical Safety Requirements")
    d.requirement("TSR-01", "Force-guided inhibit relay",
                  "The drive inhibit relay shall use force-guided contacts and its state shall be read back by the safety PLC every cycle (10 ms).",
                  plr="d", status="Approved", traces="FSR-01", allocation="Safety PLC, Drive Inhibit Relay")
    d.requirement("TSR-02", "Stopping distance",
                  "From 5 km/h on a 10 % grade the machine shall stop within 2.0 m of the stop command.",
                  plr="d", status="Approved", traces="FSR-04, FSR-05", allocation="Brake Actuator")
    d.requirement("TSR-03", "Lidar OSSD wiring",
                  "Each safety lidar OSSD pair shall be wired to a dedicated dual-channel safety input with cross-circuit detection.",
                  plr="d", status="Approved", traces="FSR-04", allocation="Safety Lidar, Safety PLC")
    d.requirement("TSR-04", "Heartbeat timeout",
                  "Loss of the autonomy heartbeat for more than 100 ms shall transition the safety PLC to SAFE_STOP.",
                  plr="d", status="Approved", traces="FSR-06, FSR-01", allocation="Safety PLC")
    d.h1("Traceability")
    d.table(["Requirement", "Traces to", "Allocated to", "Verification"],
            [["FSR-01", "SG-01", "Safety PLC", "Test TC-11"], ["FSR-02", "SG-01", "E-Stop Circuit", "Test TC-12"],
             ["FSR-03", "SG-03", "Safety PLC", "Test TC-21"], ["FSR-04", "SG-02", "Safety Lidar", "Analysis + TC-31"],
             ["FSR-05", "SG-02", "Safety PLC", "Test TC-32"], ["FSR-06", "SG-04", "Safety PLC", "Test TC-41"],
             ["TSR-01", "FSR-01", "Relay", "Inspection + TC-11"], ["TSR-02", "FSR-04, FSR-05", "Brake", "Test TC-33 (grade)"],
             ["TSR-03", "FSR-04", "Lidar", "Inspection"], ["TSR-04", "FSR-06, FSR-01", "Safety PLC", "Test TC-42"]],
            caption="Requirement trace matrix")
    d.diagram(ESCHER["req_safety_requirements"](), "Requirement diagram (asi_escher)")
    d.appendix("Open points")
    d.tasks([("Confirm 0.3 reflectivity test piece with lidar supplier", False),
             ("Agree 2 km/h DEGRADED limit with operations", True),
             ("Allocate FSR-06 to vehicle interface firmware", False)])
    return d


def icd_safety_plc() -> Document:
    d = Document("Safety PLC Interface Control Document", "icd",
                 subtitle="Signals, timing and CAN frames between the autonomy computer and the safety PLC",
                 doc_id="ASI-FS-ICD-004", revision="1.0", status="Released", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:2])
    d.revision("0.1", "2026-10-06", "Roy", "Draft signal list")
    d.revision("1.0", DATE, "Roy", "Released for integration")
    d.h1("Overview")
    d.p("The autonomy computer and the safety PLC exchange a small set of safety-relevant signals over a "
        "dedicated CAN bus (`can_s`, 500 kbit/s) plus discrete I/O for the E-stop and inhibit paths. This "
        "document is the single source for signal names, encodings and timing.")
    d.diagram(ESCHER["ibd_safety_controller"](), "Internal block diagram of the safety system (asi_escher)")
    d.h1("Discrete I/O")
    d.table(["Signal", "Direction", "Type", "Source", "Sink", "Notes"],
            [["estop_ch1", "in", "24 V DI", "E-stop loop", "PLC di0", "Channel 1, NC contact"],
             ["estop_ch2", "in", "24 V DI", "E-stop loop", "PLC di1", "Channel 2, NC contact"],
             ["lidar_ossd1/2", "in", "OSSD pair", "Safety lidar", "PLC di2:3", "Field set A"],
             ["lidar_ossd3/4", "in", "OSSD pair", "Safety lidar", "PLC di4:5", "Field set B"],
             ["inhibit_cmd", "out", "24 V DO", "PLC do0", "Relay coil", "Energise to run"],
             ["inhibit_fb", "in", "24 V DI", "Relay aux", "PLC di6", "Force-guided read-back"],
             ["brake_cmd", "out", "24 V DO", "PLC do1", "Brake solenoid", "De-energise to apply"]],
            caption="Discrete signals")
    d.h1("CAN messages")
    d.h2("Autonomy heartbeat (0x100)")
    d.table(["Byte", "Field", "Encoding", "Range"],
            [["0", "counter", "uint8, rolling", "0..255"], ["1", "mode", "enum", "0 MANUAL, 1 STANDBY, 2 AUTO, 3 DEGRADED, 4 SAFE_STOP"],
             ["2-3", "speed_cmd", "int16, 0.01 m/s", "-350..350"], ["4-5", "steer_cmd", "int16, 0.01 deg", "-4000..4000"],
             ["6", "flags", "bitfield", "b0 implement_enable, b1 perception_degraded"], ["7", "crc8", "SAE J1850", "-"]],
            caption="0x100 layout, 10 Hz, timeout 100 ms", align=["center", "left", "left", "left"])
    d.h2("Safety status (0x101)")
    d.table(["Byte", "Field", "Encoding", "Range"],
            [["0", "counter", "uint8, rolling", "0..255"], ["1", "state", "enum", "0 IDLE, 1 ARMED, 2 TRIPPED, 3 DISCREPANCY, 4 RESET_PENDING"],
             ["2-3", "speed_limit", "uint16, 0.01 m/s", "0..350"], ["4", "fault_code", "uint8", "see Appendix A"],
             ["5", "field_set", "uint8", "0 none, 1 A, 2 B"], ["7", "crc8", "SAE J1850", "-"]],
            caption="0x101 layout, 10 Hz", align=["center", "left", "left", "left"])
    d.h2("Encoding example")
    d.code("# autonomy heartbeat, AUTO mode, 1.5 m/s, 0 deg, implement enabled\n"
           "frame = bytes([0x2A, 0x02, 0x00, 0x96, 0x00, 0x00, 0x01, crc8(payload)])\n"
           "bus.send(can.Message(arbitration_id=0x100, data=frame, is_extended_id=False))", "python")
    d.warning("A heartbeat whose rolling counter repeats or skips more than one step is treated as lost. "
              "Do **not** resend a frame with the same counter after a transmit error.", "Counter discipline")
    d.h1("Timing")
    d.bullets(["Heartbeat period 100 ms, timeout 100 ms after last valid frame",
               "Safety status period 100 ms; the autonomy computer shall enter STANDBY if no status for 300 ms",
               ("Reaction chain on E-stop", ["input filter 2 ms", "PLC logic cycle 10 ms", "relay drop-out 15 ms", "pilot pressure decay 120 ms"])])
    d.appendix("Fault codes")
    d.table(["Code", "Meaning", "Reaction"],
            [["0x00", "No fault", "-"], ["0x11", "E-stop discrepancy", "TRIPPED after 500 ms"],
             ["0x12", "Inhibit relay read-back mismatch", "SAFE_STOP, lock-out"], ["0x21", "Lidar OSSD cross-circuit", "SAFE_STOP"],
             ["0x31", "Heartbeat timeout", "SAFE_STOP"], ["0x32", "Heartbeat CRC error x3", "DEGRADED"]],
            caption="Safety PLC fault codes")
    return d


# ---------------------------------------------------------- reports -------
def report_hara() -> Document:
    d = Document("Hazard Analysis and Risk Assessment", "analysis",
                 subtitle="HARA for autonomous load-and-carry operation at the pilot quarry",
                 doc_id="ASI-FS-HARA-002", revision="0.4", status="In Review", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS)
    d.revision("0.1", "2026-10-08", "Roy", "Operating scenarios drafted")
    d.revision("0.4", DATE, "Roy", "Severity / exposure / controllability re-rated with operations")
    d.h1("Method")
    d.p("Hazards were identified per operating scenario using a guide-word review (STPA control actions "
        "*not provided*, *provided incorrectly*, *too early / too late*, *stopped too soon*). Each hazardous "
        "event was rated for **severity (S)**, **exposure (E)** and **controllability (C)** per ISO 25119 and "
        "mapped to a required performance level (AgPLr).")
    d.table(["Parameter", "Class", "Definition"],
            [["S", "S0 - S3", "No injury / light / severe / life-threatening"],
             ["E", "E1 - E4", "Very low / low / medium / high probability of exposure"],
             ["C", "C0 - C3", "Controllable in general / simply / normally / difficult"]],
            caption="Rating classes")
    d.h1("Operating scenarios")
    d.numbered_list(["Load cycle at the stockpile face with a truck in the dump position",
                     "Travel between stockpile and crusher hopper along the haul road",
                     "Manual-to-autonomous hand-over at the staging area",
                     "Maintenance access while the machine is in STANDBY",
                     "Pedestrian survey crew inside the geofence (permit-controlled)"])
    d.diagram(ESCHER["activity_load_cycle"](), "Autonomous load cycle used for scenario 1 (asi_escher)")
    d.h1("Hazard register")
    rows = [
        ["H-01", "Unexpected start-up in STANDBY", "1, 4", "S3", "E3", "C2", "d", "SG-01"],
        ["H-02", "Loss of steering authority on haul road", "2", "S3", "E3", "C1", "c", "SG-03"],
        ["H-03", "Unexpected start-up during maintenance", "4", "S3", "E2", "C3", "d", "SG-01"],
        ["H-05", "Reversing into truck cab", "1", "S2", "E4", "C2", "c", "SG-02"],
        ["H-07", "Collision with pedestrian inside geofence", "5", "S3", "E2", "C3", "d", "SG-02"],
        ["H-09", "Obstacle not detected in dust", "1, 2", "S3", "E3", "C2", "d", "SG-03"],
        ["H-12", "Bucket strike on spotter during dump", "1", "S3", "E2", "C2", "c", "SG-04"],
        ["H-14", "Runaway on grade after brake fault", "2", "S3", "E2", "C3", "d", "SG-02"],
    ]
    d.table(["ID", "Hazardous event", "Scenario", "S", "E", "C", "AgPLr", "Safety goal"], rows,
            caption="Hazard register (extract)", align=["left", "left", "center", "center", "center", "center", "center", "left"])
    d.caution("Ratings for H-09 assume the dust suppression system is operational. If it is removed from the "
              "site plan the exposure rises to E4 and the goal remains PL d with reduced margin.")
    d.h1("Results")
    d.p("Four safety goals cover the register. Three are rated **PL d** and one **PL c**; the architecture "
        "in the technical safety concept therefore targets category 3 with a dual-channel safety PLC.")
    d.table(["AgPLr", "Count", "Goals"],
            [["d", "5", "SG-01, SG-02"], ["c", "3", "SG-03, SG-04"]], caption="Hazards by required performance level",
            align=["center", "right", "left"])
    d.h1("Assumptions")
    d.bullets(["Site traffic management keeps light vehicles outside the loader's active zone",
               "Survey crews enter only under a permit that puts the machine in STANDBY",
               "Operators are trained on the safety HMI reset procedure"])
    return d


def report_fmea_brake() -> Document:
    d = Document("DFMEA: Brake and Drive Inhibit Path", "analysis",
                 subtitle="Design failure mode and effects analysis for the safe-stop actuation chain",
                 doc_id="ASI-FS-FMEA-007", revision="0.2", status="Draft", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:1])
    _std(d, "Worksheet populated for relay, solenoid and pressure sensor")
    d.h1("Scope and ground rules")
    d.p("The analysis covers the components between the safety PLC output stage and the park brake / pilot "
        "pressure cut, as shown in the internal block diagram. Ratings use 1-10 scales; the risk priority "
        "number (RPN) is S x O x D. Items with **RPN >= 100** or **S >= 9** require an action.")
    d.diagram(ESCHER["bdd_safety_system"](), "Safety system block definition (asi_escher)")
    d.h1("Worksheet")
    rows = [
        ["Inhibit relay", "Contacts weld closed", "Drive stays enabled after stop command", "9", "2", "3", "54", "Force-guided aux read-back (TSR-01)"],
        ["Inhibit relay", "Coil open circuit", "Cannot enable drive (safe)", "2", "3", "2", "12", "-"],
        ["Brake solenoid", "Stuck energised", "Brake not applied on stop", "9", "3", "5", "135", "Add pressure switch feedback; test weekly"],
        ["Brake solenoid", "Stuck de-energised", "Brake applied unexpectedly (skid)", "7", "2", "4", "56", "Controlled-stop ramp before park brake"],
        ["Pilot pressure sensor", "Reads low (drift)", "False 'pilot cut' confirmation", "8", "3", "6", "144", "Plausibility vs pump pressure; dual sensor"],
        ["PLC output do1", "Short to 24 V", "Brake cannot be commanded", "9", "1", "2", "18", "Test pulse detection in PLC"],
        ["Wiring harness", "Chafed to chassis", "Loss of inhibit command (safe) or stuck 24 V", "8", "2", "4", "64", "Routing review, abrasion sleeve"],
    ]
    d.table(["Item", "Failure mode", "Effect", "S", "O", "D", "RPN", "Action"], rows,
            caption="DFMEA worksheet", align=["left", "left", "left", "center", "center", "center", "center", "left"])
    d.important("Two items exceed the RPN threshold. Both are owned by the hydraulics lead and tracked as "
                "CR-118 and CR-119 in the change register.", "Actions required")
    d.h1("Actions")
    d.tasks([("CR-118 Pressure switch feedback on brake circuit", False),
             ("CR-119 Dual pilot pressure sensor with plausibility check", False),
             ("Weekly brake function test added to maintenance schedule", True)])
    return d


def report_post_test() -> Document:
    d = Document("Power-On Self Test Report", "report",
                 subtitle="Results of the POST campaign on loader unit WA485-A02, firmware 2.3.1",
                 doc_id="ASI-FS-TR-021", revision="1.0", status="Approved", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:2])
    d.revision("1.0", DATE, "Roy", "Final results after re-test of T4")
    d.h1("Summary")
    d.p("All five POST checks were exercised 50 times each on unit A02. Four passed on every run; the brake "
        "pressure sensor test (T4) failed twice on cold starts and passed after the sensor warm-up delay was "
        "raised from 200 ms to 500 ms. The firmware was re-tested and **released**.")
    d.diagram(ESCHER["activity_fork_join_post"](), "POST activity diagram (asi_escher)")
    d.h1("Results")
    d.table(["Test", "Description", "Runs", "Pass", "Fail", "Mean time", "Verdict"],
            [["T1", "Lidar heartbeat", "50", "50", "0", "118 ms", "PASS"],
             ["T2", "Radar heartbeat", "50", "50", "0", "96 ms", "PASS"],
             ["T3", "E-stop loop continuity", "50", "50", "0", "22 ms", "PASS"],
             ["T4", "Brake pressure sensor", "50", "48", "2", "512 ms", "FAIL (cold)"],
             ["T4r", "Brake pressure sensor, 500 ms warm-up", "50", "50", "0", "640 ms", "PASS"],
             ["T5", "Steering feedback", "50", "50", "0", "74 ms", "PASS"]],
            caption="POST results by test", align=["left", "left", "right", "right", "right", "right", "center"])
    d.h2("Failure analysis T4")
    d.p("Both failures occurred at ambient temperatures below 4 C. The pressure sensor output had not settled "
        "within the 200 ms window; the raw log shows the reading crossing the 2 bar threshold at 380-420 ms.")
    d.code("2026-10-21T06:12:03.118Z POST T4 start  ambient=2.8C\n"
           "2026-10-21T06:12:03.318Z POST T4 sample p=1.62bar  (threshold 2.00)\n"
           "2026-10-21T06:12:03.319Z POST T4 FAIL   reason=below_threshold\n"
           "2026-10-21T06:12:03.520Z POST T4 sample p=2.11bar  (late)", "log")
    d.tip("Cold-start behaviour is now part of the POST acceptance checklist; see the V&V plan section 4.2.")
    d.h1("Conclusion")
    d.bullets(["Firmware 2.3.1 with the 500 ms warm-up is approved for the pilot fleet",
               "T4 cold-start case added to the regression suite as TC-POST-04c",
               "No change to the safety concept is required"])
    return d


def report_stopping_distance() -> Document:
    d = Document("Stopping Distance Test Report", "report",
                 subtitle="Verification of TSR-02 on the 10 % test grade",
                 doc_id="ASI-FS-TR-033", revision="0.9", status="In Review", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:2])
    _std(d, "Data from 24 runs, awaiting witness signature")
    d.h1("Objective")
    d.p("Demonstrate that the loader stops within **2.0 m** of the stop command from 5 km/h on a 10 % downhill "
        "grade with rated payload (TSR-02).")
    d.h1("Set-up")
    d.fields({"Unit": "WA485-A02", "Payload": "5 500 kg (rated)", "Surface": "compacted gravel, dry",
              "Grade": "10.2 % measured", "Trigger": "safety PLC SAFE_STOP via test E-stop",
              "Measurement": "RTK GNSS 20 Hz + wheel encoder"}, caption="Test configuration")
    d.h1("Results")
    rng = random.Random(33)
    rows = []
    for i in range(1, 25):
        v = round(5.0 + rng.uniform(-0.2, 0.2), 2)
        t = round(0.12 + rng.uniform(-0.02, 0.03), 3)
        dist = round(0.98 + (v - 5.0) * 0.25 + rng.uniform(-0.08, 0.12), 2)
        rows.append([str(i), f"{v:.2f}", f"{t:.3f}", f"{dist:.2f}", "PASS" if dist <= 2.0 else "FAIL"])
    d.table(["Run", "Speed (km/h)", "Reaction (s)", "Distance (m)", "Verdict"], rows,
            caption="Stopping distance per run", align=["right", "right", "right", "right", "center"])
    dists = [float(r[3]) for r in rows]
    d.table(["Statistic", "Value"],
            [["Runs", str(len(rows))], ["Mean distance", f"{sum(dists)/len(dists):.2f} m"],
             ["Max distance", f"{max(dists):.2f} m"], ["Margin to limit", f"{2.0 - max(dists):.2f} m"]],
            caption="Summary statistics", align=["left", "right"])
    d.h1("Verdict")
    d.important("All 24 runs stopped within the 2.0 m limit with a worst-case margin above 0.8 m. "
                "TSR-02 is **verified** pending witness signature.")
    return d


# ------------------------------------------------------------- plans ------
def plan_functional_safety() -> Document:
    d = Document("Functional Safety Plan", "plan",
                 subtitle="Lifecycle, roles, work products and milestones for the WA485 autonomy kit",
                 doc_id="ASI-FS-PLAN-001", revision="0.2", status="Draft", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS)
    d.revision("0.1", "2026-10-15", "Roy", "Outline from onboarding sessions")
    d.revision("0.2", DATE, "Roy", "Milestones aligned with pilot schedule")
    d.h1("Purpose")
    d.p("This plan defines how functional safety is managed over the lifecycle of the autonomy kit: who is "
        "responsible, which work products are produced at each phase, and how confirmation measures are applied.")
    d.diagram(ESCHER["flowchart_fusa_lifecycle"](), "Functional safety lifecycle (asi_escher)")
    d.h1("Roles and responsibilities")
    d.table(["Activity", "FuSa Manager", "FuSa Engineer", "Autonomy Lead", "Hydraulics Lead", "Site Ops"],
            [["HARA", "A", "R", "C", "C", "C"], ["Safety concept", "A", "R", "C", "C", "I"],
             ["Requirements", "A", "R", "C", "C", "I"], ["Hardware PL evaluation", "A", "C", "I", "R", "I"],
             ["Software verification", "A", "C", "R", "I", "I"], ["Integration testing", "A", "R", "R", "C", "C"],
             ["Safety case", "R/A", "R", "C", "C", "I"], ["Operational release", "A", "C", "C", "I", "R"]],
            caption="RACI matrix", align=["left", "center", "center", "center", "center", "center"])
    d.h1("Work products")
    d.table(["Phase", "Work product", "Document ID", "Status"],
            [["Concept", "Item definition", "ASI-FS-ID-001", "Approved"], ["Concept", "HARA", "ASI-FS-HARA-002", "In Review"],
             ["Concept", "Functional safety concept", "ASI-FS-FSC-003", "Draft"], ["Development", "Safety requirements", "ASI-FS-SRS-001", "In Review"],
             ["Development", "Safety PLC ICD", "ASI-FS-ICD-004", "Released"], ["Development", "DFMEA brake path", "ASI-FS-FMEA-007", "Draft"],
             ["Verification", "V&V plan", "ASI-FS-VVP-010", "Draft"], ["Verification", "POST report", "ASI-FS-TR-021", "Approved"],
             ["Release", "Safety case", "ASI-FS-SC-050", "Draft"]],
            caption="Work product register")
    d.h1("Milestones")
    d.table(["Milestone", "Date", "Entry criteria"],
            [["M1 Concept gate", "2026-11-14", "HARA approved, safety goals baselined"],
             ["M2 Design gate", "2026-12-19", "SRS approved, PL evaluation complete"],
             ["M3 Integration complete", "2027-02-06", "All TSR tests passed on unit A02"],
             ["M4 Pilot release", "2027-03-13", "Safety case approved by FuSa manager"]],
            caption="Milestones")
    d.h1("Confirmation measures")
    d.bullets(["Confirmation review of HARA and safety concept by an independent FuSa engineer (I1)",
               "Functional safety audit of the development process before M2",
               "Functional safety assessment of the safety case before M4"])
    d.note("Independence level I1 means a different person; I2 a different team. Assessment before M4 is I2.")
    d.h1("Change management")
    d.p("Changes to safety-relevant elements are raised as change requests (CR) and impact-assessed against "
        "the HARA and the SRS before approval. See the engineering change notice template.")
    return d


def plan_verification_validation() -> Document:
    d = Document("Verification and Validation Plan", "plan",
                 subtitle="Test strategy, environments and traceability for safety requirements",
                 doc_id="ASI-FS-VVP-010", revision="0.1", status="Draft", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:2])
    _std(d)
    d.h1("Strategy")
    d.p("Each safety requirement is verified by at least one method: **test**, **analysis**, **inspection** or "
        "**demonstration**. Tests are run at three levels - bench (HIL), vehicle on the test pad and vehicle "
        "at the pilot site.")
    d.h1("Environments")
    d.table(["Level", "Environment", "Scope", "Owner"],
            [["L1", "Safety PLC HIL bench", "Logic, timing, fault injection", "FuSa Engineer"],
             ["L2", "Test pad, unit A02", "Actuation chain, stopping distance, POST", "Autonomy Lead"],
             ["L3", "Pilot quarry", "Operational scenarios, zone monitoring", "Site Ops"]],
            caption="Test environments")
    d.h1("Test matrix")
    d.table(["Test", "Requirement", "Level", "Method", "Status"],
            [["TC-11", "FSR-01, TSR-01", "L1, L2", "Test + inspection", "Planned"], ["TC-12", "FSR-02", "L1", "Fault injection", "Planned"],
             ["TC-21", "FSR-03", "L1", "Test", "Planned"], ["TC-31", "FSR-04", "L2", "Analysis + test", "Planned"],
             ["TC-32", "FSR-05", "L2", "Test", "Planned"], ["TC-33", "TSR-02", "L2", "Test on grade", "Passed"],
             ["TC-41", "FSR-06", "L1", "Test", "Planned"], ["TC-42", "TSR-04", "L1", "Fault injection", "Planned"],
             ["TC-POST-04c", "POST T4 cold start", "L2", "Test", "Passed"]],
            caption="Test matrix")
    d.h2("Cold start acceptance")
    d.p("POST tests are repeated at ambient below 5 C with the machine cold-soaked for at least four hours.")
    d.h1("Entry and exit criteria")
    d.bullets([("Entry", ["Requirement baselined", "Test procedure reviewed", "Environment calibrated"]),
               ("Exit", ["All steps executed", "Deviations dispositioned", "Report approved"])])
    d.h1("Deliverables")
    d.tasks([("Test procedures TC-11 to TC-42", False), ("HIL bench commissioning report", False),
             ("Stopping distance report ASI-FS-TR-033", True), ("POST report ASI-FS-TR-021", True)])
    return d


# ------------------------------------------------------- procedures -------
def procedure_estop_reset() -> Document:
    d = Document("E-Stop Reset Procedure", "procedure",
                 subtitle="Work instruction for resetting the loader after an emergency stop",
                 doc_id="ASI-OPS-WI-012", revision="1.1", status="Released", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Site personnel", approvers=APPROVERS[:1])
    d.revision("1.0", "2026-09-30", "Roy", "First release")
    d.revision("1.1", DATE, "Roy", "Added discrepancy case (step 6)")
    d.h1("Scope")
    d.p("Applies to all WA485 units at the pilot site after any E-stop trip, whether from the cab button, the "
        "remote pendant or the site perimeter loop.")
    d.caution("Never reset an E-stop without knowing **why** it tripped. If the cause is not obvious, call the "
              "shift supervisor before proceeding.", "Before you start")
    d.diagram(ESCHER["activity_swimlanes_estop"](), "E-stop handling across operator, autonomy, safety controller and machine (asi_escher)")
    d.h1("Procedure")
    d.numbered_list(["Confirm the machine is stationary and the park brake indicator is lit",
                     "Identify which E-stop tripped using the safety HMI fault page",
                     "Clear the hazard that caused the trip; ensure no one is inside the active zone",
                     "Release the tripped E-stop button (twist to release)",
                     "On the safety HMI press **Reset** and hold for 2 s",
                     ("If the HMI shows *Discrepancy*", ["press **Reset** again after 5 s", "if it persists, open a maintenance ticket and leave the machine in SAFE_STOP"]),
                     "Confirm the state indicator shows STANDBY",
                     "Hand the machine back to the fleet server by acknowledging the mission prompt"])
    d.warning("Steps 4 and 5 must be done by the same person within 30 s or the reset will time out.")
    d.h1("Records")
    d.table(["Field", "Entry"], [["Date / time", ""], ["Unit", ""], ["E-stop source", ""],
                                 ["Cause", ""], ["Reset by", ""], ["Supervisor informed", "yes / no"]],
            caption="Reset log entry")
    return d


def checklist_release() -> Document:
    d = Document("Autonomy Release Checklist", "checklist",
                 subtitle="Gate checklist for releasing a firmware build to the pilot fleet",
                 doc_id="ASI-FS-CHK-003", revision="0.1", status="Draft", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:1])
    _std(d)
    d.h1("How to use")
    d.p("Every item must be checked by the named role before the release note is signed. Items marked *n/a* "
        "need a one-line justification in the comments column.")
    d.h1("Build")
    d.tasks([("Build is tagged and reproducible from the release branch", True),
             ("Static analysis clean for safety-related modules", True),
             ("Unit test coverage >= 90 % on safety PLC logic", False)])
    d.h1("Verification")
    d.tasks([("All L1 HIL tests passed", False), ("POST campaign report approved", True),
             ("Stopping distance verified on grade", True), ("Regression suite green on unit A02", False)])
    d.h1("Documentation")
    d.tasks([("SRS revision matches build", False), ("ICD revision matches build", True),
             ("Release notes drafted", False), ("Safety case impact assessed", False)])
    d.h1("Sign-off")
    d.table(["Role", "Name", "Date", "Signature"],
            [["FuSa Manager", "", "", ""], ["Autonomy Lead", "", "", ""], ["Site Ops", "", "", ""]])
    return d


# ------------------------------------------------------------ minutes -----
def minutes_safety_review() -> Document:
    d = Document("Safety Review Meeting Minutes", "minutes",
                 subtitle="Weekly FuSa review - week 44",
                 doc_id="ASI-FS-MIN-044", revision="1.0", status="Released", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal")
    d.revision("1.0", DATE, "Roy", "Issued")
    d.h1("Attendance")
    d.table(["Name", "Role", "Present"],
            [["Max Stone", "FuSa Manager (chair)", "yes"], ["Ida Moss", "Safety Mentor", "yes"],
             ["Roy", "FuSa Engineer (minutes)", "yes"], ["Autonomy Lead", "Autonomy", "yes"],
             ["Hydraulics Lead", "Hydraulics", "apologies"]])
    d.h1("Agenda")
    d.numbered_list(["Review of HARA v0.4 ratings", "DFMEA brake path - open RPN items",
                     "Stopping distance results", "AOB"])
    d.h1("Discussion and decisions")
    d.h2("HARA v0.4")
    d.p("The re-rated exposure for H-09 (dust) was accepted. **Decision:** HARA v0.4 goes to approval; "
        "safety goals SG-01 to SG-04 are baselined.")
    d.h2("DFMEA brake path")
    d.p("Two items exceed RPN 100. Hydraulics lead to confirm the dual pressure sensor option by next review. "
        "**Decision:** CR-118 and CR-119 opened.")
    d.h2("Stopping distance")
    d.p("24 runs, worst case 1.17 m against a 2.0 m limit. **Decision:** TSR-02 verified; report to be "
        "witnessed by the safety mentor.")
    d.h1("Actions")
    d.table(["#", "Action", "Owner", "Due", "Status"],
            [["A-44-1", "Approve HARA v0.4", "Max Stone", "2026-10-31", "open"],
             ["A-44-2", "Dual pressure sensor proposal (CR-119)", "Hydraulics Lead", "2026-11-05", "open"],
             ["A-44-3", "Witness stopping distance report", "Ida Moss", "2026-10-30", "open"],
             ["A-43-2", "Publish ICD 1.0", "Roy", "2026-10-27", "done"]],
            caption="Action register")
    d.h1("Next meeting")
    d.p("Tuesday 2026-11-03, 09:00, Teams.")
    return d


def memo_change_notice() -> Document:
    d = Document("Engineering Change Notice: Brake Pressure Warm-up", "memo",
                 subtitle="CR-121 - raise POST T4 sensor warm-up from 200 ms to 500 ms",
                 doc_id="ASI-ECN-121", revision="1.0", status="Approved", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:2])
    d.revision("1.0", DATE, "Roy", "Approved at safety review")
    d.h1("Change")
    d.p("Increase the brake pressure sensor warm-up window in the POST sequence from **200 ms** to **500 ms**. "
        "Firmware 2.3.0 -> 2.3.1, module `post_brake.c`, constant `BRAKE_WARMUP_MS`.")
    d.h1("Reason")
    d.p("Two cold-start POST failures on unit A02 (see ASI-FS-TR-021). The sensor output settles at 380-420 ms "
        "below 4 C ambient.")
    d.h1("Impact assessment")
    d.table(["Area", "Impact", "Assessment"],
            [["HARA", "none", "No new hazard; longer POST does not affect operation"],
             ["Safety requirements", "none", "No requirement bounds the POST duration"],
             ["Safety PLC", "none", "POST runs on the autonomy computer"],
             ["Verification", "re-test", "TC-POST-04c added, passed 50/50"],
             ["Documentation", "update", "POST report, release notes"]],
            caption="Impact matrix")
    d.h1("Approval")
    d.tasks([("Impact assessment reviewed", True), ("Re-test complete", True), ("Release notes updated", True)])
    return d


def release_notes_fw() -> Document:
    d = Document("Autonomy Firmware 2.3.1 Release Notes", "release_notes",
                 subtitle="Pilot fleet release", doc_id="ASI-RN-2.3.1", revision="1.0", status="Released",
                 date=DATE, author=AUTHOR, project=PROJECT, organisation=ORG, classification="Internal",
                 toc=False)
    d.revision("1.0", DATE, "Roy", "Issued")
    d.h1("Highlights")
    d.bullets(["POST brake pressure warm-up raised to 500 ms (CR-121)",
               "Safety status frame 0x101 now carries the active field set",
               "Heartbeat counter discipline enforced per ICD 1.0"])
    d.h1("Changes")
    d.table(["ID", "Type", "Description", "Safety relevant"],
            [["CR-121", "fix", "POST T4 warm-up 200 -> 500 ms", "yes"], ["CR-115", "feature", "field_set in 0x101", "yes"],
             ["CR-109", "fix", "GNSS time sync drift on resume", "no"], ["CR-117", "chore", "Logging buffer 4 -> 8 MB", "no"]],
            caption="Change list")
    d.h1("Known issues")
    d.bullets(["Safety HMI shows stale fault code for up to 300 ms after reset (cosmetic)",
               "Radar heartbeat test reports 96 ms mean, spec says <= 100 ms: margin is small"])
    d.h1("Upgrade")
    d.code("asi-flash --unit A02 --image autonomy-2.3.1.bin --verify\nasi-post --unit A02 --runs 3", "bash")
    return d


# -------------------------------------------------------- safety case -----
def safety_case_outline() -> Document:
    d = Document("Safety Case", "safety_case",
                 subtitle="Argument and evidence that the WA485 autonomy kit is acceptably safe for pilot operation",
                 doc_id="ASI-FS-SC-050", revision="0.1", status="Draft", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS)
    _std(d, "Argument structure drafted; evidence links to be populated")
    d.h1("Top-level claim")
    d.quote("G1: The WA485 autonomy kit is acceptably safe to operate autonomously at the pilot quarry under "
            "the operating concept in ASI-FS-ID-001.")
    d.h1("Argument structure")
    d.p("The argument follows a goal-structuring style: goals (**G**), strategies (**S**), context (**C**) and "
        "solutions / evidence (**Sn**).")
    d.bullets([("G1 Acceptably safe for pilot operation",
                [("S1 Argue over each safety goal from the HARA",
                  [("G2 SG-01 unintended motion is prevented", ["Sn1 TC-11, TC-12 results", "Sn2 TSR-01 inspection record"]),
                   ("G3 SG-02 stops before occupied zones", ["Sn3 TC-31/32 results", "Sn4 ASI-FS-TR-033 stopping distance"]),
                   ("G4 SG-03 speed limited when degraded", ["Sn5 TC-21 results"]),
                   ("G5 SG-04 implement interlock", ["Sn6 TC-41 results (planned)"])]),
                 ("S2 Argue the process was followed", ["Sn7 Confirmation review records", "Sn8 Functional safety audit report"]),
                 "C1 Operating concept ASI-FS-ID-001", "C2 Site traffic management plan"])])
    d.diagram(ESCHER["state_autonomy_modes"](), "Autonomy mode state machine referenced by G2 and G4 (asi_escher)")
    d.h1("Evidence register")
    d.table(["Evidence", "Document", "Status", "Supports"],
            [["Sn1", "TC-11, TC-12 reports", "planned", "G2"], ["Sn2", "Relay inspection record", "planned", "G2"],
             ["Sn3", "TC-31, TC-32 reports", "planned", "G3"], ["Sn4", "ASI-FS-TR-033", "in review", "G3"],
             ["Sn5", "TC-21 report", "planned", "G4"], ["Sn6", "TC-41 report", "planned", "G5"],
             ["Sn7", "Confirmation review records", "partial", "S2"], ["Sn8", "FuSa audit report", "planned", "S2"]],
            caption="Evidence register")
    d.h1("Residual risk and limitations")
    d.warning("The argument for SG-04 depends on TC-41, which has not yet been run. Until it passes, implement "
              "motion remains disabled in autonomous mode by configuration.")
    d.h1("Conclusion")
    d.p("Pending the evidence listed as *planned*, the structure of the argument is complete and has been "
        "reviewed by the safety mentor. The case will be re-issued at revision 1.0 for the M4 gate.")
    return d


# ------------------------------------------------------------- stress -----
def stress_long_document(n_sections: int = 36, seed: int = 7) -> Document:
    rng = random.Random(seed)
    d = Document(f"Stress Test: {n_sections}-section Generated Report", "report",
                 subtitle="Synthetic content exercising every block type at volume",
                 doc_id="ASI-FS-STRESS-999", revision="0.1", status="Draft", date=DATE, author=AUTHOR,
                 project=PROJECT, organisation=ORG, classification="Internal", approvers=APPROVERS[:1])
    _std(d)
    verbs = ["Validate", "Fuse", "Filter", "Estimate", "Check", "Plan", "Log", "Encode", "Arbitrate", "Schedule"]
    objs = ["pose", "zones", "payload", "tyres", "CAN bus", "lidar", "radar", "brakes", "steering", "map", "mission"]
    words = ("the safety PLC evaluates each input every cycle and the autonomy computer replans when the "
             "watchdog flags a late heartbeat while the operator observes the HMI for faults and the fleet "
             "server records every mission event for later analysis").split()
    for i in range(1, n_sections + 1):
        d.h1(f"{rng.choice(verbs)} {rng.choice(objs)}")
        for j in range(rng.randint(1, 3)):
            d.h2(f"{rng.choice(verbs)} {rng.choice(objs)} detail {j + 1}")
            n = rng.randint(25, 70)
            txt = " ".join(rng.choice(words) for _ in range(n))
            d.p(txt[0].upper() + txt[1:] + ".")
            roll = rng.random()
            if roll < 0.3:
                d.table(["Signal", "Value", "Unit", "Status"],
                        [[f"{rng.choice(objs)}_{k}", f"{rng.uniform(0, 100):.1f}", rng.choice(["ms", "bar", "m", "Hz"]),
                          rng.choice(["ok", "warn", "fault"])] for k in range(rng.randint(2, 7))],
                        caption=f"Generated table {i}.{j + 1}", align=["left", "right", "left", "center"])
            elif roll < 0.5:
                d.bullets([f"{rng.choice(verbs)} {rng.choice(objs)}" for _ in range(rng.randint(2, 5))])
            elif roll < 0.6:
                d.callout(rng.choice(["note", "tip", "important", "warning", "caution"]),
                          " ".join(rng.choice(words) for _ in range(18)) + ".")
            elif roll < 0.7:
                d.requirement(f"GEN-{i:02d}{j}", f"{rng.choice(verbs)} {rng.choice(objs)}",
                              "The system shall " + " ".join(rng.choice(words) for _ in range(12)) + ".",
                              plr=rng.choice("bcd"), status=rng.choice(["Draft", "Approved"]))
            elif roll < 0.8:
                d.code("\n".join(f"{rng.choice(verbs).lower()}_{rng.choice(objs).replace(' ', '_')}({rng.randint(0, 99)})"
                                 for _ in range(rng.randint(2, 6))), "python")
            elif roll < 0.9:
                d.numbered_list([f"Step {k + 1}: {rng.choice(verbs).lower()} {rng.choice(objs)}" for k in range(rng.randint(3, 6))])
    return d


DEMOS = [
    ("spec_safety_requirements", spec_safety_requirements),
    ("icd_safety_plc", icd_safety_plc),
    ("report_hara", report_hara),
    ("report_fmea_brake", report_fmea_brake),
    ("report_post_test", report_post_test),
    ("report_stopping_distance", report_stopping_distance),
    ("plan_functional_safety", plan_functional_safety),
    ("plan_verification_validation", plan_verification_validation),
    ("procedure_estop_reset", procedure_estop_reset),
    ("checklist_release", checklist_release),
    ("minutes_safety_review", minutes_safety_review),
    ("memo_change_notice", memo_change_notice),
    ("release_notes_fw", release_notes_fw),
    ("safety_case_outline", safety_case_outline),
    ("stress_long_document", stress_long_document),
]
