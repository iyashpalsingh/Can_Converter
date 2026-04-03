# busmaster_parser.py
import os
from openpyxl import Workbook
from helpers import hex_byte, to_signed_int16, time_to_seconds

# ===================== PARSER BUSMASTER =====================

def parse_busmaster(input_file):

    output_excel = os.path.join(
        os.path.dirname(input_file),
        os.path.splitext(os.path.basename(input_file))[0] + "_parsed.xlsx"
    )

    #Adding column headers for excel file
    headers = (
        ["Date", "Time", "Time(Sec)"] +
        [f"Cell{i}" for i in range(1, 25)] +
        ["Current", "Capacity", "SOC"] +
        [f"T{i}" for i in range(1, 15)] +
        ["IG_Status"] +
        [f"IB{i}" for i in range(1, 29)] +
        ["SwVMajor", "SwVMinor", "SwVSub"] +
        ["ActiveFaults", "ActiveWarnings", "VehicleState"] +
        ["SerialNumber"]
    )

    # Create Excel with column headers
    wb = Workbook()
    ws = wb.active
    ws.title = "Parsed_Data"
    ws.append(headers)
    first_time = None
    start_date = None
    current_row = None
    row_count = 0

    with open(input_file, "r", errors="ignore") as fin:

        for line in fin:
            line = line.strip()
            if not line:
                continue
            
            #Creating timestamp for each row
            if start_date is None and "START DATE AND TIME" in line:
                start_date = (
                    line.replace("***START DATE AND TIME ", "")
                        .replace("***", "")
                        .split()[0]
                )
                continue

            parts = line.split()
            if len(parts) < 14:
                continue

            can_id = parts[3].replace("0x", "").upper()
            b = [hex_byte(x) for x in parts[6:14]]

            # -------- NEW ROW --------
            # For Cell 1 to cell 4
            if can_id == "41FFAEA":
                if current_row:
                    ws.append(current_row)
                    row_count += 1

                current_row = [None] * 80
                current_row[0] = start_date
                current_row[1] = parts[0]
                t_sec = time_to_seconds(parts[0])
                if first_time is None:
                    first_time = t_sec
                    current_row[2] = 0
                else:
                    current_row[2] = round(t_sec - first_time, 3)
                current_row[3] = (b[1] << 8 | b[0]) / 1000
                current_row[4] = (b[3] << 8 | b[2]) / 1000
                current_row[5] = (b[5] << 8 | b[4]) / 1000
                current_row[6] = (b[7] << 8 | b[6]) / 1000

            if not current_row:
                continue

            if can_id == "51FFAEA": # Cell5–8
                current_row[7:10] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "61FFAEA": # Cell9–12
                current_row[11:14] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "71FFAEA": # Cell13–16
                current_row[15:18] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "420FAEA": # Cell17–20
                current_row[19:22] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "620FAEA": # Cell21–24
                current_row[23:26] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            # Current, SOC
            elif can_id == "821FAEA":
                c1 = to_signed_int16(b[1] << 8 | b[0]) / 10
                current_row[27] = c1
                current_row[29] = b[6]

            # Capacity
            elif can_id == "E14FBEB":
                current_row[28] = (b[7] << 8 | b[6]) / 100

            # ---------------- TEMPERATURES ----------------
            elif can_id in ("1422FAEA", "1424FAEA", "1425FAEA"):
                base_col = {
                    "1422FAEA": 30, # Temp 1-4
                    "1424FAEA": 34, # Temp 5-8
                    "1425FAEA": 38, # Temp 9-12
                }[can_id]

                for i in range(0, 8, 2):
                    current_row[base_col + i // 2] = (b[i+1] << 8 | b[i]) / 100

            elif can_id == "1426FAEA": # Temp 13-14
                current_row[42] = (b[1] << 8 | b[0]) / 100
                current_row[43] = (b[3] << 8 | b[2]) / 100

            # ---------------- Imbalance ----------------
            elif can_id in ("1402FAEA", "1502FAEA", "1603FAEA"):
                base_col = {
                    "1402FAEA": 45, # Cell 1-8
                    "1502FAEA": 53, # Cell 9-16
                    "1603FAEA": 61 # Cell 17-24
                }[can_id]

                for i in range(8):
                    current_row[base_col + i] = b[i] / 100

            elif can_id == "1702FAEA": # Cell 25 onwards
                for i in range(4):
                    current_row[69 + i] = b[i] / 10

            # ---------------- SOFTWARE VERSION ----------------
            elif can_id == "1A14FBEB":
                current_row[73] = b[1]
                current_row[74] = b[2]
                current_row[75] = b[0]

            # ---------------- FAULT + STATE + WARNING ----------------
            elif can_id == "C23FAEA":
                fault_byte = b[0]
                vehicle_state_byte = b[2]

                fault_map = {
                    0: "E001",
                    1: "E002",
                    2: "E004",
                    3: "E008",
                    4: "E016",
                    5: "E032",
                    6: "E064",
                    7: "E128",
                }
                
                warning_map = {
                    0: "Temp Grad err",
                    1: "Voltage Grad err",
                    2: "Charger Timeout cutoff",
                    3: "Thermal Runaway",
                    4: "Shunt Offset Error",
                    5: "Watchdog Reset",
                    6: "Deep Discharge warning_1Day",
                    7: "Deep Discharge warning_3Day",
                }

                vehicle_state_map = {
                    0x00: "Idle",
                    0x01: "Discharge",
                    0x02: "Charge_0_EVQ",
                    0x03: "Balancing",
                    0x04: "Error",
                    0x05: "Charging_2_GBT",
                    0x06: "Charging_3_Solterra",
                    0x07: "Charging_Ather",
                    0x08: "Low_Power_Mode",
                    0x31: "Zivan_Charging",
                }

                active_faults = [
                    code for bit, code in fault_map.items()
                    if fault_byte & (1 << bit)
                ]
                
                if active_faults:
                    current_row[76] = ",".join(active_faults)

                warning_byte = b[1]

                active_warnings = [
                    code for bit, code in warning_map.items()
                    if warning_byte & (1 << bit)
                ]

                if active_warnings:
                    current_row[77] = ",".join(active_warnings)
                    
                current_row[44] = "IGOFF" if vehicle_state_byte == 0x00 else "IGON"
                current_row[78] = vehicle_state_map.get(
                    vehicle_state_byte,
                    f"UNKNOWN(0x{vehicle_state_byte:02X})"
                )

            # ---------------- SERIAL ----------------
            elif can_id == "1914EAFA":
                serial_number = (b[6] << 8) | b[7]
                current_row[79] = f"{chr(b[0])}{b[1]:02X}{b[2]:02d}{b[3]}{b[4]:X}{b[5]:02}0{serial_number:04d}"

            if row_count % 5000 == 0 and row_count > 0:
                print(f"Parsed {row_count} rows...")

        if current_row:
            ws.append(current_row)
            row_count += 1

    wb.save(output_excel)

    print("\nParsing completed successfully.")
    print(f"Total rows written: {row_count}")
    print(f"Output saved at:\n{output_excel}")