# tabletxt_parser.py

from openpyxl import Workbook
import os
from helpers import hex_byte, time_to_seconds, to_signed_int16
import re
from datetime import datetime

def parse_tabletxt(input_file):
    filename = os.path.basename(input_file)

    file_date = None

    # Try YYYY-MM-DD
    match = re.search(r"\d{4}-\d{2}-\d{2}", filename)
    if match:
        file_date = match.group()

    # Try YYYYMMDD
    if not file_date:
        match = re.search(r"\d{8}", filename)
        if match:
            raw = match.group()
            file_date = datetime.strptime(raw, "%Y%m%d").strftime("%Y-%m-%d")

    # Try DD-MM-YYYY
    if not file_date:
        match = re.search(r"\d{2}-\d{2}-\d{4}", filename)
        if match:
            raw = match.group()
            file_date = datetime.strptime(raw, "%d-%m-%Y").strftime("%Y-%m-%d")

    # Fallback to file modified date
    if not file_date:
        timestamp = os.path.getmtime(input_file)
        file_date = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d")
    output_excel = os.path.join(
        os.path.dirname(input_file),
        os.path.splitext(os.path.basename(input_file))[0] + "_parsed.xlsx"
    )

    headers = (
    ["Date", "Time", "Time(Sec)"] +
    [f"Cell{i}" for i in range(1, 25)] +
    ["Current", "Capacity", "SOC"] +
    [f"T{i}" for i in range(1, 15)] +
    ["IG_Status"] +
    [f"IB{i}" for i in range(1, 29)] +
    ["SwVMajor", "SwVMinor", "SwVSub"] +
    ["ActiveFaults", "VehicleState"] +
    ["SerialNumber"]
)

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Parsed_Data"

    # Write header
    ws.append(headers)
    first_time = None
    current_row = None
    row_count = 0
    
    with open(input_file, "r", errors="ignore") as fin:

        for line in fin:
            line = line.strip()

            if not line or line.startswith("No"):
                continue

            parts = line.split()
            if len(parts) < 10:
                continue

            # Find Frame ID dynamically
            frame_index = None
            for i, p in enumerate(parts):
                if p.lower().startswith("0x"):
                    frame_index = i
                    break

            if frame_index is None:
                continue

            time_str = parts[2] + ":" + parts[3]
            can_id = parts[frame_index].replace("0x", "").upper()

            # Data bytes start after DLC
            data_index = frame_index + 2
            data_bytes = parts[data_index:]

            if len(data_bytes) < 8:
                continue

            b = [hex_byte(x) for x in data_bytes[:8]]

            # ---------------- NEW ROW ----------------
            if can_id == "041FFAEA":
                if current_row:
                    ws.append(current_row)
                    row_count += 1

                current_row = [None] * 79
                current_row[0] = file_date
                current_row[1] = time_str
                t_sec = time_to_seconds(time_str)
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

            # ---------------- CELLS ----------------
            if can_id == "051FFAEA":
                current_row[7:10] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "061FFAEA":
                current_row[11:14] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "071FFAEA":
                current_row[15:18] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "0420FAEA":
                current_row[19:22] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            elif can_id == "0620FAEA":
                current_row[23:26] = [(b[i+1] << 8 | b[i]) / 1000 for i in range(0, 8, 2)]

            # ---------------- CURRENT / SOC ----------------
            elif can_id == "821FAEA":
                current_row[27] = to_signed_int16(b[1] << 8 | b[0]) / 10
                current_row[29] = b[6]

            elif can_id == "E14FBEB":
                current_row[28] = (b[7] << 8 | b[6]) / 100

            # ---------------- TEMPERATURES ----------------
            elif can_id in ("1422FAEA", "1424FAEA", "1425FAEA"):
                base_col = {
                    "1422FAEA": 30,
                    "1424FAEA": 34,
                    "1425FAEA": 38,
                }[can_id]

                for i in range(0, 8, 2):
                    current_row[base_col + i // 2] = (b[i+1] << 8 | b[i]) / 100

            elif can_id == "1426FAEA":
                current_row[42] = (b[1] << 8 | b[0]) / 100
                current_row[43] = (b[3] << 8 | b[2]) / 100

            # ---------------- IB ----------------
            elif can_id in ("1402FAEA", "1502FAEA", "1603FAEA"):
                base_col = {
                    "1402FAEA": 45,
                    "1502FAEA": 53,
                    "1603FAEA": 61
                }[can_id]

                for i in range(8):
                    current_row[base_col + i] = b[i] / 100

            elif can_id == "1702FAEA":
                for i in range(4):
                    current_row[69 + i] = b[i] / 10

            # ---------------- SOFTWARE VERSION ----------------
            elif can_id == "1A14FBEB":
                current_row[73] = b[1]
                current_row[74] = b[2]
                current_row[75] = b[0]

            # ---------------- FAULT + STATE ----------------
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

                current_row[44] = "IGOFF" if vehicle_state_byte == 0x00 else "IGON"
                current_row[77] = vehicle_state_map.get(
                    vehicle_state_byte,
                    f"UNKNOWN(0x{vehicle_state_byte:02X})"
                )

            # ---------------- SERIAL ----------------
            elif can_id == "1914EAFA":
                serial_number = (b[6] << 8) | b[7]
                current_row[78] = f"{chr(b[0])}{b[1]:02X}{b[2]:02d}{b[3]}{b[4]:X}{b[5]:02}0{serial_number:04d}"

            if row_count and row_count % 5000 == 0:
                print(f"Parsed {row_count} rows...")

        # Write last row
        if current_row:
            ws.append(current_row)
            row_count += 1

    wb.save(output_excel)

    print("\nParsing completed successfully.")
    print(f"Total rows written: {row_count}")
    print(f"Output saved at:\n{output_excel}")