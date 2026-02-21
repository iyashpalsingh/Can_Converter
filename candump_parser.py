# candump_parser.py
import os
from openpyxl import Workbook
from helpers import epoch_to_date_time, to_signed_int16, time_to_seconds

# ===================== PARSER CANDUMP =====================

def parse_can_dump(input_file):

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

    wb = Workbook()
    ws = wb.active
    ws.title = "Parsed_Data"
    ws.append(headers)
    current_row = None
    row_count = 0
    start_epoch = None
    current_row = None
    row_count = 0

    with open(input_file, "r", errors="ignore") as fin:

        for line in fin:
            line = line.strip()
            if not line or "#" not in line:
                continue

            try:
                epoch = line.split(")")[0][1:]
                _, frame = line.split(" ", 1)
                can_part = frame.split()[-1]

                can_id, data = can_part.split("#")
                can_id = can_id.upper().lstrip("0")

                if len(data) < 16:
                    continue

                b = [int(data[i:i+2], 16) for i in range(0, 16, 2)]
                date_str, time_str = epoch_to_date_time(epoch)

            except Exception:
                continue
            date_str, time_str = epoch_to_date_time(epoch)
            epoch = float(epoch)           
            # -------- NEW ROW --------
            if can_id == "41FFAEA":
                if current_row:
                    ws.append(current_row)
                    row_count += 1

                current_row = [None] * 79
                current_row[0] = date_str
                current_row[1] = time_str
                if start_epoch is None:
                    start_epoch = epoch
                    current_row[2] = 0
                else:
                    current_row[2] = round(epoch - start_epoch, 3)
                current_row[3] = (b[1] << 8 | b[0]) / 1000
                current_row[4] = (b[3] << 8 | b[2]) / 1000
                current_row[5] = (b[5] << 8 | b[4]) / 1000
                current_row[6] = (b[7] << 8 | b[6]) / 1000

            if not current_row:
                continue

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

        if current_row:
            ws.append(current_row)
            row_count += 1

    wb.save(output_excel)

    print("\nParsing completed successfully.")
    print(f"Total rows written: {row_count}")
    print(f"Output saved at:\n{output_excel}")