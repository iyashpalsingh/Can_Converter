#helpers.py
from datetime import datetime

def epoch_to_date_time(epoch):
    dt = datetime.fromtimestamp(float(epoch))
    return dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M:%S")

def to_signed_int16(raw):
    return raw - 65536 if raw >= 32768 else raw

def hex_byte(s):
    return int(s, 16)

def time_to_seconds(t):
    # 11:29:01:5203
    h, m, s, ms = t.split(":")
    return int(h)*3600 + int(m)*60 + int(s) + int(ms)/10000