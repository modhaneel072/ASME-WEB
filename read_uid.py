from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import NoCardException

print("=== read_uid.py starting ===", flush=True)

GET_UID = [0xFF, 0xCA, 0x00, 0x00, 0x00]

r = readers()
print("Readers:", r, flush=True)
if not r:
    raise SystemExit("No PC/SC readers found.")

reader = r[0]
print("Using reader:", reader, flush=True)

conn = reader.createConnection()
print("Tap/hold a tag on the reader NOW...", flush=True)

try:
    conn.connect()
    data, sw1, sw2 = conn.transmit(GET_UID)
    print("SW1 SW2:", hex(sw1), hex(sw2), flush=True)

    if (sw1, sw2) == (0x90, 0x00):
        print("UID:", toHexString(data).replace(" ", ""), flush=True)
    else:
        print("UID read failed (SW not 0x9000).", flush=True)

except NoCardException:
    print("No tag present on reader.", flush=True)

print("=== read_uid.py done ===", flush=True)
