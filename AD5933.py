import ctypes
from ctypes import wintypes
import csv
import math
import time

# Use WinDLL because the vendor DLL is expected to use stdcall on Windows.
DLL_PATH = r"C:\Program Files\Analog Devices\USB Drivers\ADI_CYUSB_USB4.dll"
dll = ctypes.WinDLL(DLL_PATH)

# USB request constants used by the ADI GUI.
REQUEST = 0xDE
VALUE = 0x000D

# Device defaults for EVAL-AD5933.
DEFAULT_VID = 0x0456
DEFAULT_PID = 0xB203
DEFAULT_MCLK_HZ = 4_000_000.0

# AD5933 register addresses.
REG_CTRL_HIGH = 0x80
REG_CTRL_LOW = 0x81
REG_START_MSB = 0x82
REG_START_MID = 0x83
REG_START_LSB = 0x84
REG_DELTA_MSB = 0x85
REG_DELTA_MID = 0x86
REG_DELTA_LSB = 0x87
REG_INCR_MSB = 0x88
REG_INCR_LSB = 0x89
REG_SETTLE_MSB = 0x8A
REG_SETTLE_LSB = 0x8B
REG_STATUS = 0x8F
REG_TEMP_MSB = 0x92
REG_TEMP_LSB = 0x93
REG_REAL_MSB = 0x94
REG_REAL_LSB = 0x95
REG_IMAG_MSB = 0x96
REG_IMAG_LSB = 0x97

# Status bits.
STAT_TEMP_VALID = 0x01
STAT_DATA_VALID = 0x02
STAT_SWEEP_DONE = 0x04


# Bind C signatures once to keep function bodies clean.
dll.Search_For_Boards.restype = wintypes.DWORD
dll.Search_For_Boards.argtypes = [
    wintypes.DWORD,
    wintypes.DWORD,
    ctypes.POINTER(wintypes.DWORD),
    wintypes.CHAR * 16,
]

dll.Connect.restype = wintypes.INT
dll.Connect.argtypes = [
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.CHAR,
    ctypes.POINTER(wintypes.DWORD),
]

dll.Disconnect.restype = wintypes.INT
dll.Disconnect.argtypes = [wintypes.DWORD]

dll.Vendor_Request.restype = wintypes.INT
dll.Vendor_Request.argtypes = [
    wintypes.DWORD,
    wintypes.BYTE,
    wintypes.WORD,
    wintypes.WORD,
    wintypes.BYTE,
    wintypes.WORD,
    ctypes.POINTER(wintypes.BYTE),
]


def search_for_boards(vid=DEFAULT_VID, pid=DEFAULT_PID, max_boards=16):
    """Return (result, count, board_indices)."""
    num_boards = wintypes.DWORD(0)
    path_array = (wintypes.CHAR * max_boards)()

    result = dll.Search_For_Boards(vid, pid, ctypes.byref(num_boards), path_array)

    indices = []
    for i in range(num_boards.value):
        raw = path_array[i]
        indices.append(raw[0] if isinstance(raw, (bytes, bytearray)) else int(raw))

    return result, num_boards.value, indices


def connect(vid, pid, board_index):
    """Return (result, handle)."""
    handle = wintypes.DWORD(0)
    result = dll.Connect(vid, pid, int(board_index), ctypes.byref(handle))
    return result, handle.value


def disconnect(handle):
    """Close device handle."""
    return dll.Disconnect(handle)


def vendor_request(handle, request, value, index, direction, data_length, buffer=None):
    """Low-level USB vendor request wrapper.

    direction: 0 for OUT(write), 1 for IN(read)
    """
    if buffer is None:
        buffer = (wintypes.BYTE * max(1, data_length))()

    result = dll.Vendor_Request(
        handle,
        request,
        value,
        index,
        direction,
        data_length,
        buffer,
    )
    return result, buffer


def read_register_byte(handle, reg_addr):
    """Read one AD5933 register byte."""
    buf = (wintypes.BYTE * 1)()
    result, buf = vendor_request(
        handle,
        REQUEST,
        VALUE,
        reg_addr & 0xFF,
        1,
        1,
        buf,
    )
    if result != 0:
        raise RuntimeError(f"Read reg 0x{reg_addr:02X} failed, result={result}")
    return buf[0] & 0xFF


def write_register_byte(handle, reg_addr, reg_value):
    """Write one AD5933 register byte.

    For write, wIndex packs value/address as 0xVVAA.
    """
    index = ((reg_value & 0xFF) << 8) | (reg_addr & 0xFF)
    result, _ = vendor_request(handle, REQUEST, VALUE, index, 0, 0, None)
    if result != 0:
        raise RuntimeError(
            f"Write reg 0x{reg_addr:02X}=0x{reg_value:02X} failed, result={result}"
        )
    return result


def combine24(msb, mid, lsb):
    return ((msb & 0xFF) << 16) | ((mid & 0xFF) << 8) | (lsb & 0xFF)


def to_signed16(msb, lsb):
    value = ((msb & 0xFF) << 8) | (lsb & 0xFF)
    return value - 0x10000 if value & 0x8000 else value


def code_to_freq(code, mclk_hz):
    return float(code) * float(mclk_hz) / float(1 << 27)


def read_configuration(handle, mclk_hz=DEFAULT_MCLK_HZ):
    """Read and decode current sweep-related configuration."""
    regs = {}
    addrs = [
        REG_CTRL_HIGH,
        REG_CTRL_LOW,
        REG_START_MSB,
        REG_START_MID,
        REG_START_LSB,
        REG_DELTA_MSB,
        REG_DELTA_MID,
        REG_DELTA_LSB,
        REG_INCR_MSB,
        REG_INCR_LSB,
        REG_SETTLE_MSB,
        REG_SETTLE_LSB,
        REG_STATUS,
    ]
    for addr in addrs:
        regs[addr] = read_register_byte(handle, addr)

    start_code = combine24(regs[REG_START_MSB], regs[REG_START_MID], regs[REG_START_LSB])
    delta_code = combine24(regs[REG_DELTA_MSB], regs[REG_DELTA_MID], regs[REG_DELTA_LSB])
    n_incr = regs[REG_INCR_LSB]
    n_points = n_incr + 1

    settling_cycles = regs[REG_SETTLE_LSB] if regs[REG_SETTLE_MSB] == 0 else None

    return {
        "raw_registers": regs,
        "codes": {
            "start_code": start_code,
            "delta_code": delta_code,
            "incr_code": combine24(regs[REG_INCR_MSB], regs[REG_INCR_LSB], regs[REG_SETTLE_MSB]),
        },
        "params": {
            "mclk_hz": float(mclk_hz),
            "start_freq_hz": code_to_freq(start_code, mclk_hz),
            "delta_freq_hz": code_to_freq(delta_code, mclk_hz),
            "num_increments": n_incr,
            "num_points": n_points,
            "settling_cycles": settling_cycles,
        },
        "control_status": {
            "ctrl_high_0x80": regs[REG_CTRL_HIGH],
            "ctrl_low_0x81": regs[REG_CTRL_LOW],
            "status_0x8F": regs[REG_STATUS],
        },
    }


def run_single_measurement(handle):
    """Trigger one DFT conversion at current start frequency and read Real/Imag."""
    ctrl_high = read_register_byte(handle, REG_CTRL_HIGH)
    low_nibble = ctrl_high & 0x0F

    standby_val = (0xB << 4) | low_nibble
    init_val = (0x1 << 4) | low_nibble
    start_val = (0x2 << 4) | low_nibble

    write_register_byte(handle, REG_CTRL_HIGH, standby_val)
    write_register_byte(handle, REG_CTRL_HIGH, init_val)
    write_register_byte(handle, REG_CTRL_HIGH, start_val)

    status = 0
    for _ in range(200):
        status = read_register_byte(handle, REG_STATUS)
        if status & STAT_DATA_VALID:
            break
        time.sleep(0.005)
    else:
        raise RuntimeError("Timeout waiting DATA_VALID during single measurement")

    real = to_signed16(
        read_register_byte(handle, REG_REAL_MSB),
        read_register_byte(handle, REG_REAL_LSB),
    )
    imag = to_signed16(
        read_register_byte(handle, REG_IMAG_MSB),
        read_register_byte(handle, REG_IMAG_LSB),
    )
    mag = math.sqrt(float(real * real + imag * imag))

    return {
        "status": status,
        "real": real,
        "imag": imag,
        "magnitude": mag,
    }


def read_temperature(handle):
    """Run temperature conversion and return Celsius."""
    ctrl_high = read_register_byte(handle, REG_CTRL_HIGH)
    low_nibble = ctrl_high & 0x0F
    temp_cmd = (0x9 << 4) | low_nibble

    write_register_byte(handle, REG_CTRL_HIGH, temp_cmd)

    status = 0
    for _ in range(100):
        status = read_register_byte(handle, REG_STATUS)
        if status & STAT_TEMP_VALID:
            break
        time.sleep(0.005)

    temp_raw = ((read_register_byte(handle, REG_TEMP_MSB) & 0xFF) << 8) | (
        read_register_byte(handle, REG_TEMP_LSB) & 0xFF
    )
    if temp_raw & 0x2000:
        temp_raw -= 0x4000

    return temp_raw / 32.0


def calibrate_gain(handle, z_ref_ohm=5000.0, mclk_hz=DEFAULT_MCLK_HZ, cal_point="mid"):
    """Compute gain factor at start/mid/end point of current sweep settings."""
    cfg = read_configuration(handle, mclk_hz=mclk_hz)
    start_code = cfg["codes"]["start_code"]
    delta_code = cfg["codes"]["delta_code"]
    n_incr = cfg["params"]["num_increments"]

    if cal_point == "start":
        k = 0
    elif cal_point == "mid":
        k = n_incr // 2
    elif cal_point == "end":
        k = n_incr
    else:
        raise ValueError("cal_point must be one of: start, mid, end")

    cal_code = start_code + k * delta_code
    write_register_byte(handle, REG_START_MSB, (cal_code >> 16) & 0xFF)
    write_register_byte(handle, REG_START_MID, (cal_code >> 8) & 0xFF)
    write_register_byte(handle, REG_START_LSB, cal_code & 0xFF)

    meas = run_single_measurement(handle)
    mag = meas["magnitude"]

    if mag <= 0 or z_ref_ohm <= 0:
        gain_factor = None
    else:
        gain_factor = 1.0 / (float(z_ref_ohm) * mag)

    return {
        "config": cfg,
        "cal_point": cal_point,
        "cal_point_index": k,
        "cal_code": cal_code,
        "cal_freq_hz": code_to_freq(cal_code, mclk_hz),
        "real": meas["real"],
        "imag": meas["imag"],
        "magnitude": mag,
        "gain_factor": gain_factor,
        "z_ref_ohm": float(z_ref_ohm),
    }


def write_sweep_config(handle, cfg):
    """Program sweep registers from a configuration dict."""
    start_code = cfg["codes"]["start_code"]
    delta_code = cfg["codes"]["delta_code"]
    n_incr = cfg["params"]["num_increments"]
    settling_cycles = cfg["params"]["settling_cycles"]

    write_register_byte(handle, REG_START_MSB, (start_code >> 16) & 0xFF)
    write_register_byte(handle, REG_START_MID, (start_code >> 8) & 0xFF)
    write_register_byte(handle, REG_START_LSB, start_code & 0xFF)

    write_register_byte(handle, REG_DELTA_MSB, (delta_code >> 16) & 0xFF)
    write_register_byte(handle, REG_DELTA_MID, (delta_code >> 8) & 0xFF)
    write_register_byte(handle, REG_DELTA_LSB, delta_code & 0xFF)

    write_register_byte(handle, REG_INCR_MSB, 0x00)
    write_register_byte(handle, REG_INCR_LSB, n_incr & 0xFF)

    if settling_cycles is not None:
        write_register_byte(handle, REG_SETTLE_MSB, 0x00)
        write_register_byte(handle, REG_SETTLE_LSB, settling_cycles & 0xFF)


def sweep(handle, cfg, mclk_hz=DEFAULT_MCLK_HZ, gain_factor=None, csv_path=None):
    """Run full frequency sweep and optionally save CSV."""
    write_sweep_config(handle, cfg)

    start_code = cfg["codes"]["start_code"]
    delta_code = cfg["codes"]["delta_code"]
    n_points = cfg["params"]["num_points"]

    ctrl_high = read_register_byte(handle, REG_CTRL_HIGH)
    low_nibble = ctrl_high & 0x0F

    standby_val = (0xB << 4) | low_nibble
    init_val = (0x1 << 4) | low_nibble
    start_val = (0x2 << 4) | low_nibble
    inc_val = (0x3 << 4) | low_nibble

    write_register_byte(handle, REG_CTRL_HIGH, standby_val)
    write_register_byte(handle, REG_CTRL_HIGH, init_val)
    write_register_byte(handle, REG_CTRL_HIGH, start_val)

    results = []

    for i in range(n_points):
        status = 0
        for _ in range(200):
            status = read_register_byte(handle, REG_STATUS)
            if status & STAT_DATA_VALID:
                break
            time.sleep(0.005)
        else:
            raise RuntimeError(f"Timeout waiting DATA_VALID at point {i}")

        real = to_signed16(
            read_register_byte(handle, REG_REAL_MSB),
            read_register_byte(handle, REG_REAL_LSB),
        )
        imag = to_signed16(
            read_register_byte(handle, REG_IMAG_MSB),
            read_register_byte(handle, REG_IMAG_LSB),
        )
        mag = math.sqrt(float(real * real + imag * imag))

        freq_code = start_code + i * delta_code
        freq_hz = code_to_freq(freq_code, mclk_hz)

        z_abs = None
        if gain_factor is not None and mag > 0:
            z_abs = 1.0 / (float(gain_factor) * mag)

        results.append(
            {
                "index": i,
                "freq_hz": freq_hz,
                "real": real,
                "imag": imag,
                "magnitude": mag,
                "z_abs_ohm": z_abs,
            }
        )

        if (status & STAT_SWEEP_DONE) or i == n_points - 1:
            break

        write_register_byte(handle, REG_CTRL_HIGH, inc_val)

    if csv_path:
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["index", "freq_hz", "real", "imag", "magnitude", "z_abs_ohm"])
            for row in results:
                writer.writerow(
                    [
                        row["index"],
                        f"{row['freq_hz']:.6f}",
                        row["real"],
                        row["imag"],
                        f"{row['magnitude']:.6f}",
                        "" if row["z_abs_ohm"] is None else f"{row['z_abs_ohm']:.6f}",
                    ]
                )

    return results


def print_configuration(cfg):
    """Pretty print sweep configuration."""
    c = cfg["codes"]
    p = cfg["params"]
    s = cfg["control_status"]

    print("=== Configuration ===")
    print(f"Start code: 0x{c['start_code']:06X}  -> {p['start_freq_hz']:.6f} Hz")
    print(f"Delta code: 0x{c['delta_code']:06X}  -> {p['delta_freq_hz']:.6f} Hz")
    print(f"Increments: {p['num_increments']}  -> points: {p['num_points']}")
    print(f"Settling cycles: {p['settling_cycles']}")
    print(f"Ctrl 0x80: 0x{s['ctrl_high_0x80']:02X}")
    print(f"Ctrl 0x81: 0x{s['ctrl_low_0x81']:02X}")
    print(f"Status 0x8F: 0x{s['status_0x8F']:02X}")


def main():
    """Show all core features in one CLI demo."""
    print("AD5933 Minimal Tool")
    print("1) Search boards")
    print("2) Connect board")
    print("3) Low-level read/write demo")
    print("4) Read configuration")
    print("5) Read temperature")
    print("6) Calibrate gain")
    print("7) Sweep and optional CSV")

    result, count, indices = search_for_boards(DEFAULT_VID, DEFAULT_PID)
    print(f"Search result={result}, boards={count}, indices={indices}")
    if count == 0:
        print("No AD5933 board found.")
        return

    board_input = input("Choose board index from list (default first): ").strip()
    if board_input:
        board_index = int(board_input)
    else:
        board_index = indices[0]

    result, handle = connect(DEFAULT_VID, DEFAULT_PID, board_index)
    print(f"Connect result={result}, handle={handle}")
    if result != 0:
        return

    try:
        # Low-level communication demo: read control register and write back the same value.
        ctrl_before = read_register_byte(handle, REG_CTRL_HIGH)
        print(f"Low-level read: reg 0x80 = 0x{ctrl_before:02X}")
        write_register_byte(handle, REG_CTRL_HIGH, ctrl_before)
        ctrl_after = read_register_byte(handle, REG_CTRL_HIGH)
        print(f"Low-level write/readback: reg 0x80 = 0x{ctrl_after:02X}")

        cfg = read_configuration(handle, mclk_hz=DEFAULT_MCLK_HZ)
        print_configuration(cfg)

        temp_c = read_temperature(handle)
        print(f"Temperature: {temp_c:.2f} C")

        z_ref_str = input("Reference resistor for calibration in ohm (default 5000): ").strip()
        z_ref = float(z_ref_str) if z_ref_str else 5000.0
        cal_point = input("Calibration point [start/mid/end] (default mid): ").strip().lower() or "mid"
        if cal_point not in ("start", "mid", "end"):
            cal_point = "mid"

        cal = calibrate_gain(
            handle,
            z_ref_ohm=z_ref,
            mclk_hz=DEFAULT_MCLK_HZ,
            cal_point=cal_point,
        )
        print(
            "Calibration: "
            f"k={cal['cal_point_index']}, f={cal['cal_freq_hz']:.3f} Hz, "
            f"Real={cal['real']}, Imag={cal['imag']}, |DFT|={cal['magnitude']:.3f}, "
            f"gain_factor={cal['gain_factor']}"
        )

        csv_default = f"eit_sweep_{int(time.time())}.csv"
        csv_path = input(f"CSV output path (default {csv_default}): ").strip() or csv_default

        # Reuse pre-calibration config so sweep follows original GUI setup.
        results = sweep(
            handle,
            cfg=cal["config"],
            mclk_hz=DEFAULT_MCLK_HZ,
            gain_factor=cal["gain_factor"],
            csv_path=csv_path,
        )
        print(f"Sweep done: {len(results)} points, CSV saved to: {csv_path}")

    finally:
        result = disconnect(handle)
        print(f"Disconnect result={result}")


if __name__ == "__main__":
    main()
