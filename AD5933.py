# py -3.10-32 .\hack_board.py

import ctypes
from ctypes import wintypes
import time
import math
import csv

# ADI_CYUSB_USB4.dll 很可能使用 Windows stdcall (__stdcall) 调用约定，
# 因此这里用 WinDLL 而不是 CDLL 以避免栈不一致导致参数解析异常。
dll = ctypes.WinDLL(r"C:\Program Files\Analog Devices\USB Drivers\ADI_CYUSB_USB4.dll")

# 全局 verbose 开关：True=详细步骤日志，False=只输出关键结果
VERBOSE = False


def vprint(*args, **kwargs):
    """受 VERBOSE 控制的调试输出。"""
    if VERBOSE:
        print(*args, **kwargs)


class EVAL_AD5933:
    """EVAL-AD5933 板卡高层封装。

    提供：
      - 自动搜索 / 连接
      - 读取配置
      - 增益校准（gain factor）
      - sweep 测量并导出 CSV

    用法示例：

        with EVAL_AD5933(vid=0x0456, pid=0xB203, mclk_hz=4_000_000.0) as dev:
            cfg = dev.get_configuration()
            cal = dev.calibrate_gain(z_ref_ohm=5000.0)
            dev.sweep_to_csv("sweep.csv", gain_factor=cal["gain_factor"])
    """

    def __init__(self, vid=0x0456, pid=0xB203, board_index=0, mclk_hz=4_000_000.0):
        self.vid = vid
        self.pid = pid
        self.board_index = board_index
        self.mclk_hz = float(mclk_hz)
        self.handle = None
        self.paths = None
        self._connect()

    # ---- 低层连接 / 断开 ----
    def _connect(self):
        result, count, paths = search_for_boards(self.vid, self.pid)
        if count == 0:
            raise RuntimeError("未找到 AD5933 设备")
        if self.board_index >= count:
            raise RuntimeError(f"board_index={self.board_index} 超出范围, 实际找到 {count} 块板")
        _, handle = connect(self.vid, self.pid, paths[self.board_index])
        self.handle = handle
        self.paths = paths

    def close(self):
        if self.handle is not None:
            disconnect(self.handle)
            self.handle = None

    # 为 with 语句提供支持
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    # ---- 高层功能 ----
    def get_configuration(self):
        """读取并返回当前 sweep 配置 dict，同时打印摘要。"""
        if self.handle is None:
            raise RuntimeError("设备未连接")
        cfg = get_configuration_from_handle(self.handle, mclk_hz=self.mclk_hz)

        print("=== AD5933 配置 ===")
        codes = cfg["codes"]
        params = cfg["params"]
        flags = cfg.get("decoded_flags", {})
        ctrl = cfg["control_status"]

        print(f"Start code   = 0x{codes['start_code']:06X}  -> f_start ≈ {params['start_freq_hz']:.6f} Hz")
        print(f"Delta code   = 0x{codes['delta_code']:06X}  -> f_delta ≈ {params['delta_freq_hz']:.6f} Hz")
        print(f"Incr code    = 0x{codes['incr_code']:06X}  -> N_incr = {params['num_increments']}, points = {params['num_points']}")
        print(f"Settling     = {params['settling_cycles']} cycles (仅当 0x8C==0 时有效)")

        print(f"Control 0x80 = 0x{(ctrl['ctrl_high_0x80'] or 0):02X}")
        print(f"Control 0x81 = 0x{(ctrl['ctrl_low_0x81'] or 0):02X}")
        print(f"Status  0x8F = 0x{(ctrl['status_0x8F'] or 0):02X}")

        print("--- 解码后的控制标志 ---")
        print(f"Function     = {flags.get('function')} (code={flags.get('function_code')})")
        print(f"Range Vpp    = {flags.get('range_vpp')} (code={flags.get('range_code')})")
        print(f"Gain         = {flags.get('gain')}")
        print(f"Clock source = {flags.get('clock_source')} (code={flags.get('clock_source_code')})")
        print(f"Reset flag   = {flags.get('reset')}")

        return cfg

    def calibrate_gain(self, z_ref_ohm=5000.0, cal_point="mid"):
        """执行一次 gain factor 校准，返回 measure_gain_factor_on_handle 的结果 dict。"""
        if self.handle is None:
            raise RuntimeError("设备未连接")
        return measure_gain_factor_on_handle(
            self.handle,
            mclk_hz=self.mclk_hz,
            z_ref_ohm=z_ref_ohm,
            cal_point=cal_point,
        )

    def sweep(self, gain_factor=None, config_for_sweep=None, csv_path=None):
        """执行 sweep 并返回结果列表，可选写入 CSV。"""
        if self.handle is None:
            raise RuntimeError("设备未连接")
        return measure_sweep_on_handle(
            self.handle,
            mclk_hz=self.mclk_hz,
            gain_factor=gain_factor,
            config_for_sweep=config_for_sweep,
            csv_path=csv_path,
        )

    def read_temperature(self):
        """简单读取一次温度，返回温度值（°C）。"""
        if self.handle is None:
            raise RuntimeError("设备未连接")

        # 1) 启动温度测量
        result, _ = vendor_request(self.handle, 0xDE, 0x0D, 0x9080, 0, 0, None)
        vprint(f"温度命令发送: {result}")
        time.sleep(0.05)

        # 2) 可选：读 0x80 / 0x8F 略

        # 3) 读取温度寄存器
        read_buf = (wintypes.BYTE * 1)()
        _, read_buf = vendor_request(self.handle, 0xDE, 0x0D, 0x92, 1, 1, read_buf)
        temp_upper = read_buf[0] & 0xFF
        _, read_buf = vendor_request(self.handle, 0xDE, 0x0D, 0x93, 1, 1, read_buf)
        temp_lower = read_buf[0] & 0xFF

        temp_raw = (temp_upper << 8) | temp_lower
        if temp_raw & 0x2000:
            temp_raw = temp_raw - 0x4000
        temperature = temp_raw / 32.0
        print(f"温度: {temperature:.2f} °C")
        return temperature


# ========== 1. 搜索设备（按“索引数组”理解）==========
def search_for_boards(vid=0x0456, pid=0xB203):
    """调用 ADI_CYUSB_USB4.Search_For_Boards 枚举设备并返回路径列表。

    文档原型写成:
        Uint  Search_For_Boards (uint VID, uint PID, uint *Num_boards, char *PartPath[]);

    但后面的说明又说："PartPath is a single location from the previous array, typically PartPath(0)",
    很像是把 PartPath 当成“板卡索引数组”（每个元素是一个 char，下标 0..N-1），
    而不是字符串路径。

    结合你当前现象：找到 1 块板，但我们读取到的第一个元素为 0（解码成字符串就是空串），
    更合理的猜测是：DLL 给 PartPath[i] 写的是“内部设备表索引”，Connect 用这个索引即可。

    这里因此将第四个参数绑定为一维 char 数组：char PartPath[MAX_BOARDS]。
    """

    dll.Search_For_Boards.restype = wintypes.DWORD

    MAX_BOARDS = 16
    num_boards = wintypes.DWORD(0)

    # 一维缓冲区：每个元素是一个 char，代表板卡索引或内部句柄
    PathArrayType = wintypes.CHAR * MAX_BOARDS
    path_indices = PathArrayType()

    dll.Search_For_Boards.argtypes = [
        wintypes.DWORD,                      # VID
        wintypes.DWORD,                      # PID
        ctypes.POINTER(wintypes.DWORD),      # *Num_boards
        PathArrayType                        # char PartPath[MAX_BOARDS]
    ]

    result = dll.Search_For_Boards(vid, pid, ctypes.byref(num_boards), path_indices)

    vprint(f"搜索结果: {result} (0=成功)")
    vprint(f"找到设备数: {num_boards.value}")

    devices = []  # 这里保存的是“索引”的字符串表示，仅用于打印
    for i in range(num_boards.value):
        idx = path_indices[i]
        # wintypes.CHAR 在 Python 这边是一个单字节整数（0,1,2...）
        if isinstance(idx, bytes):
            idx_val = idx[0]
        else:
            idx_val = idx
        vprint(f"  设备 {i} 索引值: {idx_val}")
        devices.append(str(idx_val))

    return result, num_boards.value, path_indices


# ========== 2. 连接设备 ==========
def connect(vid, pid, part_path):
    """调用 Connect 连接到指定路径的板卡。

    文档原型：
        Int Connect(Uint VID, Uint PID, char PartPath, Uint *Handle);

    结合“PartPath 是前面数组中的一个元素”的说明，更合理的 C 真实原型是：
        Int Connect(Uint VID, Uint PID, char *PartPath, Uint *Handle);

    因此这里按 char* 绑定，直接传入 C 字符串即可。
    """

    dll.Connect.restype = wintypes.INT

    handle = wintypes.DWORD(0)

    # 依据上面的推断，这里的 PartPath 实际上是一个单字节索引（char），而不是字符串指针
    dll.Connect.argtypes = [
        wintypes.DWORD,           # VID
        wintypes.DWORD,           # PID
        wintypes.CHAR,            # PartPath：索引值
        ctypes.POINTER(wintypes.DWORD)  # *Handle
    ]

    # part_path 这里直接视为“索引值”（0,1,2...），从 Search_For_Boards 返回的数组中取出即可
    if isinstance(part_path, (bytes, bytearray)):
        idx_val = part_path[0]
    elif isinstance(part_path, str):
        # 如果上层误传字符串，尝试解析成整数
        idx_val = int(part_path)
    else:
        idx_val = int(part_path)

    result = dll.Connect(vid, pid, idx_val, ctypes.byref(handle))
    print(f"连接结果: {result} (0=成功), 句柄: {handle.value}")

    return result, handle.value


# ========== 3. 下载固件 ==========
def download_firmware(handle, hex_file_path):
    dll.Download_Firmware.restype = wintypes.INT
    
    dll.Download_Firmware.argtypes = [
        wintypes.DWORD,      # Handle
        wintypes.CHAR * 512  # pcFilePath（null terminated）
    ]
    
    path_buffer = (wintypes.CHAR * 512)()
    path_bytes = hex_file_path.encode('utf-8')
    ctypes.memmove(path_buffer, path_bytes, min(len(path_bytes), 511))
    
    result = dll.Download_Firmware(handle, path_buffer)
    print(f"下载固件结果: {result} (0=成功)")
    return result


# ========== 4. 厂商请求（读写寄存器）==========
def vendor_request(handle, request, value, index, direction, data_length, buffer=None):
    """
    direction: 0=写(OUT), 1=读(IN)
    """
    dll.Vendor_Request.restype = wintypes.INT
    
    if buffer is None:
        buffer = (wintypes.BYTE * 64)()
    
    dll.Vendor_Request.argtypes = [
        wintypes.DWORD,      # Handle
        wintypes.BYTE,       # Request (0xDE)
        wintypes.WORD,       # Value (0x0D)
        wintypes.WORD,       # Index
        wintypes.BYTE,       # Direction (0=写, 1=读)
        wintypes.WORD,       # DataLength
        ctypes.POINTER(wintypes.BYTE)  # Buffer
    ]
    
    result = dll.Vendor_Request(
        handle, request, value, index, 
        direction, data_length, buffer
    )
    return result, buffer


def read_register_byte(handle, reg_addr, request=0xDE, value=0x000D):
    """通过 Vendor_Request 读 AD5933 单个寄存器字节。

    reg_addr: 0x80-0xFF 之间的寄存器地址。
    """
    buf = (wintypes.BYTE * 1)()
    # 读命令: wIndex = 0x00XX，其中低字节是寄存器地址
    index = reg_addr & 0x00FF
    result, buf = vendor_request(handle, request, value, index, 1, 1, buf)
    if result != 0:
        raise RuntimeError(f"读取寄存器 0x{reg_addr:02X} 失败, result={result}")
    # 统一按无符号 0-255 处理，避免 BYTE 被当成有符号 8bit
    return buf[0] & 0xFF


def write_register_byte(handle, reg_addr, reg_value, request=0xDE, value=0x000D):
    """通过 Vendor_Request 向 AD5933 写单个寄存器字节。

    DLL 抓包显示写操作使用的 wIndex 形如 0xVVAA：
        * 高字节 VV = 要写入的寄存器值
        * 低字节 AA = 寄存器地址

    reg_addr:  0x80-0xFF 之间的寄存器地址。
    reg_value: 要写入的 8bit 数值。
    """
    index = ((reg_value & 0xFF) << 8) | (reg_addr & 0xFF)
    result, _ = vendor_request(handle, request, value, index, 0, 0, None)
    if result != 0:
        raise RuntimeError(
            f"写寄存器 0x{reg_addr:02X} = 0x{reg_value:02X} 失败, result={result}"
        )
    return result


def get_configuration_from_handle(handle, mclk_hz=4_000_000.0):
    """从当前已连接的 AD5933 上读取并解码 sweep 配置。

    返回一个 dict，包含：
      - raw_registers: 相关寄存器的原始字节值
      - codes: 频率/步数相关的 24bit 代码
      - params: 按给定 mclk_hz 解码得到的物理量（Hz、点数、settling cycles 等）

    mclk_hz: AD5933 DDS 实际使用的 MCLK 频率。
      * 若外部 16MHz 且 REFCLK/4 打开，则 mclk_hz=4e6；
      * 若直接用 16MHz，则可传入 16e6 进行对比。
    """

    regs = {}
    for addr in (0x80, 0x81, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x8B, 0x8C, 0x8F):
        try:
            regs[addr] = read_register_byte(handle, addr)
        except RuntimeError as e:
            print(e)
            regs[addr] = None

    # 频率/步数 24bit 代码
    def _combine24(msb, mid, lsb):
        """将 3 个寄存器字节组合成无符号 24bit 代码。

        注意：某些 ctypes/wintypes 在 Python 这边会把 BYTE 当成有符号 8bit
        （-128..127），因此这里统一按 &0xFF 处理，避免 0x80..0xFF 被当成负数，
        导致出现 0x-0007A 这类“负的代码值”。
        """
        if None in (msb, mid, lsb):
            return None
        msb = msb & 0xFF
        mid = mid & 0xFF
        lsb = lsb & 0xFF
        return (msb << 16) | (mid << 8) | lsb

    start_code = _combine24(regs.get(0x82), regs.get(0x83), regs.get(0x84))
    delta_code = _combine24(regs.get(0x85), regs.get(0x86), regs.get(0x87))

    # Number of increments 寄存器在 AD5933 中是一个最多 9bit 的字段，起始地址 0x88。
    # ADI 的 no-OS 驱动将其视为一个 16bit 值写入 0x88 起始地址：
    #   write_reg(AD5933_REG_INC_NUM, value, 2)
    # 在你的板卡/GUI 抓包中，典型模式是：
    #   0x88 = 0x00, 0x89 = N, 0x8A = 0x00
    # 也就是整个 24bit 字段实际等价于 (N << 8)。
    # 因此真正的 increments 数应解码为 N = code >> 8，或更直接为 regs[0x89]。
    incr_code = _combine24(regs.get(0x88), regs.get(0x89), regs.get(0x8A))

    # Settling cycles：大部分实际配置下 0x8C 为 0，仅 0x8B 有效
    # 为避免过度猜测编码细节，这里只在 0x8C==0 时简单认为 cycles=0x8B。
    settling_cycles = None
    reg_8b = regs.get(0x8B)
    reg_8c = regs.get(0x8C)
    if reg_8b is not None and reg_8c is not None:
        if reg_8c == 0:
            settling_cycles = reg_8b

    # 根据 AD5933 公式解码频率
    def _code_to_freq(code):
        if code is None:
            return None
        return code * float(mclk_hz) / float(1 << 27)

    f_start = _code_to_freq(start_code)
    f_delta = _code_to_freq(delta_code)

    # 递增次数与点数
    # 注意：根据 ADI no-OS 驱动和你当前抓包，实际的 increments N
    # 存在于以 0x88 为首地址的 16bit 字段中，且 0x88 通常为 0x00、0x8A 为 0x00，
    # 0x89 直接等于 GUI 里设置的 increments 数。
    #   例：N=100 -> 0x88=0x00, 0x89=0x64, 组合 24bit=0x006400 (=100<<8)
    # 所以真实的 N 应当解码为 regs[0x89]（或 (incr_code >> 8) & 0xFFFF）。
    if incr_code is not None and regs.get(0x89) is not None:
        n_incr = regs.get(0x89) & 0xFF
    else:
        n_incr = None
    n_points = (n_incr + 1) if n_incr is not None else None

    # Control / Status 寄存器
    ctrl_high = regs.get(0x80)
    ctrl_low = regs.get(0x81)
    status = regs.get(0x8F)

    # 进一步解码 control 的几个关键标志位，方便直接使用
    decoded_flags = {
        "function": None,
        "function_code": None,
        "range_vpp": None,
        "range_code": None,
        "gain": None,
        "clock_source": None,  # "INT" / "EXT"
        "clock_source_code": None,
        "reset": None,
    }

    if ctrl_high is not None:
        func_code = (ctrl_high >> 4) & 0x0F
        range_code = (ctrl_high >> 1) & 0x07
        gain_bit = ctrl_high & 0x01

        function_map = {
            0x0: "NOP",
            0x1: "INIT_START_FREQ",
            0x2: "START_SWEEP",
            0x3: "INC_FREQ",
            0x4: "REPEAT_FREQ",
            0x9: "MEASURE_TEMP",
            0xA: "POWER_DOWN",
            0xB: "STANDBY",
        }

        range_map = {
            0: "2.0Vpp",   # AD5933_RANGE_2000mVpp
            1: "0.2Vpp",   # AD5933_RANGE_200mVpp
            2: "0.4Vpp",   # AD5933_RANGE_400mVpp
            3: "1.0Vpp",   # AD5933_RANGE_1000mVpp
        }

        gain_map = {
            0: "x5",  # AD5933_GAIN_X5
            1: "x1",  # AD5933_GAIN_X1
        }

        decoded_flags["function_code"] = func_code
        decoded_flags["function"] = function_map.get(func_code, f"UNKNOWN({func_code})")
        decoded_flags["range_code"] = range_code
        decoded_flags["range_vpp"] = range_map.get(range_code, "UNKNOWN")
        decoded_flags["gain"] = gain_map.get(gain_bit, "UNKNOWN")

    if ctrl_low is not None:
        clk_bit = (ctrl_low >> 3) & 0x01
        reset_bit = (ctrl_low >> 4) & 0x01

        decoded_flags["clock_source_code"] = clk_bit
        decoded_flags["clock_source"] = "EXT" if clk_bit else "INT"
        decoded_flags["reset"] = bool(reset_bit)

    return {
        "raw_registers": regs,
        "codes": {
            "start_code": start_code,
            "delta_code": delta_code,
            "incr_code": incr_code,
        },
        "params": {
            "mclk_hz": mclk_hz,
            "start_freq_hz": f_start,
            "delta_freq_hz": f_delta,
            "num_increments": n_incr,
            "num_points": n_points,
            "settling_cycles": settling_cycles,
        },
        "control_status": {
            "ctrl_high_0x80": ctrl_high,
            "ctrl_low_0x81": ctrl_low,
            "status_0x8F": status,
        },
        "decoded_flags": decoded_flags,
    }


def get_configuration(vid=0x0456, pid=0xB203, board_index=0, mclk_hz=4_000_000.0):
    """搜索并连接第一块板卡，读取并打印当前 sweep 配置。

    典型用法：先在官方 GUI 中设置好参数并完成初始化，然后在同一块板上运行：

        py -3.10-32 .\hack_board.py

    或在交互式 Python 中：

        from hack_board import get_configuration
        cfg = get_configuration(mclk_hz=4_000_000.0)

    然后把 cfg['params'] 里的数值和 GUI 填入的配置进行对比验证。
    """

    result, count, paths = search_for_boards(vid, pid)
    if count == 0:
        raise RuntimeError("未找到 AD5933 设备")

    if board_index >= count:
        raise RuntimeError(f"board_index={board_index} 超出范围, 实际找到 {count} 块板")

    result, handle = connect(vid, pid, paths[board_index])
    if result != 0:
        raise RuntimeError(f"连接失败, result={result}")

    try:
        cfg = get_configuration_from_handle(handle, mclk_hz=mclk_hz)

        print("=== AD5933 配置 ===")
        codes = cfg["codes"]
        params = cfg["params"]
        flags = cfg.get("decoded_flags", {})

        print(f"Start code   = 0x{codes['start_code']:06X}  -> f_start ≈ {params['start_freq_hz']:.6f} Hz")
        print(f"Delta code   = 0x{codes['delta_code']:06X}  -> f_delta ≈ {params['delta_freq_hz']:.6f} Hz")
        print(f"Incr code    = 0x{codes['incr_code']:06X}  -> N_incr = {params['num_increments']}, points = {params['num_points']}")
        print(f"Settling     = {params['settling_cycles']} cycles (仅当 0x8C==0 时有效)")

        ctrl = cfg["control_status"]
        print(f"Control 0x80 = 0x{(ctrl['ctrl_high_0x80'] or 0):02X}")
        print(f"Control 0x81 = 0x{(ctrl['ctrl_low_0x81'] or 0):02X}")
        print(f"Status  0x8F = 0x{(ctrl['status_0x8F'] or 0):02X}")

        # 额外打印一份解码好的 flags，省得肉眼算位
        print("--- 解码后的控制标志 ---")
        print(f"Function     = {flags.get('function')} (code={flags.get('function_code')})")
        print(f"Range Vpp    = {flags.get('range_vpp')} (code={flags.get('range_code')})")
        print(f"Gain         = {flags.get('gain')}")
        print(f"Clock source = {flags.get('clock_source')} (code={flags.get('clock_source_code')})")
        print(f"Reset flag   = {flags.get('reset')}")

        return cfg
    finally:
        disconnect(handle)


def measure_gain_factor_on_handle(handle,
                                  mclk_hz=4_000_000.0,
                                  z_ref_ohm=5000.0,
                                  cal_point="mid"):
    """在已有连接的 AD5933 上执行一次 gain factor 校准。

    整个流程尽量按照你抓到的 cal_gain 段复刻：
      1. 读取当前 sweep 配置（start/delta/N）
      2. 选一个校准点（默认扫频中点），计算对应频率代码并写入 0x82-0x84
      3. 通过控制寄存器状态机 INIT_START_FREQ -> START_SWEEP 触发一次测量
      4. 轮询状态寄存器 DATA_VALID 位
      5. 读取 Real/Imag，解码成有符号 16bit 数值
      6. 计算 magnitude 与 gain factor = 1 / (|Zref| * magnitude)

    返回一个 dict，便于后续 sweep 使用同一个 gain factor。
    """

    vprint("=== Gain factor 校准开始 ===")

    # 1. 读取当前配置
    vprint("[1] 读取当前 sweep 配置...")
    cfg = get_configuration_from_handle(handle, mclk_hz=mclk_hz)
    codes = cfg["codes"]
    params = cfg["params"]
    ctrl = cfg["control_status"]

    start_code = codes["start_code"]
    delta_code = codes["delta_code"]
    n_incr = params["num_increments"]

    if start_code is None or delta_code is None or n_incr is None:
        raise RuntimeError("当前寄存器配置不完整，无法执行校准（start/delta/N 缺失）")

    def code_to_freq(code):
        return code * float(mclk_hz) / float(1 << 27)

    f_start = code_to_freq(start_code)
    f_delta = code_to_freq(delta_code)

    if cal_point == "mid":
        k = n_incr // 2
        cal_desc = "中点 (mid)"
    elif cal_point == "start":
        k = 0
        cal_desc = "起始点 (start)"
    elif cal_point == "end":
        k = n_incr
        cal_desc = "终点 (end)"
    else:
        raise ValueError(f"未知 cal_point='{cal_point}', 只支持 'start'/'mid'/'end'")

    cal_code = start_code + k * delta_code
    f_cal = code_to_freq(cal_code)

    vprint(
        f"    start≈{f_start:.3f} Hz, delta≈{f_delta:.3f} Hz, N_incr={n_incr}, "
        f"选择第 {k} 个点作为校准频率 ({cal_desc})"
    )
    vprint(f"[2] 计算校准频率: code=0x{cal_code:06X}, f_cal≈{f_cal:.6f} Hz")

    # 2. 将校准频率写入起始频率寄存器 0x82-0x84
    vprint("[3] 写入校准频率到 0x82/0x83/0x84 ...")
    msb = (cal_code >> 16) & 0xFF
    mid = (cal_code >> 8) & 0xFF
    lsb = cal_code & 0xFF

    try:
        write_register_byte(handle, 0x84, lsb)
        write_register_byte(handle, 0x83, mid)
        write_register_byte(handle, 0x82, msb)
        ok_write = True
    except RuntimeError as e:
        vprint("    写寄存器失败:", e)
        ok_write = False

    if not ok_write:
        raise RuntimeError("写入校准起始频率失败，终止校准流程")

    # 读回确认
    r_msb = read_register_byte(handle, 0x82)
    r_mid = read_register_byte(handle, 0x83)
    r_lsb = read_register_byte(handle, 0x84)
    read_back_code = ((r_msb & 0xFF) << 16) | ((r_mid & 0xFF) << 8) | (r_lsb & 0xFF)
    f_read_back = code_to_freq(read_back_code)
    vprint(
        f"    读回 start_code = 0x{read_back_code:06X}, f≈{f_read_back:.6f} Hz"
    )

    # 3. 控制寄存器状态机：STANDBY -> INIT_START_FREQ -> START_SWEEP
    vprint("[4] 通过控制寄存器触发一次单点测量...")
    ctrl_high = ctrl["ctrl_high_0x80"]
    if ctrl_high is None:
        ctrl_high = read_register_byte(handle, 0x80)

    low_nibble = ctrl_high & 0x0F  # 保留原有 range/gain 配置
    standby_val = (0xB << 4) | low_nibble   # FUNCTION = STANDBY
    init_val = (0x1 << 4) | low_nibble      # FUNCTION = INIT_START_FREQ
    start_val = (0x2 << 4) | low_nibble     # FUNCTION = START_SWEEP

    try:
        write_register_byte(handle, 0x80, standby_val)
        vprint(f"    写 0x80=0x{standby_val:02X} (STANDBY) 成功")
        # 按 no-OS 流程，一般会插一个 reset，这里保持简化，只做状态切换
        write_register_byte(handle, 0x80, init_val)
        vprint(f"    写 0x80=0x{init_val:02X} (INIT_START_FREQ) 成功")
        write_register_byte(handle, 0x80, start_val)
        vprint(f"    写 0x80=0x{start_val:02X} (START_SWEEP) 成功")
    except RuntimeError as e:
        vprint("    写控制寄存器失败:", e)
        raise

    # 4. 轮询状态寄存器，等待 DATA_VALID 置位
    vprint("[5] 轮询状态寄存器 0x8F 等待 DATA_VALID 置位...")
    AD5933_STAT_DATA_VALID = 0x02
    status = 0
    max_tries = 100
    for i in range(max_tries):
        status = read_register_byte(handle, 0x8F)
        if status & AD5933_STAT_DATA_VALID:
            vprint(
                f"    第 {i+1} 次读取状态 0x8F=0x{status:02X}, DATA_VALID=1"
            )
            break
        if i == 0:
            vprint(f"    第 1 次读取状态 0x8F=0x{status:02X}, 继续等待...")
        time.sleep(0.005)
    else:
        raise RuntimeError(
            f"轮询 {max_tries} 次后 DATA_VALID 仍未置位, 最后状态=0x{status:02X}"
        )

    # 5. 读取 Real / Imag 数据
    vprint("[6] 读取 DFT 实部/虚部寄存器...")
    real_msb = read_register_byte(handle, 0x94)
    real_lsb = read_register_byte(handle, 0x95)
    imag_msb = read_register_byte(handle, 0x96)
    imag_lsb = read_register_byte(handle, 0x97)

    def to_signed16(msb, lsb):
        val = ((msb & 0xFF) << 8) | (lsb & 0xFF)
        if val & 0x8000:
            val -= 0x10000
        return val

    real = to_signed16(real_msb, real_lsb)
    imag = to_signed16(imag_msb, imag_lsb)

    vprint(
        f"    Real: MSB=0x{real_msb:02X}, LSB=0x{real_lsb:02X} -> {real}"
    )
    vprint(
        f"    Imag: MSB=0x{imag_msb:02X}, LSB=0x{imag_lsb:02X} -> {imag}"
    )

    # 6. 计算 magnitude 与 gain factor
    vprint("[7] 计算 magnitude 与 gain factor...")
    magnitude = math.sqrt(float(real * real + imag * imag))
    if magnitude == 0 or z_ref_ohm == 0:
        gain_factor = None
        print("magnitude 或 Z_ref 为 0，无法计算 gain factor")
    else:
        gain_factor = 1.0 / (magnitude * float(z_ref_ohm))
        print(f"|DFT| = {magnitude:.6f}")
        print(f"Z_ref = {z_ref_ohm} Ω")
        print(f"Gain factor = {gain_factor:.6e}")

    vprint("=== Gain factor 校准结束 ===")

    return {
        "config": cfg,
        "cal_point_index": k,
        "cal_freq_hz": f_cal,
        "cal_code": cal_code,
        "real": real,
        "imag": imag,
        "magnitude": magnitude,
        "gain_factor": gain_factor,
        "z_ref_ohm": z_ref_ohm,
    }


def measure_single_point_on_handle(handle,
                                   mclk_hz=4_000_000.0,
                                   gain_factor=None):
    """在当前 sweep 配置不变的前提下，做一次单点测量。

    不重写 0x82~0x88 等 sweep 配置寄存器，只通过控制寄存器
    STANDBY -> INIT_START_FREQ -> START_SWEEP 触发一次 DFT，
    然后读取 Real/Imag 并返回结果。
    """

    vprint("=== 单点测量（不改配置）开始 ===")

    # 读取当前起始频率代码，用于还原测量频率
    msb = read_register_byte(handle, 0x82)
    mid = read_register_byte(handle, 0x83)
    lsb = read_register_byte(handle, 0x84)
    start_code = ((msb & 0xFF) << 16) | ((mid & 0xFF) << 8) | (lsb & 0xFF)

    def code_to_freq(code):
        return code * float(mclk_hz) / float(1 << 27)

    f_meas = code_to_freq(start_code)

    # 读取控制寄存器，保留原有 range/gain 设置
    ctrl_high = read_register_byte(handle, 0x80)
    low_nibble = ctrl_high & 0x0F

    standby_val = (0xB << 4) | low_nibble
    init_val = (0x1 << 4) | low_nibble
    start_val = (0x2 << 4) | low_nibble

    # 触发一次单点测量
    write_register_byte(handle, 0x80, standby_val)
    write_register_byte(handle, 0x80, init_val)
    write_register_byte(handle, 0x80, start_val)

    # 轮询 DATA_VALID
    AD5933_STAT_DATA_VALID = 0x02
    status = 0
    max_tries = 100
    for _ in range(max_tries):
        status = read_register_byte(handle, 0x8F)
        if status & AD5933_STAT_DATA_VALID:
            break
        time.sleep(0.005)
    else:
        raise RuntimeError(
            f"单点测量轮询 {max_tries} 次后 DATA_VALID 仍未置位, 最后状态=0x{status:02X}"
        )

    # 读取 DFT 结果
    real_msb = read_register_byte(handle, 0x94)
    real_lsb = read_register_byte(handle, 0x95)
    imag_msb = read_register_byte(handle, 0x96)
    imag_lsb = read_register_byte(handle, 0x97)

    def to_signed16(msb, lsb):
        val = ((msb & 0xFF) << 8) | (lsb & 0xFF)
        if val & 0x8000:
            val -= 0x10000
        return val

    real = to_signed16(real_msb, real_lsb)
    imag = to_signed16(imag_msb, imag_lsb)
    mag = math.sqrt(float(real * real + imag * imag))

    if gain_factor is not None and mag > 0:
        z_abs = 1.0 / (gain_factor * mag)
    else:
        z_abs = None

    vprint(
        f"单点: f≈{f_meas:.3f} Hz, Real={real}, Imag={imag}, "
        f"|DFT|={mag:.3f}, |Z|={z_abs if z_abs is not None else 'N/A'}"
    )

    return {
        "index": 0,
        "freq_hz": f_meas,
        "real": real,
        "imag": imag,
        "magnitude": mag,
        "z_abs_ohm": z_abs,
    }


def measure_sweep_on_handle(handle,
                            mclk_hz=4_000_000.0,
                            gain_factor=None,
                            config_for_sweep=None,
                            csv_path=None):
    """执行一次完整 sweep，并可选保存为 CSV。

    典型用法：
        cal = measure_gain_factor_on_handle(handle, ...)
        results = measure_sweep_on_handle(handle,
                                          mclk_hz=4e6,
                                          gain_factor=cal["gain_factor"],
                                          config_for_sweep=cal["config"],
                                          csv_path="sweep.csv")

    参数：
        handle: 已连接设备句柄
        mclk_hz: AD5933 MCLK 频率
        gain_factor: 若提供，则在 CSV 中附带 |Z| 估算
        config_for_sweep: 事先读取的配置（推荐传入 cal["config"]，
                          这样 sweep 会按 GUI 原始配置重新写回所有寄存器，
                          而不会受前面校准写入 start_code 的影响）
        csv_path: 若不为 None，则将结果写入该 CSV 文件
    """

    print("=== Sweep 开始 ===")

    if config_for_sweep is None:
        print("[1] 未提供 config_for_sweep，直接从芯片读取当前配置...")
        cfg = get_configuration_from_handle(handle, mclk_hz=mclk_hz)
    else:
        print("[1] 使用传入的 config_for_sweep 作为 sweep 配置基础...")
        cfg = config_for_sweep

    codes = cfg["codes"]
    params = cfg["params"]
    ctrl = cfg["control_status"]

    start_code = codes["start_code"]
    delta_code = codes["delta_code"]
    n_incr = params["num_increments"]
    n_points = params["num_points"]
    settling_cycles = params["settling_cycles"]

    if start_code is None or delta_code is None or n_incr is None or n_points is None:
        raise RuntimeError("配置中 start/delta/N 不完整，无法执行 sweep")

    def code_to_freq(code):
        return code * float(mclk_hz) / float(1 << 27)

    print("[2] 将配置写回 AD5933 相关寄存器...")
    # Start code: 0x82(MSB),0x83(MID),0x84(LSB)
    sc_msb = (start_code >> 16) & 0xFF
    sc_mid = (start_code >> 8) & 0xFF
    sc_lsb = start_code & 0xFF
    write_register_byte(handle, 0x82, sc_msb)
    write_register_byte(handle, 0x83, sc_mid)
    write_register_byte(handle, 0x84, sc_lsb)

    # Delta code: 0x85(MSB),0x86(MID),0x87(LSB)
    dc_msb = (delta_code >> 16) & 0xFF
    dc_mid = (delta_code >> 8) & 0xFF
    dc_lsb = delta_code & 0xFF
    write_register_byte(handle, 0x85, dc_msb)
    write_register_byte(handle, 0x86, dc_mid)
    write_register_byte(handle, 0x87, dc_lsb)

    # Increments: 0x88 高字节通常为 0，0x89 为 N_incr
    write_register_byte(handle, 0x88, 0x00)
    write_register_byte(handle, 0x89, n_incr & 0xFF)

    # Settling cycles: 简单场景下 0x8A=0, 0x8B=cycles
    if settling_cycles is not None:
        write_register_byte(handle, 0x8A, 0x00)
        write_register_byte(handle, 0x8B, settling_cycles & 0xFF)

    print(
        f"    start_code=0x{start_code:06X}, delta_code=0x{delta_code:06X}, "
        f"N_incr={n_incr}, points={n_points}, settling={settling_cycles}"
    )

    # 控制寄存器基值：保留原来的低 4 位（range/gain），只修改高 nibble function
    ctrl_high = ctrl.get("ctrl_high_0x80")
    if ctrl_high is None:
        ctrl_high = read_register_byte(handle, 0x80)
    low_nibble = ctrl_high & 0x0F

    standby_val = (0xB << 4) | low_nibble
    init_val = (0x1 << 4) | low_nibble
    start_val = (0x2 << 4) | low_nibble
    inc_val = (0x3 << 4) | low_nibble

    print("[3] 通过控制寄存器启动 sweep...")
    write_register_byte(handle, 0x80, standby_val)
    write_register_byte(handle, 0x80, init_val)
    write_register_byte(handle, 0x80, start_val)

    AD5933_STAT_DATA_VALID = 0x02
    AD5933_STAT_SWEEP_DONE = 0x04

    results = []

    for k in range(n_points):
        # 轮询 DATA_VALID
        tries = 0
        while True:
            status = read_register_byte(handle, 0x8F)
            tries += 1
            if status & AD5933_STAT_DATA_VALID:
                if tries > 1:
                    print(f"    点 {k}: 第 {tries} 次读取 0x8F=0x{status:02X}, DATA_VALID=1")
                break
            if tries == 1:
                print(f"    点 {k}: 首次读取 0x8F=0x{status:02X}, 等待 DATA_VALID...")
            if tries > 200:
                raise RuntimeError(
                    f"点 {k} 轮询 0x8F 超过 {tries} 次仍未 DATA_VALID，最后值=0x{status:02X}"
                )
            time.sleep(0.005)

        # 读取 DFT
        real_msb = read_register_byte(handle, 0x94)
        real_lsb = read_register_byte(handle, 0x95)
        imag_msb = read_register_byte(handle, 0x96)
        imag_lsb = read_register_byte(handle, 0x97)

        def to_signed16(msb, lsb):
            val = ((msb & 0xFF) << 8) | (lsb & 0xFF)
            if val & 0x8000:
                val -= 0x10000
            return val

        real = to_signed16(real_msb, real_lsb)
        imag = to_signed16(imag_msb, imag_lsb)
        mag = math.sqrt(float(real * real + imag * imag))

        code_k = start_code + k * delta_code
        f_k = code_to_freq(code_k)

        if gain_factor is not None and mag > 0:
            z_abs = 1.0 / (gain_factor * mag)
        else:
            z_abs = None

        print(
            f"    点 {k}: f≈{f_k:.3f} Hz, Real={real}, Imag={imag}, "
            f"|DFT|={mag:.3f}, |Z|={z_abs if z_abs is not None else 'N/A'}"
        )

        results.append({
            "index": k,
            "freq_hz": f_k,
            "real": real,
            "imag": imag,
            "magnitude": mag,
            "z_abs_ohm": z_abs,
        })

        # 检查是否已到 sweep 末尾
        if status & AD5933_STAT_SWEEP_DONE:
            print(f"    点 {k}: SWEEP_DONE 置位，结束 sweep")
            break

        if k == n_points - 1:
            # 理论上不会再有下一点
            break

        # 递增频率到下一点
        write_register_byte(handle, 0x80, inc_val)

    print(f"[4] Sweep 完成，共获得 {len(results)} 个点")

    if csv_path:
        print(f"[5] 写入 CSV 文件: {csv_path} ...")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "index",
                "freq_hz",
                "real",
                "imag",
                "magnitude",
                "z_abs_ohm",
            ])
            for r in results:
                writer.writerow([
                    r["index"],
                    f"{r['freq_hz']:.6f}",
                    r["real"],
                    r["imag"],
                    f"{r['magnitude']:.6f}",
                    (f"{r['z_abs_ohm']:.6f}" if r["z_abs_ohm"] is not None else ""),
                ])
        print("    CSV 写入完成")

    print("=== Sweep 结束 ===")
    return results


# ========== 5. 断开连接 ==========
def disconnect(handle):
    dll.Disconnect.restype = wintypes.INT
    dll.Disconnect.argtypes = [wintypes.DWORD]
    
    result = dll.Disconnect(handle)
    print(f"断开结果: {result}")
    return result


# ========== 使用示例：读取温度 ==========
def read_temperature_example():
    # 1. 搜索 AD5933 设备
    vid, pid = 0x0456, 0xB203
    result, count, paths = search_for_boards(vid, pid)
    
    if count == 0:
        print("未找到设备")
        return
    
    # 2. 连接第一个设备
    result, handle = connect(vid, pid, paths[0])
    if result != 0:
        print("连接失败")
        return
    
    # 3. 不再在这里下载固件，初始化全部交给 GUI 完成
    #    保证在运行本脚本前，官方 GUI 已经连接并正确加载固件。
    
    # 4. 读取温度（最小复刻 GUI 的 5 条命令）
    # 1) 向控制寄存器高字节 0x80 写入 0x90 （启动温度测量）
    #    对应 Vendor_Request 写命令：Value=0x0D, Index=0x9080
    result, _ = vendor_request(handle, 0xDE, 0x0D, 0x9080, 0, 0, None)
    print(f"温度命令发送: {result}")

    # 给 AD5933 一点时间完成温度转换（GUI 在内部也会有延时）
    time.sleep(0.05)

    # 2) 读回控制寄存器 0x80（可选，用于和抓包对比）
    ctrl_buf = (wintypes.BYTE * 1)()
    result, ctrl_buf = vendor_request(handle, 0xDE, 0x0D, 0x0080, 1, 1, ctrl_buf)
    if result == 0:
        ctrl_val = ctrl_buf[0] & 0xFF
        print(f"控制寄存器 0x80: 0x{ctrl_val:02X}")
    else:
        print(f"读取控制寄存器 0x80 失败: {result}")

    # 3) 读取状态寄存器 0x8F（只读一次，用于和 GUI 对比，不再提前放弃）
    status_buf = (wintypes.BYTE * 1)()
    result, status_buf = vendor_request(handle, 0xDE, 0x0D, 0x008F, 1, 1, status_buf)
    if result == 0:
        status_val = status_buf[0] & 0xFF
        print(f"状态寄存器 0x8F: 0x{status_val:02X}")
    else:
        print(f"读取状态寄存器 0x8F 失败: {result}")

    # 4) 无论状态位如何，都尝试读取温度寄存器：0x92 (MSB), 0x93 (LSB)
    read_buf = (wintypes.BYTE * 1)()
    result, read_buf = vendor_request(handle, 0xDE, 0x0D, 0x92, 1, 1, read_buf)
    temp_upper = read_buf[0] & 0xFF
    print(f"温度高字节 (0x92): 0x{temp_upper:02X}")

    result, read_buf = vendor_request(handle, 0xDE, 0x0D, 0x93, 1, 1, read_buf)
    temp_lower = read_buf[0] & 0xFF
    print(f"温度低字节 (0x93): 0x{temp_lower:02X}")

    # 4) 按数据手册计算温度：14 位二进制补码，单位 1/32 °C
    temp_raw = (temp_upper << 8) | temp_lower
    if temp_raw & 0x2000:  # 负数
        temp_raw = temp_raw - 0x4000
    temperature = temp_raw / 32.0
    print(f"温度: {temperature:.2f} °C")
    
    # 5. 断开
    disconnect(handle)


if __name__ == "__main__":
    print("EVAL-AD5933 命令行工具")
    print("请确保官方 GUI 已经完成固件下载和基本初始化。")

    # 建立设备连接
    try:
        dev = EVAL_AD5933(vid=0x0456, pid=0xB203, mclk_hz=4_000_000.0)
    except Exception as e:
        print(f"初始化设备失败: {e}")
        raise SystemExit(1)

    print("\n可用命令:")
    print("  CFG         - 读取并打印当前 sweep 配置")
    print("  TEMP        - 读取一次温度")
    print("  CAL         - 做一次 gain factor 校准")
    print("  SWEEP       - 执行一次 sweep 并写出 CSV")
    print("  EXIT        - 退出")

    cached_gain_factor = None
    cached_config_for_sweep = None

    try:
        while True:
            cmd = input("\n请输入命令 (CFG/TEMP/CAL/SWEEP/EXIT): ").strip().upper()

            if cmd == "CFG":
                cfg = dev.get_configuration()
                cached_config_for_sweep = cfg

            elif cmd == "TEMP":
                dev.read_temperature()

            elif cmd == "CAL":
                try:
                    z_str = input("请输入参考电阻(欧)，回车默认 5000: ").strip()
                    z_ref = float(z_str) if z_str else 5000.0
                except ValueError:
                    print("输入无效，使用默认 5000 Ω")
                    z_ref = 5000.0

                cp_str = input("选择校准点 start/mid/end，回车默认 mid: ").strip().lower()
                if cp_str not in ("start", "mid", "end"):
                    cp_str = "mid"

                cal = dev.calibrate_gain(z_ref_ohm=z_ref, cal_point=cp_str)
                cached_gain_factor = cal.get("gain_factor")
                cached_config_for_sweep = cal.get("config")

            elif cmd == "SWEEP":
                if cached_gain_factor is None:
                    ans = input("当前无 gain factor，是否先执行 CAL? (Y/n): ").strip().lower()
                    if ans in ("", "y", "yes"):
                        try:
                            z_str = input("请输入参考电阻(欧)，回车默认 5000: ").strip()
                            z_ref = float(z_str) if z_str else 5000.0
                        except ValueError:
                            print("输入无效，使用默认 5000 Ω")
                            z_ref = 5000.0
                        cal = dev.calibrate_gain(z_ref_ohm=z_ref, cal_point="mid")
                        cached_gain_factor = cal.get("gain_factor")
                        cached_config_for_sweep = cal.get("config")

                cfg_for_sweep = cached_config_for_sweep

                default_csv = f"eit_sweep_{int(time.time())}.csv"
                csv_path = input(f"请输入 CSV 文件名，回车默认 {default_csv}: ").strip()
                if not csv_path:
                    csv_path = default_csv

                results = dev.sweep(
                    gain_factor=cached_gain_factor,
                    config_for_sweep=cfg_for_sweep,
                    csv_path=csv_path,
                )
                print(f"Sweep 结果已写入 {csv_path}，共 {len(results)} 个点")

            elif cmd == "EXIT":
                break

            else:
                print("未知命令，请输入 CFG/TEMP/CAL/SWEEP/EXIT")

    finally:
        dev.close()
