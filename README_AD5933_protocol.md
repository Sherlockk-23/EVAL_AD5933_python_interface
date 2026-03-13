# AD5933 USB 协议与测量流程笔记

> 基于 ADI 官方 GUI 的 USB 抓包 (`complete_round_out.txt`) 反向工程结果，整理 FX2 板卡 + AD5933 在 Windows 下通过厂商 DLL / USB 控制传输的使用协议与典型测量流程。

本文件只覆盖我们在 GUI 中实际用到的寄存器与操作：

- 寄存器映射（子集）
- USB 控制传输封装（读 / 写 AD5933 寄存器）
- 配置读取（`get_configuration`）
- 温度测量流程
- 校准增益（`cal_gain` / `measure_gain_factor`）
- 频率扫描（`sweep`）流程

所有例子均假设：

- `bRequest = 222`
- `wValue = 0x000d`
- `bmRequestType`:
  - 0x40 = Host→Device，Vendor，Device（写寄存器）
  - 0xC0 = Device→Host，Vendor，Device（读寄存器）

---

## 1. AD5933 寄存器子集

只列出本项目中实际使用到的寄存器：

- 0x80：Control High Byte
- 0x81：Control Low Byte
- 0x82–0x84：Start Frequency Code（24-bit，LSB first）
- 0x85–0x87：Frequency Increment Code（24-bit，LSB first）
- 0x88–0x89：Number of Increments（16-bit，只用低字节）
- 0x8A–0x8B：Settling Time Cycles（+ multiplier）
- 0x8F：Status Register
- 0x92–0x93：温度原始数据
- 0x94–0x95：DFT Real（16-bit signed）
- 0x96–0x97：DFT Imag（16-bit signed）

### 1.1 Control High (0x80)

- 高 4 bit：Function Code
  - 0x1x：Initialize with start frequency (INIT_START_FREQ)
  - 0x2x：Start frequency sweep (START_SWEEP)
  - 0x3x：Increment frequency (INC/REPEAT 变体，见抓包中 0x31)
  - 0x4x：Repeat current frequency (REPEAT_FREQ 变体，见抓包中 0x41)
  - 0x9x：Measure temperature (MEASURE_TEMP)
- 低 4 bit：与量程、增益相关（具体位分布参考 ADI no-OS `ad5933.h`，此处只按抓包结果解码）：
  - 输出量程（Range）：
    - 0 → 2.0 Vpp
    - 1 → 1.0 Vpp
    - 2 → 0.4 Vpp
    - 3 → 0.2 Vpp
  - 增益：x1 / x5（在 no-OS 中对应 Gain bit）

### 1.2 Control Low (0x81)

- Bit3：时钟源
  - 0 = 内部时钟
  - 1 = 外部时钟
- Bit4：Reset
  - 1 = 复位

### 1.3 频率编码 & 个数

- Start Freq Code：三个字节（0x82 LSB，0x83 Mid，0x84 MSB）
- Delta Freq Code：三个字节（0x85 LSB，0x86 Mid，0x87 MSB）
- Number of Increments：
  - （0x88 MSB, 0x89 LSB）整体是 16-bit 字段，**但 GUI 只用低字节 0x89** 作为 N_incr
  - 实际点数 = N_points = N_incr + 1

频率公式（MCLK = 外部或内部时钟，单位 Hz）：

$$
 f = \text{Code} \times \frac{MCLK}{2^{27}}
$$

### 1.4 Settling Time（0x8A–0x8B）

- GUI 中常见组合：
  - 0x8A = 0x00
  - 0x8B = 某个 cycles 数（例如 0x0F = 15 cycles）
- 当 0x8C = 0 时，0x8B 直接表示 settling cycles（无额外倍增），抓包中我们只看到这种情况。

### 1.5 Status Register (0x8F)

- Bit0：VALID (Data valid for real/imag) – 我们看到：
  - 0x71：无 DATA_VALID
  - 0x73：DATA_VALID = 1
- Bit1：Sweep complete (SWEEP_DONE)
- Bit2：Temperature valid (TEMP_VALID)

GUI 的做法基本是循环读 0x8F，直到 DATA_VALID 或 SWEEP_DONE 置位。

### 1.6 DFT 结果寄存器

- Real:
  - 0x94：MSB
  - 0x95：LSB
- Imag:
  - 0x96：MSB
  - 0x97：LSB

都是 16-bit 有符号数（补码），Python 端要按 signed short 解包。

---

## 2. USB Vendor Request 封装

厂商 DLL 在 Windows 下通过 URB_CONTROL + Vendor Request 访问 AD5933 寄存器。抓包显示格式固定：

- 读寄存器：
  - `bmRequestType = 0xC0`（Device→Host, Vendor, Device）
  - `bRequest = 222`
  - `wValue = 0x000d`
  - `wIndex = 0x00AA`（AA = 寄存器地址）
  - `wLength = 1`
  - Data = 返回的寄存器值 1 字节

- 写寄存器：
  - `bmRequestType = 0x40`（Host→Device, Vendor, Device）
  - `bRequest = 222`
  - `wValue = 0x000d`
  - `wIndex = 0xVVAA`（高字节 VV = 要写入的值，低字节 AA = 寄存器地址）
  - `wLength = 0`
  - 无 data 阶段

对应的 Python 封装（伪代码）：

```python
# 读 1 字节寄存器
read_register_byte(handle, addr):
    wIndex = addr & 0xFF
    return vendor_request(handle,
                          bmRequestType=0xC0,
                          bRequest=222,
                          wValue=0x000D,
                          wIndex=wIndex,
                          length=1)

# 写 1 字节寄存器
write_register_byte(handle, addr, value):
    wIndex = ((value & 0xFF) << 8) | (addr & 0xFF)
    vendor_request(handle,
                   bmRequestType=0x40,
                   bRequest=222,
                   wValue=0x000D,
                   wIndex=wIndex,
                   length=0)
```

我们所有的高层操作（读配置、温度、cal_gain、sweep）都基于这两类请求构造。

---

## 3. 配置读取（get_configuration）

目标：在 GUI 配好 sweep 参数后，不改变 AD5933 当前状态，直接从寄存器读回：

- 起始频率 / 步进频率（原始 Code 与 Hz）
- 扫描步数（N_incr）与总点数（N_points）
- Settling cycles
- 当前 Control 状态：Function、Range、Gain、Clock Source、Reset
- 当前 Status：DATA_VALID / SWEEP_DONE / TEMP_VALID

### 3.1 寄存器读取

读取以下地址：

- 0x80–0x8C（根据需要）
- 0x8F

组合 24-bit Code 时要注意每个字节都按无符号处理：

```python
code = ((reg_msb & 0xFF) << 16) | ((reg_mid & 0xFF) << 8) | (reg_lsb & 0xFF)
```

### 3.2 解码规则

- 频率：
  - `f_start = start_code * MCLK / 2**27`
  - `f_delta = delta_code * MCLK / 2**27`
- 扫描点数：
  - `N_incr = regs[0x89]`（低字节），`N_points = N_incr + 1`
- Settling：
  - 简单场景中：`settling_cycles = regs[0x8B]`（当 0x8C == 0）
- Control 解码：
  - function_code = (regs[0x80] >> 4) & 0x0F
  - range/gain 等依照 no-OS `ad5933.h` 映射
- Status 解码：
  - DATA_VALID = bool(regs[0x8F] & 0x02)
  - SWEEP_DONE = bool(regs[0x8F] & 0x04)
  - TEMP_VALID = bool(regs[0x8F] & 0x01)

### 3.3 返回结构（示意）

Python 中 `get_configuration_from_handle(handle, mclk_hz)` 大致返回：

```python
{
  "raw": {  # 所有原始寄存器值
    0x80: ..., 0x81: ..., ..., 0x8F: ...
  },
  "start_code": int,
  "delta_code": int,
  "f_start_hz": float,
  "f_delta_hz": float,
  "increments": int,       # N_incr
  "num_points": int,       # N_incr + 1
  "settling_cycles": int,
  "decoded_flags": {
    "function": "INIT_START_FREQ" / "START_SWEEP" / ...,
    "range_vpp": 2.0 / 1.0 / 0.4 / 0.2,
    "gain": "x1" / "x5",
    "clock_source": "internal" / "external",
    "reset": bool,
  },
  "status": {
    "data_valid": bool,
    "sweep_done": bool,
    "temp_valid": bool,
  },
}
```

---

## 4. 温度测量流程

抓包与 ADI 文档一致，流程如下：

1. 写 Control High (0x80) 为 Measure Temperature 模式（高 nibble = 0x9）。
2. 轮询 Status (0x8F) 直到 TEMP_VALID 置位。
3. 读 0x92 (MSB) 与 0x93 (LSB) 得到 16-bit 温度原始值。
4. 根据 AD5933 数据手册将原始值换算为温度（此处略）。

Python 实现要点：

- 使用 `write_register_byte(0x80, 0x9X)` 进入温度模式（X = 保持原来的 range/gain bits）。
- 读 0x8F，检查 TEMP_VALID 位。
- TEMP_VALID 置位后，从 0x92, 0x93 读出值。

---

## 5. 校准增益（cal_gain / measure_gain_factor）

目标：在一个已知参考阻抗 $Z_\text{ref}$ 下，测一次 AD5933 的 DFT 输出 (Real, Imag)，计算 "gain factor"，用来在后续 sweep 中换算阻抗：

- $$ |Z_\text{meas}| \approx \frac{1}{\text{gain_factor} \cdot |\text{DFT}|} $$

### 5.1 高层思路

1. 读当前配置（start/delta/N_incr 等）。
2. 选择一个校准点 index `k`（例如中点）：
   - `cal_code = start_code + k * delta_code`
   - `f_cal = cal_code * MCLK / 2**27`
3. 把 `cal_code` 写回 0x82–0x84（覆盖 GUI 原来的 start_code，只用于本次校准）。
4. 启动单频测量：
   - 写 Control 0x80 = INIT_START_FREQ (0x1X)
   - 写 Control 0x80 = START_SWEEP (0x2X)
5. 轮询 Status 0x8F 直到 DATA_VALID 置位。
6. 读 DFT Real/Imag 寄存器 0x94–0x97。
7. 计算：
   - `mag = sqrt(real**2 + imag**2)`
   - `gain_factor = 1.0 / (mag * Z_ref)`

### 5.2 关键 USB 操作

以抓包中 `segment=cal_gain` 为例，控制序列为：

- 写 0x80 进入 STANDBY / 正确的 range+gain 状态
- 写 0x82–0x84 为选中的 `cal_code`
- 写 0x80 = 0x1X（INIT_START_FREQ）
- 写 0x80 = 0x2X（START_SWEEP）
- 多次读 0x8F：
  - 典型模式：先读出 0x71（无 DATA_VALID），随后 0x73（DATA_VALID=1）
- 依次读 0x94, 0x95, 0x96, 0x97：拼成 16-bit signed Real/Imag

Python 函数 `measure_gain_factor_on_handle(handle, mclk_hz, z_ref_ohm, cal_point)` 就是对上述顺序的封装，并打印每一步的寄存器值与计算结果，方便与 GUI 对照。

---

## 6. 频率扫描（sweep）流程

抓包中的 `segment=sweep1` 展示了 GUI 执行一次 sweep 的完整过程。大致可分为：

1. 编程 / 确认 sweep 参数
2. 启动 sweep（INIT_START_FREQ → START_SWEEP）
3. 对每个频点：
   - 等待 DATA_VALID
   - 读取 Real/Imag
   - 调整 function code（INC/REPEAT）进入下一个频点
4. 结束条件：SWEEP_DONE 或到达 N_incr+1 个点

### 6.1 参数编程与确认

典型顺序（对应抓包中 #126–#165 一段）：

- 写 Start Freq Code：0x82–0x84
- 写 Delta Freq Code：0x85–0x87
- 写 Number of Increments：
  - 0x88 = 0x00（高字节不使用）
  - 0x89 = N_incr（例如 0x03）
- 写 Settling Time：
  - 0x8B = cycles（例如 0x0F）
  - 0x8A = 0x00（multiplier 关闭）
- 设置 Control Low：
  - 写 0x81 = 0x08（示例：外部时钟 + 其他标志）
- 读回上述寄存器，确认写入成功：
  - 多次读 0x82–0x87, 0x88–0x8B, 0x81 等

### 6.2 启动 sweep

抓包中关键几步（#170–#181 一带）：

- 写 0x80 = 0x11
  - 高 nibble 1 = INIT_START_FREQ
- 读 0x80 = 0x11 确认
- 写 0x80 = 0x21
  - 高 nibble 2 = START_SWEEP
- 读 0x80 = 0x21 确认

此时 AD5933 开始在起始频率处进行第一次 DFT 变换。

### 6.3 单个频点测量循环

在 `segment=sweep1` 的后半段（约 #186 之后），可以看到典型模式：

1. 控制作频点或重复：
   - 例如写 0x80 = 0x41 或 0x31（高 nibble 4/3，对应 REPEAT/INC 变体）
2. 轮询 Status (0x8F)：
   - 一段时间内反复读 0x8F 返回 0x71（DATA_VALID=0）
   - 直至某次读出 0x73（DATA_VALID=1）
3. 在 DATA_VALID=1 后，读取 DFT 结果：
   - 读 0x94 → Real MSB
   - 读 0x95 → Real LSB
   - 读 0x96 → Imag MSB
   - 读 0x97 → Imag LSB
   - 拼成 signed 16-bit Real/Imag
4. 根据函数码（0x31/0x41）和 AD5933 文档，切换到下一个频点或重复当前频点，重复步骤 1–3。

多点 sweep 完成后，Status 0x8F 的 SWEEP_DONE 位会被置位，GUI 停止继续发 INC/REPEAT 命令。

### 6.4 Python 中的 sweep 伪代码

在 Python 侧，可以基于 `get_configuration_from_handle` 与 `measure_gain_factor_on_handle` 中现有的读取逻辑，实现类似：

```python
def measure_sweep_on_handle(handle, mclk_hz, gain_factor=None):
    cfg = get_configuration_from_handle(handle, mclk_hz)
    start_code = cfg["start_code"]
    delta_code = cfg["delta_code"]
    n_incr = cfg["increments"]
    n_points = cfg["num_points"]

    # 1) INIT + START
    write_register_byte(handle, 0x80, 0x10 | (cfg_low_nibble))  # INIT_START_FREQ
    write_register_byte(handle, 0x80, 0x20 | (cfg_low_nibble))  # START_SWEEP

    results = []

    for k in range(n_points):
        # 2) 等待 DATA_VALID
        while True:
            status = read_register_byte(handle, 0x8F)
            if status & 0x02:  # DATA_VALID
                break

        # 3) 读取 DFT
        real = read_s16_from_regs(handle, 0x94)
        imag = read_s16_from_regs(handle, 0x96)
        mag = math.sqrt(real*real + imag*imag)

        # 计算频率
        code_k = start_code + k * delta_code
        f_k = code_k * mclk_hz / (1 << 27)

        # 计算阻抗（若已有 gain_factor）
        z_abs = None
        if gain_factor is not None and mag > 0:
            z_abs = 1.0 / (gain_factor * mag)

        results.append({
            "index": k,
            "freq_hz": f_k,
            "real": real,
            "imag": imag,
            "mag": mag,
            "z_abs_ohm": z_abs,
        })

        # 4) 判断是否结束 sweep
        if status & 0x04:  # SWEEP_DONE
            break

        # 5) 递增频率（依据 no-OS，通常写 INC_FREQ）
        write_register_byte(handle, 0x80, 0x30 | (cfg_low_nibble))

    return results
```

> 注：上面是逻辑示意，并未严格区分 0x31 / 0x41 的所有变体，实际实现可以直接参考 ADI no-OS 的 `ad5933_start_sweep` / `ad5933_get_impedance` 等函数，并与抓包中观察到的 0x31/0x41 使用方式对齐。

---

## 7. 总结

- 寄存器访问：通过固定格式的 Vendor Request，将寄存器地址编码在 `wIndex` 的低字节，写操作则将要写的值放在高字节。
- 配置读取：直接从当前 AD5933 寄存器中解码出 GUI 设置的 sweep 参数与控制 / 状态标志。
- 温度测量：使用 0x80 的 Measure Temperature function code，轮询 0x8F 的 TEMP_VALID，再读 0x92–0x93。
- 校准增益：在参考阻抗下进行单频测量，从 DFT Real/Imag 计算 gain_factor。
- 频率扫描：遵循 INIT_START_FREQ → START_SWEEP → （轮询 DATA_VALID → 读 DFT → INC/REPEAT）×N_points → SWEEP_DONE 的序列。

本 README 旨在作为以后开发 / 调试时的协议速查表，与 Python 实现中的 `get_configuration_from_handle`、`measure_gain_factor_on_handle` 等函数配套使用。