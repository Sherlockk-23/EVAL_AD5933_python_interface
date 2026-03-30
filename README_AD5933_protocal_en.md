# AD5933 USB Protocol Notes (English)

This document summarizes the AD5933 USB protocol used by the EVAL board through the ADI Windows DLL. It is aligned with the simplified implementation in AD5933.py.

## 1. Scope

The workflow covered here includes:

- board discovery and connection
- register read/write through USB vendor requests
- configuration readback
- temperature measurement
- gain-factor calibration
- full frequency sweep

Assumed default USB request fields:

- bRequest = 0xDE (222)
- wValue = 0x000D

## 2. Key AD5933 Registers

- 0x80: Control High
- 0x81: Control Low
- 0x82-0x84: Start frequency code (24-bit)
- 0x85-0x87: Delta frequency code (24-bit)
- 0x88-0x89: Number of increments (GUI usually uses low byte 0x89)
- 0x8A-0x8B: Settling cycles
- 0x8F: Status
- 0x92-0x93: Temperature raw value
- 0x94-0x95: Real (signed 16-bit)
- 0x96-0x97: Imag (signed 16-bit)

## 3. USB Vendor Request Format

### 3.1 Read Register Byte

- direction: IN
- wIndex = 0x00AA, where AA is register address
- length = 1

### 3.2 Write Register Byte

- direction: OUT
- wIndex = 0xVVAA, where:
  - VV is register value to write
  - AA is register address
- length = 0

## 4. Frequency and Data Conversion

### 4.1 Frequency Code to Hz

f = code * MCLK / 2^27

Typical MCLK in this project is 4 MHz when the external clock path effectively provides REFCLK/4.

### 4.2 Real/Imag Decode

Real and Imag are signed 16-bit values from:

- Real: 0x94 (MSB), 0x95 (LSB)
- Imag: 0x96 (MSB), 0x97 (LSB)

Magnitude:

|DFT| = sqrt(real^2 + imag^2)

### 4.3 Impedance Estimate with Gain Factor

|Z| = 1 / (gain_factor * |DFT|)

## 5. Configuration Readback

Read registers:

- 0x80, 0x81
- 0x82-0x87
- 0x88-0x8B
- 0x8F

Decode:

- start_code from 0x82-0x84
- delta_code from 0x85-0x87
- increments from 0x89
- points = increments + 1
- settling cycles from 0x8B when 0x8A is 0

## 6. Temperature Measurement Sequence

1. Read current 0x80 and keep low nibble (range/gain bits).
2. Write function MEASURE_TEMP (high nibble = 0x9) to 0x80.
3. Poll 0x8F until TEMP_VALID bit is set (bit mask 0x01).
4. Read 0x92 and 0x93.
5. Convert raw value to Celsius:
   - 14-bit two's complement
   - temperature = raw / 32.0

## 7. Gain Calibration Sequence

Given known reference impedance Z_ref:

1. Read current sweep config.
2. Select calibration point k (start/mid/end).
3. Compute cal_code = start_code + k * delta_code.
4. Write cal_code to 0x82-0x84.
5. Trigger single conversion:
   - STANDBY (0xBx)
   - INIT_START_FREQ (0x1x)
   - START_SWEEP (0x2x)
6. Poll 0x8F until DATA_VALID (bit mask 0x02).
7. Read Real/Imag from 0x94-0x97.
8. Compute:
   - magnitude = sqrt(real^2 + imag^2)
   - gain_factor = 1 / (Z_ref * magnitude)

## 8. Sweep Sequence

1. Program sweep registers (start, delta, increments, settling).
2. Trigger sweep:
   - STANDBY
   - INIT_START_FREQ
   - START_SWEEP
3. For each point:
   - poll DATA_VALID in 0x8F
   - read Real/Imag
   - compute frequency from code
   - if not done, issue INC_FREQ (0x3x)
4. Stop when SWEEP_DONE is set (bit mask 0x04) or expected points reached.
5. Optionally write CSV columns:
   - index, freq_hz, real, imag, magnitude, z_abs_ohm

## 9. Notes

- Always mask register bytes with 0xFF when combining values.
- Keep control low nibble when switching function codes to avoid changing range/gain settings.
- If calibration changes start frequency registers, reuse a saved configuration for sweep to restore original setup.
