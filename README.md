# Python Control for EVAL-AD5933EBZ

## How it works

We hacked the usb communication between the GUI and the EVAL-AD5933EBZ evaluation board, and by calling the `ADI_CYUSB_USB4.dll` directly, we can control the board with Python. 

## Installation

You must first install the official GUI software for EVAL-AD5933EBZ, following [EVAL-AD5933 User Guide](https://www.analog.com/media/en/technical-documentation/user-guides/UG-364.pdf).

Locate `ADI_CYUSB_USB4.dll`, it is commonly at `C:\Program Files\Analog Devices\USB Drivers\ADI_CYUSB_USB4.dll`. You need to copy this path to the `AD5933.py`.

Then, since the `ADI_CYUSB_USB4.dll` is 32bit, you need to install **32bit Python**. We are using Python 3.10 (32bit) as an example. You can download it from [Python's official website](https://www.python.org/downloads/).

No more dependencies are needed.

## Usage
First open the official GUI software for EVAL-AD5933EBZ for basic setup.

Then you can run `AD5933.py` to control the board with Python. 