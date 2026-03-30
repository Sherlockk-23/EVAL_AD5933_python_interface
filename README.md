# Python Control for EVAL-AD5933EBZ

## How it works

We hacked the usb communication between the GUI and the EVAL-AD5933EBZ evaluation board, and by calling the `ADI_CYUSB_USB4.dll` directly, we can control the board with Python. 

## Installation

You must first install the official GUI software for EVAL-AD5933EBZ, following [EVAL-AD5933 User Guide](https://www.analog.com/media/en/technical-documentation/user-guides/UG-364.pdf).

Locate `ADI_CYUSB_USB4.dll`, it is commonly at `C:\Program Files\Analog Devices\USB Drivers\ADI_CYUSB_USB4.dll`. You need to copy this path to the `AD5933.py`.

Then, since the `ADI_CYUSB_USB4.dll` is 32bit, you need to install **32bit Python**. We are using Python 3.10 (32bit) as an example. You can download it from [Python's official website](https://www.python.org/downloads/).

However, the 32bit Python cannot use many of the latest versions of popular libraries. So you may need to use python 64 for later data processing and visualization. You can also build a bridge to run two python processes at the same time, one for controlling the board and the other for data processing and visualization.



## Usage
First open the official GUI software for EVAL-AD5933EBZ for basic setup.

Then you can run `py310-32 AD5933.py` to control the board with Python. 