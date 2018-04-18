# SETI@REST
SETI@REST is a tiny app that helps to reduce GDI lag while your GPU is hungrily crunching BOINC GPU tasks.
Another function is to maximize the fan speed of some notebook computers via SpeedFan.

SETI@REST能够调节BOINC的GPU使用率，解决电脑卡顿。
在GPU运算时，GPU完全分配给了BOINC，这时电脑会非常卡。SETI@REST调节BOINC的GPU计算，在用户操作电脑时降低GPU计算，闲置后立刻使GPU满负载。

SETI@REST的另一项功能是使一部分支持SpeedFan的风扇全速运转。
在DELL E6410上，SpeedFan设置风扇转速100%时，只能持续数秒。该功能可使这些风扇全速运转。

This project is also an example to use the tray icon and the tray balloontip.

## Quick Start
Tick "Enable GPU Usage Toggle" to let SETI@REST start toggling GPU process. During the first run, you should use GPU-Z to find out the appropriate toggle intensity value of your environment.

Try to adjust the "Toggle Intensity" to a value that you won't feel the lag simultaneously make the GPU usage as high as possible.

### GPU Usage Toggle
SETI@REST is able to lower the GPU usage of the boinc GPU calculation process, and meanwhile maintain the GPU calculation.
The principle is to suspend and resume the GPU process.

### GPU Process Keywords
The process name that with in these keywords would be regard as the SETI GPU app. Process names are case insensitive, and seperate by "|". 
(e.g.: cuda|opencl would help SETI@REST to recognize "Lunatics_x41zc_win32_cuda23.exe" as the GPU process)

### Idle Time Threshold
The amount of idle time before make GPU to 100%.

### Toggle Intensity
How much effort does SETI@REST try to lower the GPU usage when computer is busy. You should use GPU-Z to monitor the GPU load and find out the appropriate toggle intensity value for your computer.

### Fan Control
Note that the fan maximize function dependents on SpeedFan. (e.g. DELL Latitude E6410, which is able to control fan speed with this trick) 

### Force Fan Speed to
Enforce fan speed to desired value

### Auto Hide SpeedFan Main Window
Make SpeedFan main window hide from desktop & taskbar.

### Command-Line Usage
```
/t    Start SETI@HOME minimized in tray
```

## FAQ
### What is the difference between BOINC built in GPU control and SETI@REST?
The built in BOINC GPU control (Use GPU only after computer has been idle for XX minutes) causes more overheads, since BOINC closes the entire GPU process when user is active. When user idle the GPU process need restart and the initialization can took more than 10 seconds(can be examined from GPU usage, the initialization progress doesn't use GPU at all). SETI@REST only suspend and resume the GPU process, which wouldn't waste any extra CPU & GPU time.

### Does SETI@REST support multiple GPU process?
Currently, SETI@REST only support 1 GPU computation process.

## Release Notes
* Feb 18, 2013 - Version 0.1 Build 18 - First version  
* Jul 5, 2013 - Add 6400 Fan Full
