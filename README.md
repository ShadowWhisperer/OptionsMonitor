Compare strike price to current price, and show what happens on expiration date.  
Stock prices are pulled from Yahoo Finance.  

Stock list is saved here *C:\ProgramData\ShadowWhisperer\OptionsMonitor\data.csv*  

Double click an entry to modify it, then press Enter or click off.  

[Download](https://github.com/ShadowWhisperer/OptionsMonitor/releases/latest/download/OptionsMonitor.exe)

<br> 
 
**Build from source**  
```
pyinstaller --noconsole --onefile -i om.ico -n OptionsMonitor.exe options.py --version-file version.txt --add-data "om.ico;."
```

<img width="741" height="365" alt="Capture" src="https://github.com/user-attachments/assets/8833a94c-50c4-43ef-8752-a95866ed16a5" />
