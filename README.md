# Diag Comparer
- **Choose a folder** with 2 Diag Reports in it</br> (you will be prompted at the startup)
- **Diag Comparer** will display their Detailed-configuration-tables side by side.</br> Lines that don't match are marked with "◄", so updates are easy to spot

![image](https://github.com/DavidBevi/sisw/blob/main/CalibTxtCompare_Screenshot1.png?raw=true)

# ztl - Zip TestLab
A command to zip `DiagAsFound`+`DiagAsLeft`+`CalAsFound`+`CalAsLeft`.  **Requires 7-zip**

<details>
<summary>Install</summary>
  
 - put `ztl.bat` into a folder [listed into user-path or system-path](https://www.thewindowsclub.com/system-user-environment-variables-windows) (like C:\Windows\System32)
</details>

<details>
<summary>Simple usage</summary>
  
- use File Explorer to navigate to the folder containing the calibration files
- type `ztl` in the address bar to make the .7z file
</details>

<details>
<summary>Advanced usage</summary>
  
- Use a terminal to navigate to the target folder, and use `ztl` to see the process
</details>

![image](https://github.com/user-attachments/assets/c1da01e7-acac-4c09-9619-48c1c8ecb31c)
