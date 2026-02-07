# WorkHours

WorkHours is a simple Python-based tool for tracking and managing work hours and exporting them to Excel files.  
It is designed to be lightweight, fast, and easy to use from the command line.

A pre-built Windows executable (`.exe`) is available — no Python installation required.

---

## ⚠️ Required Files (Important)

**WorkHours.exe MUST be placed in the same folder as:**

```
objektai.txt
```

### objektai.txt format
- Each object name must be on a **separate line**
- Example:

```
1 objektas
2 objektas
3 objektas
```

If `objektai.txt` is missing or empty, the program **will not work**.

---

## ✨ Features

- Add and manage work hour entries
- Object names loaded automatically from `objektai.txt`
- Automatic Excel (`.xlsx`) file generation
- Automatic data summary table
- Supports month input by name or number
- Simple terminal-based interface
- Windows executable available

---

## 📦 Download (Windows)

Go to **Releases** and download the latest version:

➡ **WorkHours.exe**

After downloading:
1. Create a folder (for example `WorkHours`)
2. Place **WorkHours.exe** inside it
3. Place **objektai.txt** in the **same folder**
4. Run the program

---

## 🚀 Usage (Windows EXE)

1. Open the folder containing:
   ```
   WorkHours.exe
   objektai.txt
   ```
2. Double-click `WorkHours.exe`
3. Follow the on-screen instructions in the console
4. Excel files will be created in the same folder

---

## 📁 Output

- Excel files are generated automatically
- Existing files are reused and updated if found
- File naming format:
  ```
  <name>_<year>_<month>.xlsx
  ```

---


## 📌 Notes

- The Windows EXE is built using GitHub Actions
- Builds are performed on Windows runners
- Antivirus false positives are possible with PyInstaller-built executables

---

## 📜 License

MIT License
