# AberTecno Core (Distribution)

![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Excel-lightgrey.svg) ![Status](https://img.shields.io/badge/status-Production-green.svg)

**AberTecno Core** is a high-performance COM Middleware designed to bridge Microsoft Excel (VBA) with Cloud Services.

This repository hosts the **distribution binaries and installer**. The source code is proprietary and hosted in a private secure repository.

---

## 📥 Download & Install

> **Current Version:** Check the [`version.txt`](version.txt) file in the file list above.

1.  **Download:** Click on [`AberTecnoSetup.exe`](AberTecnoSetup.exe) in the file list above and click the "Download" button (save raw file).
2.  **Install:** Run the `.exe` as Administrator.
3.  **Ready:** The library is now registered and ready to use in Excel.

---

## 🚀 Key Features

* **Zero-Freeze:** Excel UI remains responsive while data is syncing.
* **Offline Mode:** Logs are stored locally if internet is lost and synced later.
* **Auto-Update:** The library self-checks this repository for new versions.

---

## 💻 Quick Start (VBA)

Copy and paste this code into your Excel Module to start using the library immediately.

### 1. Initialization

```vba
Dim AT As Object
Set AT = CreateObject("AberTecno.Controller")

' Initialize with a unique Flow ID
AT.Init "FLOW_MAIN_PROCESS"

' Set with a unique Function ID (Optional)
AT.SetFunction "FUNCTION_100"
```

### 2. Logging Data
Send data to the cloud. This method returns immediately (non-blocking).

```vba
' Log(Start, End, Status, Message)
AT.Log Format(Now, "dd/mm/yyyy hh:mm:ss"), _
       Format(Now, "dd/mm/yyyy hh:mm:ss"), _
       "SUCCESS", _
       "Data processed successfully."
```

### 3. Force Sync (Optional)
Usually not needed (auto-syncs on Init), but useful for testing.

```vba
AT.Sync
```

## ⚙️ Requirements

* **OS:** Windows 10 / 11 (x64 recommended).
* **Software:** Microsoft Excel (2010 or newer).
* **Framework:** .NET Framework 4.7.2 or higher.

## 🛠️ Architecture Overview

```mermaid
graph LR
    A[Excel VBA] -- Late Binding --> B(AberTecno Core DLL)
    B -- Async Thread --> C{Internet?}
    C -- Yes --> D[Cloud Gateway]
    C -- No --> E[Local Queue .txt]
    D --> F[AppSheet DB]
    E -- Retry Later --> B
```

---

**© 2026 AberTecno Inc.** | *Developed for internal automation processes.*
