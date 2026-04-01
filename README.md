# 🔐 BitLocker TPM + PIN Deployment Tool (UFCV)

PowerShell GUI tool for **enterprise-grade BitLocker deployment** with **TPM + PIN + Recovery Key**, designed for controlled environments with **Active Directory**, **GPO enforcement**, and **user-friendly workflow**.

---

## 📋 Table of Contents

* [📖 Overview](#-overview)
* [✨ Features](#-features)
* [🧱 Architecture](#-architecture)
* [⚙️ Requirements](#️-requirements)
* [🚀 Usage](#-usage)
* [🧠 How It Works](#-how-it-works)
* [🛡️ Security Considerations](#️-security-considerations)
* [⚠️ Troubleshooting](#️-troubleshooting)
* [📁 Project Structure](#-project-structure)
* [📌 TODO / Improvements](#-todo--improvements)
* [📜 License](#-license)

---

## 📖 Overview

This script provides a **complete and secure BitLocker deployment workflow** for enterprise environments.

It ensures that:

* The system is compliant with **BitLocker GPO policies**
* The device is connected to the **correct Active Directory domain**
* A **Recovery Key is securely backed up to AD DS**
* The user sets a **secure PIN via a graphical interface**
* The deployment is **controlled, traceable, and user-friendly**

💡 Designed for internal use within **UFCV DSI infrastructure**.

---

## ✨ Features

### 🔍 Pre-deployment validation

* GPO compliance check (`HKLM\SOFTWARE\Policies\Microsoft\FVE`)
* Registry value comparison (Expected vs Current)
* Blocking execution if non-compliant

### 🌐 Network & domain validation

* Verifies Active Directory domain (`ufcvfr.lan`)
* Detects reachable Domain Controller
* Supports LAN and VPN scenarios

### 🧑‍💻 User-friendly GUI (WPF)

* Modern interface for PIN input
* Real-time validation (length, format, sequence detection)
* Visual feedback (colors, states)

### ⏳ Postpone system

* Users can delay activation (up to 99 times)
* Persistent counter stored in:

  ```
  C:\ProgramData\BitLockerActivation\PostponeCount.txt
  ```
* Automatic enforcement when limit is reached

### 🔐 BitLocker provisioning

* Adds or reuses **Recovery Password**
* Backs up key to **Active Directory**
* Enables BitLocker with:

  * TPM + PIN
  * XTS-AES 256
  * Used Space Only mode

### 📊 Live progress UI

* Asynchronous execution (Runspace)
* Step-by-step status
* Progress bar + logs

### 🧠 Smart handling

* Detects existing encryption state
* Handles GPO delay errors (e.g. `0x80310060`)
* Prevents interruption during provisioning

---

## 🧱 Architecture

```text
User (GUI WPF)
        ↓
Validation Layer
  - GPO / Registry
  - Domain / Network
        ↓
Control Logic
  - Postpone system
  - State checks
        ↓
Provisioning Engine
  - BitLocker module
  - AD Backup
        ↓
Async Execution (Runspace)
        ↓
Progress UI Feedback
```

---

## ⚙️ Requirements

### 💻 System

* Windows 10 / 11 (Enterprise recommended)
* TPM 2.0 enabled
* Secure Boot enabled

### 🏢 Environment

* Active Directory domain joined
* GPO BitLocker configured
* Network access to Domain Controller

### 🔑 Permissions

* Execution as **SYSTEM / LocalSystem** recommended

---

## 🚀 Usage

### ▶️ Run the script

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1
```

### ⚡ Recommended deployment

* SCCM / Intune
* GPO Startup Script
* Scheduled Task (SYSTEM context)

---

## 🧠 How It Works

### 1️⃣ Environment validation

* Checks registry configuration
* Verifies domain membership

### 2️⃣ User interaction

* PIN input (6–20 digits)
* Validation rules applied

### 3️⃣ Provisioning process

* Recovery key creation
* Backup to AD DS
* BitLocker activation

### 4️⃣ Finalization

* Progress displayed
* Reboot required

---

## 🛡️ Security Considerations

* 🔒 PIN is never stored in plaintext
* 🔑 Recovery key is backed up to Active Directory
* 🚫 Script blocks execution if environment is not compliant
* 🧠 Prevents weak PINs (sequences, invalid formats)
* ⛔ Prevents interruption during encryption setup

---

## ⚠️ Troubleshooting

### ❌ GPO not applied

* Run:

```powershell
gpupdate /force
```

### ❌ Error `0x80310060`

* Reboot required after GPO application

### ❌ Not on domain

* Ensure:

```powershell
whoami /fqdn
```

### ❌ BitLocker already running

* Check:

```powershell
Get-BitLockerVolume
```

---

## 📁 Project Structure

```text
.
├── BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1
├── .gitattributes
└── README.md
```

---

## 📌 TODO / Improvements

* [ ] Logging to file (centralized logs)
* [ ] Integration with SIEM / monitoring tools
* [ ] Multi-language support
* [ ] Enhanced reporting (JSON / API)
* [ ] Silent mode (no GUI)

---

## 📜 License

This repository is intended for internal professional use within an organizational context.

All rights reserved.
Unauthorized use, reproduction, modification, or distribution outside the company or authorized scope is prohibited without prior permission.
