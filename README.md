# 🔍 Win-Validate : Outil de Diagnostic & Audit Hardware

![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg) ![License](https://img.shields.io/badge/License-MIT-green.svg)

**Win-Validate** est un outil automatisé écrit en PowerShell pour auditer, tester et noter l'état de santé des ordinateurs (PC Portables et Fixes) dans un contexte de reconditionnement ou de maintenance informatique.

## 🚀 Fonctionnalités

* **Scoring Automatique (/20) :** Note globale calculée sur CPU, RAM, GPU, Disque et Batterie.
* **Stress Test Batterie :** Analyse l'état de santé (SOH) et détecte les chutes de tension en charge.
* **Détection de Pannes :** Plafonne automatiquement la note si un composant critique est défectueux (Disque SMART, Batterie HS).
* **Historique Centralisé :** Génère un rapport TXT par machine (trié par modèle) et alimente un fichier CSV global (`Inventaire_Parc_Global.csv`).
* **100% Portable :** Conçu pour fonctionner depuis une clé USB sans installation.

## 🛠️ Prérequis

* Windows 10 ou 11.
* Exécution en tant qu'**Administrateur** (Requis pour WinSAT et BatteryReport).

## 📦 Compilation (PS2EXE)

Le script est conçu pour être compilé en `.exe`.

```powershell
Invoke-PS2EXE -InputFile ".\pc_validator.ps1" `
              -OutputFile ".\Win-Validate_v4.1.exe" `
              -icon ".\favicon.ico" `
              -requireAdmin `
              -title "Win-Validate" `
              -description "Outil de Diagnostic Hardware" `
              -version "4.1.0.0"
