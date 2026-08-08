# WordPad

A feature-rich **Microsoft Foundation Classes (MFC)** word processor demonstrating advanced document editing, Rich Edit integration, OLE embedding, printing, customization, and modern Office-style user interface components.

Originally based on the Microsoft MFC WordPad sample, this project serves as both a fully functional document editor and a comprehensive reference implementation for building professional Windows desktop applications with MFC.

---

## Overview

WordPad demonstrates how to build a complete MDI-style Windows text editor using the MFC framework. While inspired by the classic Windows WordPad application, it also showcases many advanced MFC features introduced with the Feature Pack, including Office-style toolbars, docking windows, task panes, customizable UI, and Rich Edit document handling.

The project is especially valuable as a learning resource for developers interested in:

- Modern MFC application architecture
- Rich text editing
- OLE document containers
- Printing infrastructure
- Ribbon/toolbar customization
- Dockable windows
- Office-like user interfaces
- Document/View architecture

---

# Features

## Rich Text Editing

- Rich Edit based document editor
- Character formatting
- Paragraph formatting
- Font selection
- Font size selection
- Text colors
- Background highlighting
- Bullet lists
- Indentation controls
- Tabs
- Alignment options
- Undo / Redo
- Clipboard support
- Find / Replace
- Word wrapping

---

## Document Support

Supports multiple document formats including Rich Text Format (RTF).

Features include:

- New/Open/Save
- Save As
- Print
- Print Preview
- Page Setup
- Recent files
- Document templates
- File conversion support
- Import/export infrastructure

---

## Office-Style User Interface

The application demonstrates numerous MFC Feature Pack controls, including:

- Dockable toolbars
- Menu bar
- Status bar
- Formatting toolbar
- Font toolbar
- Color picker
- Docking panes
- Customizable toolbars
- Visual managers
- Office appearance
- Dynamic menus
- Toolbar personalization

---

## Task Pane

Includes an Office-like task pane capable of hosting:

- Formatting commands
- Document operations
- Quick actions
- Application shortcuts

This illustrates how to integrate `CMFCTasksPane` into an MFC application.

---

## OLE Support

The project demonstrates MFC OLE container functionality including:

- Embedded objects
- Linked objects
- OLE in-place activation
- Container items
- OLE frame windows

Relevant classes include:

- COleCntrFrameWndEx
- COleIPFrameWndEx

---

## Printing

Comprehensive printing support:

- Print dialog
- Print preview
- Page setup
- Headers
- Margins
- Printer selection

---

## Customizable Appearance

Demonstrates MFC visual customization:

- Office visual managers
- Themes
- Docking layouts
- Toolbar customization
- User preferences
- Persisted UI state

---

# Architecture

The application follows the classic **Document/View** architecture.

```
Application
      │
      ▼
Document
      │
      ▼
Rich Edit View
      │
      ▼
MFC Framework
```

Major components include:

- application class
- main frame window
- document class
- Rich Edit view
- task pane
- formatting dialogs
- options dialogs
- printing support
- OLE container infrastructure

---

# Project Structure

| File | Purpose |
|-------|----------|
| `WordPad.sln` | Visual Studio solution |
| `mainfrm.cpp` | Main application frame |
| `MainFrm.h` | Main frame declarations |
| `wordpad.cpp` | Application entry point |
| `wordpadDoc.cpp` | Document implementation |
| `wordpadView.cpp` | Rich Edit view |
| `TaskPane.*` | Office-style task pane |
| `format*.cpp` | Formatting dialogs |
| `options*.cpp` | Application options |
| `pageset.cpp` | Page setup |
| `printdlg.cpp` | Printing support |
| `cntritem.cpp` | OLE container objects |

---

# Technologies

- C++
- Microsoft Foundation Classes (MFC)
- Win32 API
- Rich Edit Control
- OLE / COM
- GDI
- Windows Printing API

---

# Demonstrated MFC Classes

This project showcases a large number of MFC Feature Pack classes, including:

- `CWinAppEx`
- `CMDIFrameWndEx`
- `CFrameWndEx`
- `CMFCMenuBar`
- `CMFCToolBar`
- `CMFCStatusBar`
- `CMFCFontComboBox`
- `CMFCToolBarFontComboBox`
- `CMFCToolBarFontSizeComboBox`
- `CMFCColorMenuButton`
- `CMFCColorBar`
- `CMFCPopupMenu`
- `CMFCToolBarsCustomizeDialog`
- `CMFCVisualManager`
- `CMFCVisualManagerOffice`
- `CMFCTasksPane`
- `CPane`
- `CDockState`
- `COleCntrFrameWndEx`
- `COleIPFrameWndEx`

making it an excellent reference for developers working with advanced MFC applications.

---

# Building

Requirements:

- Visual Studio
- Desktop C++ workload
- Microsoft Foundation Classes (MFC)

Build steps:

```text
Open WordPad.sln

Select Debug or Release

Build Solution

Run
```

---

# Educational Value

This project is an excellent reference for learning:

- MFC application architecture
- Document/View programming
- Rich Edit programming
- OLE containers
- Docking windows
- Office-style UI
- Printing architecture
- MFC Feature Pack controls
- Toolbar customization
- Modern Windows desktop development

---

# Based On

This project originates from the **Microsoft WordPad MFC sample**, which demonstrates how to implement an application that reproduces the functionality of WordPad while showcasing numerous MFC Feature Pack components and best practices for Windows desktop development.

The original sample focuses on illustrating concepts rather than production-ready implementation, making it an ideal educational reference for MFC developers.

---

# Use Cases

- Learning MFC
- Rich Edit development
- Office-style UI implementation
- OLE programming
- Printing systems
- Document/View architecture
- Legacy Windows application maintenance
- MFC Feature Pack reference

---

# License

Please refer to the original project for licensing information.

This repository is based on Microsoft's educational MFC sample and preserves its purpose as a demonstration of MFC application development.
