# Windows-port-manager

A comprehensive PowerShell utility to diagnose and manage TCP/IP port reservations and conflicts on Windows. This tool is designed for developers, system administrators, and power users who frequently encounter "port in use" errors caused by system services, WSL, Docker, or other applications.

It provides a user-friendly, menu-driven interface to quickly identify port usage, manage reserved port ranges, and resolve common conflicts.
<p align="center">
<img width="620" height="297" alt="image" src="https://github.com/user-attachments/assets/3769fe90-6052-41f3-9818-75587afb4ba8" />
</p>

## The Problem

On Windows, especially with features like Hyper-V, WSL2, and Docker, the system can dynamically reserve large, seemingly random blocks of ports. This often leads to conflicts when you try to run a development server or application on a specific port (e.g., 8000, 5000, 3000), only to be told the port is already in use, even when no application appears to be using it.

Diagnosing this requires a series of `netsh` and `netstat` commands, and resolving it can be a tedious process of trial and error. This script automates that entire workflow.

## Features

*   **Automatic Admin Elevation**: The script detects if it's running without administrator privileges and will prompt to relaunch itself as an administrator.
*   **Show Current Status**: Display all currently reserved port ranges and the system's dynamic port range in one command.
*   **Check Specific Port**: Instantly check if a port is in use and, if so, identify the process name and PID using it.
*   **Full Port Range Diagnostics**: Before you reserve a range, run a full diagnostic to check for:
    *   Active processes using ports in the range.
    *   Conflicts with existing reserved ranges.
    *   Conflicts with the system's dynamic port range.
*   **Reserve a Port Range**:
    *   Intelligently checks for conflicts before attempting a reservation.
    *   If a port in the desired range is in use, it identifies the process and offers to stop it for you.
    *   Provides clear feedback on success or failure, diagnosing the likely cause (e.g., dynamic port conflict).
*   **Delete a Reserved Range**: Lists all custom and system-reserved ranges and allows you to easily select one to delete.
*   **Manage Dynamic Port Range**: Easily view and change the default TCP dynamic port range to prevent future conflicts.
*   **Restart WinNAT Service**: A one-click option to restart the Windows NAT service, which can often release "ghost" port reservations held by WSL or Docker.

<p align="center">
  <tr><td align="center"><img width="829" height="808" alt="Main Menu" src="https://github.com/user-attachments/assets/d90150ab-bde9-4744-90ce-5211d353efca" /></td></tr>
</p>

## How to Use

1.  **Download**: Save the `port-manager.ps1` script to your computer.
2.  **Open PowerShell**: Open a PowerShell or Windows Terminal window.
3.  **Navigate to the script**:
    ```powershell
    cd path\to\your\script
    ```
4.  **Run the script**:
    ```powershell
    .\port-manager.ps1
    ```

If you are not running as an administrator, the script will offer to restart itself with the necessary permissions.
<p align="center">
<img width="747" height="192" alt="image" src="https://github.com/user-attachments/assets/dc1ad1d7-3dde-4fde-af4b-85c8dadca5c9" />
</p>

Once running, simply follow the on-screen menu to select the action you want to perform.

## Example Workflow: Reserving ports for a project

Let's say you need to ensure ports `8000-8010` are free for your web development projects.

1.  Run the script: `.\port-manager.ps1`.
2.  Choose **Option 7** (Full diagnostics for a port range).
3.  Enter `8000` as the start port and `11` as the number of ports.
4.  The script will tell you if there are any conflicts.
    *   If a process is using a port, you can stop it.
    *   If it conflicts with the dynamic range, you can use **Option 5** to change it.
    *   If it conflicts with a Hyper-V reservation, you can try **Option 6** (Restart WinNAT).
5.  Once the diagnostics report is clear, the script will ask if you want to reserve the range. Choose 'y'.

Your ports are now protected from being used by other applications!
<p align="center">
<img width="626" height="737" alt="image" src="https://github.com/user-attachments/assets/4a8d04e5-c6f1-47d0-a1c3-61be2ab9c34f" />
</p>
