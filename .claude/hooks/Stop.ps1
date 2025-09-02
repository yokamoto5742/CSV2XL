# Stop Hook - Shows notification when user stops Claude Code
# This script runs when the user manually stops the session

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("作業が完了しました", 'Claude Code', 'OK', 'Information')