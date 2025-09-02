# SessionEnd Hook - Shows notification when Claude Code session ends
# This script runs when the session is terminated

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("Claude Code セッションが終了しました", 'Claude Code', 'OK', 'Information')