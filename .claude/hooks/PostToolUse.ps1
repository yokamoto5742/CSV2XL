# PostToolUse Hook - Shows notification after tool completion
# This script runs after every tool execution

param(
    [string]$toolName,
    [string]$result
)

# Show notification for significant tool completions
if ($toolName -match "pytest|Edit|Write|MultiEdit|Bash.*test") {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("タスクが完了しました: $toolName", 'Claude Code', 'OK', 'Information')
}