# Data Automation with Powershell
This PowerShell script processes a CSV of user actions, grouping users by program and faculty, and formats them for Papercut account import. It generates a TSV file with structured user additions and removals, excluding headers for compatibility with the Papercut system.

The project uses PowerShell for scripting, leveraging Import-Csv, Group-Object, and custom objects for data manipulation. Key concepts include conditional filtering, object pipelining, string formatting, and manual TSV generation using Set-Content to ensure compatibility with Papercut’s input requirements.
