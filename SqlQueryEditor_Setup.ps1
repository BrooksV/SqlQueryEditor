# Step 1: Navigate to your project directory (adjust the path as necessary)
$projectPath = "C:\Git\SqlQueryEditor"
if (!(Test-Path $projectPath)) {
    New-Item -ItemType Directory -Path $projectPath
}
Set-Location -Path $projectPath

# Step 2: Create the scripts and logs directories
$scriptPath = Join-Path -Path $projectPath -ChildPath "scripts"
$logPath = Join-Path -Path $projectPath -ChildPath "logs"
if (!(Test-Path $scriptPath)) {
    New-Item -ItemType Directory -Path $scriptPath
}
if (!(Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath
}

# Step 3: Create the initial script files
$psScriptFile = Join-Path -Path $scriptPath -ChildPath "SqlQueryEditor.ps1"
$xamlFile = Join-Path -Path $scriptPath -ChildPath "SqlQueryEditor.xaml"
if (!(Test-Path $psScriptFile)) {
    New-Item -ItemType File -Path $psScriptFile
}
if (!(Test-Path $xamlFile)) {
    New-Item -ItemType File -Path $xamlFile
}

# Step 4: Create the initial log files
$transcriptLogFile = Join-Path -Path $logPath -ChildPath "SqlQuery_Transcript.log"
$userActivityLogFile = Join-Path -Path $logPath -ChildPath "SqlQuery_User_Activity.log"
if (!(Test-Path $transcriptLogFile)) {
    New-Item -ItemType File -Path $transcriptLogFile
}
if (!(Test-Path $userActivityLogFile)) {
    New-Item -ItemType File -Path $userActivityLogFile
}

# Step 5: Create the README.md file
$readmeFile = Join-Path -Path $projectPath -ChildPath "README.md"
if (!(Test-Path $readmeFile)) {
    New-Item -ItemType File -Path $readmeFile
}

# Step 6: Initialize a Git repository and make the first commit
git init
git add .
git commit -m "Initial commit"
