[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ScriptFromGitHub = Invoke-WebRequest https://raw.githubusercontent.com/path/to/SPLA/SPLAScriptv1.ps1
Invoke-Expression $($ScriptFromGitHub.Content)