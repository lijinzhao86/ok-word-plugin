# Start OpenClaw Gateway with Bonjour disabled
$env:OPENCLAW_SKIP_BONJOUR = "1"
Set-Location "d:\llmprojects\ok-word-plugin\openclaw"
pnpm openclaw gateway --port 18789 --verbose
