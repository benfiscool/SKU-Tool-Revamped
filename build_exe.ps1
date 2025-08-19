# Build script: creates venv, installs requirements, and runs PyInstaller
$venv = "venv_build"
python -m venv $venv
& "$venv\Scripts\pip.exe" install --upgrade pip
& "$venv\Scripts\pip.exe" install -r requirements.txt

# Run PyInstaller to create a one-folder build with console
$pyinstallerExe = "$venv\Scripts\pyinstaller.exe"
$pyArgs = @(
	'--noconfirm',
	'--onedir',
	'--console',
	'--name', 'sku-tool',
	'--hidden-import=flask',
	'--hidden-import=werkzeug',
	'--hidden-import=click',
	'--hidden-import=itsdangerous',
	'--hidden-import=blinker',
	'--hidden-import=jinja2',
	'--hidden-import=google',
	'--hidden-import=googleapiclient',
	'--hidden-import=googleapiclient.discovery',
	'--hidden-import=googleapiclient.http',
	'--hidden-import=googleapiclient.errors',
	'--hidden-import=google_auth_oauthlib',
	'--hidden-import=google_auth_oauthlib.flow',
	'--hidden-import=google.auth',
	'--hidden-import=google.auth.transport.requests',
	'--hidden-import=httplib2',
	'--hidden-import=uritemplate',
	'SkuTool Revamped Backup.py'
)
& $pyinstallerExe @pyArgs

Write-Host "Build finished. See dist\sku-tool\ for the executable and supporting files."
