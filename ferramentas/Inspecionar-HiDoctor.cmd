@echo off
rem Inspecionar-HiDoctor.cmd: gera um relatorio SO com a ESTRUTURA do banco do HiDoctor (programas,
rem nomes de arquivos de banco, tabelas, colunas, tipos e contagens). Nao le nenhum registro de
rem paciente. O relatorio vai para a Area de Trabalho (HiDoctor-estrutura.txt) e abre no Bloco de
rem Notas: leia antes de enviar. Feche o HiDoctor antes de rodar.
rem Este bloco so abre o Windows PowerShell de 32 bits (o mesmo tipo do HiDoctor, que le bancos
rem Access antigos) com a parte PowerShell deste arquivo, conferida e copiada para um arquivo
rem temporario. Opcoes para testes: -Saida <arquivo> -SemBlocoDeNotas -PastasExtras <pasta>.
setlocal EnableExtensions DisableDelayedExpansion
title Estrutura do banco do HiDoctor
rem Windows PowerShell started from PowerShell 7 inherits its module path and loses some
rem built-in commands; use Windows PowerShell's own default module path.
set "PSModulePath=%SystemRoot%\system32\WindowsPowerShell\v1.0\Modules;%ProgramFiles%\WindowsPowerShell\Modules"
set "INSPECIONAR=%~f0"
set "INSPECIONAR_PS1=%TEMP%\Inspecionar-HiDoctor-%RANDOM%%RANDOM%.ps1"
set "INSPECIONAR_POWERSHELL=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%INSPECIONAR_POWERSHELL%" set "INSPECIONAR_POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%INSPECIONAR_POWERSHELL%" -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "$t = [IO.File]::ReadAllText($env:INSPECIONAR, [Text.Encoding]::UTF8); $m = [regex]::Match($t, '(?m)^#=== PowerShell'); if (-not $m.Success) { exit 1 }; $s = $t.Substring($m.Index); $k = $null; $e = $null; [void][Management.Automation.Language.Parser]::ParseInput($s, [ref]$k, [ref]$e); if ($e.Count -gt 0) { $e | ForEach-Object { Write-Host ('linha ' + $_.Extent.StartLineNumber + ': ' + $_.Message) }; exit 1 }; [IO.File]::WriteAllText($env:INSPECIONAR_PS1, $s, (New-Object Text.UTF8Encoding $true))"
rem No labels or goto: this file also works when downloaded with Unix line endings (LF only).
if errorlevel 1 echo  O Windows PowerShell nao conseguiu abrir este programa. Tire uma foto desta janela e envie.
if exist "%INSPECIONAR_PS1%" "%INSPECIONAR_POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%INSPECIONAR_PS1%" %*
set "INSPECIONAR_CODIGO=%errorlevel%"
del /f /q "%INSPECIONAR_PS1%" >nul 2>nul
echo.
if not defined INSPECIONAR_SEM_PAUSA pause
exit /b %INSPECIONAR_CODIGO%

#=== PowerShell: the inspector itself (the lines above only start it) ===
# Reads only STRUCTURE: installed programs, database services, database file names and formats,
# table and column names, types and row counts. It never prints a record, a file name that could be
# an attachment's, a folder name below the first level, or a configuration value other than where and
# how the data is stored.
param(
  [string]$Saida = '',
  [switch]$SemBlocoDeNotas,
  [string[]]$PastasExtras = @()
)
$ErrorActionPreference = 'Continue'
$ArquivoDoRelatorio = $Saida
if (-not $ArquivoDoRelatorio) { $ArquivoDoRelatorio = Join-Path ([Environment]::GetFolderPath('Desktop')) 'HiDoctor-estrutura.txt' }
$Linhas = New-Object System.Collections.Generic.List[string]
function W([string]$s) { $Linhas.Add($s); Write-Host $s }
# A relative path cut to the first folder and the file name: deeper folders (an attachment folder
# can be named after a patient) never appear.
function Curto([string]$rel) { $p = $rel.Split('\'); if ($p.Count -le 2) { return $rel }; return ($p[0] + '\...\' + $p[$p.Count - 1]) }
function Secao([string]$s) { W ''; W ('=' * 78); W $s; W ('=' * 78) }

# File names are listed only for database files with an unambiguous extension (ExtForte). Files with
# a generic extension (ExtFraca: .dat, .bak, .db...) are examined but their names are omitted, and any
# other file is only counted, so the name of an attachment (which could carry a patient's name)
# never appears in the report.
$ExtForte = '.mdb','.accdb','.fdb','.gdb','.sqlite','.sqlite3','.dbf','.mdf','.ibd','.frm','.myd','.myi','.sql'
$ExtFraca = '.fbk','.gbk','.db','.db3','.cdx','.fpt','.dbt','.mdx','.ndx','.ntx','.ldf','.ndf','.dat','.idx','.hdb','.bak','.udl','.dsn'
$ExtBanco = $ExtForte + $ExtFraca
$ExtConfig = '.ini','.cfg','.conf','.config','.xml','.json'
$ExtPrograma = '.exe','.dll'

W 'RELATÓRIO DA ESTRUTURA DO BANCO DO HIDOCTOR'
W 'Contém só: programas instalados, serviços de banco de dados, nomes de arquivos de banco,'
W 'nomes de tabelas e colunas, tipos e contagens. NÃO contém dados de pacientes.'
W 'Leia antes de enviar. Se aparecer algum nome de paciente ou texto clínico, NÃO envie.'
W ('Gerado em ' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + ' | Windows ' + [Environment]::OSVersion.Version + ' | 64 bits: ' + [Environment]::Is64BitOperatingSystem + ' | PowerShell de 32 bits: ' + (-not [Environment]::Is64BitProcess))

# ---------------------------------------------------------------------------------------------
Secao '1. Programas instalados (HiDoctor, Centralx e bancos de dados)'
$Pastas = New-Object System.Collections.Generic.List[string]
foreach ($extra in $PastasExtras) { if ($extra -and (Test-Path -LiteralPath $extra)) { $Pastas.Add((Resolve-Path -LiteralPath $extra).Path.TrimEnd('\')) } }
foreach ($vista in [Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32) {
  foreach ($colmeia in [Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser) {
    try {
      $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($colmeia, $vista)
      $un = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall')
      if (-not $un) { continue }
      foreach ($nome in $un.GetSubKeyNames()) {
        $k = $un.OpenSubKey($nome)
        $dn = [string]$k.GetValue('DisplayName')
        if ($dn -match 'HiDoctor|Centralx|MySQL|MariaDB|Firebird|InterBase|SQL Server|PostgreSQL|SQLite|Access Database Engine') {
          $loc = [string]$k.GetValue('InstallLocation')
          W ('  ' + $dn + ' | versão ' + [string]$k.GetValue('DisplayVersion') + ' | ' + $loc)
          if ($loc -and (Test-Path -LiteralPath $loc) -and ($dn -match 'HiDoctor|Centralx')) { $Pastas.Add($loc.TrimEnd('\')) }
        }
      }
    } catch { }
  }
}

# ---------------------------------------------------------------------------------------------
Secao '2. Serviços e programas de banco de dados'
$BinMySQL = New-Object System.Collections.Generic.List[string]
try {
  Get-WmiObject Win32_Service -ErrorAction Stop | Where-Object { ($_.Name + ' ' + $_.DisplayName) -match 'mysql|maria|firebird|interbase|MSSQL|SQL Server|postgres' } | ForEach-Object {
    W ('  serviço ' + $_.Name + ' | ' + $_.DisplayName + ' | ' + $_.State + ' | ' + $_.PathName)
    if ($_.PathName -match '^"?([^"]*\\bin)\\mysqld') { $BinMySQL.Add($Matches[1]) }
  }
} catch { W ('  (não consegui listar os serviços: ' + $_.Exception.Message + ')') }
Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match '^(mysqld|mariadbd|fbserver|fb_inet_server|firebird|ibserver|sqlservr|postgres|hidoctor|clinic|medsync)' } | ForEach-Object {
  $caminho = ''; try { $caminho = $_.Path } catch { }
  W ('  processo ' + $_.ProcessName + ' | ' + $caminho)
  if ($caminho -match '^(.*\\bin)\\mysqld') { $BinMySQL.Add($Matches[1]) }
}
try {
  Get-NetTCPConnection -State Listen -ErrorAction Stop | Where-Object { $_.LocalPort -in 1433,3050,3306,3307,5432 } | ForEach-Object { W ('  porta ' + $_.LocalPort + ' aberta (endereço ' + $_.LocalAddress + ')') }
} catch { }

# ---------------------------------------------------------------------------------------------
Secao '3. Fontes de dados ODBC (sem senhas)'
foreach ($vista in [Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32) {
  foreach ($colmeia in [Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser) {
    try {
      $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($colmeia, $vista)
      $odbc = $base.OpenSubKey('Software\ODBC\ODBC.INI')
      if (-not $odbc) { continue }
      foreach ($dsn in $odbc.GetSubKeyNames()) {
        if ($dsn -eq 'ODBC Data Sources') { continue }
        $k = $odbc.OpenSubKey($dsn)
        $partes = @()
        foreach ($v in $k.GetValueNames()) {
          if ($v -match 'pwd|pass|senha') { $partes += ($v + '=***') } else { $partes += ($v + '=' + [string]$k.GetValue($v)) }
        }
        W ('  ' + $dsn + ': ' + ($partes -join '; '))
      }
    } catch { }
  }
}

# ---------------------------------------------------------------------------------------------
Secao '4. Pastas do HiDoctor / Centralx'
$Bases = @(${env:ProgramFiles(x86)}, $env:ProgramW6432, $env:ProgramFiles, 'C:\', 'D:\', [Environment]::GetFolderPath('MyDocuments'), $env:APPDATA, $env:LOCALAPPDATA, $env:ProgramData) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
foreach ($b in $Bases) {
  foreach ($filtro in 'Centralx*', 'HiDoctor*', 'Hi Doctor*') {
    Get-ChildItem -LiteralPath $b -Directory -Filter $filtro -ErrorAction SilentlyContinue | ForEach-Object { $Pastas.Add($_.FullName.TrimEnd('\')) }
  }
}
$Pastas = @($Pastas | Sort-Object -Unique)
# Drop folders inside another listed folder.
$Raizes = @($Pastas | Where-Object { $p = $_; -not ($Pastas | Where-Object { $_ -ne $p -and $p.StartsWith($_ + '\', [StringComparison]::OrdinalIgnoreCase) }) })
if ($Raizes.Count -eq 0) { W '  Nenhuma pasta encontrada com o nome Centralx ou HiDoctor.' }

$ArquivosBanco = New-Object System.Collections.Generic.List[object]
$ArquivosConfig = New-Object System.Collections.Generic.List[object]
$Executaveis = New-Object System.Collections.Generic.List[object]
foreach ($raiz in $Raizes) {
  W ''
  W ('Pasta: ' + $raiz)
  $todos = @(Get-ChildItem -LiteralPath $raiz -Recurse -File -Force -ErrorAction SilentlyContinue)
  $porSub = $todos | Group-Object { $rel = $_.FullName.Substring($raiz.Length).TrimStart('\'); if ($rel.Contains('\')) { $rel.Split('\')[0] + '\' } else { '(raiz)' } }
  foreach ($g in ($porSub | Sort-Object Name)) {
    $tam = ($g.Group | Measure-Object Length -Sum).Sum
    $exts = ($g.Group | Group-Object { $_.Extension.ToLower() } | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object { $e = $_.Name; if (-not $e) { $e = '(sem extensão)' }; $e + ' x' + $_.Count }) -join ', '
    W ('  ' + $g.Name + '  ' + $g.Count + ' arquivos, ' + [math]::Round($tam / 1MB, 1) + ' MB  [' + $exts + ']')
  }
  foreach ($f in $todos) {
    $ext = $f.Extension.ToLower()
    $rel = $f.FullName.Substring($raiz.Length).TrimStart('\')
    $sub = '(raiz)\'; if ($rel.Contains('\')) { $sub = $rel.Split('\')[0] + '\' }
    $profundidade = $rel.Split('\').Count - 1
    if ($ExtForte -contains $ext) { $ArquivosBanco.Add(@{ Raiz = $raiz; Rel = (Curto $rel); Arquivo = $f }) }
    elseif ($ExtFraca -contains $ext) { $ArquivosBanco.Add(@{ Raiz = $raiz; Rel = ($sub + '...\(nome omitido)' + $ext); Arquivo = $f }) }
    elseif ($ExtConfig -contains $ext -and $f.Length -le 65536 -and $profundidade -le 1) { $ArquivosConfig.Add(@{ Raiz = $raiz; Rel = $rel; Arquivo = $f }) }
    elseif ($ext -eq '.exe' -and $profundidade -le 1) { $Executaveis.Add(@{ Rel = $rel; Arquivo = $f }) }
  }
  # Big files of other kinds: only the extension, the size and the first folder (never the name).
  $todos | Where-Object { $_.Length -ge 5MB -and -not (($ExtBanco + $ExtConfig + $ExtPrograma) -contains $_.Extension.ToLower()) } | ForEach-Object {
    $rel = $_.FullName.Substring($raiz.Length).TrimStart('\'); $sub = '(raiz)'; if ($rel.Contains('\')) { $sub = $rel.Split('\')[0] + '\' }
    W ('  arquivo grande: ' + $sub + ' ... ' + $_.Extension.ToLower() + ', ' + [math]::Round($_.Length / 1MB, 1) + ' MB')
  }
}

Secao '5. Programas (.exe) e versões'
foreach ($e in $Executaveis) {
  $vi = $e.Arquivo.VersionInfo
  W ('  ' + $e.Rel + ' | versão ' + $vi.FileVersion + ' | ' + $vi.ProductName + ' ' + $vi.ProductVersion)
}

# ---------------------------------------------------------------------------------------------
function Formato([System.IO.FileInfo]$f) {
  try {
    $fs = [IO.File]::Open($f.FullName, 'Open', 'Read', 'ReadWrite')
    $buf = New-Object byte[] 32
    $n = $fs.Read($buf, 0, 32); $fs.Close()
    if ($n -lt 4) { return 'vazio' }
    $txt = [Text.Encoding]::ASCII.GetString($buf, 0, $n)
    if ($txt.Contains('Standard Jet DB')) { return 'Microsoft Access 97-2003 (Jet)' }
    if ($txt.Contains('Standard ACE DB')) { return 'Microsoft Access 2007+ (ACE)' }
    if ($txt.StartsWith('SQLite format 3')) { return 'SQLite 3' }
    if ($f.Extension -match '\.(fdb|gdb)$' -and $buf[0] -eq 1) { return 'Firebird/InterBase' }
    if ($buf[0] -eq 0xFE -and $buf[1] -eq 0x01) { return 'MySQL (definição de tabela .frm)' }
    if ($f.Extension -eq '.mdf') { return 'SQL Server' }
    if ($f.Extension -eq '.dbf') { return 'dBase/FoxPro (DBF)' }
    if ($f.Extension -eq '.sql') { return 'texto SQL' }
    return ('desconhecido (início ' + (($buf[0..3] | ForEach-Object { $_.ToString('X2') }) -join ' ') + ')')
  } catch { return ('não consegui abrir: ' + $_.Exception.Message) }
}

Secao '6. Arquivos de banco de dados encontrados'
if ($ArquivosBanco.Count -eq 0) { W '  Nenhum.' }
foreach ($a in $ArquivosBanco) {
  $a.Formato = Formato $a.Arquivo
  W ('  ' + $a.Rel + ' | ' + [math]::Round($a.Arquivo.Length / 1MB, 2) + ' MB | ' + $a.Arquivo.LastWriteTime.ToString('yyyy-MM-dd') + ' | ' + $a.Formato)
}

# ---------------------------------------------------------------------------------------------
$TiposOleDb = @{ 2 = 'inteiro curto'; 3 = 'inteiro'; 4 = 'real'; 5 = 'duplo'; 6 = 'moeda'; 7 = 'data'; 11 = 'sim/não'; 17 = 'byte'; 72 = 'GUID'; 128 = 'binário'; 130 = 'texto'; 131 = 'decimal' }
function Estrutura-Access($caminho) {
  foreach ($prov in 'Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0', 'Microsoft.Jet.OLEDB.4.0') {
    $cn = $null
    try {
      $cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$caminho;Mode=Read;Persist Security Info=False")
      $cn.Open()
      W ('  (lido com ' + $prov + ')')
      $cols = $cn.GetSchema('Columns')
      $tabelas = $cols.Rows | Group-Object { $_['TABLE_NAME'] } | Where-Object { $_.Name -notmatch '^(MSys|~)' } | Sort-Object Name
      foreach ($t in $tabelas) {
        $qtd = '?'
        try { $cmd = $cn.CreateCommand(); $cmd.CommandText = 'SELECT COUNT(*) FROM [' + $t.Name + ']'; $qtd = [string]$cmd.ExecuteScalar() } catch { }
        W ('  TABELA ' + $t.Name + ' (' + $qtd + ' linhas)')
        foreach ($c in ($t.Group | Sort-Object { [int]$_['ORDINAL_POSITION'] })) {
          $tipo = $TiposOleDb[[int]$c['DATA_TYPE']]; if (-not $tipo) { $tipo = 'tipo ' + $c['DATA_TYPE'] }
          $tam = [string]$c['CHARACTER_MAXIMUM_LENGTH']; if ($tam) { $tipo = $tipo + '(' + $tam + ')' }
          W ('      ' + $c['COLUMN_NAME'] + ' : ' + $tipo)
        }
      }
      $cn.Close(); return
    } catch {
      $msg = $_.Exception.Message
      if ($cn) { try { $cn.Close() } catch { } }
      if ($msg -match 'senha|password|not a valid password') { W '  Banco protegido por senha: não abri. Me diga, que eu explico o próximo passo.'; return }
      W ('  ' + $prov + ': ' + $msg)
    }
  }
}

function Estrutura-DBF([System.IO.FileInfo]$f) {
  try {
    $fs = [IO.File]::Open($f.FullName, 'Open', 'Read', 'ReadWrite')
    $cab = New-Object byte[] 32; [void]$fs.Read($cab, 0, 32)
    $registros = [BitConverter]::ToUInt32($cab, 4); $tamCab = [BitConverter]::ToUInt16($cab, 8)
    $campos = @()
    $desc = New-Object byte[] 32
    while ($fs.Position + 32 -le $tamCab) {
      [void]$fs.Read($desc, 0, 32)
      if ($desc[0] -eq 0x0D) { break }
      $nome = [Text.Encoding]::ASCII.GetString($desc, 0, 11).Trim([char]0, ' ')
      $campos += ($nome + ':' + [char]$desc[11] + $desc[16])
    }
    $fs.Close()
    W ('  TABELA ' + $f.Name + ' (' + $registros + ' registros): ' + ($campos -join ', '))
  } catch { W ('  ' + $f.Name + ': ' + $_.Exception.Message) }
}

function Executar($exe, $argumentos) {
  try { return (& $exe @argumentos 2>&1 | Out-String) } catch { return ('erro: ' + $_.Exception.Message) }
}

Secao '7. Estrutura (tabelas e colunas, sem dados)'
foreach ($a in $ArquivosBanco) {
  $f = $a.Arquivo
  if ($a.Formato -like 'Microsoft Access*') { W ''; W ('Banco Access: ' + $a.Rel); Estrutura-Access $f.FullName }
  elseif ($a.Formato -like 'dBase*') { Estrutura-DBF $f }
  elseif ($a.Formato -eq 'texto SQL' -and $f.Length -le 20MB) {
    # Only the CREATE TABLE blocks of an SQL file (a structure file such as DataStru\MySQL.sql);
    # INSERT lines, which hold data, are never printed.
    W ''; W ('Arquivo SQL (só os CREATE TABLE): ' + $a.Rel)
    $dentro = $false; $temInsert = $false
    foreach ($l in [IO.File]::ReadLines($f.FullName, [Text.Encoding]::GetEncoding(28591))) {
      if ($l -match '^\s*INSERT\s') { $temInsert = $true; $dentro = $false; continue }
      if ($l -match '^\s*CREATE\s+TABLE') { $dentro = $true }
      if ($dentro) { W ('  ' + $l); if ($l.TrimEnd().EndsWith(';')) { $dentro = $false } }
    }
    if ($temInsert) { W '  (este arquivo também tem dados, INSERT; eles NÃO foram copiados aqui)' }
  }
}
$frm = @($ArquivosBanco | Where-Object { $_.Formato -like 'MySQL*' })
if ($frm.Count -gt 0) {
  W ''; W 'Tabelas MySQL pelos arquivos .frm (nome da pasta = banco):'
  $frm | Group-Object { Split-Path (Split-Path $_.Arquivo.FullName -Parent) -Leaf } | ForEach-Object { W ('  banco ' + $_.Name + ': ' + (($_.Group | ForEach-Object { $_.Arquivo.BaseName }) -join ', ')) }
}

# MySQL server: structure only (mysqldump --no-data) and the approximate row counts.
foreach ($pasta in @('C:\Program Files\MySQL', 'C:\Program Files (x86)\MySQL', 'C:\mysql')) {
  if (Test-Path -LiteralPath $pasta) { Get-ChildItem -LiteralPath $pasta -Recurse -Filter 'mysql.exe' -ErrorAction SilentlyContinue | ForEach-Object { $BinMySQL.Add($_.DirectoryName) } }
}
foreach ($r in $Raizes) { Get-ChildItem -LiteralPath $r -Recurse -Filter 'mysql.exe' -ErrorAction SilentlyContinue | ForEach-Object { $BinMySQL.Add($_.DirectoryName) } }
$BinMySQL = @($BinMySQL | Sort-Object -Unique)
foreach ($bin in $BinMySQL) {
  $mysql = Join-Path $bin 'mysql.exe'; $dump = Join-Path $bin 'mysqldump.exe'
  if (-not (Test-Path -LiteralPath $mysql)) { continue }
  W ''; W ('Servidor MySQL (programas em ' + $bin + ')')
  $conectou = $false
  foreach ($cred in @(@('hiuser', 'aa'), @('root', ''), @('root', 'aa'))) {
    $auth = @('-u', $cred[0]); if ($cred[1]) { $auth += ('-p' + $cred[1]) }
    $bancos = Executar $mysql ($auth + @('-N', '-B', '-e', 'SHOW DATABASES'))
    if ($bancos -match 'ERROR|erro') { continue }
    $conectou = $true
    W ('  conectado como ' + $cred[0])
    foreach ($db in ($bancos -split "`r?`n" | Where-Object { $_ -and $_ -notmatch '^(mysql|information_schema|performance_schema|sys|test)$' })) {
      W ''; W ('  BANCO ' + $db)
      $status = Executar $mysql ($auth + @('-N', '-B', '-e', ('SHOW TABLE STATUS FROM `' + $db + '`')))
      foreach ($l in ($status -split "`r?`n" | Where-Object { $_ })) { $c = $l -split "`t"; if ($c.Count -ge 5) { W ('    tabela ' + $c[0] + ' | ' + $c[1] + ' | cerca de ' + $c[4] + ' linhas') } }
      if (Test-Path -LiteralPath $dump) {
        W '  Estrutura (mysqldump --no-data, nenhum dado):'
        $ddl = Executar $dump ($auth + @('--no-data', '--skip-comments', '--compact', $db))
        foreach ($l in ($ddl -split "`r?`n")) { if ($l -notmatch '^\s*INSERT') { W ('    ' + $l) } }
      }
    }
    break
  }
  if (-not $conectou) { W '  Não consegui entrar no MySQL com as senhas padrão do manual. Me diga, que eu explico o próximo passo.' }
}

# Firebird: metadata only (isql -x).
$isql = @()
foreach ($r in ($Raizes + @('C:\Program Files\Firebird', 'C:\Program Files (x86)\Firebird'))) { if (Test-Path -LiteralPath $r) { $isql += @(Get-ChildItem -LiteralPath $r -Recurse -Filter 'isql.exe' -ErrorAction SilentlyContinue) } }
foreach ($a in ($ArquivosBanco | Where-Object { $_.Formato -like 'Firebird*' })) {
  if ($isql.Count -eq 0) { W ('  Firebird: ' + $a.Rel + ' (sem o programa isql para ler a estrutura)'); continue }
  W ''; W ('Banco Firebird: ' + $a.Rel + ' (isql -x: só a estrutura)')
  $ddl = Executar $isql[0].FullName @('-x', '-user', 'SYSDBA', '-password', 'masterkey', $a.Arquivo.FullName)
  foreach ($l in ($ddl -split "`r?`n")) { if ($l -notmatch '^\s*INSERT') { W ('  ' + $l) } }
}

# SQLite: the schema (sqlite_master) through uv's Python, if the Prontuário is installed here.
$sqlite = @($ArquivosBanco | Where-Object { $_.Formato -eq 'SQLite 3' })
if ($sqlite.Count -gt 0) {
  $uv = $null
  $cmdUv = Get-Command uv -ErrorAction SilentlyContinue
  if ($cmdUv) { $uv = $cmdUv.Source } elseif (Test-Path "$env:USERPROFILE\.local\bin\uv.exe") { $uv = "$env:USERPROFILE\.local\bin\uv.exe" }
  foreach ($a in $sqlite) {
    W ''; W ('Banco SQLite: ' + $a.Rel)
    if (-not $uv) { W '  (sem o uv para ler a estrutura)'; continue }
    # Single quotes only: Windows PowerShell 5.1 does not escape double quotes for native programs.
    $py = "import sqlite3,sys; c=sqlite3.connect(sys.argv[1]); [print(r[0]) for r in c.execute('select sql from sqlite_master where sql is not null')]"
    $esquema = Executar $uv @('run', '--no-project', 'python', '-c', $py, $a.Arquivo.FullName)
    foreach ($l in ($esquema -split "`r?`n")) { W ('  ' + $l) }
  }
}

# ---------------------------------------------------------------------------------------------
Secao '8. Arquivos de configuração (só os nomes das opções; valores só de caminhos e do banco)'
# A setting's value is shown only when its name is about where or how the data is stored (a folder,
# the database, the server, the driver); every other value is hidden, because an .ini can keep
# things like the last patient opened.
$ChavesVisiveis = 'banco|base|database|dados|data|path|caminho|dir|pasta|folder|server|servidor|host|porta|port|driver|provider|dsn|tipo|type|vers|engine|motor'
$ChavesSecretas = 'senha|pass|pwd|key|chave|token|user|usuario|login|nome|name|paciente|ultim|last|recent|hist|mru'
foreach ($c in ($ArquivosConfig | Select-Object -First 40)) {
  if ($c.Arquivo.Extension -notin '.ini', '.cfg', '.conf') { W ('  ' + $c.Rel + ' (' + $c.Arquivo.Length + ' bytes)'); continue }
  W ('  ' + $c.Rel + ':')
  $n = 0
  foreach ($l in [IO.File]::ReadLines($c.Arquivo.FullName, [Text.Encoding]::GetEncoding(28591))) {
    if ($n -ge 80) { W '      ...'; break }
    if ($l -match '^\s*\[.*\]\s*$') { W ('      ' + $l.Trim()); $n++; continue }
    if ($l -match '^\s*([^=;#]+?)\s*=(.*)$') {
      $chave = $Matches[1]; $valor = $Matches[2].Trim()
      if ($chave -match $ChavesSecretas -or $chave -notmatch $ChavesVisiveis) { $valor = '(oculto)' }
      W ('      ' + $chave + '=' + $valor); $n++
    }
  }
}

# ---------------------------------------------------------------------------------------------
Secao '9. PostgreSQL (estrutura das tabelas, sem dados)'
# HiDoctor 10 keeps its data in PostgreSQL. The structure comes from PostgreSQL's own tools, which
# are told to leave the data out: pg_restore --schema-only on a backup file (no connection, no
# password) and, when the server lets this computer in without a password, pg_dump --schema-only
# and the approximate row counts of the statistics. No data row is ever requested, and any COPY or
# INSERT line is dropped anyway.
function Filtrar-Estrutura([string]$texto) {
  $dentro = $false; $tabela = ''
  foreach ($l in ($texto -split "`r?`n")) {
    if ($l -match '^\s*(COPY|INSERT)\s' -or $l -eq '\.') { $dentro = $false; continue }
    if ($l -match '^CREATE (UNLOGGED )?TABLE ') { $dentro = $true }
    if ($dentro) { W ('  ' + $l); if ($l -match '\);\s*$') { $dentro = $false; W '' }; continue }
    if ($l -match '^ALTER TABLE (ONLY )?(\S+)') { $tabela = $Matches[2]; continue }
    if ($l -match '^\s+ADD CONSTRAINT (.*)$' -and $tabela) { W ('  ' + $tabela + ': ' + $Matches[1]); continue }
    if ($l -match '^COMMENT ON (TABLE|COLUMN) ') { W ('  ' + $l) }
  }
}
$PgBins = New-Object System.Collections.Generic.List[string]
$PgServicos = @()
try {
  Get-WmiObject Win32_Service -ErrorAction Stop | Where-Object { $_.PathName -match 'pg_ctl' } | ForEach-Object {
    $bin = ''; $dados = ''
    if ($_.PathName -match '^"?([^"]*\\bin)\\pg_ctl') { $bin = $Matches[1]; $PgBins.Add($bin) }
    if ($_.PathName -match '-D\s+"([^"]+)"') { $dados = $Matches[1] } elseif ($_.PathName -match '-D\s+(\S+)') { $dados = $Matches[1] }
    $PgServicos += @{ Nome = $_.Name; Estado = $_.State; Bin = $bin; Dados = ($dados -replace '/', '\') }
  }
} catch { }
foreach ($r in $Raizes) { Get-ChildItem -LiteralPath $r -Recurse -Filter 'pg_restore.exe' -ErrorAction SilentlyContinue | ForEach-Object { $PgBins.Add($_.DirectoryName) } }
$PgBins = @($PgBins | Where-Object { $_ -and (Test-Path -LiteralPath (Join-Path $_ 'pg_restore.exe')) } | Sort-Object -Unique)
if ($PgServicos.Count -eq 0 -and $PgBins.Count -eq 0) { W '  Nenhum PostgreSQL encontrado.' }
$env:PGCONNECT_TIMEOUT = '5'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

foreach ($sv in $PgServicos) {
  W ('  Serviço ' + $sv.Nome + ' (' + $sv.Estado + '); programas em ' + $sv.Bin + '; dados em ' + $sv.Dados)
  $porta = '5432'
  $usuarios = New-Object System.Collections.Generic.List[string]
  $usuarios.Add('postgres')
  try {
    foreach ($l in (Get-Content -LiteralPath (Join-Path $sv.Dados 'postgresql.conf') -ErrorAction Stop)) { if ($l -match '^\s*port\s*=\s*(\d+)') { $porta = $Matches[1] } }
    W ('  porta ' + $porta)
  } catch { W ('  (postgresql.conf não pôde ser lido: ' + $_.Exception.Message + '; supondo a porta 5432)') }
  try {
    W '  pg_hba.conf (quem pode entrar e como):'
    foreach ($l in (Get-Content -LiteralPath (Join-Path $sv.Dados 'pg_hba.conf') -ErrorAction Stop)) {
      if ($l -match '^\s*(#|$)') { continue }
      W ('      ' + ($l -replace '\s+', ' '))
      $c = ($l.Trim() -split '\s+'); if ($c.Count -ge 3 -and $c[2] -notin 'all', 'replication') { foreach ($u in $c[2].Split(',')) { $usuarios.Add($u) } }
    }
  } catch { W ('  (pg_hba.conf não pôde ser lido: ' + $_.Exception.Message + ')') }
  $psql = Join-Path $sv.Bin 'psql.exe'; $dump = Join-Path $sv.Bin 'pg_dump.exe'
  if (-not (Test-Path -LiteralPath $psql)) { continue }
  $entrou = $false
  foreach ($u in ($usuarios | Sort-Object -Unique)) {
    # -w: never asks for a password; without one that is accepted, this simply fails.
    $bancos = Executar $psql @('-h', '127.0.0.1', '-p', $porta, '-U', $u, '-w', '-d', 'postgres', '-At', '-c', 'select datname from pg_database where not datistemplate')
    if ($bancos -match 'FATAL|ERRO|ERROR|could not|não foi possível|password|senha') { continue }
    $entrou = $true
    W ('  Conectado sem senha como ' + $u + '.')
    foreach ($db in ($bancos -split "`r?`n" | Where-Object { $_ -and $_ -ne 'postgres' })) {
      W ''; W ('  BANCO ' + $db + ' (linhas aproximadas, pelas estatísticas do PostgreSQL):')
      $contagens = Executar $psql @('-h', '127.0.0.1', '-p', $porta, '-U', $u, '-w', '-d', $db, '-At', '-F', ' | ', '-c', 'select schemaname || ''.'' || relname, n_live_tup from pg_stat_user_tables order by 1')
      foreach ($l in ($contagens -split "`r?`n" | Where-Object { $_ })) { W ('    ' + $l) }
      if (Test-Path -LiteralPath $dump) {
        W ''; W '  Estrutura (pg_dump --schema-only, nenhum dado):'
        Filtrar-Estrutura (Executar $dump @('-h', '127.0.0.1', '-p', $porta, '-U', $u, '-w', '--schema-only', '--no-owner', '--no-privileges', $db))
      }
    }
    break
  }
  if (-not $entrou) { W '  O servidor pede senha: a estrutura vem da cópia de segurança, abaixo.' }
}

# Backup files (MedSync): their names are not shown, only how many, the newest one's date and size,
# and its format; a PostgreSQL backup's structure is read with pg_restore --schema-only.
$copias = @()
foreach ($r in $Raizes) { $copias += @(Get-ChildItem -LiteralPath $r -Recurse -File -Filter '*.backup' -ErrorAction SilentlyContinue) }
$copias = @($copias | Sort-Object LastWriteTime -Descending)
if ($copias.Count -gt 0) {
  $nova = $copias[0]
  W ''; W ('  Cópias .backup: ' + $copias.Count + '; a mais recente é de ' + $nova.LastWriteTime.ToString('yyyy-MM-dd HH:mm') + ', ' + [math]::Round($nova.Length / 1MB, 1) + ' MB')
  $cab = New-Object byte[] 8
  try { $fs = [IO.File]::Open($nova.FullName, 'Open', 'Read', 'ReadWrite'); [void]$fs.Read($cab, 0, 8); $fs.Close() } catch { }
  $magica = [Text.Encoding]::ASCII.GetString($cab, 0, 5)
  if ($magica -eq 'PGDMP') {
    W '  Formato: cópia do PostgreSQL (pg_dump, formato personalizado).'
    $lido = $false
    foreach ($bin in $PgBins) {
      $temporario = Join-Path $env:TEMP ('Inspecionar-HiDoctor-esquema-' + [guid]::NewGuid().ToString('N') + '.sql')
      $mensagens = Executar (Join-Path $bin 'pg_restore.exe') @('--schema-only', '--no-owner', '--no-privileges', '-f', $temporario, $nova.FullName)
      $estrutura = ''
      if (Test-Path -LiteralPath $temporario) { $estrutura = [IO.File]::ReadAllText($temporario, [Text.Encoding]::UTF8); Remove-Item -LiteralPath $temporario -Force -ErrorAction SilentlyContinue }
      if ($estrutura -match 'CREATE TABLE') { W ('  Estrutura (pg_restore --schema-only, nenhum dado), com ' + $bin + ':'); Filtrar-Estrutura $estrutura; $lido = $true; break }
      W ('  pg_restore em ' + $bin + ' não leu a cópia: ' + (($mensagens -split "`r?`n" | Select-Object -First 2) -join ' '))
    }
    if (-not $lido) { W '  Não consegui ler a estrutura da cópia.' }
  } elseif ($cab[0] -eq 0x50 -and $cab[1] -eq 0x4B) {
    W '  Formato: arquivo .zip. Conteúdo por extensão (sem nomes):'
    try {
      Add-Type -AssemblyName System.IO.Compression.FileSystem
      $zip = [IO.Compression.ZipFile]::OpenRead($nova.FullName)
      $zip.Entries | Group-Object { [IO.Path]::GetExtension($_.FullName).ToLower() } | Sort-Object Count -Descending | ForEach-Object { W ('      ' + $(if ($_.Name) { $_.Name } else { '(sem extensão)' }) + ' x' + $_.Count + ', ' + [math]::Round((($_.Group | Measure-Object Length -Sum).Sum) / 1MB, 1) + ' MB') }
      $zip.Dispose()
    } catch { W ('      (não consegui abrir: ' + $_.Exception.Message + ')') }
  } else {
    W ('  Formato desconhecido (início ' + (($cab[0..7] | ForEach-Object { $_.ToString('X2') }) -join ' ') + ').')
  }
}

# ---------------------------------------------------------------------------------------------
$destino = $ArquivoDoRelatorio
[IO.File]::WriteAllLines($destino, $Linhas, (New-Object Text.UTF8Encoding $true))
Write-Host ''
Write-Host ('Pronto. O relatório está em: ' + $destino)
if (-not $SemBlocoDeNotas) {
  Write-Host 'Ele vai abrir no Bloco de Notas. Leia antes de enviar.'
  Start-Process notepad.exe -ArgumentList ('"' + $destino + '"')
}
exit 0
