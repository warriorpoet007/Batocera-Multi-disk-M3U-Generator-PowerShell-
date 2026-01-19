<#
PURPOSE: Export a master list of games across Batocera platform folders by reading each platform's gamelist.xml
VERSION: 1.0
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Place this file into the ROMS folder to process all platforms, or in a platform's individual subfolder to process just that one.
- This script does NOT modify any files; it only reads gamelist.xml and produces a CSV report.
- Output CSV is written to the same directory where the PS1 script resides.
- If "Game List.csv" already exists in the output folder, it is deleted and replaced with a newly generated file.
- Multi-disk detection is inferred as:
    - Multi-M3U: the visible entry’s <path> ends in .m3u
    - Multi-XML: the visible entry has 1+ additional entries with the same group key where <hidden>true</hidden> is set
    - Single: neither of the above
- IMPORTANT (robustness):
    - Some gamelist.xml files in the wild can be malformed (mismatched tags, partial writes, etc.).
    - This script will attempt a normal XML parse first, and if that fails it will fall back to a "salvage mode"
      that extracts <game>...</game> blocks and parses them individually.
- XMLState column:
    - Normal: gamelist.xml parsed cleanly as a complete XML document
    - Malformed: gamelist.xml was malformed; entries were extracted by parsing <game> fragments
- Progress / phase output:
    - Prints only major phase steps
    - If running from ROMS root (multi-platform mode), prints per-platform start + finished lines
    - Always prints a final "finished" summary

BREAKDOWN
- Determines runtime mode based on where the script is located:
    - If a gamelist.xml exists in the script directory, treat it as a single-platform run
    - Otherwise, treat the script directory as ROMS root and scan all first-level subfolders for gamelist.xml
- Reads each discovered gamelist.xml and extracts per-game fields:
    - Name (from <name>, with filename fallback if <name> is missing)
    - Path (from <path>)
    - Hidden (from <hidden>, if present; defaults to false)
- Uses a stable group key for multi-disk inference:
    - Primary: <name>
    - Fallback: <path> when <name> is missing/blank (prevents unrelated entries from collapsing into one group)
- Groups entries by the group key to infer multi-disk sets and produce a single row per visible entry
- Generates additional derived columns:
    - Title (Title Case derived from the name; uses safe fallbacks when name is missing)
    - EntryType (Single, Multi-M3U, Multi-XML)
    - DiskCount (total entries in the group: visible + hidden)
    - XMLState (Normal or Malformed)
- Deletes any existing "Game List.csv" in the output folder and writes a newly generated "Game List.csv"
#>

# ==================================================================================================
# SCRIPT STARTUP: PATHS AND TARGET DISCOVERY
# ==================================================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------------------------------
# Script location and run roots
# PURPOSE:
# - Ensure behavior is based on where the script resides:
#     - ROMS root mode when script is placed in ROMS
#     - Single-platform mode when script is placed in a platform folder
# - Ensure output CSV is written next to the script
# --------------------------------------------------------------------------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$startDir  = $scriptDir

# --------------------------------------------------------------------------------------------------
# Phase output helper
# PURPOSE:
# - Emit clear, consistent phase/progress messages to the console
# NOTES:
# - Intentionally not chatty: major steps only + optional per-platform lines in ROMS root mode
# --------------------------------------------------------------------------------------------------
function Write-Phase {
  param([string]$Message)
  Write-Host ""
  Write-Host $Message -ForegroundColor Cyan
}

# --------------------------------------------------------------------------------------------------
# Runtime mode determination
# PURPOSE:
# - Decide whether we are scanning a single platform or the ROMS root
# NOTES:
# - Single-platform mode: script directory contains gamelist.xml
# - ROMS root mode: script directory does not contain gamelist.xml
# --------------------------------------------------------------------------------------------------
$localGamelistPath     = Join-Path $startDir 'gamelist.xml'
$isSinglePlatformMode  = (Test-Path -LiteralPath $localGamelistPath)
$isRomsRootMode        = (-not $isSinglePlatformMode)

Write-Phase "Starting export..."

# ==================================================================================================
# USER CONFIGURATION: PLATFORM NAME TRANSLATIONS
# ==================================================================================================

# Folder -> Friendly Platform Name mapping (extend as needed)
$PlatformNameMap = @{
  '3do'          = '3DO (Panasonic)'
  'abuse'        = 'Abuse SDL (Port)'
  'atom'         = 'Atom (Acorn Computers)'
  'electron'     = 'Electron (Acorn Computers)'
  'advision'     = 'Adventure Vision (Entex)'
  'amiga1200'    = 'Amiga 1200/AGA (Commodore)'
  'amiga500'     = 'Amiga 500/OCS/ECS (Commodore)'
  'amigacd32'    = 'Amiga CD32 (Commodore)'
  'amstradcpc'   = 'Amstrad CPC (Amstrad)'
  'gx4000'       = 'Amstrad GX4000 (Amstrad)'
  'apfm1000'     = 'APF-MP1000/MP-1000/M-1000 (APF Electronics Inc.)'
  'apple2'       = 'Apple II (Apple)'
  'apple2gs'     = 'Apple IIGS (Apple)'
  'arcadia'      = 'Arcadia 2001/et al. (Emerson Radio)'
  'archimedes'   = 'Archimedes (Acorn Computers)'
  'arduboy'      = 'Arduboy (Arduboy)'
  'atari2600'    = 'Atari 2600/VCS (Atari)'
  'atari5200'    = 'Atari 5200 (Atari)'
  'atari7800'    = 'Atari 7800 (Atari)'
  'atari800'     = 'Atari 800 (Atari)'
  'jaguar'       = 'Atari Jaguar (Atari)'
  'jaguarcd'     = 'Atari Jaguar CD (Atari)'
  'lynx'         = 'Atari Lynx (Atari)'
  'atarist'      = 'Atari ST (Atari)'
  'xegs'         = 'Atari XEGS (Atari)'
  'astrocde'     = 'Astrocade (Bally/Midway)'
  'bbc'          = 'BBC Micro/Master/Archimedes (Acorn Computers)'
  'bennugd'      = 'BennuGD (Game Development Suite)'
  'camplynx'     = 'Camputers Lynx (Camputers)'
  'cannonball'   = 'Cannonball (Port)'
  'casloopy'     = 'Casio Loopy (Casio)'
  'pv1000'       = 'Casio PV-1000 (Casio)'
  'catacombgl'   = 'Catacomb GL (Port)'
  'cavestory'    = 'Cave Story (Port)'
  'cdogs'        = 'C-Dogs (Port)'
  'adam'         = 'Coleco Adam (Coleco)'
  'colecovision' = 'ColecoVision (Coleco)'
  'commanderx16' = 'Commander X16 (David Murray)'
  'c128'         = 'Commodore 128 (Commodore)'
  'c64'          = 'Commodore 64 (Commodore)'
  'amigacdtv'    = 'Commodore CDTV (Commodore)'
  'pet'          = 'Commodore PET (Commodore)'
  'cplus4'       = 'Commodore Plus/4 (Commodore)'
  'c20'          = 'Commodore VIC-20/VC-20 (Commodore)'
  'cdi'          = 'Compact Disc Interactive/CD-i (Philips, et al.)'
  'crvision'     = 'CreatiVision/Educat 2002/Dick Smith Wizzard/FunVision (VTech)'
  'daphne'       = 'DAPHNE Laserdisc (Various)'
  'devilutionx'  = 'DevilutionX - Diablo/Hellfire (Port)'
  'dice'         = 'Discrete Integrated Circuit Emulator (Various)'
  'dolphin'      = 'Dolphin (GameCube/Wii Emulator)'
  'dos'          = 'DOSbox (Peter Veenstra/Sjoerd van der Berg)'
  'dxx-rebirth'  = 'DXX Rebirth - Descent/Descent 2 (Port)'
  'easyrpg'      = 'EasyRPG - RPG Maker (Port)'
  'ecwolf'       = 'Wolfenstein 3D (Port)'
  'eduke32'      = 'Duke Nukem 3D (Port)'
  'enterprise'   = 'Enterprise (Enterprise Computers)'
  'channelf'     = 'Fairchild Channel F (Fairchild)'
  'fallout1-ce'  = 'Fallout CE (Port)'
  'fallout2-ce'  = 'Fallout2 CE (Port)'
  'fds'          = 'Family Computer Disk System/Famicom (Nintendo)'
  'fbneo'        = 'FinalBurn Neo (Various)'
  'flash'        = 'Flashpoint - Adobe Flash (Bluemaxima)'
  'flatpak'      = 'Flatpak (Linux)'
  'fmtowns'      = 'FM Towns/Towns Marty (Fujitsu)'
  'fm7'          = 'Fujitsu Micro 7 (Fujitsu)'
  'fpinball'     = 'Future Pinball (Port)'
  'gamate'       = 'Gamate/Super Boy/Super Child Prodigy (Bit Corporation)'
  'gameandwatch' = 'Game & Watch (Nintendo)'
  'gb'           = 'Game Boy (Nintendo)'
  'gb2players'   = 'Game Boy 2 Players (Nintendo)'
  'gba'          = 'Game Boy Advance (Nintendo)'
  'gbc'          = 'Game Boy Color (Nintendo)'
  'gbc2players'  = 'Game Boy Color 2 Players (Nintendo)'
  'gamegear'     = 'Game Gear (Sega)'
  'gmaster'      = 'Game Master/Systema 2000/Super Game/Game Tronic (Hartung, et al.)'
  'gamepock'     = 'Game Pocket Computer (Epoch)'
  'gamecom'      = 'Game.com (Tiger Electronics)'
  'gp32'         = 'GP32 (Game Park)'
  'gzdoom'       = 'GZDoom - Boom/Chex Quest/Heretic/Hexen/Strife (Port)'
  'lcdgames'     = 'Handheld LCD Games (Various)'
  'hurrican'     = 'Hurrican (Port)'
  'hcl'          = 'Hydra Castle Labyrinth (Port)'
  'ikemen'       = 'Ikemen Go (Port)'
  'intellivision'= 'Intellivision (Mattel)'
  'fury'         = 'Ion Fury (Port)'
  'sgb-msu1'     = 'LADX-MSU1 (Nintendo)'
  'laser310'     = 'Laser 310 (Video Technology (VTech))'
  'lowresnx'     = 'Lowres NX (Timo Kloss)'
  'lutro'        = 'Lutro (Port)'
  'mugen'        = 'M.U.G.E.N (Port)'
  'macintosh'    = 'Macintosh 128K (Apple)'
  'odyssey2'     = 'Odyssey 2/Videopac G7000 (Magnavox/Philips)'
  'vgmplay'      = 'MAME Video Game Music Player (Various)'
  'megaduck'     = 'Mega Duck/Cougar Boy (Welback Holdings)'
  'msxturbor'    = 'Microsoft MSX turboR (Microsoft)'
  'msx1'         = 'Microsoft MSX1 (Microsoft)'
  'msx2'         = 'Microsoft MSX2 (Microsoft)'
  'msx2+'        = 'Microsoft MSX2plus (Microsoft)'
  'xbox'         = 'Microsoft Xbox (Microsoft)'
  'xbox360'      = 'Microsoft Xbox 360 (Microsoft)'
  'moonlight'    = 'Moonlight (Port)'
  'mrboom'       = 'Mr. Boom (Port)'
  'msu-md'       = 'MSU-MD (Sega)'
  'mame'         = 'Multiple Arcade Machine Emulator (Various)'
  'namco2x6'     = 'Namco System 246 (Sony / Namco)'
  'ports'        = 'Native ports (Linux)'
  'pc60'         = 'NEC PC-6000 (NEC)'
  'pc88'         = 'NEC PC-8800 (NEC)'
  'pc98'         = 'NEC PC-9800/PC-98 (NEC)'
  'pcfx'         = 'NEC PC-FX (NEC)'
  'neogeo'       = 'Neo Geo (SNK)'
  'neogeocd'     = 'Neo Geo CD (SNK)'
  'ngp'          = 'Neo Geo Pocket (SNK)'
  'ngpc'         = 'Neo Geo Pocket Color (SNK)'
  '3ds'          = 'Nintendo 3DS (Nintendo)'
  'n64'          = 'Nintendo 64 (Nintendo)'
  'n64dd'        = 'Nintendo 64DD (Nintendo)'
  'nds'          = 'Nintendo DS (Nintendo)'
  'nes'          = 'Nintendo Entertainment System/Famicom (Nintendo)'
  'gamecube'     = 'Nintendo GameCube (Nintendo)'
  'wii'          = 'Nintendo Wii (Nintendo)'
  'wiiu'         = 'Nintendo Wii U (Nintendo)'
  'openbor'      = 'Open Beats of Rage (Port)'
  'openjazz'     = 'Openjazz (Port)'
  'oricatmos'    = 'Oric Atmos (Tangerine Computer Systems)'
  'multivision'  = 'Othello_Multivision (Tsukuda Original)'
  'pcenginecd'   = 'PC Engine CD-ROM2/Duo R/Duo RX/TurboGrafx CD/TurboDuo (NEC)'
  'supergrafx'   = 'PC Engine/SuperGrafx/PC Engine 2 (NEC)'
  'pcengine'     = 'PC Engine/TurboGrafx-16 (NEC)'
  'pdp1'         = 'PDP-1 (Digital Equipment Corporation)'
  'videopacplus' = 'Philips Videopac+ G7400/G7420 (Philips)'
  'pico8'        = 'PICO-8 fantasy console (Lexaloffle Games)'
  'psp'          = 'PlayStation Portable (Sony)'
  'psvita'       = 'PlayStation Vita (Sony)'
  'plugnplay'    = 'Plug ''n'' Play/Handheld TV Games (Various)'
  'pokemini'     = 'Pokemon Mini (Nintendo)'
  'prboom'       = 'Proff Boom (Port)'
  'pygame'       = 'Python Games (Port)'
  'pyxel'        = 'Pyxel fantasy console (Takashi Kitao)'
  'raze'         = 'Raze (Port)'
  'reminiscence' = 'Reminiscence (Flashback Emulator) (Port)'
  'retroarch'    = 'RetroArch - Liberato Cores (Hans-Kristian "Themaister" Arntzen)'
  'xrick'        = 'Rick Dangerous (Port)'
  'samcoupe'     = 'SAM Coupe (Miles Gordon Technology)'
  'atomiswave'   = 'Sammy Atomiswave (Sammy)'
  'satellaview'  = 'Satellaview (Nintendo)'
  'scummvm'      = 'ScummVM (Ludvig Strigeus/Vincent Hamm)'
  'sdlpop'       = 'SDLPoP - Prince of Persia (Port)'
  'sega32x'      = 'Sega 32X (Sega)'
  'segacd'       = 'Sega CD/Mega CD (Sega)'
  'dreamcast'    = 'Sega Dreamcast (Sega)'
  'megadrive'    = 'Sega Genesis/Mega Drive (Sega)'
  'lindbergh'    = 'Sega Lindbergh (Sega)'
  'mastersystem' = 'Sega Master System/Mark III (Sega)'
  'mame/model1'  = 'Sega Model 1 (Sega)'
  'model2'       = 'Sega Model 2 (Sega)'
  'model3'       = 'Sega Model 3 (Sega)'
  'naomi'        = 'Sega NAOMI (Sega)'
  'naomi2'       = 'Sega NAOMI 2 (Sega)'
  'pico'         = 'Sega Pico (Sega)'
  'saturn'       = 'Sega Saturn (Sega)'
  'sg1000'       = 'Sega SG-1000/SG-1000 II/SC-3000 (Sega)'
  'x1'           = 'Sharp X1 (Sharp)'
  'x68000'       = 'Sharp X68000 (Sharp)'
  'zx81'         = 'Sinclair ZX81 (Sinclair)'
  'singe'        = 'SINGE (Various)'
  'socrates'     = 'Socrates (VTech)'
  'solarus'      = 'Solarus (Port)'
  'psx'          = 'Sony PlayStation (Sony)'
  'ps2'          = 'Sony PlayStation 2 (Sony)'
  'ps3'          = 'Sony PlayStation 3 (Sony)'
  'ps4'          = 'Sony PlayStation 4 (Sony)'
  'spectravideo' = 'Spectravideo (Spectravideo)'
  'sonicretro'   = 'Star Engine/Sonic Retro Engine (Port)'
  'steam'        = 'Steam (Valve)'
  'sufami'       = 'SuFami Turbo (Bandai)'
  'supracan'     = 'Super A''Can (Funtech Entertainment)'
  'scv'          = 'Super Cassette Vision (Epoch Co.)'
  'sgb'          = 'Super Game Boy (Nintendo)'
  'superbroswar' = 'Super Mario War (Port)'
  'snes-msu1'    = 'Super NES CD-ROM/SNES MSU-1 (Nintendo)'
  'snes'         = 'Super Nintendo Entertainment System (Nintendo)'
  'vis'          = 'Tandy Video Information System (Tandy / Memorex)'
  'thomson'      = 'Thomson MO/TO Series Computer (Thomson)'
  'ti99'         = 'TI-99/4/4A (Texas Instruments)'
  'tic80'        = 'TIC-80 fantasy console (Vadim Grigoruk)'
  'tutor'        = 'Tomy Tutor/Pyuta/Grandstand Tutor (Tomy)'
  'traider1'     = 'TR1X - Tomb Raider 1 (Port)'
  'traider2'     = 'TR2X - Tomb Rauder 2 (Port)'
  'triforce'     = 'Triforce (Namco/Sega/Nintendo)'
  'coco'         = 'TRS-80/Tandy Color Computer (Tandy/RadioShack)'
  'tyrquake'     = 'TyrQuake - Quake 1 (Port)'
  'uzebox'       = 'Uzebox Open-Source Console (Alec Bourque)'
  'vsmile'       = 'V.Smile (TV LEARNING SYSTEM) (VTech)'
  'vectrex'      = 'Vectrex (Milton Bradley)'
  'vc4000'       = 'Video Computer 4000 (Interton)'
  'vircon32'     = 'Vircon32 virtual console (Carra)'
  'virtualboy'   = 'Virtual Boy (Nintendo)'
  'vpinball'     = 'Visual Pinball (Port)'
  'voxatron'     = 'Voxatron fantasy console (Lexaloffle Games)'
  'wasm4'        = 'WASM4 fantasy console (Aduros)'
  'supervision'  = 'Watara Supervision (Watara)'
  'windows'      = 'WINE (Bob Amstadt/Alexandre Julliard)'
  'wswan'        = 'WonderSwan (Bandai)'
  'wswanc'       = 'WonderSwan Color (Bandai)'
  'xash3d_fwgs'  = 'Xash3D FWGS - Valve Games (Port)'
  'zxspectrum'   = 'ZX Spectrum (Sinclair)'
}

# ==================================================================================================
# FUNCTIONS
# ==================================================================================================

# --- FUNCTION: Get-FriendlyPlatformName ---
# PURPOSE:
# - Translate a platform folder name (e.g., "psx") into a friendly platform name (e.g., "PlayStation 1").
# NOTES:
# - Falls back to returning the folder name if no translation is present in $PlatformNameMap.
function Get-FriendlyPlatformName {
  param([string]$PlatformFolder)
  if ($PlatformNameMap.ContainsKey($PlatformFolder)) { return $PlatformNameMap[$PlatformFolder] }
  return $PlatformFolder
}

# --------------------------------------------------------------------------------------------------
# Mode banner
# PURPOSE:
# - Provide an immediate, unambiguous indication of what scope will be processed
# NOTES:
# - ROMS root platform count is printed after discovery is complete
# - Single-platform prints folder + friendly platform name immediately
# --------------------------------------------------------------------------------------------------
if ($isSinglePlatformMode) {
  $modePlatformFolder = Split-Path -Leaf $startDir
  $modePlatformName   = Get-FriendlyPlatformName $modePlatformFolder
  Write-Host ("MODE: Single-platform ({0} / {1})" -f $modePlatformFolder, $modePlatformName) -ForegroundColor Green
} else {
  Write-Host "MODE: ROMS root (discovering platforms...)" -ForegroundColor Green
}

# --- FUNCTION: Convert-ToTitleCaseSafe ---
# PURPOSE:
# - Convert a string into Title Case.
# NOTES:
# - Preserves tokens that look like acronyms/codes or contain digits.
function Convert-ToTitleCaseSafe {
  param([object]$Text)

  if ($null -eq $Text) { return '' }
  $t = ([string]$Text).Trim()
  if ([string]::IsNullOrWhiteSpace($t)) { return $t }

  $ti = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
  $tokens = $t -split '(\s+)'

  $out = foreach ($tok in $tokens) {
    if ($tok -match '^\s+$') { $tok; continue }
    if ($tok -match '\d' -or ($tok -cmatch '^[A-Z0-9&\-\+]{2,}$')) { $tok }
    else { $ti.ToTitleCase($tok.ToLowerInvariant()) }
  }

  return ($out -join '')
}

# --- FUNCTION: Get-DisplayNameFallback ---
# PURPOSE:
# - Provide a stable display name when <name> is missing/blank by falling back to the filename from <path>.
# NOTES:
# - Returns "" if neither a usable name nor a usable path is available.
function Get-DisplayNameFallback {
  param(
    [AllowNull()][AllowEmptyString()][string]$Name,
    [AllowNull()][AllowEmptyString()][string]$Path
  )

  $n = [string]$Name
  $p = [string]$Path

  if (-not [string]::IsNullOrWhiteSpace($n)) { return $n }

  if (-not [string]::IsNullOrWhiteSpace($p)) {
    try {
      return [System.IO.Path]::GetFileNameWithoutExtension($p)
    } catch {
      return $p
    }
  }

  return ""
}

# --- FUNCTION: Get-TitleForOutput ---
# PURPOSE:
# - Produce the CSV Title value with safe fallbacks for rare malformed/partial entries.
# NOTES:
# - Fallback ladder:
#     1) Title Case of the resolved display name
#     2) Raw resolved display name as-is
#     3) "(Untitled)" if nothing is available
function Get-TitleForOutput {
  param(
    [AllowNull()][AllowEmptyString()][string]$ResolvedName
  )

  $r = [string]$ResolvedName
  $tc = Convert-ToTitleCaseSafe -Text $r

  if (-not [string]::IsNullOrWhiteSpace($tc)) { return $tc }
  if (-not [string]::IsNullOrWhiteSpace($r))  { return $r }

  return "(Untitled)"
}

# --- FUNCTION: Get-PlatformTargets ---
# PURPOSE:
# - Determine which platform folder(s) to scan based on where the script resides.
# NOTES:
# - If the script directory contains gamelist.xml, treat it as a single-platform run.
# - Otherwise, treat the script directory as ROMS root and scan all first-level subfolders for gamelist.xml.
function Get-PlatformTargets {
  param([string]$StartDir)

  $start = (Resolve-Path -LiteralPath $StartDir).Path
  $localGamelist = Join-Path $start 'gamelist.xml'

  if (Test-Path -LiteralPath $localGamelist) {
    return @([pscustomobject]@{
      PlatformFolder = (Split-Path -Leaf $start)
      GamelistPath   = $localGamelist
    })
  }

  $targets = @()
  Get-ChildItem -LiteralPath $start -Directory | ForEach-Object {
    $g = Join-Path $_.FullName 'gamelist.xml'
    if (Test-Path -LiteralPath $g) {
      $targets += [pscustomobject]@{
        PlatformFolder = $_.Name
        GamelistPath   = $g
      }
    }
  }

  return $targets
}

# --- FUNCTION: Get-XmlNodeText ---
# PURPOSE:
# - Safely read an XML child node's InnerText without throwing when missing.
# NOTES:
# - Returns "" when the requested child node does not exist.
# - Uses local-name() matching so default namespaces do not break child selection.
# - Always returns a string.
function Get-XmlNodeText {
  param(
    [Parameter(Mandatory=$true)][System.Xml.XmlNode]$Node,
    [Parameter(Mandatory=$true)][string]$ChildName
  )

  if ($null -eq $Node) { return '' }
  if ([string]::IsNullOrWhiteSpace($ChildName)) { return '' }

  $child = $null
  try {
    $child = $Node.SelectSingleNode("*[local-name()='$ChildName']")
  } catch {
    $child = $null
  }

  if ($null -eq $child) { return '' }

  $text = ''
  try { $text = [string]$child.InnerText } catch { $text = '' }
  return $text.Trim()
}

# --- FUNCTION: Read-Gamelist ---
# PURPOSE:
# - Read a gamelist.xml and return a flat list of entry objects with Name/Path/Hidden fields.
# NOTES:
# - StrictMode-safe:
#     - Missing <hidden> defaults to $false
#     - Missing <name> falls back to filename from <path> for display purposes
# - Normal XML parse:
#     - Uses XmlDocument + XPath and correctly enumerates each <game> node
# - Malformed gamelist handling:
#     - If the XML document is not well-formed, normal XML parsing will fail.
#     - In that case, this function falls back to a fragment parse of each <game> block.
# - GroupKey:
#     - Used for grouping entries into sets
#     - Primary: resolved name
#     - Fallback: path when name is missing/blank (prevents unrelated entries collapsing into one group)
# - XMLState:
#     - Returned objects include XMLState = Normal or Malformed so the CSV can show broken gamelist sources.
# - Always returns an array (even if 0 or 1 item).
function Read-Gamelist {
  param([string]$PlatformFolder, [string]$GamelistPath)

  $raw = Get-Content -LiteralPath $GamelistPath -Raw
  if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

  $out = @()

  $doc = $null
  $parsedOk = $false
  try {
    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.LoadXml($raw)
    $parsedOk = $true
  } catch {
    $parsedOk = $false
  }

  if ($parsedOk -and $null -ne $doc) {

    $nodes = $null
    try {
      $nodes = $doc.SelectNodes("//*[local-name()='game']")
    } catch {
      $nodes = $null
    }

    if ($null -ne $nodes) {
      foreach ($node in $nodes) {

        $name   = Get-XmlNodeText -Node $node -ChildName 'name'
        $path   = Get-XmlNodeText -Node $node -ChildName 'path'
        $hidden = Get-XmlNodeText -Node $node -ChildName 'hidden'

        $resolvedName = Get-DisplayNameFallback -Name $name -Path $path
        $groupKey     = if (-not [string]::IsNullOrWhiteSpace($resolvedName)) { $resolvedName } else { [string]$path }

        $out += [pscustomobject]@{
          PlatformFolder = $PlatformFolder
          NameResolved   = [string]$resolvedName
          PathRaw        = [string]$path
          Hidden         = ($hidden -match '^(true|1|yes)$')
          XMLState       = 'Normal'
          GroupKey       = [string]$groupKey
        }
      }
    }

    return $out
  }

  $gameBlocks = [regex]::Matches($raw, '(?is)<game\b[^>]*>.*?</game>')

  foreach ($m in $gameBlocks) {

    $block = $m.Value
    if ([string]::IsNullOrWhiteSpace($block)) { continue }

    $fragDoc = $null
    try {
      $fragDoc = New-Object System.Xml.XmlDocument
      $fragDoc.XmlResolver = $null
      $fragDoc.LoadXml("<root>$block</root>")
    } catch {
      continue
    }

    $node = $null
    try {
      $node = $fragDoc.SelectSingleNode("//*[local-name()='game']")
    } catch {
      $node = $null
    }

    if ($null -eq $node) { continue }

    $name   = Get-XmlNodeText -Node $node -ChildName 'name'
    $path   = Get-XmlNodeText -Node $node -ChildName 'path'
    $hidden = Get-XmlNodeText -Node $node -ChildName 'hidden'

    $resolvedName = Get-DisplayNameFallback -Name $name -Path $path
    $groupKey     = if (-not [string]::IsNullOrWhiteSpace($resolvedName)) { $resolvedName } else { [string]$path }

    $out += [pscustomobject]@{
      PlatformFolder = $PlatformFolder
      NameResolved   = [string]$resolvedName
      PathRaw        = [string]$path
      Hidden         = ($hidden -match '^(true|1|yes)$')
      XMLState       = 'Malformed'
      GroupKey       = [string]$groupKey
    }
  }

  return $out
}

# --- FUNCTION: Write-CsvUtf8NoBom ---
# PURPOSE:
# - Write a comma-delimited CSV as UTF-8 without BOM for reliable Excel opening
# NOTES:
# - Avoids UTF-8 BOM characters appearing as ï»¿ when the file is interpreted incorrectly
# - Uses ConvertTo-Csv to guarantee comma delimiter and consistent quoting
function Write-CsvUtf8NoBom {
  param(
    [Parameter(Mandatory=$true)]$Rows,
    [Parameter(Mandatory=$true)][string]$Path
  )

  $csvLines = @($Rows | ConvertTo-Csv -NoTypeInformation)

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllLines($Path, $csvLines, $utf8NoBom)
}

# ==================================================================================================
# PHASE 1: TARGET DISCOVERY
# ==================================================================================================

Write-Phase "Discovering platform folders and locating gamelist.xml files..."

$targets = @(Get-PlatformTargets -StartDir $startDir)

if (@($targets).Count -eq 0) {
  Write-Warning "No gamelist.xml found. Run from /roms or a platform folder that contains gamelist.xml."
  return
}

Write-Host "Found $($targets.Count) platform(s) with gamelist.xml"

if ($isRomsRootMode) {
  Write-Host ("MODE: ROMS root ({0} platform(s))" -f $targets.Count) -ForegroundColor Green
}

# ==================================================================================================
# PHASE 2: READ + GROUP ENTRIES / BUILD OUTPUT ROWS
# ==================================================================================================

Write-Phase "Reading gamelist.xml and collecting game entries..."

$rows = @()

foreach ($t in $targets) {

  if ($isRomsRootMode) {
    Write-Host ""
    Write-Host "PLATFORM: $($t.PlatformFolder)" -ForegroundColor Green
  }

  $platformFolder = [string]$t.PlatformFolder
  $platformName   = Get-FriendlyPlatformName $platformFolder

  $entries = @(Read-Gamelist $platformFolder $t.GamelistPath)
  if ($entries.Count -eq 0) {
    if ($isRomsRootMode) {
      Write-Host "No entries found (skipping)." -ForegroundColor Yellow
      Write-Host "Finished platform: $($t.PlatformFolder)" -ForegroundColor Cyan
    }
    continue
  }

  foreach ($group in @($entries | Group-Object GroupKey)) {

    $items        = @($group.Group)
    $hiddenItems  = @($items | Where-Object { $_.Hidden })
    $visibleItems = @($items | Where-Object { -not $_.Hidden })

    if ($visibleItems.Count -eq 0 -and $items.Count -gt 0) {
      $visibleItems = @($items | Select-Object -First 1)
    }

    foreach ($g in $visibleItems) {

      $pathStr = [string]$g.PathRaw
      $nameStr = [string]$g.NameResolved

      $entryType = 'Single'
      if ($pathStr -match '(?i)\.m3u$') {
        $entryType = 'Multi-M3U'
      }
      elseif ($hiddenItems.Count -gt 0) {
        $entryType = 'Multi-XML'
      }
      elseif ($items.Count -gt 1) {
        $entryType = 'Multi-XML'
      }

      $title = Get-TitleForOutput -ResolvedName $nameStr

      # Build output row in the exact column order expected for the CSV export
      $rows += [pscustomobject]@{
        Title          = $title
        PlatformName   = [string]$platformName
        EntryType      = [string]$entryType
        DiskCount      = [int]$items.Count
        PlatformFolder = $platformFolder
        FilePath       = $pathStr
        XMLState       = [string]$g.XMLState
      }
    }
  }

  if ($isRomsRootMode) {
    Write-Host "Finished platform: $($t.PlatformFolder)" -ForegroundColor Cyan
  }
}

Write-Phase "Finished collecting game entries."

# ==================================================================================================
# PHASE 3: SORT + EXPORT CSV
# ==================================================================================================

Write-Phase "Sorting and exporting CSV..."

$rows = @($rows | Sort-Object PlatformName, Title)

$outPath = Join-Path $startDir 'Game List.csv'

# --------------------------------------------------------------------------------------------------
# Output file handling
# PURPOSE:
# - Ensure any prior output file is removed so the generated CSV is a clean overwrite
# NOTES:
# - If the output file is locked (e.g., open in Excel), print a friendly message and abort cleanly
# --------------------------------------------------------------------------------------------------
if (Test-Path -LiteralPath $outPath) {
  try {
    Remove-Item -LiteralPath $outPath -Force -ErrorAction Stop
  } catch {
    Write-Host ""
    Write-Host "Cannot write 'Game List.csv' because the existing file is locked (likely open in Excel)." -ForegroundColor Yellow
    Write-Host "Script aborted." -ForegroundColor Yellow
    return
  }
}

Write-CsvUtf8NoBom -Rows $rows -Path $outPath

# --------------------------------------------------------------------------------------------------
# Malformed gamelist warning and gamelist path linkage
# PURPOSE:
# - Alert the user when one or more platforms required malformed XML handling
# - Include the gamelist.xml path for quick remediation
# NOTES:
# - This does not affect CSV output; it is console diagnostics only
# --------------------------------------------------------------------------------------------------
$malformedPlatforms = @(
  $rows |
    Where-Object { $_.XMLState -eq 'Malformed' } |
    Select-Object -ExpandProperty PlatformFolder -Unique |
    Sort-Object
)

if ($malformedPlatforms.Count -gt 0) {

  $gamelistPathByPlatform = @{}
  foreach ($t in $targets) {
    if ($null -ne $t -and $null -ne $t.PlatformFolder -and $null -ne $t.GamelistPath) {
      $gamelistPathByPlatform[[string]$t.PlatformFolder] = [string]$t.GamelistPath
    }
  }

  Write-Host ""
  Write-Host "WARNING: Malformed gamelist.xml detected for the following platform(s):" -ForegroundColor Yellow

  foreach ($p in $malformedPlatforms) {
    if ($gamelistPathByPlatform.ContainsKey([string]$p)) {
      Write-Host ("  - {0} ({1})" -f $p, $gamelistPathByPlatform[[string]$p]) -ForegroundColor Yellow
    } else {
      Write-Host "  - $p" -ForegroundColor Yellow
    }
  }

  Write-Host ""
  Write-Host "Game entries were recovered successfully, but the source files" -ForegroundColor Yellow
  Write-Host "should be regenerated or repaired (e.g. via a gamelist update/scrape)." -ForegroundColor Yellow
}

# --------------------------------------------------------------------------------------------------
# Final summary output
# PURPOSE:
# - Provide a clear completion line and key output facts
# NOTES:
# - @($rows).Count forces array semantics under StrictMode even when only one row exists
# --------------------------------------------------------------------------------------------------
Write-Phase "Finished."
Write-Host "CSV written to: $outPath" -ForegroundColor Green
Write-Host "Total games exported: $(@($rows).Count)" -ForegroundColor Green
