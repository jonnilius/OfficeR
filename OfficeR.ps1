
# OfficeR - Office Deployment Tool GUI
# Credits: https://github.com/gravesoft

$env:PSModulePath += ";$PSScriptRoot\Modules"
Import-Module FormBuilder

# ════ Konfiguration ═══════════════════════════════════════════════════════════
# Programmname und Version
$name       = "OfficeR "
$version    = "1.2.0"

# Einstellungen
$RunAsAdmin = $false
$hideShell  = $true

# Programmvariablen
$global:restartScript   = $false
$isAdmin                = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
$TempPath        = $env:TEMP
$TempFile        = Join-Path $TempPath "OfficeSetup.exe"
$TempImage       = Join-Path $TempPath "OfficeSetup.img"

# ════ Office Produkte und Keys ════════════════════════════════════════════════
$Offices = [ordered]@{
    "Retail" = [ordered]@{
        "Office365"  = @("O365AppsBasic", "O365Business", "O365EduCloud", "O365HomePrem", "O365ProPlus", "O365SmallBusPrem")
        "Office2024" = @("Access2024", "Excel2024", "Home2024", "HomeBusiness2024", "Outlook2024", "PowerPoint2024", "ProjectPro2024", "ProjectStd2024", "ProPlus2024", "VisioPro2024", "VisioStd2024", "Word2024")
        "Office2021" = @("Access2021", "AccessRuntime2021", "Excel2021", "HomeBusiness2021", "HomeStudent2021", "OneNoteFree2021", "OneNote2021", "Outlook2021", "Personal2021", "PowerPoint2021", "ProPlus2021", "Professional2021", "ProjectPro2021", "ProjectStd2021", "Publisher2021", "SkypeforBusiness2021", "Standard2021", "VisioPro2021", "VisioStd2021", "Word2021")
        "Office2019" = @("Access2019", "AccessRuntime2019", "Excel2019", "HomeBusiness2019", "HomeStudentARM2019", "HomeStudentPlusARM2019", "HomeStudent2019", "Outlook2019", "Personal2019", "PowerPoint2019", "ProPlus2019", "Professional2019", "ProjectPro2019", "ProjectStd2019", "Publisher2019", "SkypeforBusiness2019", "SkypeforBusinessEntry2019", "Standard2019", "VisioPro2019", "VisioStd2019", "Word2019")
        "Office2016" = @("Access", "AccessRuntime", "Excel", "HomeBusinessPipc", "HomeBusiness", "HomeStudentARM", "HomeStudentPlusARM", "HomeStudent", "HomeStudentVNext", "Mondo", "OneNoteFree", "OneNote", "Outlook", "PersonalPipc", "Personal", "PowerPoint", "ProPlus", "ProfessionalPipc", "Professional", "ProjectPro", "ProjectStd", "Publisher", "SkypeServiceBypass", "SkypeforBusinessEntry", "SkypeforBusiness", "Standard", "VisioPro", "VisioStd", "Word")
    }
    "Volume" = [ordered]@{
        "Office2024" = @("Access2024", "Excel2024", "Outlook2024", "PowerPoint2024", "ProjectPro2024", "ProjectStd2024", "ProPlus2024", "SkypeforBusiness2024", "Standard2024", "VisioPro2024", "VisioStd2024", "Word2024")
        "Office2021" = @("Access2021", "Excel2021", "OneNote2021", "Outlook2021", "PowerPoint2021", "ProPlus2021", "ProPlusSPLA2021", "ProjectPro2021", "ProjectStd2021", "Publisher2021", "SkypeforBusiness2021", "Standard2021", "StandardSPLA2021", "VisioPro2021", "VisioStd2021", "Word2021")
        "Office2019" = @("Access2019", "Excel2019", "Outlook2019", "PowerPoint2019", "ProPlus2019", "ProjectPro2019", "ProjectStd2019", "Publisher2019", "SkypeforBusiness2019", "Standard2019", "VisioPro2019", "VisioStd2019", "Word2019")
        "Office2016" = @("Access", "Excel", "Mondo", "OneNote", "Outlook", "PowerPoint", "ProPlus", "ProjectProX", "ProjectPro", "ProjectStdX", "ProjectStd", "Publisher", "SkypeforBusiness", "Standard", "VisioProX", "VisioPro", "VisioStdX", "VisioStd", "Word")
    }
}
$Keys = @{
    "Office365" = @{
        "O365AppsBasicRetail"       = "3HYJN-9KG99-F8VG9-V3DT8-JFMHV"
        "O365BusinessRetail"        = "Y9NF9-M2QWD-FF6RJ-QJW36-RRF2T"
        "O365EduCloudRetail"        = "W62NQ-267QR-RTF74-PF2MH-JQMTH"
        "O365HomePremRetail"        = "3NMDC-G7C3W-68RGP-CB4MH-4CXCH"
        "O365ProPlusRetail"         = "H8DN8-Y2YP3-CR9JT-DHDR9-C7GP3"
        "O365SmallBusPremRetail"    = "2QCNB-RMDKJ-GC8PB-7QGQV-7QTQJ"
    }
    "Office2016" = @{
        "AccessRetail"                  = "WHK4N-YQGHB-XWXCC-G3HYC-6JF94"
        "AccessRuntimeRetail"           = "RNB7V-P48F4-3FYY6-2P3R3-63BQV"
        "AccessVolume"                  = "JJ2Y4-N8KM3-Y8KY3-Y22FR-R3KVK"
        "ExcelRetail"                   = "RKJBN-VWTM2-BDKXX-RKQFD-JTYQ2"
        "ExcelVolume"                   = "FVGNR-X82B2-6PRJM-YT4W7-8HV36"
        "HomeBusinessPipcRetail"        = "2WQNF-GBK4B-XVG6F-BBMX7-M4F2Y"
        "HomeBusinessRetail"            = "HM6FM-NVF78-KV9PM-F36B8-D9MXD"
        "HomeStudentARMRetail"          = "PBQPJ-NC22K-69MXD-KWMRF-WFG77"
        "HomeStudentPlusARMRetail"      = "6F2NY-7RTX4-MD9KM-TJ43H-94TBT"
        "HomeStudentRetail"             = "PNPRV-F2627-Q8JVC-3DGR9-WTYRK"
        "HomeStudentVNextRetail"        = "YWD4R-CNKVT-VG8VJ-9333B-RC3B8"
        "MondoRetail"                   = "VNWHF-FKFBW-Q2RGD-HYHWF-R3HH2"
        "MondoVolume"                   = "FMTQQ-84NR8-2744R-MXF4P-PGYR3"
        "OneNoteFreeRetail"             = "XYNTG-R96FY-369HX-YFPHY-F9CPM"
        "OneNoteRetail"                 = "FXF6F-CNC26-W643C-K6KB7-6XXW3"
        "OneNoteVolume"                 = "9TYVN-D76HK-BVMWT-Y7G88-9TPPV"
        "OutlookRetail"                 = "7N4KG-P2QDH-86V9C-DJFVF-369W9"
        "OutlookVolume"                 = "7QPNR-3HFDG-YP6T9-JQCKQ-KKXXC"
        "PersonalPipcRetail"            = "9CYB3-NFMRW-YFDG6-XC7TF-BY36J"
        "PersonalRetail"                = "FT7VF-XBN92-HPDJV-RHMBY-6VKBF"
        "PowerPointRetail"              = "N7GCB-WQT7K-QRHWG-TTPYD-7T9XF"
        "PowerPointVolume"              = "X3RT9-NDG64-VMK2M-KQ6XY-DPFGV"
        "ProPlusRetail"                 = "GM43N-F742Q-6JDDK-M622J-J8GDV"
        "ProPlusVolume"                 = "FNVK8-8DVCJ-F7X3J-KGVQB-RC2QY"
        "ProfessionalPipcRetail"        = "CF9DD-6CNW2-BJWJQ-CVCFX-Y7TXD"
        "ProfessionalRetail"            = "NXFTK-YD9Y7-X9MMJ-9BWM6-J2QVH"
        "ProjectProRetail"              = "WPY8N-PDPY4-FC7TF-KMP7P-KWYFY"
        "ProjectProVolume"              = "PKC3N-8F99H-28MVY-J4RYY-CWGDH"
        "ProjectProXVolume"             = "JBNPH-YF2F7-Q9Y29-86CTG-C9YGV"
        "ProjectStdRetail"              = "NTHQT-VKK6W-BRB87-HV346-Y96W8"
        "ProjectStdVolume"              = "4TGWV-6N9P6-G2H8Y-2HWKB-B4G93"
        "ProjectStdXVolume"             = "N3W2Q-69MBT-27RD9-BH8V3-JT2C8"
        "PublisherRetail"               = "WKWND-X6G9G-CDMTV-CPGYJ-6MVBF"
        "PublisherVolume"               = "9QVN2-PXXRX-8V4W8-Q7926-TJGD8"
        "SkypeServiceBypassRetail"      = "6MDN4-WF3FV-4WH3Q-W699V-RGCMY"
        "SkypeforBusinessEntryRetail"   = "4N4D8-3J7Y3-YYW7C-73HD2-V8RHY"
        "SkypeforBusinessRetail"        = "PBJ79-77NY4-VRGFG-Y8WYC-CKCRC"
        "SkypeforBusinessVolume"        = "DMTCJ-KNRKR-JV8TQ-V2CR2-VFTFH"
        "StandardRetail"                = "2FPWN-4H6CM-KD8QQ-8HCHC-P9XYW"
        "StandardVolume"                = "WHGMQ-JNMGT-MDQVF-WDR69-KQBWC"
        "VisioProRetail"                = "NVK2G-2MY4G-7JX2P-7D6F2-VFQBR"
        "VisioProVolume"                = "NRKT9-C8GP2-XDYXQ-YW72K-MG92B"
        "VisioProXVolume"               = "G98Q2-B6N77-CFH9J-K824G-XQCC4"
        "VisioStdRetail"                = "NCRB7-VP48F-43FYY-62P3R-367WK"
        "VisioStdVolume"                = "XNCJB-YY883-JRW64-DPXMX-JXCR6"
        "VisioStdXVolume"               = "B2HTN-JPH8C-J6Y6V-HCHKB-43MGT"
        "WordRetail"                    = "P8K82-NQ7GG-JKY8T-6VHVY-88GGD"
        "WordVolume"                    = "YHMWC-YN6V9-WJPXD-3WQKP-TMVCV"
    }
    "Office2019" = @{
        "Access2019Retail"              = "WRYJ6-G3NP7-7VH94-8X7KP-JB7HC"
        "Access2019Volume"              = "6FWHX-NKYXK-BW34Q-7XC9F-Q9PX7"
        "AccessRuntime2019Retail"       = "FGQNJ-JWJCG-7Q8MG-RMRGJ-9TQVF"
        "Excel2019Retail"               = "KBPNW-64CMM-8KWCB-23F44-8B7HM"
        "Excel2019Volume"               = "8NT4X-GQMCK-62X4P-TW6QP-YKPYF"
        "HomeBusiness2019Retail"        = "QBN2Y-9B284-9KW78-K48PB-R62YT"
        "HomeStudentARM2019Retail"      = "DJTNY-4HDWM-TDWB2-8PWC2-W2RRT"
        "HomeStudentPlusARM2019Retail"  = "NM8WT-CFHB2-QBGXK-J8W6J-GVK8F"
        "HomeStudent2019Retail"         = "XNWPM-32XQC-Y7QJC-QGGBV-YY7JK"
        "Outlook2019Retail"             = "WR43D-NMWQQ-HCQR2-VKXDR-37B7H"
        "Outlook2019Volume"             = "RN3QB-GT6D7-YB3VH-F3RPB-3GQYB"
        "Personal2019Retail"            = "NMBY8-V3CV7-BX6K6-2922Y-43M7T"
        "PowerPoint2019Retail"          = "HN27K-JHJ8R-7T7KK-WJYC3-FM7MM"
        "PowerPoint2019Volume"          = "29GNM-VM33V-WR23K-HG2DT-KTQYR"
        "ProPlus2019Retail"             = "BN4XJ-R9DYY-96W48-YK8DM-MY7PY"
        "ProPlus2019Volume"             = "T8YBN-4YV3X-KK24Q-QXBD7-T3C63"
        "Professional2019Retail"        = "9NXDK-MRY98-2VJV8-GF73J-TQ9FK"
        "ProjectPro2019Retail"          = "JDTNC-PP77T-T9H2W-G4J2J-VH8JK"
        "ProjectPro2019Volume"          = "TBXBD-FNWKJ-WRHBD-KBPHH-XD9F2"
        "ProjectStd2019Retail"          = "R3JNT-8PBDP-MTWCK-VD2V8-HMKF9"
        "ProjectStd2019Volume"          = "RBRFX-MQNDJ-4XFHF-7QVDR-JHXGC"
        "Publisher2019Retail"           = "4QC36-NW3YH-D2Y9D-RJPC7-VVB9D"
        "Publisher2019Volume"           = "K8F2D-NBM32-BF26V-YCKFJ-29Y9W"
        "SkypeforBusiness2019Retail"    = "JBDKF-6NCD6-49K3G-2TV79-BKP73"
        "SkypeforBusiness2019Volume"    = "9MNQ7-YPQ3B-6WJXM-G83T3-CBBDK"
        "SkypeforBusinessEntry2019Retail" = "N9722-BV9H6-WTJTT-FPB93-978MK"
        "Standard2019Retail"            = "NDGVM-MD27H-2XHVC-KDDX2-YKP74"
        "Standard2019Volume"            = "NT3V6-XMBK7-Q66MF-VMKR4-FC33M"
        "VisioPro2019Retail"            = "2NWVW-QGF4T-9CPMB-WYDQ9-7XP79"
        "VisioPro2019Volume"            = "33YF4-GNCQ3-J6GDM-J67P3-FM7QP"
        "VisioStd2019Retail"            = "263WK-3N797-7R437-28BKG-3V8M8"
        "VisioStd2019Volume"            = "BGNHX-QTPRJ-F9C9G-R8QQG-8T27F"
        "Word2019Retail"                = "JXR8H-NJ3MK-X66W8-78CWD-QRVR2"
        "Word2019Volume"                = "9F36R-PNVHH-3DXGQ-7CD2H-R9D3V"
    }
    "Office2021" = @{
        "Access2021Retail"              = "P286B-N3XYP-36QRQ-29CMP-RVX9M"
        "AccessRuntime2021Retail"       = "MNX9D-PB834-VCGY2-K2RW2-2DP3D"
        "Access2021Volume"              = "JBH3N-P97FP-FRTJD-MGK2C-VFWG6"
        "Excel2021Retail"               = "V6QFB-7N7G9-PF7W9-M8FQM-MY8G9"
        "Excel2021Volume"               = "WNYR4-KMR9H-KVC8W-7HJ8B-K79DQ"
        "HomeBusiness2021Retail"        = "JM99N-4MMD8-DQCGJ-VMYFY-R63YK"
        "HomeStudent2021Retail"         = "N3CWD-38XVH-KRX2Y-YRP74-6RBB2"
        "OneNoteFree2021Retail"         = "CNM3W-V94GB-QJQHH-BDQ3J-33Y8H"
        "OneNote2021Retail"             = "NB2TQ-3Y79C-77C6M-QMY7H-7QY8P"
        "OneNote2021Volume"             = "THNKC-KFR6C-Y86Q9-W8CB3-GF7PD"
        "Outlook2021Retail"             = "4NCWR-9V92Y-34VB2-RPTHR-YTGR7"
        "Outlook2021Volume"             = "JQ9MJ-QYN6B-67PX9-GYFVY-QJ6TB"
        "Personal2021Retail"            = "RRRYB-DN749-GCPW4-9H6VK-HCHPT"
        "PowerPoint2021Retail"          = "3KXXQ-PVN2C-8P7YY-HCV88-GVM96"
        "PowerPoint2021Volume"          = "39G2N-3BD9C-C4XCM-BD4QG-FVYDY"
        "ProPlus2021Retail"             = "8WXTP-MN628-KY44G-VJWCK-C7PCF"
        "ProPlus2021Volume"             = "RNHJY-DTFXW-HW9F8-4982D-MD2CW"
        "ProPlusSPLA2021Volume"         = "JRJNJ-33M7C-R73X3-P9XF7-R9F6M"
        "Professional2021Retail"        = "DJPHV-NCJV6-GWPT6-K26JX-C7PBG"
        "ProjectPro2021Retail"          = "QKHNX-M9GGH-T3QMW-YPK4Q-QRWMV"
        "ProjectPro2021Volume"          = "HVC34-CVNPG-RVCMT-X2JRF-CR7RK"
        "ProjectStd2021Retail"          = "2B96V-X9NJY-WFBRC-Q8MP2-7CHRR"
        "ProjectStd2021Volume"          = "3CNQX-T34TY-99RH4-C4YD2-KW6WH"
        "Publisher2021Retail"           = "CDNFG-77T8D-VKQJX-B7KT3-KK28V"
        "Publisher2021Volume"           = "2KXJH-3NHTW-RDBPX-QFRXJ-MTGXF"
        "SkypeforBusiness2021Retail"    = "DVBXN-HFT43-CVPRQ-J89TF-VMMHG"
        "SkypeforBusiness2021Volume"    = "R3FCY-NHGC7-CBPVP-8Q934-YTGXG"
        "Standard2021Retail"            = "HXNXB-J4JGM-TCF44-2X2CV-FJVVH"
        "Standard2021Volume"            = "2CJN4-C9XK2-HFPQ6-YH498-82TXH"
        "StandardSPLA2021Volume"        = "BQWDW-NJ9YF-P7Y79-H6DCT-MKQ9C"
        "VisioPro2021Retail"            = "T6P26-NJVBR-76BK8-WBCDY-TX3BC"
        "VisioPro2021Volume"            = "JNKBX-MH9P4-K8YYV-8CG2Y-VQ2C8"
        "VisioStd2021Retail"            = "89NYY-KB93R-7X22F-93QDF-DJ6YM"
        "VisioStd2021Volume"            = "BW43B-4PNFP-V637F-23TR2-J47TX"
        "Word2021Retail"                = "VNCC4-CJQVK-BKX34-77Y8H-CYXMR"
        "Word2021Volume"                = "BJG97-NW3GM-8QQQ7-FH76G-686XM"
    }
    "Office2024" = @{
        "Access2024Retail"              = "P6NMW-JMTRC-R6MQ6-HH3F2-BTHKB"
        "Access2024Volume"              = "CXNJT-98HPP-92HX7-MX6GY-2PVFR"
        "Excel2024Retail"               = "82CNJ-W82TW-BY23W-BVJ6W-W48GP"
        "Excel2024Volume"               = "7Y287-9N2KC-8MRR3-BKY82-2DQRV"
        "Home2024Retail"                = "N69X7-73KPT-899FD-P8HQ4-QGTP4"
        "HomeBusiness2024Retail"        = "PRKQM-YNPQR-77QT6-328D7-BD223"
        "Outlook2024Retail"             = "2CFK4-N44KG-7XG89-CWDG6-P7P27"
        "Outlook2024Volume"             = "NQPXP-WVB87-H3MMB-FYBW2-9QFPB"
        "PowerPoint2024Retail"          = "CT2KT-GTNWH-9HFGW-J2PWJ-XW7KJ"
        "PowerPoint2024Volume"          = "RRXFN-JJ26R-RVWD2-V7WMP-27PWQ"
        "ProjectPro2024Retail"          = "GNJ6P-Y4RBM-C32WW-2VJKJ-MTHKK"
        "ProjectPro2024Volume"          = "WNFMR-HK4R7-7FJVM-VQ3JC-76HF6"
        "ProjectStd2024Retail"          = "C2PNM-2GQFC-CY3XR-WXCP4-GX3XM"
        "ProjectStd2024Volume"          = "F2VNW-MW8TT-K622Q-4D96H-PWJ8X"
        "ProPlus2024Retail"             = "VWCNX-7FKBD-FHJYG-XBR4B-88KC6"
        "ProPlus2024Volume"             = "4YV2J-VNG7W-YGTP3-443TK-TF8CP"
        "SkypeforBusiness2024Volume"    = "XKRBW-KN2FF-G8CKY-HXVG6-FVY2V"
        "Standard2024Volume"            = "GVG6N-6WCHH-K2MVP-RQ78V-3J7GJ"
        "VisioPro2024Retail"            = "HGRBX-N68QF-6DY8J-CGX4W-XW7KP"
        "VisioPro2024Volume"            = "GBNHB-B2G3Q-G42YB-3MFC2-7CJCX"
        "VisioStd2024Retail"            = "VBXPJ-38NR3-C4DKF-C8RT7-RGHKQ"
        "VisioStd2024Volume"            = "YNFTY-63K7P-FKHXK-28YYT-D32XB"
        "Word2024Retail"                = "XN33R-RP676-GMY2F-T3MH7-GCVKR"
        "Word2024Volume"                = "WD8CQ-6KNQM-8W2CX-2RT63-KK3TP"
    }
}


# ══════════════════════════════════════════════════════════════════════════════
# ▓▓ Funktionen  ▓
# ══════════════════════════════════════════════════════════════════════════════
# Shell Funktionen
function Show-Shell {
    param(
        [switch]$Hide,
        [switch]$Show
    )
    if ($Hide) {
    Add-Type -Name Win -Namespace Console -MemberDefinition '
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'
    $consolePtr = [Console.Win]::GetConsoleWindow()
    [Console.Win]::ShowWindow($consolePtr, 0)  # 0 = SW_HIDE
    } elseif ($Show) {
    Add-Type -Name Win -Namespace Console -MemberDefinition '
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'
    $consolePtr = [Console.Win]::GetConsoleWindow()
    [Console.Win]::ShowWindow($consolePtr, 5)  # 5 = SW_SHOW
    }
}
function restart-Script($Admin) {
    if ($Admin) {
        Write-Host "Starte Skript als Administrator neu ..."
        Start-Process -FilePath powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    } else {
        Write-Host "Starte Skript neu ..."
        Start-Process -FilePath powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    }
    Exit 0
}

# Download Funktionen
function Get-Downloader {
    <#
    .SYNOPSIS
        Erstellt ein WebClient-Objekt mit Proxy-Unterstützung.

    .DESCRIPTION
        Diese Funktion erzeugt ein vorkonfiguriertes .NET-WebClient-Objekt, das
        Proxy-Einstellungen und Anmeldeinformationen aus der Systemumgebung übernimmt.
        Sie dient als Grundlage für Datei-Downloads über HTTP/HTTPS.

    .PARAMETER Url
        Optionale Ziel-URL, um zu prüfen, ob der Proxy den Zugriff umgehen sollte.

    .PARAMETER ProxyUrl
        Optionaler Proxyserver (z.B. "http://proxy.firma.local:8080").

    .PARAMETER ProxyCredential
        Anmeldeinformationen für den Proxy (vom Typ [PSCredential]).

    .OUTPUTS
        Gibt ein konfiguriertes [System.Net.WebClient]-Objekt zurück.

    .EXAMPLE
        $downloader = Get-Downloader -Url "https://example.com/datei.zip"
        $download.DownloadFile("https://example.com/datei.zip", "C:\temp\datei.zip")
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string] $Url,
        [Parameter(Mandatory = $false)][string] $ProxyUrl,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential] $ProxyCredential
    )

    $downloader = New-Object System.Net.WebClient

    $defaultCreds = [System.Net.CredentialCache]::DefaultCredentials
    if ($defaultCreds) { $downloader.Credentials = $defaultCreds }

    if ($ProxyUrl) {
        Write-Host "Verwendung des übergebenden Proxy-Servers '$ProxyUrl'."
        $proxy = New-Object System.Net.WebProxy -ArgumentList $ProxyUrl, $true

        $proxy.Credentials = if ($ProxyCredential) {
            $ProxyCredential.GetNetworkCredential()
        } elseif ($defaultCreds) {
            $defaultCreds
        } else {
            Write-Warning "Keine Proxy-Anmeldedaten gefunden - manuelle Eingabe erforderlich."
            (Get-Credential).GetNetworkCredential()
        }

        if (-not $proxy.IsBypassed($Url)) {
            $downloader.Proxy = $proxy
        }
    } 

    return $downloader
}
function Request-File {
    <#
    .SYNOPSIS
        Lädt eine Datei von einer URL herunter.

    .DESCRIPTION
        Lädt eine Datei über HTTP oder HTTPS von der angegebenen Quelle herunter.
        Unterstützt optionale Proxy-Konfigurationen und Fehlerbehandlung.

    .PARAMETER Url
        Die vollständige Download-URL der Datei.

    .PARAMETER File
        Der lokale Speicherpfad, unter dem die Datei gespeichert werden soll.

    .PARAMETER ProxyConfiguration
        Optionales Hashtable mit Proxy-Parametern (ProxyUrl, ProxyCredential).

    .EXAMPLE
        Request-File -Url "https://example.com/file.zip" -File "C:\Temp\file.zip"

    .NOTES
        Diese Funktion nutzt Get-Downloader zur automatischen Proxy-Erkennung.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string] $Url,
        [Parameter(Mandatory = $false)][string] $File,
        [Parameter(Mandatory = $false)][hashtable] $ProxyConfiguration
    )

    Write-Host "Herunterladen: $Url in $File"
    $dl = Get-Downloader -Url $Url @ProxyConfiguration
    try {
        $dl.DownloadFile($Url, $File)
    }
    catch {
        throw "Download fehlgeschlagen: $_"
    }
}

# Office Lizenz Funktionen
function activateOfficeProduct {
    param(
        [string]$ProductID,
        [string]$ProductKey
    )
    # Ohook herunterladen
    $ohookDir = Join-Path $env:SystemDrive "ohook"
    $ohookLink = "https://github.com/asdcorp/ohook/releases/download/0.5/ohook_0.5.zip"
    $tempDir = [System.IO.Path]::GetTempPath()
    $tempZip = [System.IO.Path]::Combine($tempDir, "ohook.zip")
    
    # Alte Dateien aufräumen
    if (Test-Path $tempZip) { 
        Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $ohookDir) { 
        Remove-Item $ohookDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    # Verzeichnis erstellen
    New-Item -Path $ohookDir -ItemType Directory -Force | Out-Null
    Write-Host "Erstelle Verzeichnis: $ohookDir" -ForegroundColor Cyan
    
    Write-Host "Lade Ohook herunter ..." -ForegroundColor Cyan
    Request-File -Url $ohookLink -File $tempZip
    Write-Host "Entpacke Ohook ..." -ForegroundColor Cyan
    
    # Zielordner komplett löschen falls vorhanden
    if (Test-Path $ohookDir) {
        Remove-Item $ohookDir -Recurse -Force
    }
    
    Expand-Archive -Path $tempZip -DestinationPath $ohookDir
    Remove-Item -Path $tempZip -Force

    # Ohook verlinken
    $targetDir      = Join-Path $env:ProgramFiles "Microsoft Office\root\vfs\System"
    $symlinkTarget  = Join-Path $env:windir "System32\sppc.dll"
    $symlinkLink    = Join-Path $targetDir "sppc.dll"
    $symlinkDir     = Split-Path -Path $symlinkLink -Parent

    # Symbolischen Link erstellen
    if (-not (Test-Path -Path $symlinkDir)) { 
        New-Item -Path $symlinkDir -ItemType Directory -Force | Out-Null 
    }

    # Symbolischen Link erstellen
    try { 
        cmd /c mklink "$symlinkLink" "$symlinkTarget"
    } catch {
        Write-Error "` Fehler beim Erstellen des symbolischen Links: $_"
        exit 1
    }

    # Quelle und Ziel für OHook definieren
    $sourceDir = "C:\ohook"
    $sourceFile = Join-Path $sourceDir "sppc64.dll"
    $targetFile = Join-Path $targetDir "sppc.dll"

    # Prüfen ob die Quelldatei existiert
    if (-not (Test-Path -Path $sourceFile)) {
        Write-Error "Fehler: Quelldatei für OHook nicht gefunden: $sourceFile"
        exit 1
    }

    # Zielordner bei Bedarf anlegen
    $targetDir   = Join-Path $env:ProgramFiles "Microsoft Office\root\vfs\System"
    $sourceDir   = "C:\ohook"
    $sourceFile  = Join-Path $sourceDir "sppc64.dll"
    $targetFile  = Join-Path $targetDir "sppc.dll"
    
    if (-not (Test-Path -Path $targetDir)) {
        try {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Host "Zielordner erstellt: $targetDir" -ForegroundColor Cyan
        }
        catch {
            Write-Error "Konnte Zielordner nicht erstellen: $_"
            return
        }
    }

    # Office-Prozesse beenden
    $officeProcesses = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSACCESS", "MSPUB", "lync", "TEAMS")
    Write-Host "Beende Office-Prozesse ..." -ForegroundColor Yellow
    foreach ($proc in $officeProcesses) {
        Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2

    # Datei kopieren (überschreiben falls vorhanden)
    try {
        Write-Host "Kopiere OHook Datei: $sourceFile nach $targetFile" -ForegroundColor Green
        
        # Alte Datei entfernen mit attrib
        if (Test-Path $targetFile) {
            Write-Host "Entferne alte sppc.dll ..." -ForegroundColor Yellow
            & attrib -r -s -h "$targetFile"
            Remove-Item $targetFile -Force
        }
        
        Copy-Item -Path $sourceFile -Destination $targetFile -Force
        Write-Host "Datei erfolgreich kopiert." -ForegroundColor Green
    }
    catch {
        Write-Error "Fehler beim Kopieren der Datei: $_"
        return
    }

    # Registrierungswert setzen
    $regPath    = "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency"
    $valueName  = "TimeOfLastHeartbeatFailure"
    $valueDate  = "2040-01-01T00:00:00Z"

    # Backup des existierenden Werts erstellen
    try {
        if (-not (Test-Path -Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
            Write-Host "Erstelle Registrierungspfad: $regPath" -ForegroundColor Cyan
        }
        $existing = Get-ItemProperty -Path $regPath -Name $valueName -ErrorAction SilentlyContinue
        if ($null -ne $existing) {
            $backupname = "$valueName-Backup-" + (Get-Date -Format "yyyyMMddHHmmss")
            New-ItemProperty -Path $regPath -Name $backupname -Value $existing.$valueName -PropertyType String -Force | Out-Null
            Write-Host "Existierender Wert gesichert als: $backupname" -ForegroundColor Cyan
        }
    } catch {
        Write-Error "Fehler beim Zugriff auf die Registrierung: $_"
        exit 1
    }

    try {
        Write-Host "`t  Installiere Produktschlüssel für $ProductID" -ForegroundColor Green
        $ipkResult = & cscript.exe //nologo "$env:SystemRoot\System32\slmgr.vbs" /ipk $ProductKey 2>&1
        Write-Host $ipkResult -ForegroundColor Gray
        
        Write-Host "`t  Starte Aktivierung (kann 30-60 Sekunden dauern) ..." -ForegroundColor Green
        $atoResult = & cscript.exe //nologo "$env:SystemRoot\System32\slmgr.vbs" /ato 2>&1
        Write-Host $atoResult -ForegroundColor Gray
        
        Write-Host "`t  Aktivierung abgeschlossen!" -ForegroundColor Green
    } catch {
        Write-Error "Fehler beim Installieren/Aktivieren: $_"
        return
    }

}
function Get-C2RProducts {
    $reg = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    $products = @()
    
    if (Test-Path $reg) {
        $props = Get-ItemProperty $reg
        $ids   = $props.ProductReleaseIds -split ","

        foreach ($id in $ids) {
            # Lizenzstatus per WMI abrufen
            $licenseStatus = Get-OfficeLicenseStatusWMI
            
            # ProductKey ermitteln
            $productKey = $null
            $productBundle = $null
            
            # Durchsuche alle Office-Pakete nach dem ProductID
            foreach ($bundle in $Keys.Keys) {
                if ($Keys[$bundle].ContainsKey($id)) {
                    $productKey = $Keys[$bundle][$id]
                    $productBundle = $bundle
                    break
                }
            }
                       
            $products += [PSCustomObject]@{
                ProductID       = $id
                ProductKey      = $productKey
                ProductBundle   = $productBundle
                Version         = $props.ClientVersionToReport
                Edition         = $id -replace '(Volume|Retail)$', ''
                LicenseStatus   = $licenseStatus
                IsActivated     = ($licenseStatus -eq "Licensed")
            }
        }
        
        return $products
    }
}
function Get-OfficeLicenseStatus {
    $osppPath = Join-Path $env:ProgramFiles "Microsoft Office\Office16\OSPP.VBS"

    if (Test-Path $osppPath) {
        # Status aller Lizenzen abfragen
        $result = cscript //Nologo $osppPath /dstatus

        # Nach Lizenzstatus suchen
        if ($result -match "LICENSE STATUS:\s*---LICENSED---") {
            return "Licensed"
        } elseif ($result -match "LICENSE STATUS:\s*---NOTIFICATIONS---") {
            return "Grace Period"
        } elseif ($result -match "LICENSE STATUS:\s*---NOTIFIED---") {
            return "Notified"
        } else {
            return "Unlicensed"
        }
    }
    return "Unknown"
}
function Get-OfficeLicenseStatusWMI {
    try {
        # Versuche WMI-Abfrage mit SoftwareLicensingProduct (Windows 10+)
        $licenses = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663' AND PartialProductKey IS NOT NULL" -ErrorAction SilentlyContinue
        
        if (-not $licenses) {
            # Fallback: Versuche alte WMI-Klasse
            $licenses = Get-WmiObject -Query "SELECT LicenseStatus FROM SoftwareLicensingProduct WHERE ApplicationId='0ff1ce15-a989-479d-af46-f275c6370663' AND PartialProductKey IS NOT NULL" -ErrorAction SilentlyContinue
        }
        
        if (-not $licenses) {
            return "No License Found"
        }
        
        # Priorisiere lizenzierte Einträge
        foreach ($license in $licenses) {
            if ($license.LicenseStatus -eq 1) {
                return "Licensed"
            }
        }
        
        # Falls keine lizenzierte gefunden, gib den Status der ersten zurück
        $firstLicense = $licenses | Select-Object -First 1
        switch ($firstLicense.LicenseStatus) {
            0 { return "Unlicensed" }
            1 { return "Licensed" }
            2 { return "OOB Grace" }
            3 { return "OOT Grace" }
            4 { return "Non-Genuine Grace" }
            5 { return "Notification" }
            6 { return "Extended Grace" }
            default { return "Unknown" }
        }
    }
    catch {
        Write-Warning "WMI-Abfrage fehlgeschlagen: $_"
        return "Unknown"
    }
}
function Get-OfficeLicenseStatusRegistry {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot"
    )
    
    foreach ($path in $regPaths) {
        if (Test-Path $path) {
            $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            
            # ProductReleaseIds prüfen
            if ($props.ProductReleaseIds) {
                # Lizenzstatus in anderem Schlüssel
                $licensePath = "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext"
                if (Test-Path $licensePath) {
                    $licenseProps = Get-ItemProperty -Path $licensePath -ErrorAction SilentlyContinue
                    
                    # Verschiedene Lizenzindikatoren
                    if ($licenseProps) {
                        return "Licensed (Registry)"
                    }
                }
            }
        }
    }
    return "Unknown"
}

# Office Ohook
function linkOhook {

}
function Build-Ohook {
    param (
        [string]$productID,
        [string]$licenseKey
    )
    $targetDir   = Join-Path $env:ProgramFiles "Microsoft Office\root\vfs\System"
    $sourceDir   = "C:\ohook"
    $sourceFile  = Join-Path $sourceDir "sppc64.dll"
    $targetFile  = Join-Path $targetDir "sppc.dll"

    # Datei kopieren (überschreiben falls vorhanden)
    try {
        Write-Host "Kopiere OHook Datei: $sourceFile nach $targetFile" -ForegroundColor Green
        Copy-Item -Path $sourceFile -Destination $targetFile -Force
        Write-Host "Datei erfolgreich kopiert." -ForegroundColor Green
    }
    catch {
        Write-Error "Fehler beim Kopieren der Datei: $_"
        exit 1
    }

    # Registrierungswert setzen
    $regPath    = "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency"
    $valueName  = "TimeOfLastHeartbeatFailure"
    $valueDate  = "2040-01-01T00:00:00Z"

    # Backup des existierenden Werts erstellen
    try {
        if (-not (Test-Path -Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
            Write-Host "Erstelle Registrierungspfad: $regPath" -ForegroundColor Cyan
        }
        $existing = Get-ItemProperty -Path $regPath -Name $valueName -ErrorAction SilentlyContinue
        if ($existing -ne $null) {
            $backupname = "$valueName-Backup-" + (Get-Date -Format "yyyyMMddHHmmss")
            New-ItemProperty -Path $regPath -Name $backupname -Value $existing.$valueName -PropertyType String -Force | Out-Null
            Write-Host "Existierender Wert gesichert als: $backupname" -ForegroundColor Cyan
        }
    } catch {
        Write-Error "Fehler beim Zugriff auf die Registrierung: $_"
        exit 1
    }

    try {
        Write-Host "`t  Installiere Produktschlüssel für $productID" -ForegroundColor Green
        slmgr /ipk $licenseKey
        Write-Host "`t  Starte Aktivierung ..." -ForegroundColor Green
    } catch {
        Write-Error "Fehler beim Installieren des Produktschlüssels: $_"
        exit 1
    } except {
        Write-Error "Fehler bei der Aktivierung: $_"
        exit 1
    }
}

# Office Installation
function Get-OfficeDownloadURL {
    param (
        [string]$productID,
        [string]$language   = "de-de",
        [string]$version    = "O16GA",
        [string]$platform   = "x64",
        [switch]$Offline
    )
    if ($Offline) {
        return "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/" + $language + "/" + $productID + ".img"
    }
    # Anpassung der Plattformbezeichnung
    if ($platform -eq "x86") { $platform = "x32" }
    return "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=" + $productID + "&platform=" + $platform + "&language=" + $language + "&version=" + $version
}
function installOfficeOnline {
    param (
        [string]$License,
        [string]$Version,
        [string]$Edition,
        [string]$Platform = "x64",
        [string]$Language = "de-de"
    )

    # ProductID erstellen
    if ($Offices.Contains($License)) { $productID = $Edition + $License } else { $productID = $Edition }
    
    
    # Download vorbereiten
    Write-Host "Bereite Download für: $License $Version $Edition ($Platform)" -ForegroundColor Green
    $downloadURL = Get-OfficeDownloadURL -productID $productID -language $Language -platform $Platform
    Write-Host "Download-URL: $downloadURL" -ForegroundColor Cyan
    $tempFile    = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$($productID)_installer.exe")
    Write-Host "Temp File: $tempFile" -ForegroundColor Cyan

    # Download starten
    Write-Host "Starte Download ..." -ForegroundColor Green
    Request-File -Url $downloadURL -File $tempFile
    Write-Host "Download abgeschlossen." -ForegroundColor Green

    # Installation starten
    Write-Host "Starte Office Installation: $License $Version $Edition ($Platform)" -ForegroundColor Green

    # Installationsprozess
    if ($script:quietInstall) {
        Start-Process -FilePath $tempFile -ArgumentList "/quiet" -Wait
    } else {
        Start-Process -FilePath $tempFile -Wait
    }
    Write-Host "Office Installation abgeschlossen." -ForegroundColor Green
    Start-Sleep -Seconds 5

    # Temporäre Datei löschen
    Write-Host "Lösche temporäre Datei: $tempFile" -ForegroundColor Cyan
    Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
}
function installOfficeOffline {
    param (
        [string]$License,
        [string]$Version,
        [string]$Edition,
        [string]$Platform = "x64",
        [string]$Language = "de-de"
    )
    if ($Offices.Contains($License)) { $productID = $Edition + $License } else { $productID = $Edition }

    # Download vorbereiten
    Write-Host "Erstelle Offline-Installer für: $License $Version $Edition ($Platform)" -ForegroundColor Green
    $downloadURL = Get-OfficeDownloadURL -productID $productID -language $Language -platform $Platform -Offline
    Write-Host "Download-URL: $downloadURL" -ForegroundColor Cyan
    $tempFile    = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$($productID)_installer.img")
    Write-Host "Temp File: $tempFile" -ForegroundColor Cyan

    # Download starten
    Write-Host "Starte Download des Offline-Installers ..." -ForegroundColor Green
    Request-File -Url $downloadURL -File $tempFile
    Write-Host "Download des Offline-Installers abgeschlossen." -ForegroundColor Green

    # Installation starten
    Write-Host "Starte Office Installation: $License $Version $Edition ($Platform)" -ForegroundColor Green
    $mount = Mount-DiskImage -ImagePath $tempFile -PassThru
    $driveLetter = ($mount | Get-Volume).DriveLetter
    $setupFile = "$($driveLetter):\setup.exe"

    # Installationsprozess
    if ($script:quietInstall) {
        Start-Process -FilePath $setupFile -ArgumentList "/quiet" -Wait
    } else {
        Start-Process -FilePath $setupFile -Wait
    }
    Write-Host "Office Installation abgeschlossen." -ForegroundColor Green
    Dismount-DiskImage -ImagePath $tempFile
    Start-Sleep -Seconds 5

    # Temporäre Datei löschen
    Write-Host "Lösche temporäre Datei: $tempFile" -ForegroundColor Cyan
    Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
}

# ══════════════════════════════════════════════════════════════════════════════
# ▓▓ Initialisierung ▓
# ══════════════════════════════════════════════════════════════════════════════
if ($RunAsAdmin -and -not $isAdmin) { restart-Script -Admin $true }
if ($hideShell) { Show-Shell -Hide }
# Invoke-RestMethod "https://aporie.me/scripts/windows.form.ps1" | Invoke-Expression
. .\windows.form.ps1

# ══════════════════════════════════════════════════════════════════════════════
# ▓▓ GUI  ▓
# ══════════════════════════════════════════════════════════════════════════════

# Heights



# ════ Header ═════════════════════════════════════════════════════════════════
$headerPanel = createPanel -Width 400 -Height 100 -Left 5 -Top 5
$headerLabel  = createLabel -Width 400 -Height 50 -FontSize 42 -ForeColor $script:AccentColor -Top 20 -Text "$name" -FontStyle "Bold"
$versionLabel = createLabel -Width 400 -Height 20 -FontSize 10 -ForeColor $script:AccentColor -Top 70 -Text "Version $version"
$headerPanel.Controls.AddRange(@($headerLabel, $versionLabel))
# Event-Listener
$headerLabel.Add_DoubleClick({
    $form.Close()
    $global:restartScript = $true
})


# ════ Install-Office ══════════════════════════════════════════════════════════
# Dropdown-Listen
$installPanel = createPanel        -Width 400 -Height 65 -Top 110 -Left 5
$listLabel    = createLabel        -Width 400 -Height 20 -Top 10 -FontSize 12 -Text "Office Installieren" -FontStyle "Bold" -ForeColor $script:AccentColor
$listLicenses = createDropDownList -Width 80  -Height 30 -Top 35 -Left 15  -SelectedIndex 0 -Items @($Offices.Keys)
$listVersions = createDropDownList -Width 100 -Height 30 -Top 35 -Left 100 -SelectedIndex 0 -Items @($Offices[$listLicenses.SelectedItem].Keys)
$listEditions = createDropDownList -Width 180 -Height 30 -Top 35 -Left 205 -SelectedIndex 0 -Items @($Offices[$listLicenses.SelectedItem][$listVersions.SelectedItem])
$installPanel.Controls.AddRange(@($listLabel, $listLicenses, $listVersions, $listEditions ))
$listLicenses.Add_SelectedIndexChanged({
    $license    = $listLicenses.SelectedItem
    $versions   = $Offices[$license].Keys
    $listVersions.Items.Clear()
    $listVersions.Items.AddRange($versions)
    $listVersions.SelectedIndex = 0
}.GetNewClosure())
$listVersions.Add_SelectedIndexChanged({
    $version    = $listVersions.SelectedItem
    $editions   = $Offices[$listLicenses.SelectedItem][$version]
    $listEditions.Items.Clear()
    $listEditions.Items.AddRange($editions)
    $listEditions.SelectedIndex = 0
}.GetNewClosure())

# Install-Buttons
$buttonPanel             = createPanel  -Width 400 -Height 70 -Top 155 -Left 5
$onlineInstaller32Button = createButton -Width 110 -Height 30 -Top 27  -Left 20  -Text "Online (32 Bit)"
$onlineInstaller64Button = createButton -Width 110 -Height 30 -Top 27  -Left 145 -Text "Online (64 Bit)"
$offlineInstallerButton  = createButton -Width 110 -Height 30 -Top 27  -Left 270 -Text "Offline"
$buttonPanel.Controls.AddRange(@( $onlineInstaller32Button, $onlineInstaller64Button, $offlineInstallerButton ))
# Event-Listener
Add-Events -Control $onlineInstaller64Button -Events @{
    "Click"      = { installOfficeOnline -License $listLicenses.SelectedItem -Version $listVersions.SelectedItem -Edition $listEditions.SelectedItem -Platform "x64" }
}
Add-Events -Control $offlineInstallerButton  -Events @{
    "Click"      = { installOfficeOffline -License $listLicenses.SelectedItem -Version $listVersions.SelectedItem -Edition $listEditions.SelectedItem -Platform "x64" }
}


# ════ Installierte Produkte ═══════════════════════════════════════════════════════════
$installedPanel   = createPanel        -Width 400 -Height 80 -Top 230 -Left 5
$installedLabel   = createLabel        -Width 400 -Height 20 -Top 10 -FontSize 12 -Text "Installierte Produkte" -FontStyle "Bold" -ForeColor $script:AccentColor
$installedText    = createLabel        -Width 400 -Height 30 -Top 35 -FontSize 12 -TextAlign "TopCenter"
$installedPanel.Controls.AddRange(@( $installedLabel, $installedText ))
$installedOffices = Get-C2RProducts

if ($installedOffices.Count -eq 0) {
    # Keine installierten Produkte gefunden
    $installedText.Text = "Keine installierten Office-Produkte gefunden."
    $footerTop = 0
} else {
    # Produkte gefunden
    $installed = $installedOffices[0]
    $ProductID = $installed.ProductID
    $ProductKey = $installed.ProductKey

    $installedPanel.Height += 100
    
    $installedText.Top += 40
    $installedText.Height += 70
    $footerTop += 100
    $formHeight += 100

    $installedText.Text = switch ($installed.LicenseStatus) {
        "Licensed"          { "✓ Office ist aktiviert und lizenziert" }
        "Unlicensed"        { "✗ Office ist nicht lizenziert" }
        "OOB Grace"         { "⚠ Testphase aktiv (Out-of-Box Grace Period)" }
        "OOT Grace"         { "⚠ Toleranzphase nach Hardware-Änderung (Out-of-Tolerance Grace)" }
        "Non-Genuine Grace" { "⚠ Nicht-genuine Lizenz in Kulanzphase" }
        "Notification"      { "✗ Aktivierung erforderlich `nLizenz abgelaufen oder ungültig (nur Lesemodus)" }
        "Extended Grace"    { "⚠ Erweiterte Kulanzphase aktiv" }
        "Unknown"           { "? Lizenzstatus konnte nicht ermittelt werden" }
        default             { "Lizenzstatus: $($installed.LicenseStatus)" }
    }

    $installedVersion = createDropDownList -Width 200 -Height 30 -Top 35 -Left 15 -SelectedIndex 0 -Items @($installedOffices | ForEach-Object { $_.ProductID })
    $activateButton   = createButton       -Width 150 -Height 30 -Top 35 -Left 225 -Text "Aktivieren"
    Add-Events -Control $activateButton -Events @{
        "Click"      = { 
            $activateButton.Text = "Aktiviere ..."
            $activateButton.Enabled = $false
            $installedText.Text = "⏳ Aktivierung läuft (30-60 Sekunden) ..."
            
            activateOfficeProduct -ProductID $ProductID -ProductKey $ProductKey
            
            Write-Host "Warte 10 Sekunden auf Aktivierung ..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
            
            # Lizenzstatus aktualisieren
            $installed.LicenseStatus = Get-OfficeLicenseStatusWMI
            if ($installed.LicenseStatus -eq "Licensed") {
                $installed.IsActivated = $true
                $activateButton.Text    = "Bereits aktiviert"
                $activateButton.Enabled = $false
                $installedText.Text     = "✓ Office ist aktiviert und lizenziert"
            } else {
                $activateButton.Text    = "Erneut versuchen"
                $activateButton.Enabled = $true
                $installedText.Text     = "✗ Aktivierung fehlgeschlagen. Lizenzstatus: $($installed.LicenseStatus)"
            }
        }
    }

    if ($installed.IsActivated) {
        $activateButton.Text    = "Bereits aktiviert"
        $activateButton.Enabled = $false
    }
    $installedPanel.Controls.AddRange(@( $installedVersion, $activateButton, $installedText ))

    $installedVersion.Add_SelectedIndexChanged({
        $installedText.Text = switch ($installed.LicenseStatus) {
            "Licensed"          { "✓ Office ist aktiviert und lizenziert" }
            "Unlicensed"        { "✗ Office ist nicht lizenziert" }
            "OOB Grace"         { "⚠ Testphase aktiv (Out-of-Box Grace Period)" }
            "OOT Grace"         { "⚠ Toleranzphase nach Hardware-Änderung (Out-of-Tolerance Grace)" }
            "Non-Genuine Grace" { "⚠ Nicht-genuine Lizenz in Kulanzphase" }
            "Notification"      { "✗ Aktivierung erforderlich `nLizenz abgelaufen oder ungültig (nur Lesemodus)" }
            "Extended Grace"    { "⚠ Erweiterte Kulanzphase aktiv" }
            "Unknown"           { "? Lizenzstatus konnte nicht ermittelt werden" }
            default             { "Lizenzstatus: $($installed.LicenseStatus)" }
        }
        $productID = $installedVersion.SelectedItem
        $installed = $installedOffices | Where-Object { $_.ProductID -eq $productID }
        if ($installed.IsActivated) {
            $activateButton.Enabled = $false
            $activateButton.Text    = "Bereits aktiviert"
        } else {
            $activateButton.Enabled = $true
            $activateButton.Text    = "Aktivieren"
        }
    }.GetNewClosure())
}


# ═══ Footer ═════════════════════════════════════════════════════════════════════
$footerPanel    = createPanel -Width 400 -Height 20 -Left 5   -BackColor $script:AccentColor  -Top (310 + $footerTop)
$aporieLabel    = createLabel -Width 100 -Height 20 -Left 5 -Text "APORIE" -config @{
    "FontSize"  = 8
    "FontStyle" = "Bold"
    "TextAlign" = "MiddleLeft"
    "ForeColor" = $script:DarkColor
    "BackColor" = $script:AccentColor
}
$showShellLabel = createLabel -Width 200 -Height 20 -Left 195 -Hand -Visible $hideShell -config @{
    "Text"      = "Konsole anzeigen"
    "FontSize"  = 8
    "FontStyle" = "Bold"
    "TextAlign" = "MiddleRight"
    "ForeColor" = $script:DarkColor
    "BackColor" = $script:AccentColor
    "Visible"   = $hideShell
}
$HideShellLabel = createLabel -Width 200 -Height 20 -Left 195 -Hand -Visible (-not $hideShell) -config @{
    "Text"      = "Konsole verbergen"
    "FontSize"  = 8
    "FontStyle" = "Bold"
    "TextAlign" = "MiddleRight"
    "ForeColor" = $script:DarkColor
    "BackColor" = $script:AccentColor
}
$footerPanel.Controls.AddRange(@( $aporieLabel, $showShellLabel, $HideShellLabel ))

Add-Events -Control $showShellLabel -Events @{
    "Click"      = { Show-Shell -Show; $showShellLabel.Visible = $false; $HideShellLabel.Visible = $true }
}
Add-Events -Control $hideShellLabel -Events @{
    "Click"      = { Show-Shell -Hide; $showShellLabel.Visible = $true; $HideShellLabel.Visible = $false }
}
# ════ Form ════════════════════════════════════════════════════════════════════
$form   = createForm -Width 410 -Height (330 + $formHeight) -Text "$name - APORIE Skript" 
$form.Controls.AddRange(@( $headerPanel, $installPanel, $buttonPanel, $installedPanel, $footerPanel ))
$form.ShowDialog() | Out-Null

# ══════════════════════════════════════════════════════════════════════════════
## Skript-Neustart
if ($global:restartScript) {
    restart-Script -Admin $RunAsAdmin
}