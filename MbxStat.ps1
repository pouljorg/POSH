
<#
.Synopsis
   Scriptet henter mailbox statistik for specifik server eller database
.DESCRIPTION
   Scriptet henter mailbox statistik for specifik server eller database. Data gemmes i en ';' separeret csv fil, med oplysningerne:
   LastLogonTime, DisplayName, ItemCount, Størrelse i MB, Databasenavn og servernavn
   Mailboskstatistik er sorteret i aftagende størrelse efter 
.PARAMETER Database
    Database Parameteren benytte til at specificere hvilken database, der trækkes statistik ud af
.PARAMETER Server
    Server Parameteren benytte til at specificere hvilken servers databaser, der trækkes statistik ud af
.PARAMETER Path
    Path Benyttes til at angive sti og navn for ønsket fil der skal gemmes i.
.PARAMETER Unit
    Unit Benyttes til at angive enheden størrelser skal returneres i
.EXAMPLE
   .\Mbxstat.ps1 -Database mbx1 -Path c:\mbxstatDBMbx.csv
   Ovenstående trækker mailboksstatistik for databasen MBX1 og gemmer i file på stien: c:\mbxstatDBMbx.csv
.EXAMPLE
   .\Mbxstat.ps1 -Server Lon-mbx1 -Path c:\mbxstatSvrLon-Mbx1.csv
   Ovenstående trækker mailboksstatistik for alle databaser på servern Lon-MBX1 og gemmer i file på stien: c:\mbxstatSvrLon-Mbx1.csv
#>
[CmdLetBinding()]
Param
(
    [Parameter(Position = 0, Mandatory, 
               ParameterSetName = "DataBase")]
    [String]$Database,
    [Parameter(Position = 0, Mandatory, 
               ParameterSetName = "Server")]
    [String]$Server,
    [Parameter(Position = 1, Mandatory)] 
    [String]$Path,
    [Parameter(Position = 2)]
    [ValidateSet('KB','MB','GB')]
    [String]$Unit = 'MB'
)
switch($Unit) {
    'KB' { $Val = 1KB}
    'MB' { $val = [System.Math]::Pow(1KB,2) }
    'GB' { $val = [System.Math]::Pow(1KB,3) }
}

if ($Database){ 
    $mbx = get-mailbox -database $Database -ResultSize Unlimited 
} else {
    $mbx = get-mailbox -server $Server -ResultSize Unlimited 
}
$mbx| 
Get-MailboxStatistics |
Select-Object LastLogonTime,DisplayName,ItemCount,
    @{label="Total Item Size($Unit)";expression={([int](100*([double]($_.totalitemsize.tostring().replace('(',' ').split()[3] ) )/$Val))/100}},
    ServerName,DatabaseName | 
Sort-Object "Total Item Size(MB)" -Descending | 
Export-Csv -NoTypeInformation -Encoding Unicode -Path $Path -Delimiter ';'

