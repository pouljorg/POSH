
#create 64 bit odbc dsn
Add-OdbcDsn -Name "Kunde64" -DriverName 'ODBC Driver 13 for SQL Server' -DsnType "System" -Platform '64-bit' -SetPropertyValue @("Server=Basse", "Trusted_Connection=Yes", "Database=Kunde")
#Add-OdbcDsn -Name Kunde64  -Platform '64-bit' -DriverName 'ODBC Driver 13 for SQL Server' -DsnType System
#create 32 bit odbc dsn
Add-OdbcDsn -Name "Kunde32" -DriverName 'ODBC Driver 13 for SQL Server' -DsnType "System" -Platform '32-bit' -SetPropertyValue @("Server=Basse", "Trusted_Connection=Yes", "Database=Kunde")

# 32 bit drivers require a 32-bit edition of Powershell running!!!!!!!!!!!!!

$dsn = 'Kunde32'
$query = " SELECT TOP (10) * FROM dbo.person;"
   $conn = New-Object System.Data.Odbc.OdbcConnection
   $conn.ConnectionString = "DSN=$dsn;"
   $conn.open()
   $cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
   $ds = New-Object system.Data.DataSet
   (New-Object system.Data.odbc.odbcDataAdapter($cmd)).fill($ds) | out-null
   $conn.close()
   $ds.Tables[0]



Function Test-ODBCConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,
                    HelpMessage="DSN name of ODBC connection")]
                    [string]$DSN =""
    )
    $conn = new-object system.data.odbc.odbcconnection
    $conn.connectionstring = "(DSN=$DSN)"
    
    try {
        if (($conn.open()) -eq $true) {
            $conn.Close()
            $true
        }
        else {
            $false
        }
    } catch {
        Write-Host $_.Exception.Message
        $false
    }
}

function Get-ODBC-Data{
   param(
   [string]$query=$(throw 'query is required.'),
   [string]$dsn
   )
   $conn = New-Object System.Data.Odbc.OdbcConnection
   $conn.ConnectionString = "DSN=$dsn;"
   $conn.open()
   $cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
   $ds = New-Object system.Data.DataSet
   (New-Object system.Data.odbc.odbcDataAdapter($cmd)).fill($ds) | out-null
   $conn.close()
   $ds.Tables[0]
}
 


function Set-ODBC-Data{
  param(
  [string]$query=$(throw 'query is required.'),
  [string]$dsn
  )
  $conn = New-Object System.Data.Odbc.OdbcConnection
  $conn.ConnectionString= "DSN=$dsn;"
  $cmd = new-object System.Data.Odbc.OdbcCommand($query,$conn)
  $conn.open()
  $cmd.ExecuteNonQuery()
  $conn.close()
}


function Set-ODBC-Data{
  param(
  [string]$query=$(throw 'query is required.'),
  [string]$dsn
  )
  $conn = New-Object System.Data.Odbc.OdbcConnection
  $conn.ConnectionString= "DSN=$dsn;"
  $cmd = new-object System.Data.Odbc.OdbcCommand($query,$conn)
  $conn.open()
  $cmd.ExecuteNonQuery()
  $conn.close()
}
 

 Test-ODBCConnection -DSN Kunde64


#
# https://www.andersrodland.com/working-with-odbc-connections-in-powershell/