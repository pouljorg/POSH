Function LogIt {
[cmdLetBinding()]
Param (
    [String]$Message,
    [String]$Destination ,
    [switch]$Append
)
    $outString = "{0} - {1} - {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), "$($env:USERDOMAIN)\$($env:USERNAME)", $Message
    if ( $Append ) {
        $outString | Out-File -Append -FilePath $Destination
    } else {
        $outString | Out-File -FilePath $Destination
    }
}

Function Get-LocationNames {
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,Position=0)]
    [String]$Coordinates,
    [string]$APIKey='AlkHm-0g08LqlDkYg2T47AxUAOgHVFDfDjJUddLS-AKe1uJSjkFT-JyU2qwagwFK'

)

#$Coordiantes = '55.4860267638889, 9.74988079055555'
$data =Invoke-RestMethod  "http://dev.virtualearth.net/REST/v1/Locations/$($Coordinates)?key=$APIKey"
$d = ($data.resourceSets.Resources[0].address) | Select @{Label='LocationInfo';expression={"$($_.CountryRegion) - $($_.AdminDistrict) - $($_.AdminDistrict2) - $($_.PostalCode) - $($_.AddressLine)"}}
#Write-Host -Object $d -ForegroundColor Cyan
return $d
<#
addressLine      : Hjarnøvænget 3
adminDistrict    : Southern Denmark Region
adminDistrict2   : Middelfart
countryRegion    : Denmark
formattedAddress : Hjarnøvænget 3, 5500 Middelfart, Denmark
locality         : Middelfart
postalCode       : 5500
#>

}

<#
.Synopsis
   The Get-ExifProperty retrieves informatin in  property not directly being readable
.DESCRIPTION
   The Get-ExifProperty retrieves informatin in  property not directly being readable. Depending on the Type of the property it is handled in different ways.
    This function is only supposed to be called from Get-ExifPictureData
.PARAMETER Image
    The Image ( ComObject ) to process on 
.PARAMETER ExifPropId
    The Property to return information about
#>
Function Get-EXIFProperty {
[CmdletBinding()]
Param(
  [__ComObject]$Image,
  [string]$ExifPropId
)
    Try {    
        $Prop =$Image.Properties.Item("$ExifPropId") 
    }
   
    Catch {
        Logit -Message "Unable to read property: $ExifPropId of picture $pic" -Append -Destination $logdestination
        $Prop
        Return #no data to process
    }
     If ($prop.Type -in '1006','1007') {
        if ($Prop.Name -eq 'ExifExposureTime') {
           $o=  "1/$([int]($Prop.Value.Denominator / $Prop.Value.Numerator))"
        } else {
           $o=  $Prop.Value.Value
        }
    }
    elseif ($ExifPropId -eq '7') {
        $o =  ($prop.value | Select -ExpandProperty Value) -join ':'
    }
    elseif ($ExifPropId -in '2','4') {
        $o =  ($prop.value | Select -ExpandProperty Value) -join '; '
    }
    elseif ($Prop.Type -in '1100','1101'){
        $o =  #[System.Text.Encoding]::ASCII.GetString($Prop.Value) -replace "`0$"
              #  $Prop.value.tostring()
            if (($prop.value | measure -Maximum | select -ExpandProperty Maximum) -gt 20 -and ($prop.value | select -First 1 ) -gt 10) {
                $b=''
                foreach ($a in $prop.Value) { $b+=if( $a -gt 0) { [char]$a} }
                $b
             }
             else { $Prop.value -join '' } 
              #  $o | gm
    } 
    else {
        $o= $Prop.Value
    }
    Write-Verbose -Message "Prop: $o"
    return $o
}
    
<#
.Synopsis
   The Get-ExifPictureData retrieves properties of the picture
.DESCRIPTION
   The Get-ExifPictureData retrieves properties of the picture. Around 70 porperties with image information is returned
.PARAMETER Picture
    This Parameter is used to specify the picture to process. It can be a path to a directory. I so all files within thaat directory will be processed
.PARAMETER Recurse
    If this switch is specified all files in subdirectories will be processed too
.EXAMPLE
   'C:\Users\ThisUser\Pictures\2016-12 | Get-ExifPicturedata -recurse
   This retrieves informatin about all pictures in the C:\Users\ThisUser\Pictures\2016-12 including files in subdirectories

.EXAMPLE
   Get-ExifPicturedata -Picture C:\Users\ThisUser\Pictures\2016-12 
   This retrieves informatin about all pictures in the C:\Users\ThisUser\Pictures\2016-12 directory
.EXAMPLE
    Get-ExifPicturedata -Picture C:\Users\ThisUser\Pictures\2016-12\MyPic.jpg
    This retrieves picture information for the picture: C:\Users\ThisUser\Pictures\2016-12\MyPic.jpg
EXIF info from here: 
#>
Function Get-ExifPicturedata {
[CmdLetBinding()]
Param (
    [Parameter(ValueFromPipeline=$True,Mandatory=$true,Position=0)]
    [String[]]$Picture,
    [switch]$Recurse,
    [switch]$IncludeLocationWithName,
    [string]$LogDestination = "$env:USERPROFILE\desktop\$(Get-date -Format 'yyyy-MM-dd HH:mm').txt"
)
    Begin {
        $image = New-Object -ComObject WIA.ImageFile
        $ReplaceName = @{'27' = 'GPS Processing Method';
                         '29' = 'GPSDate';
                         '41985' = 'Custom Rendered'; 
                         '41986' = 'ExposureMode'; 
                         '41987' = 'WhiteBalance'; 
                         '41988' = 'DigitalZoomRatio'; 
                         '41989' = 'FocalLengthIn35mmFormat'; 
                         '41990' = 'SceneCaptureMode'; 
                         '41991' = 'Gain Control'; 
                         '41992' = 'Contrast'; 
                         '41993' = 'Saturation'; 
                         '41994' = 'Sharpness'; 
                         '41995' = 'Device Settings Description'; 
                         '41996' = 'SubjectRange'} 
    }
    Process {
            if (Test-Path -Path $Picture -PathType Container) {
                if ($Recurse) {
                    $Picture = (Get-ChildItem -Path $Picture -Recurse -File).FullName
                } else {
                    $Picture = (Get-ChildItem -Path $Picture -File).FullName
                }
            }        Foreach ($Pic in $Picture ) {
            Try {
                $image.LoadFile($Pic)
            }
            Catch {
                Write-Information "Unable to read picture: $pic"
                LogIt -Message "Unable to read picture: $pic" -Append -Destination $LogDestination
            }
            #$ReplaceName
            $BaseProp = $image.Properties | where { $_.Value -isnot [System.__ComObject] -and $_.Name -notlike 'ThumbNail*' }
            $hash = @{}
            $hash.Add('Path',$Pic)
            ForEach ($BProp in $BaseProp ) {
                if ($BProp.Name -like '419??' -or $BProp.Name -in 27,29) { $Label = $ReplaceName.item($BProp.Name) } else { $Label = $BProp.Name }
                $Label = $label -replace '^ExifID','' -replace '^Exif', ''

                $value = if ( $BProp.Value  -match '^\d{4}:\d\d:\d\d \d\d:\d\d:\d\d$') {
                     [DateTime]::ParseExact($BProp.Value,"yyyy:MM:dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture)     
                } else {
                    $BProp.Value
                }
                $Prid = [int]($BProp.PropertyID)
                if ( $PrID -in 274,34850,37383,37384, 37385,40961,41728,41986,41987,41990,41991,41992,41993,41994,41996) {
    
                    switch ($PrID)  {
                        274 { #Orientation
                        Switch ([int]($BProp.Value)) {
                            1 { $value = 'Horizontal (normal)' ; break }
                            2 { $value = 'Mirror horizontal' ; break }
                            3 { $value = 'Rotate 180' ; break }
                            4 { $value = 'Mirror vertical' ; break }
                            5 { $value = 'Mirror horizontal and rotate 270 CW' ; break }
                            6 { $value = 'Rotate 90 CW' ; break }
                            7 { $value = 'Mirror horizontal and rotate 90 CW' ; break }
                            8 { $value = 'Rotate 270 CW' ; break }
                            }
                        }
                        34850 { #Exposure Program
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Not Defined' ; break }
                            1 { $value = 'Manual' ; break }
                            2 { $value = 'Program AE' ; break }
                            3 { $value = 'Aperture-priority AE' ; break }
                            4 { $value = 'Shutter speed priority AE' ; break }
                            5 { $value = 'Creative (Slow speed) ' ; break }
                            6 { $value = 'Action (High speed)' ; break }
                            7 { $value = 'Portrait' ; break }
                            8 { $value = 'Landscape' ; break }
                            9 { $value = 'Bulb' ; break }
                            }
                        }
                        37383 { # Metering Mode
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Unknown' ; break }
                            1 { $value = 'Average' ; break }
                            2 { $value = 'Center-weighted average' ; break }
                            3 { $value = 'Spot' ; break }
                            4 { $value = 'Multi-spot' ; break }
                            5 { $value = 'Multi-segment' ; break }
                            6 { $value = 'Partial' ; break }
                            255 { $value = 'Other' ; break }
                            }
                        }
                        37384 { # Light Source
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Auto' ; break }
                            1 { $value = 'Daylight' ; break }
                            2 { $value = 'Fluorescent' ; break }
                            3 { $value = 'Tungsten' ; break }
                            4 { $value = 'Flash' ; break }
                            9 { $value = 'Fine Weather' ; break }
                            10 { $value = 'Cloudy Weather' ; break }
                            11 { $value = 'Shade' ; break }
                            12 { $value = 'Daylight Fluorescent' ; break }
                            13 { $value = 'Day White Fluorescent' ; break }
                            14 { $value = 'Cool White Fluorescent' ; break }
                            15 { $value = 'White Fluorescent' ; break }
                            17 { $value = 'Standard Light A' ; break }
                            18 { $value = 'Standard Light B' ; break }
                            19 { $value = 'Standard Light C' ; break }
                            20 { $value = 'D55' ; break }
                            21 { $value = 'D65' ; break }
                            22 { $value = 'D75' ; break }
                            23 { $value = 'D50' ; break }
                            24 { $value = 'ISO Studio Tungsten'} 
                            }
                        }
                        37385 { # Flash
                       Switch ([int]($BProp.Value)) {
                        0	{ $Value = 'No Flash'; break }
                        1	{ $Value = 'Fired'; break }
                        5	{ $Value = 'Fired, Return not detected'; break }
                        7	{ $Value = 'Fired, Return detected'; break }
                        8	{ $Value = 'On, Did not fire'; break }
                        9	{ $Value = 'On, Fired'; break }
                        12	{ $Value = 'On, Return not detected'; break }
                        15	{ $Value = 'On, Return detected'; break }
                        16	{ $Value = 'Off, Did not fire'; break }
                        20	{ $Value = 'Off, Did not fire, Return not detected'; break }
                        24	{ $Value = 'Auto, Did not fire'; break }
                        25	{ $Value = 'Auto, Fired'; break }
                        28	{ $Value = 'Auto, Fired, Return not detected'; break }
                        31	{ $Value = 'Auto, Fired, Return detected'; break }
                        32	{ $Value = 'No flash function'; break }
                        48	{ $Value = 'Off, No flash function'; break }
                        49	{ $Value = 'Fired, Red-eye reduction'; break }
                        69	{ $Value = 'Fired, Red-eye reduction, Return not detected'; break }
                        71	{ $Value = 'Fired, Red-eye reduction, Return detected'; break }
                        73	{ $Value = 'On, Red-eye reduction'; break }
                        76	{ $Value = 'On, Red-eye reduction, Return not detected'; break }
                        79	{ $Value = 'On, Red-eye reduction, Return detected'; break }
                        80	{ $Value = 'Off, Red-eye reduction'; break }
                        88	{ $Value = 'Auto, Did not fire, Red-eye reduction'; break }
                        89	{ $Value = 'Auto, Fired, Red-eye reduction'; break }
                        92	{ $Value = 'Auto, Fired, Red-eye reduction, Return not detected'; break }
                        95	{ $Value = 'Auto, Fired, Red-eye reduction, Return detected' ; break }
                        }
                    }
                        40961 { # Color Space
                        Switch ([int]($BProp.Value)) {
                            1 { $value = 'sRGB' ; break }
                            2 { $value = 'Adobe RGB' ; break }
                            }
                        }
                        41495{ # Sensing Method
                        Switch ([int]($BProp.Value)) {
                            1	{ $Value = 'Not Defined'; break }
                            2	{ $Value = 'One-Chip Color area'; break }
                            3	{ $Value = 'Two-Chip Color area'; break }
                            4	{ $Value = 'Three-Chip Color aread'; break }
                            5	{ $Value = 'Color sequential area'; break }
                            7	{ $Value = 'Trilinear '; break }
                            8	{ $Value = 'Color sequential linea'; break }
                            }
                        }
                        41728 { #File Source
                        Switch ([int]($BProp.Value)) {
                            1 { $value = 'Film Scanner' ; break }
                            2 { $value = 'Reflection Print Scanner' ; break }
                            3 { $value = 'Digital Camera' ; break }
                            }
                        }
                        41985 { # Custom Rendered
                        Switch ([int]($BProp.Value)) {
                            0	{ $Value = 'Normal'; break }
                            1	{ $Value = 'Custom'; break }
                            3	{ $Value = 'HDR'; break }
                            6	{ $Value = 'Panorama'; break }
                            8	{ $Value = 'PorTrait'; break }
 
                            }
                        } 
                        41986 { #Exposure mode
    	                 Switch ([int]($BProp.Value)) {
                            0 { $value = 'Auto' ; break }
                            1 { $value = 'Manual' ; break }
                            2 { $value = 'Auto bracket' ; break }
                            }
                        }
                        41987 { # white Balance
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Auto' ; break }
                            1 { $value = 'manual' ; break }
                            }
                        }
                        41990 { # Scene Capture Mode
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Standard' ; break }
                            1 { $value = 'Landscape' ; break }
                            2 { $value = 'Portrait' ; break }
                            3 { $value = 'Night' ; break }
                            }
                    }
                        41991 { #Gani Control
	                     Switch ([int]($BProp.Value)) {
                            0  { $value = 'None' ; break }
                            1  { $value = 'Low gain up' ; break }
                            2  { $value = 'High gain up' ; break }
                            3  { $value = 'Low gain down' ; break }
                            4  { $value = 'High gain down' ; break }
                           }
                        }
                        41992 { #Contrast
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Normal' ; break }
                            1 { $value = 'Low' ; break }
                            2 { $value = 'High' ; break }
                           }
                        }
                        41993 { #Saturation
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Normal' ; break }
                            1 { $value = 'Low' ; break }
                            2 { $value = 'High' ; break }
                            }
                        }
                        41994 { #Sharpness
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Normal' ; break }
                            1 { $value = 'Soft' ; break }
                            2 { $value = 'Hard' ; break }
                            }
                        }
                        41996 { #subject Range
	                     Switch ([int]($BProp.Value)) {
                            0 { $value = 'Unknown' ; break }
                            1 { $value = 'Macro' ; break }
                            2 { $value = 'Close' ; break }
                            3 { $value = 'Distant' ; break }
                        }
                    }
                    } #end outer switch
    
                } #end if                
                $hash.Add($Label,$Value)
                Write-Verbose "l:$label - val:$($BProp.Value) - ip: $($BProp.PropertyID)"
            } # end foreach prop base
          #$hash
          #Read-Host
            $BaseProp = $image.Properties | where { $_.Value -is [System.__ComObject] -and $_.Name -notlike 'ThumbNail*' -and $_.Name -notlike '*table' -and  $_.PropertyId -ne '37500'}
            ForEach ($BProp in $BaseProp ) {
                if ($BProp.Name -like '419??'){ $Label = $ReplaceName.Item($BProp.Name) } else { $Label = $BProp.Name }
                $propId = $BProp.PropertyID
                $data = Get-EXIFProperty -Image $image -ExifPropId $propId
                $value = if ( $data  -match '^\d{4}:\d\d:\d\d \d\d:\d\d:\d\d$') {
                     [DateTime]::ParseExact($data,"yyyy:MM:dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture)     
                } else {
                    $data
                }                
                $Label = $label -replace '^ExifID','' -replace '^Exif', ''
                Write-Verbose "her :$data"
                 $hash.Add( $Label , $value )
                # 1075 -1155  & 1179 -1191
            }# End foreach custom Prop
                $q = $hash['GPSLatitude'] -split '; '
                if ($q -ne $null ) {
                    $coor = "$($q[0])°$($q[1])'$($q[2])"" "
                    $coord = [int]($q[0])+([int]($q[1])/60) + ([double]($q[2] -replace ',', '.')/3600)
                    if ( $hash['GPSLatitudeRef'] -eq 'S') { $coord = -$coord }
                        
                    #write-host $coord
                    $coor +=  $hash['GPSLatitudeRef']
                    $q = $hash['GPSLongitude'] -split '; '
                    $coor += " $($q[0])°$($q[1])'$($q[2])"" "
                    $coor +=  $hash['GPSLongitudeRef']
                    $coorc = [int]($q[0])+([int]($q[1])/60) + ([double]($q[2] -replace ',', '.')/3600)
                    if ( $hash['GPSLongitudeRef'] -eq 'W') { $coorc = -$coorc }
                    $hash.Add('GPSCoord.',($coor -replace ',', '.'))
                    $coo = "$($coord -replace ',','.'), $($coorc -replace ',','.')"
                    $hash.Add('GPSGeoLoc',$coo )
                    if ($IncludeLocationWithName) {
                        $LocName= (Get-LocationNames -Coordinates $coo).LocationInfo
                        $hash.Add('LocNameInfo',$LocName)
                    }
                }
            $obj = New-Object -TypeName PSObject -Property $hash
            Write-Output $obj
        }#end Foreach Pic

    }
    End {
    }
}
#(Get-ChildItem -Path C:\Users\W530\Pictures\2016-12 -File).FullName | Get-ExifPicturedata
#Get-ExifPicturedata -Picture C:\Users\W530\Pictures\2016-12
#$img = 'C:\Users\W530\Pictures\NYC APR 2017\IMG_20170415_055839.jpg'
#$img = 'C:\Users\W530\Pictures\2016-12\IMG_20161224_111516.jpg'
$img = 'C:\Users\poulj\Pictures\IMG_20161224_144513.jpg' # 'C:\Users\W530\OneDrive\SkyDrive camera roll\WP_20131207_23_13_10_Pro.jpg','C:\Users\W530\Pictures\2016-12\IMG_20161224_111516.jpg'
Get-ExifPicturedata -Picture $img -IncludeLocationWithName # | select *gps*
<#
$GPSGeoLoc = '55.4860267638889, 9.74988079055555'
$PJKey= 'AlkHm-0g08LqlDkYg2T47AxUAOgHVFDfDjJUddLS-AKe1uJSjkFT-JyU2qwagwFK'
$data =Invoke-RestMethod  "http://dev.virtualearth.net/REST/v1/Locations/$($GPSGeoLoc)?key=$PJKey"
$data.resourceSets.Resources[0].address
addressLine      : Hjarnøvænget 3
adminDistrict    : Southern Denmark Region
adminDistrict2   : Middelfart
countryRegion    : Denmark
formattedAddress : Hjarnøvænget 3, 5500 Middelfart, Denmark
locality         : Middelfart
postalCode       : 5500
#>

#Get-LocationNames -Coordiantes '55.4860267638889, 9.74988079055555' | select -ExpandProperty Locationinfo