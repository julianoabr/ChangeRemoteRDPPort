#Requires -Version 4.0
#Requires -RunAsAdministrator 
<#
.Synopsis
   Change RDP Port
.DESCRIPTION
   Change RDP Port on computer
.EXAMPLE
   Set-RDPPort -rServerName $rServerName -rdpPortNumber 33389
.EXAMPLE
   Set-RDPPort -rServerName $rServerName -rdpPortNumber 33889
.AUTHOR
  Juliano Alves de Brito Ribeiro (Find me at https://github.com/julianoabr or jaribeiro@uoldiveo.com or julianoalvesbr@live.com)
.VERSION
  0.2
.ENVIRONMENT
  PROD
.NEXT IMPROVEMENTS
  0.3 Create Menu to Change RDP Port in one, two or more computers from a list or in all AD Computers
  0.4 Insert command to automatically open mstsc
  0.5 Rollback function
    
#BASED ON http://woshub.com/change-rdp-port-3389-windows/

.TOTHINK

This is was written more than 2000 years ago, and we are so close to this. Revelation Chapter 13. 

15 The second beast was given power to give breath to the image of the first beast, so that the image could speak and cause all who refused to worship the image to be killed.
16 It also forced all people, great and small, rich and poor, free and slave, to receive a mark on their right hands or on their foreheads, 
17 so that they could not buy or sell unless they had the mark, which is the name of the beast or the number of its name.

#>

###### GET OS VERSION ######
$rWmiOSVersion = {param([string]$rServer) 
try { $rOsVersion = ((Get-WmiObject -ComputerName $rServer -Class win32_operatingsystem -ErrorAction Stop).Version).ToString() }
catch{$rOsVersion = "Failure" }
return $rOsVersion
}


function Set-RDPPort
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $rServerName=$env:COMPUTERNAME,

        # Param2 help description
         [Parameter(Mandatory=$false,
                   HelpMessage='When choosing a non-standard RDP port, please note that it is not recommended to use port 1-1023 (known ports) and dynamic RPC port range 49152-65535',
                   Position=1)]
        [ValidateRange(1024,49151)]
        [System.Int32]$rdpPortNumber,

        [Parameter(Mandatory=$false,
                   HelpMessage='If this value is present, test remote port after change it', 
                   Position=2)]
        [switch]$testRemoteRDPPort
        
    )

    if ($rServerName -eq $env:COMPUTERNAME){
    
    
        $localRDPPort = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name PortNumber).PortNumber

        Write-Host "Server: $rServerName. Actual RDP Port Number: " -ForegroundColor White -NoNewline; Write-Host $localRDPPort -ForegroundColor Green; 

        #IF PORT IS NOT NULL ASK IF YOU WANT TO CHANGE
        if ($rdpPortNumber -ne $null){
         
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-TCP\" -Name PortNumber -Value $rdpPortNumber

            Restart-Service -Name TermService -Force -Verbose
            
            Write-host "The number of the RDP port has been changed to $rdpPortNumber " -ForegroundColor Magenta
                    
            $fwStatus = (Get-service -Name mpssvc).Status

            #CREATE RULE IF FW SERVICE IS RUNNING
            if ($fwStatus -eq 'Running'){
                                    
                $localOSVersion = ((Get-WmiObject -Class Win32_OperatingSystem).Version).ToString()

                if ($localOSVersion -match '^6\.[0-1].\.*'){
                    
                    netsh advfirewall firewall add rule name= "Remote Desktop (TCP Port $rdpPortNumber)" dir=in action=allow protocol=TCP localport=$rdpPortNumber

                    netsh advfirewall firewall add rule name= "Remote Desktop (UDP Port $rdpPortNumber)" dir=in action=allow protocol=UDP localport=$rdpPortNumber
                    
                }#end of If
                elseif(($localOSVersion -match '^6.[2-3].\.*') -or ($localOSVersion -match '^10.0.\.*')){
                    
                    
                    New-NetFirewallRule -DisplayName "Remote Desktop (TCP Port $rdpPortNumber)" -Direction Inbound –LocalPort $rdpPortNumber -Protocol TCP -Action Allow -Enabled True -Verbose
                                
                    New-NetFirewallRule -DisplayName "Remote Desktop (UDP Port $rdpPortNumber)" -Direction Inbound –LocalPort $rdpPortNumber -Protocol UDP -Action Allow -Enabled True -Verbose
                                        
                }#end of ElseIf
                else{
                    
                    Write-Host "I can't run on OS Version: $localOSVersion" -ForegroundColor White -BackgroundColor Red
                    
                }#end of Else
            
                        
            }#end of If
            else{
        
                Write-Host "Firewall Service is stopped. Please verify if this service can be started and run this script again" -ForegroundColor White -BackgroundColor DarkBlue

            }#end of Else
            #ASK IF YOU WANT TO CHANGE PORT
                             
        
        }#end of if rdpPortNull
        else{
                
                $YesNoAnswer = ""

                do
                {
            
                    $YesNoAnswer = Read-Host "Do you want to change RDP Port?" 
        
                }
                while ($YesNoAnswer -notmatch "^(?:Y\b|N\b)")


                if ($YesNoAnswer -match "^(\bY\b)"){
            
                    Write-host "Specify the number of your new RDP port:(1024-49151)" -ForegroundColor Yellow -NoNewline;$rdpPortNumber = Read-Host

                    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-TCP\" -Name PortNumber -Value $rdpPortNumber

                    Restart-Service -Name TermService -Force -Verbose
            
                    Write-host "The number of the RDP port has been changed to $rdpPortNumber " -ForegroundColor Magenta
                                     
                    $fwStatus = (Get-service -Name mpssvc).Status

                    #CREATE RULE IF FW SERVICE IS RUNNING
                    if ($fwStatus -eq 'Running'){
                                        
                        $localOSVersion = ((Get-WmiObject -Class Win32_OperatingSystem).Version).ToString()

                        if ($localOSVersion -match '^6\.[0-1].\.*'){
                    
                            netsh advfirewall firewall add rule name= "Remote Desktop (TCP Port $rdpPortNumber)" dir=in action=allow protocol=TCP localport=$rdpPortNumber

                            netsh advfirewall firewall add rule name= "Remote Desktop (UDP Port $rdpPortNumber)" dir=in action=allow protocol=UDP localport=$rdpPortNumber
                    
                        }#end of If
                        elseif(($localOSVersion -match '^6.[2-3].\.*') -or ($localOSVersion -match '^10.0.\.*')){
                    
                            New-NetFirewallRule -DisplayName "Remote Desktop (TCP Port $rdpPortNumber)" -Direction Inbound –LocalPort $rdpPortNumber -Protocol TCP -Action Allow -Enabled True -Verbose
                                
                            New-NetFirewallRule -DisplayName "Remote Desktop (UDP Port $rdpPortNumber)" -Direction Inbound –LocalPort $rdpPortNumber -Protocol UDP -Action Allow -Enabled True -Verbose
                                        
                        }#end of ElseIf
                        else{
                    
                            Write-Host "I can't run on OS Version: $localOSVersion" -ForegroundColor White -BackgroundColor Red
                    
                        }#end of Else
                    
                
                    }#end of If Verify Firewall service is running
                    else{
                
                        Write-Host "Firewall Service is stopped. Please verify if this service can be started and run this script again" -ForegroundColor White -BackgroundColor DarkBlue
                
                    }#end of Else Verify Firewall service is running
                                
            
                }#end of IF Actions if answer is yes
                elseif ($YesNoAnswer -match "^(\bN\b)"){
            
                    Write-Host "You selected $YesNoAnswer. I will not do anything" -ForegroundColor White -BackgroundColor Red                
                        
                }#end of elseIF Actions if answer is NO
                else{
            
                    Write-Host "I can't recognize the option selected" -ForegroundColor White -BackgroundColor Red          
            
                }#end of Else no Answer
        
        }#end of Else RDP Port is Not Null
          
    
    }#end of IF ComputerName
    else{
              
        
        $rSession = New-PSSession -ComputerName $rServerName
        
        $rRDPPort = Invoke-Command -Session $rSession -ScriptBlock {(Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name PortNumber).PortNumber}

        Write-Host "Server: $rServerName. Actual RDP Port Number: " -ForegroundColor White -NoNewline; Write-Host $rRDPPort -ForegroundColor Green; 

        #IF PORT IS NOT NULL ASK IF YOU WANT TO CHANGE
        if ($rdpPortNumber -ne $null){
         
            Invoke-Command -Session $rSession -ScriptBlock {Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-TCP\" -Name PortNumber -Value $using:rdpPortNumber}

            Get-Service -ComputerName $rServerName -Name TermService | Restart-Service -Force -Verbose
                              
            Write-host "The number of the RDP port has been changed to $rdpPortNumber " -ForegroundColor Magenta
                    
            $fwStatus = (Get-service -ComputerName $rServerName -Name mpssvc).Status

            #CREATE RULE IF FW SERVICE IS RUNNING
            if ($fwStatus -eq 'Running'){
                                    
                $rOSVersion = (Get-WmiObject -ComputerName $rServerName -Class Win32_operatingSystem).Version.ToString()

                if ($rOSVersion -match '^6\.[0-1].\.*'){
                    
                    Invoke-Command -Session $rSession -ScriptBlock {netsh advfirewall firewall add rule name= "Remote Desktop (TCP Port $using:rdpPortNumber)" dir=in action=allow protocol=TCP localport=$using:rdpPortNumber}

                    Invoke-Command -Session $rSession -ScriptBlock {netsh advfirewall firewall add rule name= "Remote Desktop (UDP Port $using:rdpPortNumber)" dir=in action=allow protocol=UDP localport=$using:rdpPortNumber}
                    
                }#end of If
                elseif(($localOSVersion -match '^6.[2-3].\.*') -or ($localOSVersion -match '^10.0.\.*')){
                    
                    Invoke-Command -Session $rSession -ScriptBlock {New-NetFirewallRule -DisplayName "Remote Desktop (TCP Port $using:rdpPortNumber)" -Direction Inbound –LocalPort $using:rdpPortNumber -Protocol TCP -Action Allow -Enabled True -Verbose}
                                                   
                    Invoke-Command -Session $rSession -ScriptBlock {New-NetFirewallRule -DisplayName "Remote Desktop (UDP Port $using:rdpPortNumber)" -Direction Inbound –LocalPort $using:rdpPortNumber -Protocol UDP -Action Allow -Enabled True -Verbose}
                                        
                }#end of ElseIf
                else{
                    
                    Write-Host "I can't run on OS Version: $rOSVersion" -ForegroundColor White -BackgroundColor Red
                    
                }#end of Else
            
                        
            }#end of If
            else{
        
                Write-Host "Firewall Service is stopped on remote computer: $rServerName. Please verify if this service can be started and run this script again" -ForegroundColor White -BackgroundColor DarkBlue

            }#end of Else
            #ASK IF YOU WANT TO CHANGE PORT
                             
        
        }#end of if rdpPortNull
        else{
                
                $YesNoAnswer = ""

                do
                {
            
                    $YesNoAnswer = Read-Host "Do you want to change RDP Port?" 
        
                }
                while ($YesNoAnswer -notmatch "^(?:Y\b|N\b)")


                if ($YesNoAnswer -match "^(\bY\b)"){
            
                    Write-host "Specify the number of your new RDP port:(1024-49151)" -ForegroundColor Yellow -NoNewline;$rdpPortNumber = Read-Host

                    Invoke-Command -Session $rSession -ScriptBlock {Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-TCP\" -Name PortNumber -Value $using:rdpPortNumber}

                    Get-Service -ComputerName $rServerName -Name TermService | Restart-Service -Force -Verbose

                    Write-host "The number of the RDP port on server: $rServerName has been changed to $rdpPortNumber " -ForegroundColor Magenta
                                     
                    $fwStatus = (Get-service -ComputerName $rServerName -Name mpssvc).Status

                    #CREATE RULE IF FW SERVICE IS RUNNING
                    if ($fwStatus -eq 'Running'){
                                        
                        $rOSVersion = (Get-WmiObject -ComputerName $rServerName -Class Win32_operatingSystem).Version.ToString()

                        if ($rOSVersion -match '^6\.[0-1].\.*'){
                    
                            Invoke-Command -Session $rSession -ScriptBlock {netsh advfirewall firewall add rule name= "Remote Desktop (TCP Port $using:rdpPortNumber)" dir=in action=allow protocol=TCP localport=$using:rdpPortNumber}

                            Invoke-Command -Session $rSession -ScriptBlock {netsh advfirewall firewall add rule name= "Remote Desktop (UDP Port $using:rdpPortNumber)" dir=in action=allow protocol=UDP localport=$using:rdpPortNumber}
                    
                        }#end of If
                        elseif(($rOSVersion -match '^6.[2-3].\.*') -or ($rOSVersion -match '^10.0.\.*')){
                    
                            Invoke-Command -Session $rSession -ScriptBlock {New-NetFirewallRule -DisplayName "Remote Desktop (TCP Port $using:rdpPortNumber)" -Direction Inbound –LocalPort $using:rdpPortNumber -Protocol TCP -Action Allow -Enabled True -Verbose}
                                                   
                            Invoke-Command -Session $rSession -ScriptBlock {New-NetFirewallRule -DisplayName "Remote Desktop (UDP Port $using:rdpPortNumber)" -Direction Inbound –LocalPort $using:rdpPortNumber -Protocol UDP -Action Allow -Enabled True -Verbose}
                                        
                        }#end of ElseIf
                        else{
                    
                            Write-Host "I can't run on Remote Server: $rServerName with OS Version: $rOSVersion" -ForegroundColor White -BackgroundColor Red
                    
                        }#end of Else
                    
                
                    }#end of If Verify Firewall service is running
                    else{
                
                        Write-Host "Firewall Service is stopped on server: $rServerName. Please verify if this service can be started and run this script again" -ForegroundColor White -BackgroundColor DarkBlue
                
                    }#end of Else Verify Firewall service is running
                                
            
                }#end of IF Actions if answer is yes
                elseif ($YesNoAnswer -match "^(\bN\b)"){
            
                    Write-Host "You selected $YesNoAnswer. I will not do anything" -ForegroundColor White -BackgroundColor Red                
                        
                }#end of elseIF Actions if answer is NO
                else{
            
                    Write-Host "I can't recognize the option selected" -ForegroundColor White -BackgroundColor Red          
            
                }#end of Else no Answer
        
        }#end of Else RDP Port is Not Null
        
        #VALIDATE CHANGE ON REMOTE PORT
        if ($testRemoteRDPPort.IsPresent){
        
            $connectTestResult = Test-NetConnection -ComputerName $rServerName -Port $rdpPortNumber

            if ($connectTestResult.TcpTestSucceeded) {
            
                 Write-Host "Server: $rServerName. Actual RDP Port Number: " -ForegroundColor White -NoNewline; Write-Host -NoNewline $rdpPortNumber -ForegroundColor Green;Write-Host -NoNewline " is reachable remotely" -ForegroundColor White      

            }
            else {
            
            Write-Error -Message "Unable to reach server: $rServerName through Port: $rdpPortNumber. Check to make sure your organization or ISP is not blocking port $rdpPortNumber"
            
            }

        
        }#end of IF
        else{
        
           Write-Host "Remote Port $rdpPortNumber on server: $rServerName will not tested" -ForegroundColor DarkBlue -BackgroundColor White
        
        }#end of Else
        
    
    }#end of Else Computername


}#end of Function


Set-RDPPort -rServerName $rServerName -rdpPortNumber 33389