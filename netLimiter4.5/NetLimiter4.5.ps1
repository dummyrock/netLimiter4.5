if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

function Get-NetworkStatistics
{
 $properties = 'Protocol','LocalAddress','LocalPort'
 $properties += 'RemoteAddress','RemotePort','State','ProcessName','PID'

 netstat -ano |Select-String -Pattern '\s+(TCP|UDP)' | ForEach-Object {

 $item = $_.line.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries)

 if($item[1] -notmatch '^\[::')
 {
 if (($la = $item[1] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6')
 {
 $localAddress = $la.IPAddressToString
 $localPort = $item[1].split('\]:')[-1]
 }
 else
 {
 $localAddress = $item[1].split(':')[0]
 $localPort = $item[1].split(':')[-1]
 }

 if (($ra = $item[2] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6')
 {
 $remoteAddress = $ra.IPAddressToString
 $remotePort = $item[2].split('\]:')[-1]
 }
 else
 {
 $remoteAddress = $item[2].split(':')[0]
 $remotePort = $item[2].split(':')[-1]
 }

New-Object PSObject -Property @{
 PID = $item[-1]
 ProcessName = (Get-Process -Id $item[-1] -ErrorAction SilentlyContinue).Name
 Protocol = $item[0]
 LocalAddress = $localAddress
 LocalPort = $localPort
 RemoteAddress =$remoteAddress
 RemotePort = $remotePort
 State = if($item[0] -eq 'tcp') {$item[3]} else {$null}
 } |Select-Object -Property $properties
 }
 }
}

Hide-Console

$global:isOnUDP = $true
$global:isOnTCP1 = $true
$global:isOnTCP2 = $true
$global:isOnTCP3 = $true

$formObject = [System.Windows.Forms.Form]
$labelObject = [System.Windows.Forms.Label]
$buttonObject = [System.Windows.Forms.Button]
$textBoxObject = [System.Windows.Forms.TextBox]
$dataGridObject = [System.Windows.Forms.DataGrid]

$limit3074Form = New-Object $formObject 
$limit3074Form.ClientSize = '1250,300'
$limit3074Form.BackColor = '#2B2B32'
$limit3074Form.FormBorderStyle = 'Fixed3D'
$limit3074Form.MaximizeBox = $false
$limit3074Form.ControlBox = $false

#fileLocation
$tBoxFileLocation = New-Object $textBoxObject
$tBoxFileLocation.Location = New-Object System.Drawing.Point(50,10)
$tBoxFileLocation.Size = New-Object System.Drawing.Size(200,20)
$tBoxFileLocation.BackColor = "#2B2B32"

$lblFileLocation = New-Object $labelObject
$lblFileLocation.Text = 'Enter D2 File Location without quotes'
$lblFileLocation.AutoSize = $true
$lblFileLocation.Font='verdana,10'
$lblFileLocation.Location = New-Object System.Drawing.Point(30,30)


#title text
$lblTitle = New-Object $labelObject
$lblTitle.Text = 'Limiter!!'
$lblTitle.AutoSize = $true
$lblTitle.Font='verdana,36,style=Bold'
$lblTitle.Location = New-Object System.Drawing.Point(55,50)

#subtitle text
$lblSubTitle1 = New-Object $labelObject
$lblSubTitle1.Text = 'To use this all you have to do is if the ports are UDP'
$lblSubTitle1.AutoSize = $true
$lblSubTitle1.Font='verdana,10'
$lblSubTitle1.Location = New-Object System.Drawing.Point(20,150)

$lblSubTitle2 = New-Object $labelObject
$lblSubTitle2.Text = 'enter the port you want to use, and then press'
$lblSubTitle2.AutoSize = $true
$lblSubTitle2.Font='verdana,10'
$lblSubTitle2.Location = New-Object System.Drawing.Point(20,170)

$lblSubTitle3 = New-Object $labelObject
$lblSubTitle3.Text = 'press the box to activate and deactivate. However'
$lblSubTitle3.AutoSize = $true
$lblSubTitle3.Font='verdana,10'
$lblSubTitle3.Location = New-Object System.Drawing.Point(20,190)

$lblSubTitle4 = New-Object $labelObject
$lblSubTitle4.Text = 'if it is a TCP port, enter the port and enter the'
$lblSubTitle4.AutoSize = $true
$lblSubTitle4.Font='verdana,10'
$lblSubTitle4.Location = New-Object System.Drawing.Point(20,210)

$lblSubTitle5 = New-Object $labelObject
$lblSubTitle5.Text = 'bits per second you want to limit and press'
$lblSubTitle5.AutoSize = $true
$lblSubTitle5.Font='verdana,10'
$lblSubTitle5.Location = New-Object System.Drawing.Point(20,230)

$lblSubTitle6 = New-Object $labelObject
$lblSubTitle6.Text = 'the box to turn on and off'
$lblSubTitle6.AutoSize = $true
$lblSubTitle6.Font='verdana,10'
$lblSubTitle6.Location = New-Object System.Drawing.Point(20,250)

#grid view button
$dataGridRP = New-Object $dataGridObject
$dataGridRP.Location = New-Object System.Drawing.Point(710,20)
$dataGridRP.Size = New-Object System.Drawing.Size(506,250)
$dataGridRP.BackColor = "#2B2B32"
$dataGridRP.ReadOnly = $true

#grid view button
$btnRP = New-Object $buttonObject
$btnRP.Location = New-Object System.Drawing.Point(670,100)
$btnRP.Size = New-Object System.Drawing.Size(25,25)


#UDP things
$lblUDPTitle = New-Object $labelObject
$lblUDPTitle.Text = 'UDP'
$lblUDPTitle.AutoSize = $true
$lblUDPTitle.Font = 'verdana,12,style=Bold'
$lblUDPTitle.Location = New-Object System.Drawing.Point(570,20)

$tBoxUDPPort = New-Object $textBoxObject
$tBoxUDPPort.Location = New-Object System.Drawing.Point(500,50)
$tBoxUDPPort.Size = New-Object System.Drawing.Size(100,20)
$tBoxUDPPort.BackColor = "#2B2B32"

$btnUDP=New-Object $buttonObject
$btnUDP.Location=New-Object System.Drawing.Point(620,47.5)
$btnUDP.Size = New-Object System.Drawing.Size(25,25)


#TCP things
$lblTCPTitle = New-Object $labelObject
$lblTCPTitle.Text = 'TCP'
$lblTCPTitle.AutoSize = $true
$lblTCPTitle.Font = 'verdana,12,style=Bold'
$lblTCPTitle.Location = New-Object System.Drawing.Point(570,110)

#TCP ports
$tBoxTCPPort1 = New-Object $textBoxObject
$tBoxTCPPort1.Location = New-Object System.Drawing.Point(440,150)
$tBoxTCPPort1.Size = New-Object System.Drawing.Size(100,20)
$tBoxTCPPort1.BackColor = "#2B2B32"

$tBoxTCPPort2 = New-Object $textBoxObject
$tBoxTCPPort2.Location = New-Object System.Drawing.Point(440,200)
$tBoxTCPPort2.Size = New-Object System.Drawing.Size(100,20)
$tBoxTCPPort2.BackColor = "#2B2B32"

$tBoxTCPPort3 = New-Object $textBoxObject
$tBoxTCPPort3.Location = New-Object System.Drawing.Point(450,250)
$tBoxTCPPort3.Size = New-Object System.Drawing.Size(100,20)
$tBoxTCPPort3.BackColor = "#2B2B32"

$tBoxTCPPort4 = New-Object $textBoxObject
$tBoxTCPPort4.Location = New-Object System.Drawing.Point(330,250)
$tBoxTCPPort4.Size = New-Object System.Drawing.Size(100,20)
$tBoxTCPPort4.BackColor = "#2B2B32"

#TCP bits per second
$tBoxTCPBts1 = New-Object $textBoxObject
$tBoxTCPBts1.Location = New-Object System.Drawing.Point(570,150)
$tBoxTCPBts1.Size = New-Object System.Drawing.Size(60,20)
$tBoxTCPbts1.BackColor = "#2B2B32"

$tBoxTCPBts2 = New-Object $textBoxObject
$tBoxTCPBts2.Location = New-Object System.Drawing.Point(570,200)
$tBoxTCPBts2.Size = New-Object System.Drawing.Size(60,20)
$tBoxTCPbts2.BackColor = "#2B2B32"

$tBoxTCPBts3 = New-Object $textBoxObject
$tBoxTCPBts3.Location = New-Object System.Drawing.Point(570,250)
$tBoxTCPBts3.Size = New-Object System.Drawing.Size(60,20)
$tBoxTCPbts3.BackColor = "#2B2B32"

#TCP button activator
$btnTCP1=New-Object $buttonObject
$btnTCP1.Location=New-Object System.Drawing.Point(670,147.5)
$btnTCP1.Size = New-Object System.Drawing.Size(25,25)

$btnTCP2=New-Object $buttonObject
$btnTCP2.Location=New-Object System.Drawing.Point(670,197.5)
$btnTCP2.Size = New-Object System.Drawing.Size(25,25)

$btnTCP3=New-Object $buttonObject
$btnTCP3.Location=New-Object System.Drawing.Point(670,247.5)
$btnTCP3.Size = New-Object System.Drawing.Size(25,25)

$limit3074Form.Controls.AddRange(@($lblTitle,$btnUDP,$lblSubTitle1,$lblSubTitle2,$lblSubTitle3,$lblSubTitle4,
$lblSubTitle5,$lblSubTitle6,$lblUDPTitle,$tBoxUDPPort,$lblTCPTitle,$tBoxTCPPort1,$tBoxTCPPort2,$tBoxTCPPort3,
$tBoxTCPPort4,$tBoxTCPBts1,$tBoxTCPBts2,$tBoxTCPBts3,$btnTCP1,$btnTCP2,$btnTCP3,$tBoxFileLocation,$lblFileLocation,
$dataGridRP,$btnRP))

function startLimitUDP{
    if($tBoxUDPPort.Text -ne '' -and $tBoxFileLocation.Text -ne ''){
        if($global:isOnUDP)
        {
        
            $targetPort11 = $tBoxUDPPort.Text
            $d2Location = $tBoxFileLocation.Text

            New-NetFirewallRule -DisplayName "Destiny-2-3074-1UDP" -Direction Inbound -Protocol UDP -RemotePort $targetPort11 -Program $d2Location -Action Block -EdgeTraversalPolicy Block
            New-NetFirewallRule -DisplayName "Destiny-2-3074-2UDP" -Direction Outbound -Protocol UDP -RemotePort $targetPort11 -Program $d2Location -Action Block -EdgeTraversalPolicy Block
            $global:isOnUDP = $false
            $btnUDP.BackColor = '#00FF00'
        }
        else
        {
            Remove-NetFirewallRule -DisplayName "Destiny-2-3074-1UDP" 
            Remove-NetFirewallRule -DisplayName "Destiny-2-3074-2UDP" 
            $global:isOnUDP = $true
            $btnUDP.BackColor = '#2B2B32'
        }
    }
}

function startLimitTCP1{
    if($tBoxTCP1Port.Text -ne '' -and $tBoxTCPBts1.Text -ne ''){
        if($global:isOnTCP1)
        {
        
            $targetPort11 = $tBoxTCP1Port.Text
            $d2Location = $tBoxFileLocation.Text
            
            New-NetFirewallRule -DisplayName "Destiny-2-3074-2TCP" -Direction Outbound -Protocol UDP -RemotePort $targetPort11 -Program $d2Location -Action Block -EdgeTraversalPolicy Block
            $global:isOnTCP1 = $false
            $btnTCP1.BackColor = '#00FF00'
        }
        else
        {
            Remove-NetFirewallRule -DisplayName "Destiny-2-3074-2TCP" 
            $global:isOnTCP1 = $true
            $btnTCP1.BackColor = '#2B2B32'
        }
    }
}

function startLimitTCP2{
    if($tBoxTCP2Port.Text -ne '' -and $tBoxTCPBts2.Text -ne ''){
        if($global:isOnTCP2)
        {
        
            $targetPort2 = $tBoxTCP2Port.Text
            $targetBTS2 = $tBoxTCPBts2.Text

            New-NetQosPolicy -Name "Destiny-2-3074-2TCP2" -IPProtocol TCP  -AppPathNameMatchCondition "destiny2.exe" -ThrottleRateActionBitsPerSecond $targetBTS2 -PolicyStore "ActiveStore" -IPPortMatchCondition $targetPort2
            $global:isOnTCP2 = $false
            $btnTCP1.BackColor = '#00FF00'
        }
        else
        {
           Remove-NetQosPolicy -Name "Destiny-2-3074-2TCP2" -PolicyStore "ActiveStore" -Confirm:$false 
            $global:isOnTCP2 = $true
            $btnTCP1.BackColor = '#2B2B32'
        }
    }
}

function startLimitTCP3{
    if($tBoxTCP3Port.Text -ne '' -and $tBoxTCPBts3.Text -ne '' -and $tBoxTCP4Port.Text -ne ''){
        if($global:isOnTCP3)
        {
            $targetPort3 = $tBoxTCPPort4.Text
            $endTargetPort = $tBoxTCPPort3.Text
            $targetBTS3 = $tBoxTCPBts3.Text
            
            New-NetQosPolicy -Name "Destiny-2-3074-2TCP3" -IPProtocol TCP  -AppPathNameMatchCondition "destiny2.exe" -ThrottleRateActionBitsPerSecond $targetBTS3 -PolicyStore "ActiveStore" -IPDstPortStartMatchCondition $targetPort3 -IPDstPortEndMatchCondition $endTargetPort
            $global:isOnTCP3 = $false
            $btnTCP3.BackColor = '#00FF00'
        }
        else
        {
            Remove-NetQosPolicy -Name "Destiny-2-3074-2TCP3" -PolicyStore "ActiveStore" -Confirm:$false
            $global:isOnTCP3 = $true
            $btnTCP3.BackColor = '#2B2B32'
        }
    }
}

function Get-RemotePorts{
    $array = New-Object System.Collections.ArrayList
    $Script:remInfo = Get-NetworkStatistics | Where-Object -Property ProcessName -EQ destiny2 | Select Protocol,LocalAddress,LocalPort,RemoteAddress,RemotePort,State| sort -Property Name
    $array.AddRange($remInfo)
    $dataGridRP.DataSource = $array
    $limit3074Form.Refresh()
}

$OnLoadForm_UpdateGrid=
{
    Get-RemotePorts
}

$limit3074Form.add_load($OnLoadForm_UpdateGrid)

$btnUDP.Add_Click({startLimitUDP})
$btnTCP1.Add_Click({startLimitTCP1})
$btnTCP2.Add_Click({startLimitTCP2})
$btnTCP3.Add_Click({startLimitTCP3})
$btnRP.Add_Click({Get-RemotePorts})

$limit3074Form.ShowDialog()

$limit3074Form.Dispose()