$InterfaceServer = "SYDWOIFU01"
$InterfacePort = "5041"
$tcpConnection = New-Object System.Net.Sockets.TcpClient($InterfaceServer,$InterfacePort)
$tcpStream = $tcpConnection.GetStream()
$reader = New-Object System.IO.StreamReader($tcpStream)
$writer = New-Object System.IO.StreamWriter($tcpStream)
$writer.AutoFlush = $true

$buffer = new-object System.Byte[] 1024
$encoding = new-object System.Text.AsciiEncoding 
# Record start <STX> Record end <ETX>
#[string]$STX="<STX>"
#[string]$ETX="<ETX>"
# OR
[char]$STX = [char]2
[char]$eTX = [char]3
#OR
#[byte]$STX = 0x02
#[byte]$ETX = 0x03

if ($tcpConnection.Connected) #Else log or write could not connect
{
$Testcom = '<LinkStart Date="070818" Time="190649" />'  
$command = $STX + '<LinkDescription Date="280818" Time="121655" VerNum="1.0" />' + $ETX
$comArray = $command.ToCharArray()
foreach ($Element in $comArray)
    {
        $tempCom = $tempcom + " " + [system.string]::format("{0:X}",[system.convert]::Touint32($element))
    }
    #$CByte = [system.Text.Encoding]::UTF8
    #$CommandByte = $CByte.getbytes($command)
    #$CommandByte = $CByte.getbytes($TempCom)
   # $writer.WriteLine($commandByte) | Out-Null
    $writer.WriteLine($Tempcom) | Out-Null
    $writer.Flush()
    start-sleep -Milliseconds 900 # can we lop until stream is avail with a 5 second timeout
    while ($tcpStream.DataAvailable)
    {
        $rawresponse = $reader.Read($buffer, 0, 1024)
        $response = $encoding.GetString($buffer, 0, $rawresponse)   
    } 
    if ($response -contains "LinkAlive Date")
    {
        Write-Host "Link Alive received " + $response
        $command = $STX +'<PostInquiry InquiryInformation="M" MaximumReturnedMatches="16" SequenceNumber="1235" RequestType="8" PaymentMethod="16" Date="070905" Time="194121" RevenueCenter="1" WaiterId="Waiter1" WorkstationId="POS1" />' + $ETX
        $writer.WriteLine($command) | Out-Null
        start-sleep -Milliseconds 3000
        while ($tcpStream.DataAvailable)
        {
            $rawresponse = $reader.Read($buffer, 0, 1024)
            $response = $encoding.GetString($buffer, 0, $rawresponse)   
        } 
        Write-Host "Response received " + $response
    }
   
}

$reader.Close()
$writer.Close()
$tcpConnection.Close()