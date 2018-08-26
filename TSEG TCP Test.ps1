$InterfaceServer = "SYDWOIFU01"
$InterfacePort = "5002"
$tcpConnection = New-Object System.Net.Sockets.TcpClient($InterfaceServer,$InterfacePort)
$tcpStream = $tcpConnection.GetStream()
$reader = New-Object System.IO.StreamReader($tcpStream)
$writer = New-Object System.IO.StreamWriter($tcpStream)
$writer.AutoFlush = $true

$buffer = new-object System.Byte[] 1024
$encoding = new-object System.Text.AsciiEncoding 
# Record start <STX> Record end <ETX>
# [char]$STX = [char]2
# [char]$eTX = [char]3
#OR
[byte]$STX = 0x02
[byte]$ETX = 0x03

if ($tcpConnection.Connected) #Else log or write could not connect
{
    $command = $STX + '<LinkDescription Date="070818" Time="190649" VerNum="1.0" />' + $ETX
    $writer.WriteLine($command) | Out-Null
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