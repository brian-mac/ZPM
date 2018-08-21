$FTPServer = "192.168.0.160"
$FTPPort = "32400"
$tcpConnection = New-Object System.Net.Sockets.TcpClient($FTPServer, $FTPPort)
$tcpStream = $tcpConnection.GetStream()
$reader = New-Object System.IO.StreamReader($tcpStream)
$writer = New-Object System.IO.StreamWriter($tcpStream)
$writer.AutoFlush = $true

$buffer = new-object System.Byte[] 1024
$encoding = new-object System.Text.AsciiEncoding 

if ($tcpConnection.Connected)
{
    $command = '<LinkDescription Date="070818" Time="190649" VerNum="1.0" />'
    $writer.WriteLine($command) | Out-Null
    start-sleep -Milliseconds 900
    while ($tcpStream.DataAvailable)
    {
        $rawresponse = $reader.Read($buffer, 0, 1024)
        $response = $encoding.GetString($buffer, 0, $rawresponse)   
    } 
    if ($response -contains "LinkAlive Date")
    {
        Write-Host "Link Alive received " + $response
        $command = '<PostInquiry InquiryInformation="M" MaximumReturnedMatches="16" SequenceNumber="1235" RequestType="8" PaymentMethod="16" Date="070905" Time="194121" RevenueCenter="1" WaiterId="Waiter1" WorkstationId="POS1" />'
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