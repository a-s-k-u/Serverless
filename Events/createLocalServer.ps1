#* FileName: createLocalServer.ps1
 #*=============================================================================
 #* Script Name: createLocalServer
 #* Created:     [27/11/2015]
 #* Author:      Arun Sree Kumar
 #* Email:       arun.kumar@allianz.com.au
 #* Email2:      arun.sreekumar@allianzcornhill.co.in
 #*
 #*=============================================================================
 
#*=============================================================================
 #* REVISION HISTORY
 #*=============================================================================
 #* Date: [27/11/2015]
 #* Description: Initial Version for Digitalization Campaign.
 #*
 #*=============================================================================
"Control reached powershell - setting up a local server/listener"
$listener = New-Object Net.HttpListener
$listener.Prefixes.Add("http://localhost:8081/")
"Starting the local server at port 8081"
$listener.Start()
"Local server started and server listening to port 8081"
$path=$args[0]
$timelineContentTemplate = '<div class="item" data-id="tDate" data-description="tName">
			<a class="image_rollover_bottom con_borderImage" data-description="ZOOM IN" href="images/light/timeline_content/tImage" rel="lightbox[timeline]">
			<img src="images/light/thumbnails/tThumbnailImage" alt=""/>
			</a>
			<h2>tDateMonth</h2>
			<span>tSummary</span>
			<div class="read_more" data-id="tDate">Read more</div>
		</div>
		<div class="item_open" data-id="tDate">
			tReadmore
		</div>';
"Opening the browser"
Start-Process -FilePath "http://localhost:8081/home/"
"Browser Opened with the url to local server"
While ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    #"Start Hello" |Out-File -append \\S83DCP02TJWDC2\data\users\I84446\Profile_data\Desktop\log.txt
    #$context.Request | Out-String | Out-File -append \\S83DCP02TJWDC2\data\users\I84446\Profile_data\Desktop\log.txt
    #"End Hello" |Out-File -append \\S83DCP02TJWDC2\data\users\I84446\Profile_data\Desktop\log.txt   
    #$d = $request.RawUrl.ToLower().StartsWith("/hello/")
    #Below is how you could implement the rest functionality.
    if ($request.RawUrl.ToLower().StartsWith("/readxl/")) 
    {
    "Raw URL starts wih readXL" |Out-File -append \\S83DCP02TJWDC2\data\users\I84446\Profile_data\Desktop\log.txt
     $page = Import-csv \\S83DCP02TJWDC2\data\users\I84446\Profile_data\Desktop\MeetingRoomBlocking.csv| ConvertTo-Json
     $page |Out-File -append \\S83DCP02TJWDC2\data\users\I84446\Profile_data\Desktop\log.txt
     invoke-item .
    }
    elseif ($request.RawUrl.ToLower().StartsWith("/measurexl/"))
    {
     # not implemented.   
    }
    else
    {
        "Opening event.xlsx file from path - " + $path 
        $strPath= $path + '\events.xlsx'
        $objExcel=New-Object -ComObject Excel.Application
        $objExcel.Visible=$false
        $WorkBook=$objExcel.Workbooks.Open($strPath)
        $worksheet = $workbook.sheets.item("Events")
        $intRowMax =  ($worksheet.UsedRange.Rows).count
        $Columnnumber = 1
        "Reading excel contents"
        $timelineContent = '';
        for($intRow = 2 ; $intRow -le $intRowMax ; $intRow++)
             {
                "Progress - reading " + ($intRow - 1 ) + " of the total " + ($intRowMax-1) + " records..."
                $tDate = $worksheet.cells.item($intRow,1).value2
                $tDate
                $tDateMonth = Get-Date $tDate -format "MMMM,d"
                $tDateMonth = $tDateMonth.ToUpper()
                $tName = $worksheet.cells.item($intRow,2).value2
                $tSummary = $worksheet.cells.item($intRow,3).value2
                $tDescription = $worksheet.cells.item($intRow,4).value2
                $tReadmore = $worksheet.cells.item($intRow,5).value2
                $tImage = $worksheet.cells.item($intRow,6).value2
                $tThumbnailImage = $worksheet.cells.item($intRow,7).value2
                $timelineContentItem = $timelineContentTemplate
                $timelineContentItem = $timelineContentItem.Replace('tDateMonth',$tDateMonth);
                $timelineContentItem = $timelineContentItem.Replace('tDate',$tDate);
                $timelineContentItem = $timelineContentItem.Replace('tName',$tName);
                $timelineContentItem = $timelineContentItem.Replace('tSummary',$tSummary);
                $timelineContentItem = $timelineContentItem.Replace('tDescription',$tDescription);
                $timelineContentItem = $timelineContentItem.Replace('tReadmore',$tReadmore);
                $timelineContentItem = $timelineContentItem.Replace('tThumbnailImage',$tThumbnailImage);
                $timelineContentItem = $timelineContentItem.Replace('tImage',$tImage);
                $timelineContent = $timelineContent + $timelineContentItem
              }
        $objexcel.quit()
        "Completed - " + ($intRowMax - 1) + " records from excel read"

    $page = Get-Content -Path ($path + 'light.html') -Raw
    $page = $page.Replace('_timelineContent_',$timelineContent)
    $page = $page.Replace('js/',$path + '/js/');
    $page = $page.Replace('css/',$path + '/css/');
    $page = $page.Replace('images/',$path + '/images/')   
    }
    $response = $context.Response
    $response.Headers.Add("Content-Type","text/html")
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($page)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer,0,$buffer.Length)
    $response.Close()
    "Response sent to browser"
}
$listener.Stop()
