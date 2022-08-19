
#This source one may/may not cover most items: Advisory, Industry News, Partner News, etc.
$cyware = 'https://cyware.com/allnews/feed'
#News
$bleepingcomputer = 'https://www.bleepingcomputer.com/feed/'
$krebsonsecurity = 'https://krebsonsecurity.com/feed/'
$zdnet = 'https://www.zdnet.com/topic/security/rss.xml'
$darkreading = 'https://www.darkreading.com/rss_simple.asp'
$TheHackersNews = 'https://feeds.feedburner.com/TheHackersNews'
$securityweek = 'http://feeds.feedburner.com/securityweek?format=xml'
$threatpost = 'https://threatpost.com/feed/'
$grahamcluley = 'https://grahamcluley.com/feed/'
$trendmicro = 'http://feeds.trendmicro.com/TrendMicroResearch'
$cso="https://www.csoonline.com/index.rss"
$cybersecuritynews= 'https://cybersecuritynews.com/feed/'
$citrix= 'https://support.citrix.com/feed/products/all/securitybulletins.rss'
$microsoft= 'https://msrc-blog.microsoft.com/feed/'
$tenable= 'https://www.tenable.com/blog/cyber-exposure-alerts/feed'
$ibm= 'https://www.ibm.com/blogs/psirt/feed/atom/'
#Vendor and other Blogs/Analysis
$cisco = 'https://blogs.cisco.com/security/feed'
$amazon = 'https://aws.amazon.com/blogs/security/feed/'
$qualys = 'https://blog.qualys.com/feed'
$avast = 'https://blog.avast.com/rss.xml'
$tripwire = 'https://www.tripwire.com/state-of-security/feed/'
$sentinelone = 'https://www.sentinelone.com/feed/'
$securityintelligence = 'https://securityintelligence.com/feed/'
$trailofbits = 'https://blog.trailofbits.com/feed/'
$google= 'http://feeds.feedburner.com/GoogleOnlineSecurityBlog'
$Trendmicro= 'http://feeds.trendmicro.com/TrendMicroResearch'



#CERTs around the World
$us = 'https://us-cert.cisa.gov/ncas/current-activity.xml'
$canada1= 'https://cyber.gc.ca/webservice/en/rss/alerts'
$canada2= 'https://cyber.gc.ca/webservice/en/rss/news'
$belgium='https://cert.be/en/rss'
$hongkong='https://www.hkcert.org/getrss/security-news'
$carnegiemellon='https://www.kb.cert.org/vuls/atomfeed/'
$austrailia= 'https://www.auscert.org.au/rss/bulletins/?fmt=xml'

#Other Languages
$france='https://www.cert.ssi.gouv.fr/feed/'
$germany='https://cert.at/cert-at.de.warnings.rss_2.0.xml'
$germany2= 'https://www.cert-bund.de/feed/advisoryshort.rdf'
$finland='https://www.kyberturvallisuuskeskus.fi/feed/rss/fi'
$netherlands='https://feeds.ncsc.nl/nieuws.rss'

$CERTUrls= $us,$canada1,$canada2,$belgium,$hongkong,$carnegiemellon,$austrailia,
$france,$germany,$germany2,$finland,$netherlands

$TotalArticles= 0
$Today = Get-Date -format "MM-dd-yyyy"
$location= "C:\Users\$EmpID\Documents\ThreatBytes\$Today ThreatBytes.xlsx"

#Creating Excel File
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$workbook.Worksheets.Add()


$Data0= $workbook.Worksheets.Item(1)
$Data0.Range("A1:Q1").WrapText = "True"

# Worksheet Tab 1
$Data0.Name = 'CERTs'
$Sourcecount=2
#Mixed sources 
$Data0.Cells.Item(1,1) = "CERTs around the world"
$Data0.Cells.Item(1,1).Font.Bold = "True"


ForEach($CERTurl in $CERTUrls)
{
        #placeholder file to download and read from
        $filepath="C:\Users\$EmpID\Documents\ThreatBytes\CERTfeeds.xml"
        Invoke-RestMethod -Uri $CERTurl -OutFile $filepath
        [xml]$Content = Get-Content $filepath
        $Feed = $Content.rss.channel
        if ($CERTurl -eq $france){
         $Data0.Cells.Item($Sourcecount,1) = 'In French'
         $Sourcecount++
        }
        
        if ($CERTurl -eq $germany){
         $Data0.Cells.Item($Sourcecount,1) = 'In German'
        $Sourcecount++
        }
        
        if ($CERTurl -eq $netherlands){
         $Data0.Cells.Item($Sourcecount,1) = 'In Dutch'
        $Sourcecount++
        }
        
        if ($CERTurl -eq $finland){
         $Data0.Cells.Item($Sourcecount,1) = 'In Finnish'
        $Sourcecount++
        }
        $Data0.Cells.Item($Sourcecount,1) = $CERTurl
        Write-Output $CERTurl
       
        $Time = Get-Date -UFormat "%d %b %Y %T"
        $EndDate= $Time
        $msgcount=2

        ForEach ($msg in $Feed.Item)
        {
            $StartDate=[datetime]$msg.pubDate.Substring(0,25)
            $diff= NEW-TIMESPAN –Start $StartDate –End $EndDate
            #Getting Articles from last 24 hours
            if ($diff.Days -le 1){
                #Adding Title
                $Data0.Cells.Item($Sourcecount,$msgcount) = $msg.title
                #Adding Hyperlink
                $Data0.Hyperlinks.Add(
                $Data0.Cells.Item($Sourcecount,$msgcount),
                $msg.link,"",$msg.pubDate,$Data0.Cells.Item($Sourcecount,$msgcount).Text) | Out-Null
                $msgcount++
            }
        } #EndForEach
        $Sourcecount++
}









  # Format, save and quit excel
  $usedRange0 = $Data0.UsedRange                                                                                              
  $usedRange0.EntireColumn.AutoFit() | Out-Null

$Data= $workbook.Worksheets.Item(2)
$Data.Range("A1:Q1").WrapText = "True"

#2nd Worksheet Tab
$Data.Name = 'Mixed Sources'
$Sourcecount=2
#Mixed sources 
$Data.Cells.Item(1,1) = "Sources"
$Data.Cells.Item(1,1).Font.Bold = "True"

$newsUrls = $Cyware,$Bleepingcomputer,$Krebsonsecurity,$Zdnet,$Darkreading,$TheHackersNews,$Securityweek,$Threatpost,$Grahamcluley,$cybersecuritynews,$citrix,$microsoft,$cso,$tenable, $ibm,
$cisco,$amazon,$qualys,$avast,$tripwire,$sentinelone,$securityintelligence,$trailofbits,$google,$Trendmicro

ForEach($url in $newsUrls)
{
        #placeholder file to download and read from
        $filepath="C:\Users\$EmpID\Documents\ThreatBytes\feeds.xml"
        Invoke-RestMethod -Uri $url -OutFile $filepath
        [xml]$Content = Get-Content $filepath
        $Feed = $Content.rss.channel
        $Data.Cells.Item($Sourcecount,1) = $url
        Write-Output $url
       
        $Time = Get-Date -UFormat "%d %b %Y %T"
        $EndDate= $Time
        $msgcount=2

        ForEach ($msg in $Feed.Item)
        {
            $StartDate=[datetime]$msg.pubDate.Substring(0,25)
            $diff= NEW-TIMESPAN –Start $StartDate –End $EndDate
            #Getting Articles from last 24 hours
            if ($diff.Days -le 1){
                #Adding Title
                $Data.Cells.Item($Sourcecount,$msgcount) = $msg.title
                #Adding Hyperlink
                $Data.Hyperlinks.Add(
                $Data.Cells.Item($Sourcecount,$msgcount),
                $msg.link,"",$msg.pubDate,$Data.Cells.Item($Sourcecount,$msgcount).Text) | Out-Null
                $TotalArticles++
                $msgcount++
            }
        } #EndForEach
        $Sourcecount++
}
  Write-Output 'Total Number of Articles Today: '$TotalArticles
  # Format, save and quit excel
  $usedRange = $Data.UsedRange                                                                                              
  $usedRange.EntireColumn.AutoFit() | Out-Null

#Latest Vulnerabilities CVS 7 and above
$early1url = 'https://nvd.nist.gov/feeds/xml/cve/misc/nvd-rss.xml'
#3rd Worksheet Tab
$Data2= $workbook.Worksheets.Item(3)
$Data2.Name = 'Latest CVEs Today'

        $filepath="C:\Users\$EmpID\Documents\ThreatBytes\latestcve.xml"
        Invoke-RestMethod -Uri $early1url -OutFile $filepath
        [xml]$Content = Get-Content $filepath
        $Feed = $Content.rdf.item
        $msgcount=1
        $Sourcecount=1

        $Data2.Cells.Item($Sourcecount,$msgcount) = 'CVSS Score'
        $msgcount++
        $Data2.Cells.Item($Sourcecount,$msgcount) = 'Source'
        $msgcount++
        $Data2.Cells.Item($Sourcecount,$msgcount) = 'CVE Name'
        $msgcount++
        $Data2.Cells.Item($Sourcecount,$msgcount) = 'Description'
        $msgcount++
        $Sourcecount++
        $msgcount= 1
        ForEach ($msg in $Feed)
        {
            $StartDate=[datetime]$msg.date.Substring(0,20)
            $diff= NEW-TIMESPAN –Start $StartDate –End $EndDate

            if (($diff.Days -le 1) -and ($msg.description -notlike '*DO NOT USE THIS CANDIDATE NUMBER*')){
                $res = invoke-webrequest $msg.link -ErrorAction Stop
                $val = $res.parsedhtml.getelementsbytagname('a') |Where-Object {$_.id -like "Cvss3NistCalculatorAnchor"}
                    $Data2.Cells.Item($Sourcecount,$msgcount) = $val.innerText
                    $msgcount++
                    $val3 = $res.parsedhtml.getelementsbytagname('span') |Where-Object {$_.outerHTML -like "*vuln-current-description-source*"}
                    $Data2.Cells.Item($Sourcecount,$msgcount) = $val3.innerText
                    $msgcount++
                    #Adding Title
                    $Data2.Cells.Item($Sourcecount,$msgcount) = $msg.title
                    #Adding Hyperlink
                    $Data2.Hyperlinks.Add(
                    $Data2.Cells.Item($Sourcecount,$msgcount),
                    $msg.link,"",$Data2.Cells.Item($Sourcecount,$msgcount).Text) | Out-Null
                    $msgcount++
                    $Data2.Cells.Item($Sourcecount,$msgcount) = $msg.description
                    $msgcount++

                    if ($msgcount -ge 3){
                        $msgcount= 1
                        $Sourcecount++
                    }
               
            }
        } #EndForEach
  $usedRange2 = $Data2.UsedRange                                                                                              
  $usedRange2.EntireColumn.AutoFit() | Out-Null

$analysed1url = 'https://nvd.nist.gov/feeds/xml/cve/misc/nvd-rss-analyzed.xml'

#4th Worksheet Tab
$Data3= $workbook.Worksheets.Item(4)
$Data3.Name = 'Analyzed CVEs this week'

        $filepath="C:\Users\$EmpID\Documents\ThreatBytes\analyzedcve.xml"
        Invoke-RestMethod -Uri $analysed1url -OutFile $filepath
        [xml]$Content = Get-Content $filepath
        $Feed = $Content.rdf.item
        $msgcount=1
        $Sourcecount=1
        $Data3.Cells.Item($Sourcecount,$msgcount) = 'CVSS Score'
        $msgcount++
        $Data3.Cells.Item($Sourcecount,$msgcount) = 'Last Updated'
        $msgcount++
        $Data3.Cells.Item($Sourcecount,$msgcount) = 'Source'
        $msgcount++
        $Data3.Cells.Item($Sourcecount,$msgcount) = 'CVE Name'
        $msgcount++
        $Data3.Cells.Item($Sourcecount,$msgcount) = 'Date'
        $msgcount++
        $Data3.Cells.Item($Sourcecount,$msgcount) = 'Description'
        $msgcount++
        $Sourcecount++
        $msgcount= 1
        ForEach ($msg in $Feed)
        {
            $StartDate=[datetime]$msg.date.Substring(0,20)
            $diff= NEW-TIMESPAN –Start $StartDate –End $EndDate
            #Getting Articles from last 7 days
            if ($diff.Days -le 7){
                $res = invoke-webrequest $msg.link -ErrorAction Stop
                $val = $res.parsedhtml.getelementsbytagname('a') |Where-Object {$_.id -like "Cvss3NistCalculatorAnchor"}
                $score = [double]$val.innerText.Substring(0,3)
                if ($score -ge 7.0) {
                    $Data3.Cells.Item($Sourcecount,$msgcount) = $score
                    $msgcount++
                    $val2 = $res.parsedhtml.getelementsbytagname('span') |Where-Object {$_.outerHTML -like "*vuln-last-modified-on*"}
                    $Data3.Cells.Item($Sourcecount,$msgcount) = $val2.innerText
                    $msgcount++
                    $val3 = $res.parsedhtml.getelementsbytagname('span') |Where-Object {$_.outerHTML -like "*vuln-current-description-source*"}
                    $Data3.Cells.Item($Sourcecount,$msgcount) = $val3.innerText
                    $msgcount++
                    #Adding Title
                    $Data3.Cells.Item($Sourcecount,$msgcount) = $msg.title
                    #Adding Hyperlink
                    $Data3.Hyperlinks.Add(
                    $Data3.Cells.Item($Sourcecount,$msgcount),
                    $msg.link,"",$Data3.Cells.Item($Sourcecount,$msgcount).Text) | Out-Null
                    $msgcount++
                    $Data3.Cells.Item($Sourcecount,$msgcount) = [datetime]$msg.date
                    $msgcount++
                    $Data3.Cells.Item($Sourcecount,$msgcount) = $msg.description
                    $msgcount++
                    if ($msgcount -ge 4){
                        $msgcount= 1
                        $Sourcecount++
                    }
                }
            }
        } #EndForEach

  $usedRange3 = $Data3.UsedRange                                                                                              
  $usedRange3.EntireColumn.AutoFit() | Out-Null

  #5th Worksheet Tab
$Data4= $workbook.Worksheets.Item(5)
$Data4.Name = 'Manual Sources'
$Data4.Cells.Item(1,1) = "Manual Sources to Check"
$Data4.Cells.Item(1,1).Font.Bold = "True"
  
  $Data4.Cells.Item(2,1) = "DataBreachToday"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(2,1),
  "https://www.databreachtoday.com/latest-news","",$Data4.Cells.Item(2,1).Text) | Out-Null

  $Data4.Cells.Item(2,3) = "Good Articles from a Researcher-Paganini"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(2,3),
  "https://securityaffairs.co/wordpress/","",$Data4.Cells.Item(2,3).Text) | Out-Null

  

  $Data4.Cells.Item(3,1) = "Cisco Advisory"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(3,1),
  "https://tools.cisco.com/security/center/publicationListing.x?product=Cisco&sort=-day_sir#~Vulnerabilities","",$Data4.Cells.Item(3,1).Text) | Out-Null
  $Data4.Cells.Item(3,2) = "VMware Advisory"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(3,2),
  "https://www.vmware.com/security/advisories.html","",$Data4.Cells.Item(3,2).Text) | Out-Null
  $Data4.Cells.Item(3,3) = "Mccafe Advisory"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(3,3),
  "https://www.mcafee.com/enterprise/en-us/threat-center/product-security-bulletins.html","",$Data4.Cells.Item(3,3).Text) | Out-Null
  
  $Data4.Cells.Item(3,4) = "Indian CERT"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(3,4),
  "https://cert-in.org.in/","",$Data4.Cells.Item(3,4).Text) | Out-Null  
  
  $Data4.Cells.Item(2,2) = "Source for Strictly Oil/Gas News"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(2,2),
  "https://www.hazardexonthenet.net/industry-news.aspx?sid=96&n=Oil-Gas","",$Data4.Cells.Item(2,2).Text) | Out-Null

  $Data4.Cells.Item(4,1) = "Partner Related News"

  $Data4.Cells.Item(5,1) = "Barracuda"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,1),
  "https://www.barracuda.com/news/?type=1&search= ","",$Data4.Cells.Item(5,1).Text) | Out-Null  

  $Data4.Cells.Item(5,2) = "Azure"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,2),
  "https://www.microsoft.com/azure/partners/news","",$Data4.Cells.Item(5,2).Text) | Out-Null  

$Data4.Cells.Item(5,3) = "FireEye"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,3),
  "https://www.fireeye.com/blog.html/","",$Data4.Cells.Item(5,3).Text) | Out-Null  

$Data4.Cells.Item(5,4) = "Nessus/Tenable"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,4),
  "https://docs.tenable.com/releasenotes/Content/nessus/nessus8150.htm","",$Data4.Cells.Item(5,4).Text) | Out-Null  

$Data4.Cells.Item(5,5) = "Cisco ASA 5500"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,5),
  "https://www.cisco.com/c/en/us/support/security/asa-5500-series-next-generation-firewalls/series.html#~tab-documents","",$Data4.Cells.Item(5,5).Text) | Out-Null  

$Data4.Cells.Item(6,5) = "Cisco Firepower"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,5),
  "https://www.cisco.com/c/en/us/support/security/defense-center/series.html#~tab-documents","",$Data4.Cells.Item(6,5).Text) | Out-Null
    
$Data4.Cells.Item(5,6) = "QRadar"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,6),
  "https://www.ibm.com/community/qradar/home/software/","",$Data4.Cells.Item(5,6).Text) | Out-Null  

$Data4.Cells.Item(5,7) = "MalwareBytes"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,7),
  "https://support.malwarebytes.com/hc/en-us/categories/360002458014-Malwarebytes-for-Windows","",$Data4.Cells.Item(5,7).Text) | Out-Null  

$Data4.Cells.Item(5,8) = "Mcafee Endpoint Security"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(5,8),
  "https://docs.mcafee.com/bundle?labelkey=prod-endpoint-security-v10-7-x&labelkey=cat-release-notes","",$Data4.Cells.Item(5,8).Text) | Out-Null  

$Data4.Cells.Item(6,8) = "Mcafee MVision Cloud Edge"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,8),
  "https://docs.mcafee.com/bundle/mvision-unified-cloud-edge-release-notes/page/GUID-E0E666D4-2947-40D1-A998-1D118E12AD05.html","",$Data4.Cells.Item(6,8).Text) | Out-Null  

$Data4.Cells.Item(7,8) = "Mcafee Press Releases"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(7,8),
  "https://www.mcafee.com/enterprise/en-us/about/newsroom/press-releases.html","",$Data4.Cells.Item(7,8).Text) | Out-Null
    
$Data4.Cells.Item(7,5) = "Cisco WSA"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(7,5),
  "https://www.cisco.com/c/en/us/support/security/web-security-appliance/products-release-notes-list.html","",$Data4.Cells.Item(7,5).Text) | Out-Null  

$Data4.Cells.Item(6,1) = "Proofpoint"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,1),
  "https://help.proofpoint.com/Proofpoint_Essentials/Release_Notes/20210706","",$Data4.Cells.Item(6,1).Text) | Out-Null
    
$Data4.Cells.Item(6,2) = "Symantec Endpoint Protection"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,2),
  "https://knowledge.broadcom.com/external/article/154575/versions-system-requirements-release-dat.html","",$Data4.Cells.Item(6,2).Text) | Out-Null  

$Data4.Cells.Item(6,3) = "Websense/Forcepoint"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,3),
  "https://www.google.com/search?q=siteurl:www.websense.com+websense+forcepoint+release+notes","",$Data4.Cells.Item(6,3).Text) | Out-Null
    
$Data4.Cells.Item(6,4) = "Checkpoint"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,4),
  "https://supportcenter.checkpoint.com/supportcenter/portal?eventSubmit_doGoviewsolutiondetails=&solutionid=sk133174","",$Data4.Cells.Item(6,4).Text) | Out-Null  

$Data4.Cells.Item(8,5) = "Cisco AMP"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(8,5),
  "https://docs.amp.cisco.com/Release%20Notes.pdf","",$Data4.Cells.Item(8,5).Text) | Out-Null  

$Data4.Cells.Item(6,6) = "Nexpose"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,6),
  "https://docs.rapid7.com/release-notes/nexpose/","",$Data4.Cells.Item(6,6).Text) | Out-Null  

  
$Data4.Cells.Item(6,7) = "Flexera"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(6,7),
  "https://docs.flexera.com/?product=Software%20Vulnerability%20Research&version=Current","",$Data4.Cells.Item(6,7).Text) | Out-Null  
  
$Data4.Cells.Item(8,8) = "Mcafee Application Control"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(8,8),
  "https://docs.mcafee.com/bundle?labelkey=prod-application-control-change-control&labelkey=prod-application-change-control-v8-3-x&labelkey=prod-application-change-control-v6-4-x&labelkey=prod-application-change-control-v6-3-x&labelkey=prod-application-change-control&labelkey=prod-application-control-v8-3-x&labelkey=prod-application-control-v8-2-6&labelkey=prod-application-control-v8-2-1&labelkey=prod-application-control-v8-2-0&labelkey=prod-application-control-v8-1-1&labelkey=prod-application-control-v8-1-0&labelkey=prod-application-control-v8-1-x&labelkey=prod-application-control-v8-0-2&labelkey=prod-application-control-v8-0-1&labelkey=prod-application-control-v8-0-0&labelkey=prod-application-control-v8-0-x&labelkey=prod-application-control-v7-0-2&labelkey=prod-application-control-v7-0-1&labelkey=prod-application-control-v7-0-0&labelkey=prod-application-control-v7-0-x&labelkey=prod-application-control-v6-4-x&labelkey=prod-application-control-v6-3-x&labelkey=prod-application-control-v6-2-2&labelkey=prod-application-control-v6-2-1&labelkey=prod-application-control-v6-2-0&labelkey=prod-application-control-v6-2-x&labelkey=prod-application-control-v6-1-7&labelkey=prod-application-control-v6-1-4&labelkey=prod-application-control-v6-1-3&labelkey=prod-application-control-v6-1-2&labelkey=prod-application-control-v6-1-1&labelkey=prod-application-control-v6-1-0&labelkey=prod-application-control-v6-1-x&labelkey=prod-application-control-v5-1-x&labelkey=prod-application-control&labelkey=prod-change-control-v8-3-x&labelkey=prod-change-control-v8-2-6&labelkey=prod-change-control-v8-2-1&labelkey=prod-change-control-v8-2-0&labelkey=prod-change-control-v8-1-1&labelkey=prod-change-control-v8-1-0&labelkey=prod-change-control-v8-1-x&labelkey=prod-change-control-v8-0-2&labelkey=prod-change-control-v8-0-0&labelkey=prod-change-control-v8-0-1&labelkey=prod-change-control-v8-0-x&labelkey=prod-change-control-v7-0-2&labelkey=prod-change-control-v7-0-1&labelkey=prod-change-control-v7-0-0&labelkey=prod-change-control-v7-0-x&labelkey=prod-change-control-v6-4-x&labelkey=prod-change-control-v6-3-x&labelkey=prod-change-control-v6-2-2&labelkey=prod-change-control-v6-2-1&labelkey=prod-change-control-v6-2-0&labelkey=prod-change-control-v6-2-x&labelkey=prod-change-control-v6-1-7&labelkey=prod-change-control-v6-1-4&labelkey=prod-change-control-v6-1-3&labelkey=prod-change-control-v6-1-2&labelkey=prod-change-control-v6-1-1&labelkey=prod-change-control-v6-1-0&labelkey=prod-change-control-v6-1-x&labelkey=prod-change-control-v5-1-x&labelkey=prod-change-control
  ","",$Data4.Cells.Item(8,8).Text) | Out-Null  
  
$Data4.Cells.Item(7,1) = "Thycotic Secret Server"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(7,1),
  "https://docs.thycotic.com/ss/10.9.0/release-notes","",$Data4.Cells.Item(7,1).Text) | Out-Null  
  
$Data4.Cells.Item(9,5) = "Cisco AnyConnect"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(9,5),
  "https://www.cisco.com/c/en/us/support/security/anyconnect-secure-mobility-client-v4-x/model.html#~tab-documents","",$Data4.Cells.Item(9,5).Text) | Out-Null  

  $Data4.Cells.Item(7,2) = "Pulse Secure"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(7,2),
  "https://www-prev.pulsesecure.net/techpubs/pulse-connect-secure/pcs/9.1rx/9.1r11","",$Data4.Cells.Item(7,2).Text) | Out-Null  
  
  $Data4.Cells.Item(9,8) = "Mcafee DAM"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(9,8),
  "https://docs.mcafee.com/bundle?labelkey=prod-database-activity-monitoring&labelkey=prod-database-activity-monitoring-v5-2-4&labelkey=prod-database-activity-monitoring-v5-2-3&labelkey=prod-database-activity-monitoring-v5-2-2&labelkey=prod-database-activity-monitoring-v5-2-0&labelkey=prod-database-activity-monitoring-v5-2-x&labelkey=prod-database-activity-monitoring-v4-6-x&name_filter.field=title&name_filter.value=&rpp=20&sort.field=created&sort.value=dec
","",$Data4.Cells.Item(9,8).Text) | Out-Null
  
  $Data4.Cells.Item(7,3) = "OneLogin"
  $Data4.Hyperlinks.Add(
  $Data4.Cells.Item(7,3),
  "https://www.onelogin.com/blog/categories/whats-new","",$Data4.Cells.Item(7,3).Text) | Out-Null  

  $usedRange4= $Data4.UsedRange                                                                                              
  $usedRange4.EntireColumn.AutoFit() | Out-Null


$workbook.SaveAs($location)
$excel.Quit()   


$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable excel