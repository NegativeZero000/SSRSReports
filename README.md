# SSRSReports

This is a PowerShell module for interacting with Microsoft SQL Server Reporting Services. Primarily this is used to query the server for information but can be used to upload completed reports and change report datasources. 

I personally use this for day to day report uploads. 

## Usage Examples

#### Connect to the service

At its core all the functions reference the SSRS Serivce Connection. The function `Connect-SSRSService` initiats that connection. Once saved, it can be passed to other functions seemlessly without having to reconnect all of the time. 

    $credentials = [pscredential]::new($username,$password)
    $SSRSservice = Connect-SSRSService "http://7fssrsreports/ReportServer/ReportService2005.asmx" -Credential $credentials
    $PSDefaultParameterValues.Add("*SSRS*:ReportService",$SSRSservice)

The last line makes it so you do not have to constantly specify the `-ReportService` parameter on all the function calls. It is assumed now.

#### Find all reports in a folder

Return all of the report entities that are found in the folder /application/testing

    Find-SSRSEntities -EntityType Report -SearchPath "/application/testing" 
    
#### Export selected reports to a folder
    
    Find-SSRSEntities -EntityType Report -SearchPath "/application/testing" | Export-SSRSReport -DestinationPath C:\Temp\test

## Support

Current list of supported and tested SSRS versions with this module. 

Platfrom | Result |
---- | ---- | 
SSRS 2208r2  | Passed | 
