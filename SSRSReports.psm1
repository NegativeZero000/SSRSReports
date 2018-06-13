function Split-SSRSPath{
     <#
    .SYNOPSIS
        Split-Path for SSRS paths.  
 
    .DESCRIPTION
        Functions the same as Split-Path but is more SSRS friendly. Split-Path returns backslashes.
        This will return forward slashes. 
 
    .PARAMETER Path
        Specifies the SSRS path to be split. If the path includes spaces, enclose it in quotation marks. You can also pipe a path to this cmdlet.

    .PARAMETER Leaf
        Indicates that this cmdlet returns only the last item or container in the path.

    .PARAMETER Parent
        Indicates that this cmdlet returns only the parent containers of the item or of the container specified by the path. The Parent parameter is the default split location parameter.

    .EXAMPLE
        Get the parent container/folder from an SSRS path. 

        Split-SSRSPath -Path "/Marketing/Sales" -Parent
    
    .EXAMPLE
        Get just the container name/folder name from an SSRS path.  

        Split-SSRSPath "/Marketing/Quarterly Sales Report" -Leaf
    #>
    [CmdletBinding(DefaultParameterSetName="Parent")] 
    param(
        [Parameter(Position=0,Mandatory,ValueFromPipeline)]
        [string]$Path,

        [Parameter(Mandatory=$false,ParameterSetName="Leaf")]
        [switch]$Leaf=$False,

        [Parameter(Mandatory=$false,ParameterSetName="Parent")]
        [switch]$Parent=$False
    )
    return (Split-Path @PSBoundParameters).replace("\","/")
}

function Connect-SSRSService{
    <#
    .SYNOPSIS
        Creates a connection to a SQL Reporting Server using a Web Service Proxy
 
    .DESCRIPTION
        Creates a connection to a SQL Reporting Server using a Web Service Proxy. An object will be returned that 
        represents the connection to be used in other operations
 
    .PARAMETER URL
        Name of report wherein the rdl file will be save as in Report Server.
        If this is not specified it will get the name from the file (rdl) exluding the file extension.

    .PARAMETER Credentials
        Credentials used to connect to the SQL Reporting Server if -UseDefaultCredentials is not specified

    .PARAMETER UseDefaultCredentials
        Specifies that the credentials of the running user are to be passed for authentication

    .EXAMPLE
        Connect-SSRSService -Url "http://[ServerName]/ReportServer/ReportService2005.asmx?WSDL" -Credentials $credentials
 
    .EXAMPLE
        Connect-SSRSService -Url "http://[ServerName]/ReportServer/ReportService2005.asmx?WSDL" -UseDefaultCredentials
    #>

    [CmdletBinding(DefaultParameterSetName="DefaultCredentials")]    
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("URL")]
        [string]$URI,
 
        [Parameter(Position=1,ParameterSetName=’Credentials’)]
        [PSCredential]$Credential,
  
        [Parameter(Position=1,ParameterSetName=’DefaultCredentials’)]
        [switch]$UseDefaultCredential
    )
    $parameters = @{URI = $URI}

    # Prepare parameters to be splatted
    if($psCmdlet.ParameterSetName -eq "Credentials"){   
        $parameters.Credential = $Credential
    } else {
        $parameters.UseDefaultCredential = $true
    }

    if($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
        $parameters.GetEnumerator() | ForEach-Object{
            Write-Verbose "Parameter ($($_.Key)): '$($_.Value)'"
        }
    }

    # Create the proxy connection
    try{
        return New-WebServiceProxy @parameters -ErrorAction Stop
    } catch {
        Write-Error ("Unable to connect to proxy: `r`n $_")
    }
}

function Find-SSRSEntities{
    <#
    .SYNOPSIS
        Searches environment for SSRS entities.
 
    .DESCRIPTION
        Searches environment for SSRS entities. It will return as many items as it can.
        Search criteria can be used to reduced the set of found items.
 
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy

    .PARAMETER SearchPath
        Starting path of the SQL Server Reporting Services entity i.e. folder, data source etc. to search for
        By default this is root of the SSRS Service

    .PARAMETER EntityType
        Switch used to define EntityPath as a path to a folder, report or other SSRS Object.
        Current supported values are 

        All
        Folder
        Report
        Resource
        LinkedReport
        DataSource
        Model

        This parameter is optional and All is assumed by default

    .PARAMETER Match
        String used to reduce the found set of items. This will match against the name of the entity.
    
    .PARAMETER Partial
        Used in conjuction with Match. It will determine if the string in Match represents the whole 
        name of the entity or if partial matches be allowed. 

    .EXAMPLE
        Find-SSRSEntities -ReportService $ssrsService -EntityType Folder
    
    .EXAMPLE
        Find-SSRSEntities -ReportService $ssrsService -SearchPath "/Marketing/Sales" -EntityType Folder -Match "Test"

    .EXAMPLE
        Find-SSRSEntities -ReportService $ssrsService -SearchPath "/Marketing/Sales" -EntityType Folder -Match "Test" -Partial

    #>
    [CmdletBinding()] 
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=1)]
        [Alias("Path")]
        [string]$SearchPath="/",

        [Parameter(Position=2)]
        [ValidateSet("All", "Folder", "Report", "Resource", "LinkedReport", "DataSource", "Model")]
        [Alias("Type")]
        [String]$EntityType = "All",

        [Parameter(Position=3)]
        [String]$Match,

        [Parameter(Position=4)]
        [Switch]$Partial=$false
    )
    # Get all of the catalog items that match the criteria passed
    # https://msdn.microsoft.com/en-us/library/reportservice2005.reportingservice2005.listchildren.aspx
    $recursive = $true
    $catalogItems = $ReportService.ListChildren($SearchPath,$recursive)
    Write-Verbose "$($catalogItems.Count) item(s) located in the root path $SearchPath"

    # Limit the results to the catalog types requested
    if($EntityType -ne "All"){$catalogItems = $catalogItems | Where-Object{$_.Type -eq $EntityType}}
    Write-Verbose "$($catalogItems.Count) item(s) found matching the type $EntityType"

    # Set the match string based on parameters
    if(-not $Partial.isPresent -and $Match){$Match = "^$Match$"}
    Write-Verbose "Returning all items matching: '$Match'"

    # If the regex is an empty string all object will be returned.
    return $catalogItems | Where-Object{$_.Name -match $Match}
}

function Test-SSRSPath{
    <#
    .SYNOPSIS
        Verifies if a given entity path is a valid report entity path. 
    .DESCRIPTION
        Verifies if a given entity path exists on the target Reporting Services proxy that is provided via 
        the ReportPath. This relies heavily on the Find-SSRSEntities cmdlet and is a wrapper for it
        Returns true of false unless -PassThru is used
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy
    .PARAMETER EntityPath
        Destination path of the SQL Server Reporting Services entity i.e. folder, data source etc.
    .PARAMETER EntityType
        Switch used to define EntityPath as a path to a folder, report or other SSRS Object.
        Current supported values are 

        Folder
        Report
        Resource
        LinkedReport
        DataSource
        Model

        Report is assumed by default
    .PARAMETER PassThru
        Tells the function to return the item found down the pipeline. By default a boolean result is returned.
    .EXAMPLE
        Test-SSRSPath -ReportService $ssrsService -EntityPath "/Marketing/Sales" -EntityType Folder
    .EXAMPLE
        Test-SSRSPath -ReportService $ssrsService -EntityPath "/Marketing/Quarterly Sales Report"
    .EXAMPLE
        Test-SSRSPath -ReportService $ssrsService -Path "/Marketing/Quarterly Sales Report" -EntityType Report
    #>
    [CmdletBinding()] 
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=1,Mandatory=$true)]
        [Alias("Path")]
        [string]$EntityPath,

        [Parameter(Position=2)]
        [ValidateSet("Folder", "Report", "Resource", "LinkedReport", "DataSource", "Model")]
        [Alias("Type")]
        [String]$EntityType = "Report",

        [switch]$PassThru=$false
    )
    # Split the path into its folder and entity parts
    $SearchPath = Split-SSRSPath $EntityPath -Parent
    $EntityName = Split-Path $EntityPath -Leaf

    $findSSRSEntriesParameters = @{
		ReportService = $ReportService 
		SearchPath    = $SearchPath 
		EntityType    = $EntityType 
		Match         = $EntityName 
		Partial       = $false
    }

    $result = Find-SSRSEntities @findSSRSEntriesParameters
    if($PassThru.IsPresent){
        $result
    } else {
        $result -as [Boolean]
    }
}

function Move-SSRSEntity{
    <#
    .SYNOPSIS
        Moves an existing entity from one location to another
 
    .DESCRIPTION
        Validates an entity passed and a target. Attempts to move it to the new location.
        Full destination path must be specified and can be used to rename the entity
 
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy 

    .PARAMETER EntitySourcePath
        Path to the entity that will be moved in the Reporting Service

    .PARAMETER EntityDestinationPath
        Destination path to the new location in the Reporting Service

    .PARAMETER EntityType
        Validates that the EntitySourcePath is a specific entity type.
        Current supported values are 

        All
        Folder
        Report
        Resource
        LinkedReport
        DataSource
        Model

        This parameter is optional and All is assumed by default meaning that it does not care.
 
    .EXAMPLE
        Move-SSRSEntity -ReportService $ssrsService -EntitySourcePath "/folder/report 01" -EntityDestinationPath "/new folder/report 01"

        Move a report from /folder to /new folder

    .EXAMPLE
        Move-SSRSEntity -ReportService $ssrsService -EntitySourcePath "/folder/" -EntityDestinationPath "/folder 01/folder 02" -EntityType Folder

        Move a folder from / to /folder 01 and rename it to /folder 02
    #>
    [CmdletBinding()] 
    param(
        [Parameter(Mandatory=$true)]
        [Alias("Proxy","SSRSService")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=0,Mandatory,ValueFromPipelineByPropertyName)]
        [Alias("SourcePath","Path")]
        [string]$EntitySourcePath,

        [Parameter(Position=1,Mandatory)]
        [Alias("DestinationPath")]
        [string]$EntityDestinationPath,

        [Parameter(Position=2)]
        [ValidateSet("Folder", "Report", "Resource", "LinkedReport", "DataSource", "Model")]
        [Alias("Type")]
        [String]$EntityType = "Report"
    )

    # Validate the item to move actually exists
    if(-not (Test-SSRSPath -ReportService $ReportService -EntityPath $EntitySourcePath -EntityType $EntityType)){
        # Check for an existing folder and create if not found
        Write-Error "The entity at path '$EntitySourcePath' does not exist or is not valid."
        return
    }
 
    try
    {
        Write-Verbose "Moving '$EntitySourcePath' to '$EntityDestinationPath'"
 
        # Call proxy service to move item
        # Parameters ItemPath, Target
        # https://msdn.microsoft.com/en-us/library/reportservice2010.reportingservice2010.moveitem.aspx
        Write-Verbose "Source Path      : $EntitySourcePath"
        Write-Verbose "Destination Path : $ReportFolder"
        $7fssservice.MoveItem($EntitySourcePath,$EntityDestinationPath)
    } catch [IO.IOException]{
        Write-Error ("Error while reading report`r`nMessage: '{0}'" -f $SourceFile, $_.Exception.Message)
    } catch [Web.Services.Protocols.SoapException]{
        $errorText = $_.Exception.Detail.InnerText
        if($errorText -match "rsItemAlreadyExists"){
            Write-Error ("Error while uploading report file`r`nMessage:`r'{0}'" -f "Report already exists. Delete or use -Force" )
        } else {
            Write-Error ("Error while uploading report file`r`nMessage:`r`n'{0}'" -f $_.Exception.Detail.InnerText)
        }
    }
 
}

function Publish-SSRSReport{
    <#
    .SYNOPSIS
        Publishes an RDL file to SQL Reporting Server using a Web Service Proxy
 
    .DESCRIPTION
        Publishes an RDL file to SQL Reporting Server using a Web Service Proxy
 
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy 

    .PARAMETER SourceFile
        Path to the RDL file that will be uploaded to the Reporting Service

    .PARAMETER ReportFolder
        Destination path of the RDL -SourceFile on the target ReportService

    .PARAMETER ReportName
        Name that will be assigned to the report once it is uploaded. 
        If ommited the file name will be used
 
    .PARAMETER Force
        If -Force is specified it will create the report overwriting any existing report of the same name in the same folder.
 
    .EXAMPLE
        Publish-SSRSReport -ReportService $ssrsService -SourceFile "C:\Report.rdl" -Force

    .EXAMPLE
        Publish-SSRSReport -ReportService $ssrsService -SourceFile "C:\Report.rdl" -Folder "/reports/marketing" -ReportName "Quarterly Marketing Report"
    #>
    [CmdletBinding(DefaultParameterSetName="File")] 
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy","SSRSService")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
        [Parameter(Position=1,Mandatory,ParameterSetName="File",ValueFromPipelineByPropertyName)]
        [Alias("File")]
        [string]$SourceFile,

        [Parameter(Position=1,Mandatory,ParameterSetName="Bytes",ValueFromPipelineByPropertyName)]
        [byte[]]$Bytes,
 
        [Parameter(Position=2,Mandatory,ParameterSetName="Bytes",ValueFromPipelineByPropertyName)]
        [Alias("Path")]
        [string]$ReportPath,

        [Parameter(Position=2,Mandatory,ParameterSetName="File",ValueFromPipelineByPropertyName)]
        [Alias("Folder")]
        [string]$ReportFolder,
 
        [Parameter(Position=3,Mandatory=$false)]
        [switch]$Force=$false
    )


    process{
        # The action we perform on the report differs depending on the parameter set.
        switch ($PSCmdlet.ParameterSetName){
            "File"{
                # Get the report content in bytes
                Write-Verbose "Reading file: $SourceFile"
                $Bytes = [System.IO.File]::ReadAllBytes($SourceFile)
                Write-Verbose "File length in bytes: $($byteArray.Count)"
                $ReportName = [io.path]::GetFileNameWithoutExtension($SourceFile)
            }
            "Bytes"{
                # ReportPath should be the complete path to the report object. Need to split it out 
                $ReportName = Split-Path $ReportPath -Leaf
                $ReportFolder = Split-SSRSPath $ReportPath -Parent   
            }
        }


        if(!(Test-SSRSPath -ReportService $ReportService -EntityPath $ReportFolder -EntityType Folder)){
            # Check for an existing folder and create if not found
            Write-Error "The path $ReportFolder is does not exist or is not valid."
            return
        }
 
        try
        {
            Write-Verbose "Uploading to: $reportFolder"
 
            # Call proxy service to upload report
            # Parameters Report ,Parent, Overwrite, Definition, Properties
            # https://msdn.microsoft.com/en-us/library/reportservice2005.reportingservice2005.createreport.aspx
            Write-Verbose "Report Name  : $ReportName"
            Write-Verbose "Report Folder: $ReportFolder"
            Write-Verbose "Force switch : $force"
            $results = $ReportService.CreateReport($ReportName,$ReportFolder,$Force,$Bytes,$null)
            if(!$results){ 
                Write-Verbose "Report uploaded."
            } else { 
                # Results would contain any upload warnings. Display them now.
                $results | ForEach-Object {Write-Warning "Upload Message: $($_.Message)" }
            }

            Write-Information "Uploaded report $ReportName to $ReportFolder"
        }
        catch [IO.IOException]
        {
        
            Write-Error ("Error while reading report`r`nMessage: '{0}'" -f $SourceFile, $_.Exception.Message)
        }
        catch [Web.Services.Protocols.SoapException]
        {
            $errorText = $_.Exception.Detail.InnerText
            if($errorText -match "rsItemAlreadyExists"){
                Write-Error ("Error while uploading report file`r`nMessage:`r'{0}'" -f "Report already exists. Delete or use -Force" )
            } else {
                Write-Error ("Error while uploading report file`r`nMessage:`r`n'{0}'" -f $_.Exception.Detail.InnerText)
            }
        }
    }
 
}

function Unpublish-SSRSReport{
    <#
    .SYNOPSIS
        Unpublishes an RDL file to SQL Reporting Server using a Web Service Proxy
 
    .DESCRIPTION
        Unpublishes an RDL file to SQL Reporting Server using a Web Service Proxy. By default 
        it will ask for confirmation first.
         
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy 
    
    .PARAMETER ReportPath
        Path on the SSRS Server to a particular report to be removed. 

    .PARAMETER Confirm
        Switch governing the request for deletion confirmation from an end user. Defaults to true.
 
    .EXAMPLE
        Unpublish-SSRSReport -ReportService $ssrsService -ReportPath "/Sales/Marketing Final"
 
    .EXAMPLE
        Unpublish-SSRSReport -ReportService $ssrsService -ReportPath "/Sales/Marketing Final" -Confirm
 
    #>
    [CmdletBinding()] 
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy","SSRSService")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=1,Mandatory=$true,ValueFromPipelineByPropertyName)]
        [Alias("Path")]
        [string]$ReportPath,

        [Parameter(Position=2)]
        [switch]$Confirm=$true
    )
 
    # Check if the report already exists
    if(Test-SSRSPath -ReportService $ReportService -Path $ReportPath -EntityType Report){
        Write-Verbose "Removing the report $ReportPath"
        try {
            if(Get-Confirmation -Title "The thing" -Message "Go for it?"){
                $ReportService.DeleteItem($ReportPath)
                Write-Information "The report '$ReportPath' was removed from SSRS"
            } else {
                Write-Verbose "Aborted removing report $ReportPath"
            }
        }
        catch [System.Web.Services.Protocols.SoapException] {
            Write-Error ("Error while removing report file : '{0}'`r`nMessage:`r`n'{1}'" -f $ReportPath, $_.Exception.Detail.InnerText)
        }
    } else {
        Write-Warning "$ReportPath is not a valid report or the path does not exist"
    }
}

function Export-SSRSReport{
    <#
    .SYNOPSIS
        Downloads an SSRS Report to file. 
 
    .DESCRIPTION
        Download the SSRS Report in a given path to a local file path. When successful a fileinfo
        object is returned by the cmdlet.
 
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy 
    
    .PARAMETER ReportPath
        Path on the SSRS Server to a particular report to be exported. 

    .PARAMETER DestinationPath
        Path to export the report to. If the path is a directory the report will be saved with its name.
        Otherwise it will be saved to the full directory. 
        
    .PARAMETER IgnorePath
        Switch that will decide whether the report path folder structure will be recreated when exporting.
        If set to true the report will be placed in the root of the DestinationPath if the
        DestinationPath is to a folder. Defaults to False.
    
    .PARAMETER BytesOnly
        Tells Export-SSRSReport that it will not be saving to file but just outputing the byte stream
        of the downloaded report down the pipe.

    .PARAMETER Force
        If -Force is specified it will overwrite an existing file with the same name.
 
    .EXAMPLE
        Export-SSRSReport -ReportService $ssrsService -ReportPath "/reports/marketing" -DestinationPath "C:\Report.rdl" -Force

    .EXAMPLE
        Export-SSRSReport -ReportService $ssrsService -ReportPath "/reports/marketing" -DestinationPath "C:\Report Directory" 
    #>
    [CmdletBinding(DefaultParameterSetName="File")]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy","SSRSService")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=1,Mandatory=$true,ValueFromPipelineByPropertyName)]
        [Alias("Path")]
        [string]$ReportPath,

        [Parameter(Position=2,Mandatory=$true,ParameterSetName="File")]
        [string]$DestinationPath,

        [Parameter(Mandatory=$false,ParameterSetName="File")]
        [switch]$Force=$false,

        [Parameter(Mandatory=$false,ParameterSetName="File")]
        [switch]$IgnorePath=$false,

        [Parameter(Position=1,Mandatory=$true,ParameterSetName="Bytes")]
        [switch]$BytesOnly=$true
    )

    process{
        # Check if the report to be exported exists. 
        If(Test-SSRSPath -Proxy $ReportService -Path $ReportPath -EntityType Report){
            Write-Verbose "Report Path: $ReportPath is to a valid report"
        } else {
            Write-Error "The path '$ReportPath' is not a valid report path. Ensure the report exists and you have read access to it."
        }  

        # The action we perform on the report differs depending on the parameter set.
        switch ($PSCmdlet.ParameterSetName){
            "File"{
                Write-Verbose "Output Selection: File"
                # Verify that target location to save the report is good.
                # Assume the name is the last element of the DestinationPath
                $outputFileName = Split-Path $DestinationPath -Leaf

                # Determine if the $DestinationPath is a valid directory or contains a valid directory
                if(Test-Path $DestinationPath -PathType Container){
                    $outputDirectory = $DestinationPath
                    # Since this is a valid directory change the output file name to be that of the report to be exported
                    $outputFileName = Split-Path $ReportPath -Leaf
                    Write-Verbose "DestinationPath: '$DestinationPath' is detected as a directory."
                } else {
                    # It is either not valid path or also contains a file name. 
                    # Treat the last element in the path as the new file name.
                    $outputDirectory = Split-Path $DestinationPath
                    if(Test-Path $outputDirectory -PathType Container){
                        $outputFileName = Split-Path $DestinationPath -Leaf
                        Write-Verbose "DestinationPath: '$DestinationPath' is detected as a full file path."
                    } else {
                        Write-Error "The path '$DestinationPath' is not a valid output destination. Make sure the directory exists and is writable."
                    }
                }

                # Append the .rdl extension if not already present
                if($outputFileName -notmatch "\.rdl$"){$outputFileName += ".rdl"}
                if($IgnorePath.IsPresent){
                    Write-Verbose "Output root directory designated as: '$outputDirectory'"
                    [string[]]$pathElements = $outputDirectory,$outputFileName
                } else {
                    $outputDirectory = [io.path]::Combine([string[]](@($outputDirectory) + (Split-Path $ReportPath).TrimStart("\").Split("\")))
                    Write-Verbose "Output root directory designated as: '$outputDirectory'"
                    # Build the path if it does not already exist
                    if(-not (Test-Path $outputDirectory)){
                        Write-Verbose "Building path: '$outputDirectory'"
                        New-Item $outputDirectory -ItemType Directory | Out-Null
                    }
                    [string[]]$pathElements = $outputDirectory,$outputFileName
                }
    
                # Combine the directory and filename into a complete path. 
                $outputPath = [io.path]::Combine($pathElements)

                Write-Verbose "Report Path: $ReportPath"
                Write-Verbose "Output Path: $outputPath"

                # Fail if the file exists already and Force is not set
                if(!$Force.IsPresent -and (Test-path $outputPath -PathType Leaf)){
                    Write-Error "The file '$outputPath' already exists. Either change the DestinationPath or use the -Force switch."
                    return
                }

                # Attempt to export the report to file.
                try{
                    [IO.File]::WriteAllBytes($outputPath, $ReportService.GetReportDefinition($ReportPath))
                    Write-Information "Report saved to '$outputPath'"
                } catch {
                    Write-Error ("Unable to write the report '{0}' to file`r`nMessage:`r`n'{1}'" -f $ReportPath, $_)
                }
        
            }

            "Bytes"{
                Write-Verbose "Output Selection: Bytes"
                # Attempt to export the report to memory.  
                try{
                    return [pscustomobject]@{
                        Name = Split-Path $ReportPath -Leaf
                        Path = $ReportPath
                        Bytes = [byte[]]($ReportService.GetReportDefinition($ReportPath))
                    }
                } catch {
                    Write-Error ("Unable to export the report '{0}'`r`nMessage:`r`n'{1}'" -f $ReportPath, $_)
                }
 
            }
        }
    }
}

function Get-SSRSReportDataSources{
    <#
    .SYNOPSIS
        Returns all data sources associated to a SQL Server Reporting Server report using a Web Service Proxy
 
    .DESCRIPTION
        Returns all data sources associated to a SQL Server Reporting Server report using a Web Service Proxy
         
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy 
    
    .PARAMETER ReportPath
        Path on the SSRS Server to a particular report to be queried. 
 
    .EXAMPLE
        Get-SSRSReportDataSources -ReportService $ssrsService -ReportPath "/Sales/Marketing Final
    #>

    [CmdletBinding()] 
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy","SSRSService")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=1,Mandatory=$true,ValueFromPipelineByPropertyName)]
        [Alias("Path")]
        [string]$ReportPath
    )
 
    process{
        # Test the report path to be sure it is for a valid report
        if(Test-SSRSPath -ReportService $ReportService -EntityPath $ReportPath -EntityType Report){
            $ReportService.GetItemDataSources($reportPath) | ForEach-Object{
                [pscustomobject][ordered]@{
                    ReportPath = $reportPath
                    DataSourceName = $_.name
                    Reference = $_.item.reference
                }
            }
        } else {
            Write-Error "$ReportPath is not a valid report path"
        }
    }
}

function Update-SSRSReportDataSource{
    <#
    .SYNOPSIS
        Updates the datasource in a report assuming there is only one. 
 
    .DESCRIPTION
        Changes the datasource of a given SSRS report at ReportPath on a SQL 
        Server Reporting Server via a Web Service Proxy
         
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy 
    
    .PARAMETER ReportPath
        Path on the SSRS Server to a particular report to be queried. 

    .PARAMETER DataSourcePath
        Path to the datasource on the report server.
 
    .EXAMPLE
        Update-SSRSReportDataSource -ReportService $ssrsService -ReportPath "/Sales/Marketing Final" -DataSourcePath "/DataSource/DS1"
    #>

    [CmdletBinding()] 
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy","SSRSService")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=1,Mandatory=$true)]
        [Alias("Path")]
        [string]$ReportPath,

        [Parameter(Position=2,Mandatory=$true)]
        [string]$DataSourcePath
    )
    
    # Get the namespace of the current service. 
    $ssrsServiceNamespace = $ReportService.Gettype().Namespace
    Write-Verbose "Namespace: $ssrsServiceNamespace"

    # Get the datasource name from the current datasource on the report
    $currentDatasources = Get-SSRSReportDataSources -ReportService $ReportService -ReportPath $ReportPath

    # Check the datasource count. If there is more than one error since we don't want to risk changing too much
    if($currentDatasources.Count -eq 1){
        try{
            # Using the datasource path build a datasource object to add to the report using the current datasource name
            $newDataSource = New-Object "$ssrsServiceNamespace.DataSource"
            $newDataSource.Name = $currentDatasources.DataSourceName
            $newDataSource.Item = New-Object ("$ssrsServiceNamespace.DataSourceReference")
            $newDataSource.Item.Reference = $DataSourcePath
            Write-Verbose "New Datasource Name     : $($newDataSource.Name)"
            Write-Verbose "New Datasource Reference: $($newDataSource.Reference)"
        
            $ReportService.SetItemDataSources($reportPath, $newDataSource)
            Write-Information "The report '$reportpath' datasource was updated to '$DataSourcePath'" 
        } catch {
            Write-Error ("Unable to update report datasource: `r`n $_")
        }
    } else {
        Write-Error "There are too many datasources. Failing as to not risk changes."
    }
}

function Get-Confirmation{
    <#
    .SYNOPSIS
        Gets confirmation from a user for a requested action. 
 
    .DESCRIPTION
        Gets confirmation from a user for a requested action using the Host choice system. Returns true or false.
         
    .PARAMETER Title
        Title used for the menu. Appears a dialog title in ISE.
    
    .PARAMETER Message
        Question to be posed to the user about the action they need to approve. 
 
    .EXAMPLE
        Get-Confirmation -Title "Move on to next step" -Message "Are you sure you want to do this?"
    #>
    param(
        [Parameter(Position=0)]
        [string]$Title="Confirmation",
        
        [Parameter(Position=1,Mandatory=$true)]
        [string]$Message
    )
    $yesChoice = [Management.Automation.Host.ChoiceDescription]::New("&Yes", "Yes I approve this action. Proceed.")
    $noChoice = [Management.Automation.Host.ChoiceDescription]::New("&No", "No I do not approve this action. Abort.")
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yesChoice, $noChoice)
    # Last option sets default choice if user was to hit enter. Defaults to noChoice. The second choice in array
    $result = $host.ui.PromptForChoice($Title, $Message, $choices, 1)

    # Result contains the index of the choice chosen. Since YES and NO are to represent booleans
    # we switch their integer values for their respective boolen equivelents.
    return !$result
}

function Get-SSRSDatasourceDetails{
    <#
    .SYNOPSIS
        Gets datasource details given a path to a datasource. 
    .DESCRIPTION
        Gets all details of a datasource from an SSRS path string provided 
    .PARAMETER ReportService
        Report Service object that would have been created using Connect-SSRSService 
        or New-WebServiceProxy
    .PARAMETER EntityPath
        Destination path of the SQL Server Reporting Services entity i.e. folder, data source etc.
    .EXAMPLE
        Get-SSRSDatasourceDetails -ReportService $ssrsService -EntityPath "/Marketing/Sales/Data Source 1"
    #>
    [CmdletBinding()] 
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [Alias("Proxy")]
        [Web.Services.Protocols.SoapHttpClientProtocol]$ReportService,
 
        [Parameter(Position=1,Mandatory=$true,ValueFromPipelineByPropertyName)]
        [Alias("Path")]
        [string]$EntityPath
    )

    process{
        # Split the path into its folder and entity parts
        $SearchPath = Split-SSRSPath $EntityPath -Parent
        $EntityName = Split-Path $EntityPath -Leaf

        # Verify the path provided is to a valid datasource
        if((Find-SSRSEntities -ReportService $ReportService -SearchPath $SearchPath -EntityType DataSource -Match $EntityName -Partial:$false) -as [boolean]){
            Add-Member -InputObject ($ReportService.GetDataSourceContents($EntityPath)) -MemberType NoteProperty -Name "Path" -Value $EntityPath -PassThru
        } else {
            Write-Warning "Could not find a datasource at path: $EntityPath"
        }
    }
} 