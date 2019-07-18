function Get-PublicKey
{
    [OutputType([byte[]])]
    PARAM (
        [Uri]$Uri
    )

    if (-Not ($uri.Scheme -eq "https"))
    {
        Write-Error "You can only get keys for https addresses"
        return
    }

    $request = [System.Net.HttpWebRequest]::Create($uri)

    try
    {
        #Make the request but ignore (dispose it) the response, since we only care about the service point
        $request.GetResponse().Dispose()
    }
    catch [System.Net.WebException]
    {
        if ($_.Exception.Status -eq [System.Net.WebExceptionStatus]::TrustFailure)
        {
            #We ignore trust failures, since we only want the certificate, and the service point is still populated at this point
        }
        else
        {
            #Let other exceptions bubble up, or write-error the exception and return from this method
            throw
        }
    }

    #The ServicePoint object should now contain the Certificate for the site.
    $servicePoint = $request.ServicePoint
    $key = $servicePoint.Certificate.GetPublicKey()

	#Exporting certificate
	$bytes = $servicePoint.Certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert)
	Set-content -value $bytes -encoding byte -path $certificate_path
}


function Set-Recovery{
    param
    (
        [string] [Parameter(Mandatory=$true)] $ServiceDisplayName,
        [string] [Parameter(Mandatory=$true)] $Server,
        [string] $action1 = "restart",
        [int] $time1 =  30000, # in miliseconds
        [string] $action2 = "restart",
        [int] $time2 =  30000, # in miliseconds
        [string] $actionLast = "restart",
        [int] $timeLast = 30000, # in miliseconds
        [int] $resetCounter = 4000 # in seconds
    )
    $serverPath = "\\" + $server
    $services = Get-CimInstance -ClassName 'Win32_Service' | Where-Object {$_.DisplayName -imatch $ServiceDisplayName}
    $action = $action1+"/"+$time1+"/"+$action2+"/"+$time2+"/"+$actionLast+"/"+$timeLast
    foreach ($service in $services){
        $output = sc.exe $serverPath failure $($service.Name) actions= $action reset= $resetCounter
    }
}

try
{
	$base_url         = ""	#URL of Bamboo Server.
	$java_home        = ""	#JAVA_HOME in the machine where this script will be executed.
	$key_store_pass   = ""	#JAVA Certificates key_store_pass to import in the machine where this script will be executed.
	$jar_agent_path   = ""	#PATH of the "BAMBOO REMOTE AGENT JAR" downloaded using the Bamboo Administration (needed to install the remote agents).
	$jar_agent_name   = ""	#E.G.:atlassian-bamboo-agent-installer-6.7.1.jar.  Jar name/version in the Bamboo Server.
	$agents_quantity  = ""	#Number of remote agents that will be installed.
	$agents_destiny   = ""	#Path folder using for the workingdirectory of Bamboo Agents.
	$tools_folder	  = ""	#Aux folder used for location of some executables.
	$hostname_bamboo  = ""	#Hostname of Bamboo Server (usually is the same of base_url. But it can be different if the server is behind a proxy server.)
	$bamboo_user	  = ""	#using to execute Bamboo Agents as a Windows Service running as user account.
	$bamboo_pass      = ""

	$certificate_path = "C:\Windows\Temp\bambooCertificate.cer"

	#Getting certificate from URL
	Get-PublicKey -Uri $base_url

	#Importing certificate to the local keystore
	$key_store_path =	Join-Path -Path $java_home -ChildPath "jre\lib\security\cacerts"
	$keytoolPath 	=	Join-Path -Path $java_home -ChildPath "bin\keytool.exe"
	& $keytoolPath -genkey -noprompt -importcert -file $certificate_path -keystore $key_store_path -alias "bambooCertificate4" -storepass $key_store_pass -keypass $key_store_pass

	#Copy PATH environment variable from system to user (to prevent/solution the problem of environments variables)
	#In old versions of the Java wrapper, Windows remote agents used to present problems with the environment varibles.
	#Atlassian reference error:  BAM-16205 --  https://jira.atlassian.com/browse/BAM-16205?_ga=2.239235845.701156863.1556300330-1052637895.1533827297 --
	#The problem can be resolved updating the Java wrapper version (as I did in line 156) or copying the PATH variable value from system to user.
	#$userPath = [System.Environment]::GetEnvironmentVariable("PATH", "User")
	#$systemPath = [Environment]::GetEnvironmentVariable("PATH","Machine")
	#[System.Environment]::SetEnvironmentVariable("PATH", "", "User")
	#[System.Environment]::SetEnvironmentVariable("PATH", $userPath + ";" + $systemPath, "User")
	############################################################################################################

	#Installing agents
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$agentServicePath = $agents_destiny + "bamboo-agent-$($i)"
		
		$jarFile = Join-Path -Path $jar_agent_path -ChildPath $jar_agent_name
		
		$java_exe = Join-Path -Path $java_home -ChildPath "bin\java.exe"
		
		#If service exists, first is necessary uninstall it
		If (Get-Service "bamboo-remote-agent-$($i)" -ErrorAction SilentlyContinue) {
			$uninstallServiceCommand = $agents_destiny + "bamboo-agent-$($i)" + "\bin\UninstallBambooAgent-NT.bat"
			& $uninstallServiceCommand	
		}
		
		#ignoreServerCertName=false if you can use the SSL connection method with the server.
		& $java_exe "-Dbamboo.agent.ignoreServerCertName=true" "-Dbamboo.home=$($agentServicePath)" -jar "$($jarFile)" "$($base_url)/agentServer" install
	}


	#Changing properties for each service
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$wrapperConfPath = $agents_destiny + "bamboo-agent-$($i)" + "\conf\wrapper.conf"
		
		$replaceText = "wrapper.ntservice.name=bamboo-remote-agent-$($i)"
		(Get-Content $wrapperConfPath) -replace "^wrapper.ntservice.name=.*$", $replaceText | Set-Content $wrapperConfPath
		
		$replaceText = "wrapper.ntservice.displayname=Bamboo Remote Agent $($i)"
		(Get-Content $wrapperConfPath) -replace "^wrapper.ntservice.displayname=.*$", $replaceText | Set-Content $wrapperConfPath
		
		<#$replaceText = "wrapper.java.command=C:\Program Files (x86)\Java\jdk1.8.0_191"
		(Get-Content $wrapperConfPath) -replace "^wrapper.java.command=.*$", $replaceText | Set-Content $wrapperConfPath #>
		
		$replaceText = "wrapper.java.initmemory=512"
		(Get-Content $wrapperConfPath) -replace "^wrapper.java.initmemory=.*$", $replaceText | Set-Content $wrapperConfPath
		
		$replaceText = "wrapper.java.maxmemory=1024"
		(Get-Content $wrapperConfPath) -replace "^wrapper.java.maxmemory=.*$", $replaceText | Set-Content $wrapperConfPath
		
		$replaceText = "wrapper.app.parameter.2=$($hostname_bamboo)/agentServer"
		(Get-Content $wrapperConfPath) -replace "^wrapper.app.parameter.2=.*$", $replaceText | Set-Content $wrapperConfPath
		
		$replaceText = "wrapper.java.additional.3=-Djava.io.tmpdir=""$($tools_folder)\tmpdir"""
		(Get-Content $wrapperConfPath) -replace "^#wrapper.java.additional.3=-Dlog4j.configuration.*$", $replaceText | Set-Content $wrapperConfPath
		
		######The commands below only will work if the wrapper is update to version 3.5.5######
		#Add-Content $wrapperConfPath "`n"
		#Add-Content $wrapperConfPath "wrapper.ntservice.recovery.1.delay=60"
		#Add-Content $wrapperConfPath "wrapper.ntservice.recovery.1.failure=RESTART"
		#######################################################################################
		
		$bamboo_pass_escaped = $bamboo_pass -replace '#','##'
		
		Add-Content $wrapperConfPath "`n"
		Add-Content $wrapperConfPath "wrapper.ntservice.account=$($bamboo_user)"
		Add-Content $wrapperConfPath "wrapper.ntservice.password=$($bamboo_pass_escaped)"
	}


	<#
	##########################################################################################################
	#Updating Java Service Wrapper for each agent (to prevent/solution the problem of environments variables)#
	##########################################################################################################

	#Download the wrapper
	$url = "https://download.tanukisoftware.com/wrapper/3.5.35/wrapper-windows-x86-32-3.5.35.zip"
	$output = "C:\Windows\Temp\wrapper-windows-x86-32-3.5.35.zip"
	$start_time = Get-Date

	Invoke-WebRequest -Uri $url -OutFile $output

	#Unzip it
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	function Unzip
	{
		param([string]$zipfile, [string]$outpath)

		if (Test-Path $outpath) 
		{
			Remove-Item $outpath -Force
		}
		
		[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
	}
	Unzip $output $tools_folder

	#Replace the old versions of the wrapper in each agent
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$targetFolder = $agents_destiny + "bamboo-agent-$($i)"
			
		$source = "$($tools_folder)\wrapper-windows-x86-32-3.5.35\bin\wrapper.exe"
		Copy-Item $source -Destination $targetFolder\bin -Recurse -Force

		$source = "$($tools_folder)\wrapper-windows-x86-32-3.5.35\lib\wrapper.dll"
		Copy-Item $source -Destination $targetFolder\lib -Recurse -Force

		$source = "$($tools_folder)\wrapper-windows-x86-32-3.5.35\lib\wrapper.jar"
		Copy-Item $source -Destination $targetFolder\lib -Recurse -Force
	}

	##########################################################################################################
	#>

	#Creating service for each agent
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$installServicePath = $agents_destiny + "bamboo-agent-$($i)" + "\bin\InstallBambooAgent-NT.bat"
		
		& $installServicePath
	}


	#Creating bamboo-capabilities.properties for each agent.
	#Those are some classic capabilities that you could use in an Windows environment.
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$pathFile = $agents_destiny + "bamboo-agent-$($i)" + "\bin\bamboo-capabilities.properties"
		
		if (Test-Path $pathFile) 
		{
			Remove-Item $pathFile
		}

		New-Item $pathFile -ItemType file
		
		Add-Content $pathFile "###########################"
		Add-Content $pathFile "# Executable capabilities #"
		Add-Content $pathFile "###########################"
		Add-Content $pathFile "#MSBuild"
		Add-Content $pathFile "system.builder.MSBuild.MSBuild\ v15.0\ (32bit)=C:\\Program Files (x86)\\Microsoft Visual Studio\\2017\\BuildTools\\MSBuild\\15.0\\Bin\\Msbuild.exe"
		Add-Content $pathFile "system.builder.MSBuild.MSBuild\ v15.0\ (64bit)=C:\\Program Files (x86)\\Microsoft Visual Studio\\2017\\BuildTools\\MSBuild\\15.0\\Bin\\amd64\\Msbuild.exe"

		Add-Content $pathFile "`n"
		Add-Content $pathFile "#NodeJs"
		Add-Content $pathFile "system.builder.node.NodeJs=C:\\Program Files\\nodejs\\node.exe"
		
		Add-Content $pathFile "`n"
		Add-Content $pathFile "#AspNetCompiler"
		Add-Content $pathFile "system.builder.Command.AspNetCompiler\ (32bit)=C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319\\aspnet_compiler.exe"
		Add-Content $pathFile "system.builder.Command.AspNetCompiler\ (64bit)=C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\aspnet_compiler.exe"

		Add-Content $pathFile "`n"
		Add-Content $pathFile "#Aspnet_regiis"
		Add-Content $pathFile "system.builder.Command.Aspnet_regiis\ (32bit)=C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319\\aspnet_regiis.exe"
		Add-Content $pathFile "system.builder.Command.Aspnet_regiis\ (64bit)=C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\aspnet_regiis.exe"

		Add-Content $pathFile "`n"
		Add-Content $pathFile "#CMD"
		Add-Content $pathFile "system.builder.Command.CMD=C:\\Windows\\System32\\cmd.exe"
		Add-Content $pathFile "system.builder.Command.DotCover=C:\\tools\\dotcover-cli\\dotcover.exe"
		Add-Content $pathFile "system.builder.Command.Dotnet=C:\\Program Files\\dotnet\\dotnet.exe"
		Add-Content $pathFile "system.builder.Command.Lc=C:\\Program Files (x86)\\Microsoft SDKs\\Windows\\v10.0A\\bin\\NETFX 4.6.1 Tools\\Lc.exe"
		Add-Content $pathFile "system.builder.MSBuild.MSBuild-Sonar-Scanner=C:\\tools\\sonar-scanner-msbuild\\"
		Add-Content $pathFile "system.builder.Command.Maven=C:\\tools\\apache-maven-3.5.4"
		Add-Content $pathFile "system.builder.Command.NUnit=C:\\tools\\nunit\\nunit3-console.exe"
															
		Add-Content $pathFile "system.builder.Command.Ng=C:\\tools\\npm\\ng.cmd"
		Add-Content $pathFile "system.builder.Command.Nuget=C:\\tools\\Nuget\\nuget.exe"
		Add-Content $pathFile "system.builder.Command.Robocopy=C:\\Windows\\System32\\Robocopy.exe"
		Add-Content $pathFile "system.builder.Command.Unzip=C:\\ProgramData\\chocolatey\\lib\\unzip\\tools\\unzip.exe"
		Add-Content $pathFile "system.builder.Command.Versioner=C:\\tools\\versioner\\versioner.exe"
		Add-Content $pathFile "system.builder.Command.XUnit=C:\\tools\\xunit\\xunit.console.exe"
		Add-Content $pathFile "system.builder.Command.Yarn=C:\\Program Files (x86)\\Yarn\\bin\\Yarn.cmd"
		Add-Content $pathFile "system.builder.Command.Zip=C:\\ProgramData\\chocolatey\\lib\\zip\\tools\\zip.exe"
		
		
		Add-Content $pathFile "`n"
		Add-Content $pathFile "#JDK Capabilities"
		Add-Content $pathFile "system.jdk.JDK\ 1.8.0_191=C:\\Program Files\\Java\\jdk1.8.0_191\\"

		Add-Content $pathFile "`n"
		Add-Content $pathFile "#######################"
		Add-Content $pathFile "# Custom capabilities #"
		Add-Content $pathFile "#######################"
		Add-Content $pathFile "OS=Windows"
	}


	#Starting service for each agent
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$installServicePath = $agents_destiny + "bamboo-agent-$($i)" + "\bin\StartBambooAgent-NT.bat"
		& $installServicePath
	}

	#Changing the working directory for each service (only is possible after the agent is authenticated on the server so, the first time for this script in a new server will not work)
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$agentConfPath = $agents_destiny + "bamboo-agent-$($i)" + "\bamboo-agent.cfg.xml"
				
		if (Test-Path $agentConfPath) 
		{
			[xml]$xml = get-content $agentConfPath;

			if (Test-Path D:)
			{
				$xml.configuration.buildWorkingDirectory = "D:\"
			}
			else
			{
				if (-Not (Test-Path "$($tools_folder)\workingDirectoryBamboo"))
				{
					New-Item -ItemType directory -Path "$($tools_folder)\workingDirectoryBamboo"
				}
				$xml.configuration.buildWorkingDirectory = "$($tools_folder)\workingDirectoryBamboo"
			}
			$xml.Save($agentConfPath)
		}
	}
	
	#Deleting capabilities file for each agent
	For ($i=1; $i -le $agents_quantity; $i++) { 
		$pathFile = $agents_destiny + "bamboo-agent-$($i)" + "\bin\bamboo-capabilities.properties"
		
		if (Test-Path $pathFile) 
		{
			Remove-Item $pathFile
		}
	}

	#Restarting windows service for each agent (to previous change take effect)
	For ($i=1; $i -le $agents_quantity; $i++) {
		Restart-Service -Name "bamboo-remote-agent-$($i)"
	}
	
	#With the current version of the wrapper, this for is needed to configure the recovery options for each service
	For ($i=1; $i -le $agents_quantity; $i++) {
		Set-Recovery -ServiceDisplayName "Bamboo Remote Agent $($i)" -Server "localhost"
	}	
}
catch [Exception] {
   Write-Error $_.Exception.Message;
}