<#
.Synopsis
Exchange Receipient Management Tools Local Web Server
.Description
Starts webserver as powershell process as the current user
Navigate to the web site to use Exchange Admin Tools
.Inputs
None
.Outputs
None
.Example
Start-ExAdminWeb.ps1
.Notes
Based on: WebServer - Version 1.2.2, 2022-01-19
Based on Author: Markus Scholtes
Based on: https://github.com/MScholtes/WebServer
.LINK

#>


if (!(Get-PSSnapin Microsoft.Exchange.Management.PowerShell.RecipientManagement -Registered -ErrorAction SilentlyContinue))
{
	throw "Please install the Exchange 2019 CU12 and above Management Tools-Only install. See: https://docs.microsoft.com/en-us/Exchange/manage-hybrid-exchange-recipients-with-management-tools"
	break
}

# Load Recipient Management PowerShell Tools
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.RecipientManagement

# Define webserver details
$BASEDIR = $PSScriptRoot + "/web"
$BINDING = "http://localhost:$(Get-Random -Minimum 4000 -Maximum 10000)/"
$INDEX = "\index.html"

# MIME hash table for static content
$MIMEHASH = @{".avi" = "video/x-msvideo"; ".crt" = "application/x-x509-ca-cert"; ".css" = "text/css"; ".der" = "application/x-x509-ca-cert"; ".doc" = "application/msword"; ".flv" = "video/x-flv"; ".gif" = "image/gif"; ".htm" = "text/html"; ".html" = "text/html"; ".ico" = "image/x-icon"; ".jar" = "application/java-archive"; ".jpeg" = "image/jpeg"; ".jpg" = "image/jpeg"; ".js" = "application/javascript"; ".json" = "application/json"; ".mjs" = "application/javascript"; ".mov" = "video/quicktime"; ".mp3" = "audio/mpeg"; ".mp4" = "video/mp4"; ".mpeg" = "video/mpeg"; ".mpg" = "video/mpeg"; ".pdf" = "application/pdf"; ".pem" = "application/x-x509-ca-cert"; ".pl" = "application/x-perl"; ".png" = "image/png"; ".rss" = "application/rss+xml"; ".shtml" = "text/html"; ".txt" = "text/plain"; ".war" = "application/java-archive"; ".wmv" = "video/x-ms-wmv"; ".xml" = "application/xml"; ".xsl" = "application/xml" }

# Result Message Placeholders
$HTML_SUCCESS = "<div class=`"alert alert-success d-flex align-items-center`" role=`"alert`">{result}</div>"
$HTML_WARN = "<div class=`"alert alert-warning  d-flex align-items-center`" role=`"alert`">{result}</div>"

# Starting the powershell webserver
"$(Get-Date -Format s) Starting Exchange Recipient Admin Webserver at: $($BINDING)"
$LISTENER = New-Object System.Net.HttpListener
$LISTENER.Prefixes.Add($BINDING)
$LISTENER.Start()
$Error.Clear()

Start-Process -FilePath $BINDING

function Get-AcceptedDomains
{
	Param
	(
		[switch]$TableView,
		[switch]$RemoteRouting
	)

	$AcceptedDomains = Get-AcceptedDomain | Select-Object Name, DomainName, DomainType

	if ($RemoteRouting)
	{
		$AcceptedDomains = $AcceptedDomains | Where-Object DomainName -Match "mail.onmicrosoft.com"
	}

	if ($TableView)
	{
		foreach ($Item in $AcceptedDomains)
		{
			$AD_TableView += "
			<tr>
			<th scope=`"row`">
			<a href=`"#`">$($Item.Name)</a></th>
			<td>$($Item.DomainName)</td>
			<td>$($Item.DomainType)</td>
			</tr>"
		}

		Return $AD_TableView
	}

	else
	{
		foreach ($Item in $AcceptedDomains)
		{
			if ($Item.Default)
			{
				$AD_DropDown += "`n<option selected value=`"$($Item.DomainName)`">$($Item.DomainName)</option>"
			}

			else
			{
				$AD_DropDown += "`n<option value=`"$($Item.DomainName)`">$($Item.DomainName)</option>"
			}
		}

		Return $AD_DropDown
	}
}

function Get-NonMailboxUsers
{
	$NonMBXUsr = ""
	foreach ($Item in (Get-User -Filter "RecipientType -eq 'User' -and RecipientTypeDetails -ne 'DisabledUser'" | Where-Object { $_.UserPrincipalName }))
	{
		$NonMBXUsr += "`n<option value=`"$($Item.UserPrincipalName)`">$($Item.UserPrincipalName)</option>"
	}

	Return $NonMBXUsr

}

function Get-RemoteMailboxes
{
	$RemoteMBX = ""
	foreach ($Item in (Get-RemoteMailbox | Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenChanged))
	{
		$RemoteMBX += "
		<tr>
		<th scope=`"row`">
		<a href=`"/editremotemailbox?id=$($Item.PrimarySMTPAddress)`">$($Item.DisplayName)</a></th>
		<td>$($Item.PrimarySMTPAddress)</td>
		<td>$($Item.RecipientTypeDetails)</td>
		<td>$($Item.WhenChanged)</td>
		</tr>"
	}
}

function Get-EmailAddressPolicies
{
	foreach ($Item in (Get-EmailAddressPolicy | Select-Object Name, Priority, RecipientFilter, RecipientContainer, EnabledEmailAddressTemplates))
	{
		$EAP += "
		<tr>
		<th scope=`"row`">
		<a href=`"#`">$($Item.Name)</a></th>
		<td>$($Item.Priority)</td>
		<td>$($item.EnabledEmailAddressTemplates.AddressTemplateString)</td>
		<td>$($Item.RecipientFilter)</td>
		<td>$($Item.RecipientContainer)</td>
		</tr>"
	}

	Return $EAP

}

function Get-ExternalRecipients
{
	Param
	(
		[switch]$MailUsers
	)

	if ($MailUsers)
	{
		$MailUsr = ""
		foreach ($Item in (Get-MailUser | Select-Object DisplayName, PrimarySMTPAddress))
		{
			$MailUsr += "
		<tr>
		<th scope=`"row`">
		<a href=`"#`">$($Item.DisplayName)</a></th>
		<td>$($Item.PrimarySMTPAddress)</td>
		</tr>"
		}
	
		Return $MailUsr
	}

	else
	{
		$MailContacts = ""
		foreach ($Item in (Get-MailContact | Select-Object DisplayName, PrimarySMTPAddress))
		{
			$MailContacts += "
		<tr>
		<th scope=`"row`">
		<a href=`"#`">$($Item.DisplayName)</a></th>
		<td>$($Item.PrimarySMTPAddress)</td>
		</tr>"
		}

		Return $MailContacts
	}
}

function Get-DistributionGroups
{
	Param
	(
		[switch]$Security
	)

	$AllGroups = (Get-DistributionGroup | Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenCreated)

	if ($Security)
	{
		$SecurityGroups = $AllGroups | Where-Object RecipientTypeDetails -EQ "MailUniversalSecurityGroup"

		$SecGroups = ""
		foreach ($group in $SecurityGroups)
		{
			$SecGroups += "
			<tr>
			<th scope=`"row`">
			<a href=`"#`">$($group.DisplayName)</a></th>
			<td>$($group.PrimarySMTPAddress)</td>
			<td>$($group.WhenCreated)</td>
			</tr>"
		}

		Return $SecGroups
	}

	else
	{
		$DistributionGroups = $AllGroups | Where-Object RecipientTypeDetails -EQ "MailUniversalDistributionGroup"

		$DistGroups = ""
		foreach ($group in $DistributionGroups)
		{
			$DistGroups += "
			<tr>
			<th scope=`"row`">
			<a href=`"#`">$($group.DisplayName)</a></th>
			<td>$($group.PrimarySMTPAddress)</td>
			<td>$($group.WhenCreated)</td>
			</tr>"
		}

		Return $DistGroups
	}
}

function Get-NonMailGroups
{
	$NonMailGroups = Get-Group -Filter {(RecipientType -eq "Group") -and (RecipientTypeDetails -ne "NonUniversalGroup") -and (RecipientTypeDetails -ne "RoleGroup")}
	$NonMailGroupsList
	foreach ($group in $NonMailGroups)
	{
		$NonMailGroupsList += "`n<option value=`"$($Item.UserPrincipalName)`">$($Item.UserPrincipalName)</option>"
	}

	Return $NonMailGroupsList
}

$HTMLROWS_AD = Get-AcceptedDomains
$HTMLLIST_AD = Get-AcceptedDomains -TableView
$HTMLROWS_RRA = Get-AcceptedDomains -RemoteRouting
$HTMLROWS_USERS = Get-NonMailboxUsers
$HTMLROWS_MBX = Get-RemoteMailboxes
$HTMLROWS_EAP = Get-EmailAddressPolicies
$HTMLROWS_Contacts = Get-ExternalRecipients
$HTMLROWS_Mailusers = Get-ExternalRecipients -MailUsers
$HTMLROWS_DL = Get-DistributionGroups
$HTMLROWS_MES = Get-DistributionGroups -Security
$HTMLROWS_GROUPS	
try
{
	"$(Get-Date -Format s) Powershell webserver started."
	$WEBLOG = "$(Get-Date -Format s) Powershell webserver started.`n"
	while ($LISTENER.IsListening)
	{
		# analyze incoming request
		$CONTEXT = $LISTENER.GetContext()
		$REQUEST = $CONTEXT.Request
		$RESPONSE = $CONTEXT.Response
		$RESPONSEWRITTEN = $FALSE

		# log to console
		"$(Get-Date -Format s) $($REQUEST.RemoteEndPoint.Address.ToString()) $($REQUEST.httpMethod) $($REQUEST.Url.PathAndQuery)"
		# and in log variable
		$WEBLOG += "$(Get-Date -Format s) $($REQUEST.RemoteEndPoint.Address.ToString()) $($REQUEST.httpMethod) $($REQUEST.Url.PathAndQuery)`n"
		$RECEIVED = '{0} {1}' -f $REQUEST.httpMethod, $REQUEST.Url.LocalPath
		# check for known commands
		switch ($RECEIVED)
		{

			"GET /"
			{ 
				# Return the dashboard homepage
				$HTMLRESPONSE = Get-Content -Path "$($BASEDIR)\index.html"
				break
			}

			"GET /remotemailboxes"
			{ 
				# Remote Mailbox Section
				
				# Process submitted form
				if ($REQUEST.Url.Query)
				{
					$Table = @{}
					foreach ($Item in [URI]::UnescapeDataString(($REQUEST.Url.Query.Replace("?", ""))).Split("&"))
					{
						$Table.Add($Item.Split("=")[0], $Item.Split("=")[1])
					}
					try
					{
						$Result = Enable-RemoteMailbox -Identity $Table['username'] -PrimarySMTPAddress "$($Table['primarysmtpaddress_local'])@$($Table['primarysmtpaddress_accepteddomain'])" -RemoteRoutingAddress "$($Table['remoteroutingaddress_local'])@$($Table['remoteroutingaddress_accepteddomain'])" -ErrorAction Stop
						$HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "User $($Table['username']) enabled as Remote Mailbox")
					}
					catch
					{
						$HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
					}
					
				}

				# Prepare user list for non-Exchange users
				#$HTMLROWS_USERS = ""
				#foreach ($Item in (Get-User -Filter "RecipientType -eq 'User' -and RecipientTypeDetails -ne 'DisabledUser'" | Where-Object { $_.UserPrincipalName }))
				#{
				#	$HTMLROWS_USERS += "`n<option value=`"$($Item.UserPrincipalName)`">$($Item.UserPrincipalName)</option>"
				#}

				# Prepare accepted domain list
				#$HTMLROWS_AD = Get-AcceptedDomains
				#				foreach ($Item in (Get-AcceptedDomain))
				#				{
				#					
				#					if ($Item.Default)
				#     {
				#						$HTMLROWS_AD += "`n<option selected value=`"$($Item.Name)`">$($Item.DomainName)</option>"
				#					}
				#					else
				#					{
				#						$HTMLROWS_AD += "`n<option value=`"$($Item.Name)`">$($Item.DomainName)</option>"
				#					}
				#				}

				# Prepare remote routing domain list
				#$HTMLROWS_RRA = ""
				#foreach ($Item in (Get-AcceptedDomain))
				#{
				#	
				#	if ($Item.DomainName -like "*.mail.onmicrosoft.com")
				#	{
				#		$HTMLROWS_RRA += "`n<option selected value=`"$($Item.Name)`">$($Item.DomainName)</option>"
				#	}
				#	else
				#	{
				#		$HTMLROWS_RRA += "`n<option value=`"$($Item.Name)`">$($Item.DomainName)</option>"
				#	}
				#}

				# Return remote mailbox list
				#$HTMLROWS_MBX = ""
				#foreach ($Item in (Get-RemoteMailbox | Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenChanged))
				#{
				#	$HTMLROWS_MBX += "
				#	<tr>
				#	<th scope=`"row`">
				#	<a href=`"/editremotemailbox?id=$($Item.PrimarySMTPAddress)`">$($Item.DisplayName)</a></th>
				#	<td>$($Item.PrimarySMTPAddress)</td>
				#	<td>$($Item.RecipientTypeDetails)</td>
				#	<td>$($Item.WhenChanged)</td>
				#	</tr>"
				#}

				# Create response and replace template placeholders
				$HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\remotemailboxes.html").Replace("<!-- {row_mbx} -->", $HTMLROWS_MBX).Replace("<!-- {row_ad} -->", $HTMLROWS_AD).Replace("<!-- {row_user} -->", $HTMLROWS_USERS).Replace("<!-- {row_rra} -->", $HTMLROWS_RRA).Replace("<!-- {result} -->", $HTML_RESULT)
				break
			}

			"GET /distributiongroups"
			{ 
				# Distribution Groups Section

				# Prepare Distibution Group lists split into tabs
				#$HTMLROWS_DL = ""
				#$HTMLROWS_MES = ""
				#foreach ($Item in (Get-DistributionGroup | Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenCreated))
				#{
					if ($Item.RecipientTypeDetails -eq "MailUniversalDistributionGroup")
					{
						$HTMLROWS_DL += "
						<tr>
						<th scope=`"row`">
						<a href=`"#`">$($Item.DisplayName)</a></th>
						<td>$($Item.PrimarySMTPAddress)</td>
						<td>$($Item.WhenCreated)</td>
						</tr>"
					}
					elseif ($Item.RecipientTypeDetails -eq "MailUniversalSecurityGroup")
					{
						$HTMLROWS_MES += "
						<tr>
						<th scope=`"row`">
						<a href=`"#`">$($Item.DisplayName)</a></th>
						<td>$($Item.PrimarySMTPAddress)</td>
						<td>$($Item.WhenCreated)</td>
						</tr>"
					}
				#}

				# Create response and replace template placeholders
				$HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\distributiongroups.html").Replace("<!-- {row_dl} -->", $HTMLROWS_DL).Replace("<!-- {row_mes} -->", $HTMLROWS_MES).Replace("<!-- {result} -->", $HTML_RESULT)
				break
			}
			
			"GET /mailusers"
			{ 
				# Mail Users Section
				# Process submitted form
				if ($REQUEST.Url.Query)
				{
					$Table = @{}
					foreach ($Item in [URI]::UnescapeDataString(($REQUEST.Url.Query.Replace("?", ""))).Split("&"))
					{
						$Table.Add($Item.Split("=")[0], $Item.Split("=")[1])
					}
					try
					{
						$Result = Enable-MailUser -Identity $Table['mailenableuser_username'] -ExternalEmailAddress $Table['mailenableuser_address'] -ErrorAction Stop
						$HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "User $($Table['mailenableuser_username']) was enabled as Mail User with external E-Mail Address $($Table['mailenableuser_address'])")
					}
					catch
					{
						$HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
					}
					
				}

				# Prepare contacts list
				#$HTMLROWS = ""
				#foreach ($Item in (Get-MailContact | Select-Object DisplayName, PrimarySMTPAddress, RecipientType))
				#{
				#	$HTMLROWS += "
				#	<tr>
				#	<th scope=`"row`">
				#	<a href=`"#`">$($Item.DisplayName)</a></th>
				#	<td>$($Item.PrimarySMTPAddress)</td>
				#	<td>$($Item.RecipientType)</td>
				#	</tr>"
				#}

				# Create response and replace template placeholders
				$HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\mailusers.html").Replace("<!-- {row} -->", $HTMLROWS_Mailusers).Replace("<!-- {row_user} -->", $HTMLROWS_USERS).Replace("<!-- {result} -->", $HTML_RESULT)
				break
			}

			"GET /contacts"
			{ 
				# Mail Contacts Section
				if ($REQUEST.Url.Query)
				{
					$Table = @{}
					foreach ($Item in [URI]::UnescapeDataString(($REQUEST.Url.Query.Replace("?", ""))).Split("&"))
					{
						$Table.Add($Item.Split("=")[0], $Item.Split("=")[1])
					}
					try
					{
						$Result = New-MailContact -Name $Table['newmailcontact_username'] -DisplayName $Table['newmailcontact_username'] -ExternalEmailAddress $Table['newmailcontact_address'] -ErrorAction Stop
						$HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Successfully created Mail Contact $($Table['newmailcontact_username']) with external E-Mail Address $($Table['newmailcontact_address'])")
					}
					catch
					{
						$HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
					}
					
				}

				# Prepare contacts list
				#$HTMLROWS = ""
				#foreach ($Item in (Get-MailContact | Select-Object DisplayName, PrimarySMTPAddress, RecipientType))
				#{
				#	$HTMLROWS += "
				#	<tr>
				#	<th scope=`"row`">
				#	<a href=`"#`">$($Item.DisplayName)</a></th>
				#	<td>$($Item.PrimarySMTPAddress)</td>
				#	<td>$($Item.RecipientType)</td>
				#	</tr>"
				#}

				# Create response and replace template placeholders
				$HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\contacts.html").Replace("<!-- {row} -->", $HTMLROWS_Contacts).Replace("<!-- {result} -->", $HTML_RESULT)
				break
			}

			"GET /emailaddresspolicies"
			{ 
				# Email Address Policies Section

				# Process submitted form
				if ($REQUEST.Url.Query)
				{
					$Table = @{}
					foreach ($Item in [URI]::UnescapeDataString(($REQUEST.Url.Query.Replace("?", ""))).Split("&"))
					{
						$Table.Add($Item.Split("=")[0], $Item.Split("=")[1])
					}
					try
					{
						$Result = New-EmailAddressPolicy -Name $Table['newemailaddresspolicy_friendlyname'] -EnabledPrimarySMTPAddressTemplate "SMTP:$($Table['newemailaddresspolicy_emailgenformat'])@$($Table['newemailaddresspolicy_domain'])" -IncludedRecipients $Table['newemailaddresspolicy_includedrecipients'] -RecipientContainer $Table['newemailaddresspolicy_recipientcontainer'] -ErrorAction Stop
						$HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Email address policy $($Table['newemailaddresspolicy_friendlyname']) was successfully added.")
					}
					catch
					{
						$HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
					}
					
				}

				# Prepare email address policies list
				#$HTMLROWS = ""
				#foreach ($Item in (Get-EmailAddressPolicy | Select-Object Name, Priority, RecipientFilter, RecipientContainer))
				#{
				#	$HTMLROWS += "
				#	<tr>
				#	<th scope=`"row`">
				#	<a href=`"#`">$($Item.Name)</a></th>
				#	<td>$($Item.Priority)</td>
				#	<td>$($Item.RecipientFilter)</td>
				#	<td>$($Item.RecipientContainer)</td>
				#	</tr>"
				#}
				#
				# Prepare accepted domain list
				#$HTMLROWS_AD = ""
				#foreach ($Item in (Get-AcceptedDomain))
				#{
				#	
				#	if ($Item.Default)
				#	{
				#		$HTMLROWS_AD += "`n<option selected value=`"$($Item.DomainName)`">$($Item.DomainName)</option>"
				#	}
				#	else
				#	{
				#		$HTMLROWS_AD += "`n<option value=`"$($Item.DomainName)`">$($Item.DomainName)</option>"
				#	}
				#}
				# Create response and replace template placeholders
				$HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\emailaddresspolicies.html").Replace("<!-- {row} -->", $HTMLROWS_EAP).Replace("<!-- {row_domains} -->", $HTMLROWS_AD).Replace("<!-- {result} -->", $HTML_RESULT)
				break
			}

			"GET /accepteddomains"
			{ 
				# Accepted Domains section

				# Process submitted form
				if ($REQUEST.Url.Query)
				{
					$Table = @{}
					foreach ($Item in [URI]::UnescapeDataString(($REQUEST.Url.Query.Replace("?", ""))).Split("&"))
					{
						$Table.Add($Item.Split("=")[0], $Item.Split("=")[1])
					}
					try
					{
						$Result = New-AcceptedDomain -Name $Table['addacceptedomain_friendlyname'] -DomainName $Table['addacceptedomain_domainname'] -DomainType $Table['addacceptedomain_domaintype'] -ErrorAction Stop
						$HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Domain $($Table['addacceptedomain_domainname']) was successfully added as Accepted Domain of type $($table['addacceptedomain_domaintype']) ")
					}
					catch
					{
						$HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
					}
					
				}

				# Prepare list of accepted domains
				#$HTMLROWS = ""
				#foreach ($Item in (Get-AcceptedDomain))
				#{
				#	$HTMLROWS += "
				#	<tr>
				#	<th scope=`"row`">
				#	<a href=`"#`">$($Item.Name)</a></th>
				#	<td>$($Item.DomainName)</td>
				#	<td>$($Item.DomainType)</td>
				#	</tr>"
				#}

				# Create response and replace template placeholders
				$HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\accepteddomains.html").Replace("<!-- {row} -->", $HTMLLIST_AD).Replace("<!-- {result} -->", $HTML_RESULT)
				break
			}

			"GET /exit"
			{ 
				# Create response preparing for webserver shutdown
				$HTMLRESPONSE = "<!doctype html><html><body>Please close the browser window</body></html>"
				break
			}

			default
			{	
					
				# PowerShell webserver main code - this section should be updated if the main project is
					
				# unknown command, check if path to file
 
				# create physical path based upon the base dir and url
				$CHECKDIR = $BASEDIR.TrimEnd("/\") + $REQUEST.Url.LocalPath
				$CHECKFILE = ""
				if (Test-Path $CHECKDIR -PathType Container)
				{
					# physical path is a directory
					$INDEX = "/index.html"
					$CHECKFILE = $CHECKDIR.TrimEnd("/\") + $INDEX
					if (Test-Path $CHECKFILE -PathType Leaf)
					{
						# index file found, path now in $CHECKFILE
						break
					}
					$CHECKFILE = ""
						
					if ($CHECKFILE -eq "")
					{
						# do not generate directory listing - 404 
						# no file to serve found, return error
						$RESPONSE.StatusCode = 404
						$HTMLRESPONSE = "<!doctype html><html><body>Page $($RECEIVED) not found</body></html>"
					}
				}
				else
				{
					# no directory, check for file
					if (Test-Path $CHECKDIR -PathType Leaf)
					{
						# file found, path now in $CHECKFILE
						$CHECKFILE = $CHECKDIR
					}
				}

				if ($CHECKFILE -ne "")
				{
					# static content available
					try
					{
						# ... serve static content
						$BUFFER = [System.IO.File]::ReadAllBytes($CHECKFILE)
						$RESPONSE.ContentLength64 = $BUFFER.Length
						$RESPONSE.SendChunked = $FALSE
						$EXTENSION = [IO.Path]::GetExtension($CHECKFILE)
						if ($MIMEHASH.ContainsKey($EXTENSION))
						{
							# known mime type for this file's extension available
							$RESPONSE.ContentType = $MIMEHASH.Item($EXTENSION)
						}
						else
						{
							# no, serve as binary download
							$RESPONSE.ContentType = "application/octet-stream"
							$FILENAME = Split-Path -Leaf $CHECKFILE
							$RESPONSE.AddHeader("Content-Disposition", "attachment; filename=$FILENAME")
						}
						$RESPONSE.AddHeader("Last-Modified", [IO.File]::GetLastWriteTime($CHECKFILE).ToString('r'))
						$RESPONSE.AddHeader("Server", "Powershell Webserver/1.2 on ")
						$RESPONSE.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
						# mark response as already given
						$RESPONSEWRITTEN = $TRUE
					}
					catch
					{
						# just ignore. Error handling comes afterwards since not every error throws an exception
					}
					if ($Error.Count -gt 0)
					{
						# retrieve error message on error
						$RESULT += "`nError while downloading '$CHECKFILE'`n`n"
						$RESULT += $Error[0].Exception.InnerException
						$Error.Clear()
					}
				}
				else
				{
					# no file to serve found, return error
					if (!(Test-Path $CHECKDIR -PathType Container))
					{
						$RESPONSE.StatusCode = 404
						$HTMLRESPONSE = "<!doctype html><html><body>Page $($RECEIVED) not found</body></html>"
					}
				}
			}
		}

		# only send response if not already done
		if (!$RESPONSEWRITTEN)
		{
			# return HTML answer to caller
			$BUFFER = [Text.Encoding]::UTF8.GetBytes($HTMLRESPONSE)
			$RESPONSE.ContentLength64 = $BUFFER.Length
			$RESPONSE.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
			$RESPONSE.AddHeader("Server", "Powershell Webserver/1.2 on localhost")
			$RESPONSE.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
		}

		# and finish answer to client
		$RESPONSE.Close()

		# If exit was chosen, break out of loop and exit
		if ($RECEIVED -eq 'GET /exit')
		{
			# then break out of while loop
			"$(Get-Date -Format s) Stopping powershell webserver..."
			break
		}
	}
}
finally
{
	# Stop powershell webserver
	$LISTENER.Stop()
	$LISTENER.Close()
	"$(Get-Date -Format s) Powershell webserver stopped."
}