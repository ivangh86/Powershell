
Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -eq $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}




$AllApplications = Get-BrokerApplication

	
ForEach($Application in $AllApplications)
{



	$xTags = @()
	ForEach($Tag in $Application.Tags)
	{
		$xTags += "$($Tag)"
	}

	$xVisibility = @()
	If($Application.UserFilterEnabled)
	{
		$cnt = -1
		ForEach($tmp in $Application.AssociatedUserFullNames)
		{
			$cnt++
			$xVisibility += "$($tmp) ($($Application.AssociatedUserNames[$cnt]))"
		}
		
	}
	Else
	{
		$xVisibility = {Users inherited from Delivery Group}
	}
	
	$DeliveryGroups = @()
	If($Application.AssociatedDesktopGroupUids.Count -gt 1)
	{
		$cnt = -1
		ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
		{
			$cnt++
			$Results = Get-BrokerDesktopGroup  -Uid $DGUid
			If($? -and $Null -ne $Results)
			{
				$DeliveryGroups += "$($Results.Name) Priority: $($Application.AssociatedDesktopGroupPriorities[$cnt])"
			}
		}
	}
	Else
	{
		ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
		{
			$Results = Get-BrokerDesktopGroup  -Uid $DGUid
			If($? -and $Null -ne $Results)
			{
				$DeliveryGroups += $Results.Name
			}
		}
	}
	
	$RedirectedFileTypes = @()
	#V2.10 22-jan=2018 change from xdparams1 to xdparams2 to add maxrecordcount
	$Results = Get-BrokerConfiguredFTA -ApplicationUid $Application.Uid -MaxRecordCount 1500
	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$RedirectedFileTypes += $Result.ExtensionName
		}
	}
	

	
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Name (for administrator)"; Value = $Application.Name; }
		$ScriptInformation += @{Data = "Name (for user)"; Value = $Application.PublishedName; }
		$ScriptInformation += @{Data = "Description"; Value = $Application.Description; }
		$ScriptInformation += @{Data = "Delivery Group"; Value = $DeliveryGroups[0]; }
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $Group; }
			}
		}
		$ScriptInformation += @{Data = "Folder (for administrators)"; Value = $Application.AdminFolderName; }
		$ScriptInformation += @{Data = "Folder (for user)"; Value = $Application.ClientFolder; }
		$ScriptInformation += @{Data = "Visibility"; Value = $xVisibility[0]; }
		$cnt = -1
		ForEach($tmp in $xVisibility)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $xVisibility[$cnt]; }
			}
		}
		$ScriptInformation += @{Data = "Application Path"; Value = $Application.CommandLineExecutable; }
		$ScriptInformation += @{Data = "Command line arguments"; Value = $Application.CommandLineArguments; }
		$ScriptInformation += @{Data = "Working directory"; Value = $Application.WorkingDirectory; }
		If($Null -eq $RedirectedFileTypes)
		{
			$ScriptInformation += @{Data = "Redirected file types"; Value = ""; }
		}
		Else
		{
			$tmp1 = ""
			ForEach($tmp in $RedirectedFileTypes)
			{
				$tmp1 += "$($tmp); "
			}
			$ScriptInformation += @{Data = "Redirected file types"; Value = $tmp1; }
			$tmp1 = $Null
		}
		$ScriptInformation += @{Data = "Tags"; Value = $xTags[0]; }
		$cnt = -1
		ForEach($tmp in $xTags)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		If($Application.Visible -eq $False)
		{
			$ScriptInformation += @{Data = "Hidden"; Value = "Application is hidden"; }
		}
		
	<#If((Get-BrokerServiceAddedCapability  ) -contains "ApplicationUsageLimits")
		{
			
			$tmp = ""
			If($Application.MaxTotalInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxTotalInstances.ToString()
			}
			$ScriptInformation += @{Data = "Maximum concurrent instances"; Value = $tmp; }
			
			$tmp = ""
			If($Application.MaxPerUserInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxPerUserInstances.ToString()
			}
			$ScriptInformation += @{Data = "Maximum instances per user"; Value = $tmp; }
		}#>

}	
		$Table = AddWordTable -Hashtable $ScriptInformation 


		