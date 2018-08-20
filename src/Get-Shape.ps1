#Requires -Modules @{ ModuleName = 'biz.dfch.PS.System.Logging'; ModuleVersion = '1.4.1' }

Function Get-Shape {
<#
.SYNOPSIS

Retrieves one or more shapes from the specified page.

.DESCRIPTION

Retrieves one or more shapes from the specified page.

The page has to be provided as input by either positional or named parameter.

.EXAMPLE

$shapes = Get-Shape $page

Page is passed as a positional parameter to the Cmdlet.

.LINK

GitHub Repository: https://github.com/dfensgmbh/biz.dfch.PS.Visio.Scripts

#>

[CmdletBinding(
    SupportsShouldProcess = $true
	,
    ConfirmImpact = "Low"
	,
	DefaultParameterSetName = 'list'
)]
PARAM
(
	[ValidateNotNull()]
	[Parameter(Mandatory = $true, Position = 0)]
	$Page
	,
	# Full name or part of it, for the shape you want to search - this is not case sensitive
	[Parameter(Mandatory = $false, ParameterSetName = 'searchByName')]
	[ValidateNotNullOrEmpty()]
	[String] $Name
	,
	[Parameter(Mandatory = $false, ParameterSetName = 'searchByEaGuid')]
	[guid] $EaGuid
	,
	# Lists all available shapes of the specified page of the given Microsoft Visio document
	[Parameter(Mandatory = $false, ParameterSetName = 'list')]
	[Switch] $ListAvailable
)

BEGIN
{
	trap { Log-Exception $_; break; }
}

PROCESS
{
	trap { Log-Exception $_; break; }
	
	if($PSCmdlet.ParameterSetName -eq 'list') 
	{
		$result = $Page.Shapes;
	}
	else
	{
		if ($PSCmdlet.ParameterSetName -eq 'SearchByName')
		{
			$result = $Page.Shapes |? Name -match $Name;
		}
		if ($PSCmdlet.ParameterSetName -eq 'searchByEaGuid')
		{
			$result = $Page.Shapes |? Data1 -match $EaGuid;
		}
	}
	
	$OutputParameter = $result;
}

END
{	
	return $OutputParameter;
}
}

#
# Copyright 2018 d-fens GmbH
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
