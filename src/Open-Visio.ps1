#Requires -Modules @{ ModuleName = 'biz.dfch.PS.System.Logging'; ModuleVersion = '1.4.1' }

Function Open-Visio {
<#
.SYNOPSIS

Opens the specified Microsoft Visio document.

.DESCRIPTION

Opens the specified Microsoft Visio document.

The path to the Microsoft Visio document has to be provided as input by either positional or named parameter of type string.

.EXAMPLE

$visio = Open-Visio C:\PATH\TO\MyVisio.vsdx

Path to the Microsoft Visio document is passed as a positional parameter to the Cmdlet.

.LINK

GitHub Repository: https://github.com/dfensgmbh/biz.dfch.PS.Visio.Scripts

#>

[CmdletBinding(
    SupportsShouldProcess = $true
	,
    ConfirmImpact = "Low"
)]
PARAM
(
	[ValidateScript( { Test-Path($_); } )]
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $true, Position = 0)]
	[string] $Path
)

BEGIN
{
	trap { Log-Exception $_; break; }

	$OutputParameter = $null;
	
	# instantiating the COM object
	# NOTE: this will not work reliably
	# $ea = [Visio.Application]::new();
	$visio = New-Object -ComObject Visio.Application;
}

PROCESS
{
	trap { Log-Exception $_; break; }
	
	$result = $visio.Documents.Open($Path);
	
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
