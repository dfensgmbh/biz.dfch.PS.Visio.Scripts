#Requires -Modules @{ ModuleName = 'biz.dfch.PS.System.Logging'; ModuleVersion = '1.4.1' }

Function Save-VisioDocument {
<#
.SYNOPSIS

Saves the specified Visio document.

.DESCRIPTION

Saves the specified Visio document.

The Visio document to has to be provided as input by either pipe, positional parameter or named parameter.

.EXAMPLE

$result = Save-VisioDocument $visioDoc

Visio document is passed as a positional parameter to the Cmdlet.

.EXAMPLE

$result = Save-VisioDocument $visioDoc -Path "C:\PATH\TO\MyVisio.vsdx"

Visio document and path are passed as a positional parameters to the Cmdlet.

.EXAMPLE

$result = $visioDoc | Save-VisioDocument

Visio document is piped to the Cmdlet.

.LINK

GitHub Repository: https://github.com/dfensgmbh/biz.dfch.PS.Visio.Scripts

#>

[CmdletBinding(
    SupportsShouldProcess = $true
	,
    ConfirmImpact = "High"
)]
PARAM
(
	[ValidateNotNull()]
	[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
	$VisioDoc
	,
	[ValidateScript( { Test-Path($_); } )]
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $false)]
	[string] $Path
)

BEGIN
{
	trap { Log-Exception $_; break; }
}

PROCESS
{
	trap { Log-Exception $_; break; }
	
	$OutputParameter = $false;
	
	if ($PSBoundParameters.ContainsKey('Path'))
	{
		$null = $VisioDoc.SaveAs($Path);
	}
	else 
	{
		$null = $VisioDoc.Save();
	}
	
	$OutputParameter = $true;
}

END
{
	trap { Log-Exception $_; break; }
	
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
