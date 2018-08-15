#Requires -Modules @{ ModuleName = 'biz.dfch.PS.System.Logging'; ModuleVersion = '1.4.1' }

Function Add-ShapeToPage {
<#
.SYNOPSIS

Adds a shape to the specified page of the given Microsoft Visio document at the given position.

.DESCRIPTION

Adds a shape to the specified page of the given Microsoft Visio document at the given position.

Visio document, page, position and size have to be provided as input by either positional or named parameters.

.EXAMPLE

Add-Shape $visioDoc $page 1 1 10 20

Visio document, page, position and size are passed as a positional parameters to the Cmdlet.

.EXAMPLE

Add-Shape $visioDoc "page-1" 1.0 1.0 2.0 3.0

Visio document, page name, position and size are passed as a positional parameters to the Cmdlet.

.LINK

GitHub Repository: https://github.com/dfensgmbh/biz.dfch.PS.Visio.Scripts

#>

[CmdletBinding(
    SupportsShouldProcess = $true
	,
    ConfirmImpact = "Low"
	,
	DefaultParameterSetName = 'pageObject'
)]
PARAM
(
	[ValidateNotNull()]
	[Parameter(Mandatory = $true, Position = 0)]
	$VisioDoc
	,
	[ValidateNotNull()]
	[Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'pageObject')]
	$Page
	,
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'pageName')]
	[string] $PageName
	,
	[ValidateRange(0, [double]::MaxValue)]
	[Parameter(Mandatory = $true, Position = 2)]
	[double] $PositionX
	,
	[ValidateRange(0, [double]::MaxValue)]
	[Parameter(Mandatory = $true, Position = 3)]
	[double] $PositionY
	,
	[ValidateRange(0.5, [double]::MaxValue)]
	[Parameter(Mandatory = $true, Position = 4)]
	[double] $Height
	,
	[ValidateRange(0.5, [double]::MaxValue)]
	[Parameter(Mandatory = $true, Position = 5)]
	[double] $Width
	,
	[ValidateNotNull()]
	[Parameter(Mandatory = $false)]
	[string] $ShapeText = ''
	,
	[ValidateSet('Rectangle')]
	[Parameter(Mandatory = $false)]
	[string] $ShapeType = 'Rectangle'
	,
	[Parameter(Mandatory = $false)]
	[guid] $EaGuid
	# ,
	# [ValidateNotNullOrEmpty]
	# [Parameter(Mandatory = $false)]
	# [string] $LayerName
)

BEGIN
{
	trap { Log-Exception $_; break; }

	$OutputParameter = $null;
}

PROCESS
{
	trap { Log-Exception $_; break; }
	
	if($PSCmdlet.ParameterSetName -eq 'pageName') 
	{
		$Page = $VisioDoc.Pages |? Name -match $PageName;
		Contract-Assert($Page);
	}
	
	if ($ShapeType -eq 'Rectangle')
	{
		$shape = $Page.DrawRectangle($PositionX, $PositionY, $PositionX + $Width, $PositionY + $Height);
	}
	
	$shape.Text = $ShapeText;
	
	if ($null -ne $EaGuid)
	{
		$shape.Data1 = $EaGuid.ToString();
	}
	
	$OutputParameter = $shape;
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
