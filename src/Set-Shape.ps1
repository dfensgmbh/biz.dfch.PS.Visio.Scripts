#Requires -Modules @{ ModuleName = 'biz.dfch.PS.System.Logging'; ModuleVersion = '1.4.1' }

Function Set-Shape {
<#
.SYNOPSIS

Updates a given shape.

.DESCRIPTION

Updates a given shape.

Shape has to be provided as input by either positional or named parameter. Properties like position, size and text can be provided as input by named parameters.

.EXAMPLE

$shape = Set-Shape $shape -PositionX 1.0 -PositionY 1.0 -Height 10.0 -Width 20.0 -Text "arbitraryText";

Shape is passed as a positional parameter to the Cmdlet. Position, size and text are passed as named parameters to the Cmdlet.

.LINK

GitHub Repository: https://github.com/dfensgmbh/biz.dfch.PS.Visio.Scripts

#>

[CmdletBinding(
    SupportsShouldProcess = $true
	,
    ConfirmImpact = "Medium"
)]
PARAM
(
	[ValidateNotNull()]
	[Parameter(Mandatory = $true, Position = 0)]
	$Shape
	,
	[ValidateRange([double]::MinValue, [double]::MaxValue)]
	[Parameter(Mandatory = $false)]
	[double] $PositionX
	,
	[ValidateRange([double]::MinValue, [double]::MaxValue)]
	[Parameter(Mandatory = $false)]
	[double] $PositionY
	,
	[ValidateRange(0.1, [double]::MaxValue)]
	[Parameter(Mandatory = $false)]
	[double] $Height
	,
	[ValidateRange(0.1, [double]::MaxValue)]
	[Parameter(Mandatory = $false)]
	[double] $Width
	,
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $false)]
	[string] $Text
	,
	[ValidateSet('Rectangle')]
	[Parameter(Mandatory = $false)]
	[string] $ShapeType = 'Rectangle'
	,
	[Parameter(Mandatory = $false)]
	[guid] $EaGuid
)

BEGIN
{
	trap { Log-Exception $_; break; }

	$OutputParameter = $null;
}

PROCESS
{
	trap { Log-Exception $_; break; }
	
	if ($ShapeType -eq 'Rectangle')
	{
		$pinX = $Shape.Cells("PinX").ResultIU;
		$pinY = $Shape.Cells("PinY").ResultIU;
		$h = $Shape.Cells("Height").ResultIU;
		$w = $Shape.Cells("Width").ResultIU;

		if ($PositionX -eq $null)
		{
			$PositionX = $pinX - ($w / 2.0);
		}
		if ($PositionY -eq $null)
		{
			$PositionY = $pinY - ($h / 2.0);
		}
		if ($Height -eq $null)
		{
			$Height = $h;
		}
		if ($Width -eq $null)
		{
			$Width = $w;
		}
		
		$Shape.Cells("Height") = $Height;
		$Shape.Cells("Width") = $Width;
		$Shape.Cells("PinX") = $PositionX + ($Width / 2.0);
		$Shape.Cells("PinY") = $PositionY + ($Height / 2.0);
	}
	
	$Shape.Text = $Text;
	
	if ($null -ne $EaGuid)
	{
		$Shape.Data1 = $EaGuid.ToString();
	}
	
	$OutputParameter = $Shape;
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
