#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Set-Shape" {
	
	. "$here\$sut";
	. "$here\Open-VisioDocument.ps1";
	. "$here\Add-ShapeToPage.ps1"
	. "$here\Close-VisioDocument.ps1";
	
	Context "Set-Shape-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullShape" {
			
			# Arrange

			# Act
			{ Set-Shape $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithHeightLess0.1" {
			
			# Arrange
			$shape = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Set-Shape -Shape $shape -PositionX 1.0 -PositionY 1.0 -Height 0.0 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithWidthLess0.1" {
			
			# Arrange
			$shape = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Set-Shape -Shape $shape -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 0.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullText" {
			
			# Arrange
			$shape = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Set-Shape -Shape $shape -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0 -Text $null; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithEmptyText" {
			
			# Arrange
			$shape = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Set-Shape -Shape $shape -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0 -Text ""; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
	}
	
	Context "Set-Shape-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		
		BeforeEach {
			$visioDoc = Open-VisioDocument -Path $pathToVisioDoc;
			$shape = Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 2.0;
		}
		
		It "UpdatesShapeAndReturnsShapeWhenInvokingWithValidPosition" {
			
			# Arrange
			
			# Act
			$result = Set-Shape $shape -PositionX 2.0 -PositionY 3.0;
			
			# Assert
			$result | Should Not Be $null;
			$result.Cells("PinX").ResultIU | Should Be 3.0;
			$result.Cells("PinY").ResultIU | Should Be 3.5;
			$result.Cells("Height").ResultIU | Should Be 1.0;
			$result.Cells("Width").ResultIU | Should Be 2.0;
		}
		
		It "UpdatesShapeAndReturnsShapeWhenInvokingWithValidSize" {
			
			# Arrange
			
			# Act
			$result = Set-Shape $shape -Height 2.0 -Width 3.0;
			
			# Assert
			$result | Should Not Be $null;
			$result.Cells("PinX").ResultIU | Should Be 2.5;
			$result.Cells("PinY").ResultIU | Should Be 2.0;
			$result.Cells("Height").ResultIU | Should Be 2.0;
			$result.Cells("Width").ResultIU | Should Be 3.0;
		}
		
		It "UpdatesShapeAndReturnsShapeWhenInvokingWithValidSizeAndPosition" {
			
			# Arrange
			
			# Act
			$result = Set-Shape $shape -PositionX 2.0 -PositionY 3.0 -Height 4.0 -Width 5.0;
			
			# Assert
			$result | Should Not Be $null;
			$result.Cells("PinX").ResultIU | Should Be 4.5;
			$result.Cells("PinY").ResultIU | Should Be 5.0;
			$result.Cells("Height").ResultIU | Should Be 4.0;
			$result.Cells("Width").ResultIU | Should Be 5.0;
		}
		
		It "UpdatesShapeAndReturnsShapeWhenInvokingWithValidText" {
			
			# Arrange
			$text = "arbitrary";
			
			# Act
			$result = Set-Shape $shape -Text $text;
			
			# Assert
			$result | Should Not Be $null;
			$result.Cells("PinX").ResultIU | Should Be 2.0;
			$result.Cells("PinY").ResultIU | Should Be 1.5;
			$result.Cells("Height").ResultIU | Should Be 1.0;
			$result.Cells("Width").ResultIU | Should Be 2.0;
			$result.Text | Should Be $text;
		}
		
		It "UpdatesShapeAndReturnsShapeWhenInvokingWithValidEaGuid" {
			
			# Arrange
			$eaGuid = [guid]::newGuid();
			
			# Act
			$result = Set-Shape $shape -EaGuid $eaGuid;
			
			# Assert
			$result | Should Not Be $null;
			$result.Cells("PinX").ResultIU | Should Be 2.0;
			$result.Cells("PinY").ResultIU | Should Be 1.5;
			$result.Cells("Height").ResultIU | Should Be 1.0;
			$result.Cells("Width").ResultIU | Should Be 2.0;
			$result.Data1 | Should Be $eaGuid.ToString();
		}
		
		AfterEach {
			$null = Close-VisioDocument $visioDoc;
		}
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
