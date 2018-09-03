#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Add-ShapeToPage" {
	
	. "$here\$sut";
	. "$here\Open-VisioDocument.ps1";
	. "$here\Close-VisioDocument.ps1";
	
	Context "Add-ShapeToPage-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullVisioDoc" {
			
			# Arrange

			# Act
			{ Add-ShapeToPage -VisioDoc $null -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullPage" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -Page $null -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullPageName" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -PageName $null -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithEmptyPageName" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -PageName "" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithPositionXLessZero" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX -0.1 -PositionY 1.0 -Height 1.0 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithPositionYLessZero" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY -0.1 -Height 1.0 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithHeightLess0.5" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 0.4 -Width 1.0; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithWidthLess0.5" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 0.4; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullText" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0 -Text $null; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
	}
	
	Context "Add-ShapeToPage-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		
		BeforeEach {
			$visioDoc = Open-VisioDocument -Path $pathToVisioDoc;
		}
		
		It "AddsShapeToPageAndReturnsShapeWhenInvokingWithValidMandatoryParameters" {
			
			# Arrange
			
			# Act
			$result = Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0;
			
			# Assert
			$result | Should Not Be $null;
		}
		
		It "AddsShapeToPageAndReturnsShapeWithTextSetWhenInvokingWithValidMandatoryParametersAndOptionalTextParameter" {
			
			# Arrange
			$text = "arbitrary";
			
			# Act
			$result = Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0 -Text $text;
			
			# Assert
			$result | Should Not Be $null;
			$result.Text | Should Be $text;
		}
		
		It "AddsShapeToPageAndReturnsShapeWithData1SetWhenInvokingWithValidMandatoryParametersAndOptionalEaGuidParameter" {
			
			# Arrange
			$eaGuid = [guid]::newGuid();
			
			# Act
			$result = Add-ShapeToPage -VisioDoc $visioDoc -PageName "Page-1" -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0 -EaGuid $eaGuid;
			
			# Assert
			$result | Should Not Be $null;
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
