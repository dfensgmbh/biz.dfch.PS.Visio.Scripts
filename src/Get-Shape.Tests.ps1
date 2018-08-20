#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Get-Shape" {
	
	. "$here\$sut";
	. "$here\Open-VisioDocument.ps1";
	. "$here\Get-Page.ps1";
	. "$here\Add-ShapeToPage.ps1";
	. "$here\Close-VisioDocument.ps1";
	
	Context "Get-Shape-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullPage" {
			
			# Arrange

			# Act
			{ Get-Shape -Page $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullName" {
			
			# Arrange
			$page = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Get-Shape -Page $page -Name $null; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithEmptyName" {
			
			# Arrange
			$page = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Get-Shape -Page $page -Name ""; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
	}
	
	Context "Get-Shape-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		$pageName = "Page-1";
		$eaGuid = [guid]::newGuid();
		
		BeforeEach {
			$visioDoc = Open-VisioDocument -Path $pathToVisioDoc;
			$page = Get-Page $visioDoc -Name $pageName;
			Add-ShapeToPage -VisioDoc $visioDoc -Page $page -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0 -EaGuid $eaGuid;
			$shape = Add-ShapeToPage -VisioDoc $visioDoc -Page $page -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0;
		}
		
		It "RetrievesAndReturnsListOfAvailableShapesOfVisioPageWhenInvokingWithValidVisioPage" {
			
			# Arrange
			
			# Act
			$result = Get-Shape $page;
			
			# Assert
			$result | Should Not Be $null;
			$result.Count | Should Be 2;
		}

		It "RetrievesShapeByNameAndReturnsSpecifiedShapeOfVisioPageWhenInvokingWithValidVisioPage" {
			
			# Arrange
			
			# Act
			$result = Get-Shape $page -Name $shape.Name;
			
			# Assert
			$result | Should Not Be $null;
			$result.Name | Should Be $shape.Name;
		}
		
		It "RetrievesShapeByEaGuidAndReturnsSpecifiedShapeOfVisioPageWhenInvokingWithValidVisioPage" {
			
			# Arrange
			
			# Act
			$result = Get-Shape $page -EaGuid $eaGuid;
			
			# Assert
			$result | Should Not Be $null;
			$result.Data1 | Should Be $eaGuid;
		}
		
		It "ReturnsNullWhenInvokingWithValidVisioPageAndNotExistingShapeName" {
			
			# Arrange
			
			# Act
			$result = Get-Shape $page -Name "arbitrary";
			
			# Assert
			$result | Should Be $null;
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
