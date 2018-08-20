#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Remove-Shape" {
	
	. "$here\$sut";
	. "$here\Open-VisioDocument.ps1";
	. "$here\Get-Page.ps1";
	. "$here\Get-Shape.ps1";
	. "$here\Add-ShapeToPage.ps1";
	. "$here\Close-VisioDocument.ps1";
	
	Context "Remove-Shape-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullShape" {
			
			# Arrange

			# Act
			{ Remove-Shape $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
	}
	
	Context "Remove-Shape-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		$pageName = "Page-1";
		
		BeforeEach {
			$visioDoc = Open-VisioDocument -Path $pathToVisioDoc;
			$page = Get-Page $visioDoc -Name $pageName;
			$shape = Add-ShapeToPage -VisioDoc $visioDoc -Page $page -PositionX 1.0 -PositionY 1.0 -Height 1.0 -Width 1.0;
		}
		
		It "RemovesShapeFromItsPageAndReturnsTrueWhenInvokingWithValidShape" {
			
			# Arrange
			
			# Act
			$result = Remove-Shape $shape;
			
			# Assert
			$result | Should Be $true;
			
			$shapes = Get-Shape $page;
			$shapes.Count | Should Be 0;
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
