#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Get-Page" {
	
	. "$here\$sut";
	. "$here\Open-VisioDocument.ps1";
	. "$here\Close-VisioDocument.ps1";
	
	Context "Get-Page-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullVisioDoc" {
			
			# Arrange

			# Act
			{ Get-Page -VisioDoc $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullName" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Get-Page -VisioDoc $visioDoc -Name $null; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithEmptyName" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;
			
			# Act
			{ Get-Page -VisioDoc $visioDoc -Name ""; } | Should ThrowException 'ParameterBindingValidationException';
			
			#Assert
		}
	}
	
	Context "Get-Page-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		
		BeforeEach {
			$visioDoc = Open-VisioDocument -Path $pathToVisioDoc;
		}
		
		It "RetrievesAndReturnsListOfAvailablePagesOfVisioDocumentWhenInvokingWithValidOpenedVisioDocument" {
			
			# Arrange
			
			# Act
			$result = Get-Page $visioDoc;
			
			# Assert
			$result | Should Not Be $null;
			$result.Count | Should Be 2;
		}

		It "RetrievesPageByNameAndReturnsSpecifiedPageOfVisioDocumentWhenInvokingWithValidOpenedVisioDocumentAndExistingPageName" {
			
			# Arrange
			
			# Act
			$result = Get-Page $visioDoc -Name "Page-2";
			
			# Assert
			$result | Should Not Be $null;
			$result.Name | Should Be "Page-2";
		}
		
		It "ReturnsNullWhenInvokingWithValidOpenedVisioDocumentAndNotExistingPageName" {
			
			# Arrange
			
			# Act
			$result = Get-Page $visioDoc -Name "Page-42";
			
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
