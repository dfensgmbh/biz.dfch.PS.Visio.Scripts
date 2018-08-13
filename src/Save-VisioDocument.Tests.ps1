#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Save-VisioDocument" {
	
	. "$here\$sut";
	. "$here\Open-VisioDocument.ps1";
	. "$here\Close-VisioDocument.ps1";
	
	Context "Save-VisioDocument-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullVisioDoc" {
			
			# Arrange

			# Act
			{ Save-VisioDocument -VisioDoc $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullPath" {
			
			# Arrange
			$visioDoc = New-Object -ComObject Scripting.Dictionary;

			# Act
			{ Save-VisioDocument -VisioDoc $visioDoc -Path $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
	}
	
	Context "Save-VisioDocument-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		$pathToSaveTemporaryVisioDocTo = "$here\TestVisio.vsdx";
		
		BeforeEach {
			$visioDoc = Open-VisioDocument -Path $pathToVisioDoc;
		}
		
		It "SavesVisioDocAndReturnsTrueWhenInvokingWithExistingOpenedVisioDoc" {
			
			# Arrange
			
			# Act
			$result = Save-VisioDocument $visioDoc;
			
			# Assert
			$result | Should Be $true;
		}
		
		It "SavesVisioDocAndReturnsTrueWhenInvokingWithNotYetExistingButOpenedVisioDocAndPath" {
			
			# Arrange
			
			# Act
			$result = Save-VisioDocument $visioDoc -Path $pathToSaveTemporaryVisioDocTo;
			
			# Assert
			$result | Should Be $true;
		}
		
		It "SavesVisioDocAndReturnsTrueWhenInvokingByPipingExistingOpenedVisioDoc" {
			
			# Arrange
			
			# Act
			$result = $visioDoc | Save-VisioDocument;
			
			# Assert
			$result | Should Be $true;
		}
		
		AfterEach {
			$visioDoc | Close-VisioDocument;
			
			if (Test-Path $pathToSaveTemporaryVisioDocTo) 
			{
				Remove-Item $pathToSaveTemporaryVisioDocTo;
			}
		}
	}

	Context "Save-VisioDocument-NegativeTests" {
		
		BeforeEach {
			$visio = New-Object -ComObject Visio.Application;
			$visioDocs = $visio.Documents;
			$visioDoc = $visioDocs.Add("");
		}
		
		It "ThrowsComExceptionWhenInvokingWithNotYetExistingOpenedVisioDocWithoutPath" {
			
			# Arrange
			
			# Act
						
			# Assert
			{ Save-VisioDocument $visioDoc; } | Should ThrowException 'COMException';
		}
		
		AfterEach {
			$visioDoc | Close-VisioDocument;
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
