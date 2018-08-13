#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Close-VisioDocument" {
	
	. "$here\$sut";
	. "$here\Open-VisioDocument.ps1";
	
	Context "Close-VisioDocument-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullVisioDoc" {
			
			# Arrange

			# Act
			{ Close-VisioDocument -VisioDoc $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
	}
	
	Context "Close-VisioDocument-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		
		BeforeEach {
			$visioDoc = Open-VisioDocument -Path $pathToVisioDoc;
		}
		
		It "ClosesVisioDocAndApplicationAndReturnsTrueWhenInvokingWithValidOpenedVisioDoc" {
			
			# Arrange
			
			# Act
			$result = Close-VisioDocument $visioDoc;
			
			# Assert
			$result | Should Be $true;
		}
		
		It "ClosesVisioDocAndApplicationAndReturnsTrueWhenInvokingByPipingValidOpenedVisioDoc" {
			
			# Arrange
			
			# Act
			$result = $visioDoc | Close-VisioDocument;
			
			# Assert
			$result | Should Be $true;
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
