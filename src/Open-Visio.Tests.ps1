#Requires -Modules @{ ModuleName = 'biz.dfch.PS.Pester.Assertions'; ModuleVersion = '1.1.1' }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");

Describe "Open-Visio" {
	
	. "$here\$sut";
	# . "$here\Close-Visio.ps1";
	
	Context "Open-Visio-ValidationTests" {
		It "Warmup" -Test {
			
			# Arrange
			
			# Act
			
			# Assert
			$true | Should Be $true;
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNullPath" {
			
			# Arrange

			# Act
			{ Open-Visio -Path $null; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
		
		It "ThrowsParameterBindingValidationExceptionWhenInvokingWithNotExistingPath" {
			
			# Arrange

			# Act
			{ Open-Visio -Path "C:\arbitrary"; } | Should ThrowException 'ParameterBindingValidationException';

			# Assert
		}
	}
	
	Context "Open-Visio-PositiveTests" {
		
		$pathToVisioDoc = "$here\SampleVisio.vsdx";
		
		BeforeEach {
			$visioDoc = Open-Visio -Path $pathToVisioDoc;
		}
		
		It "OpensAndReturnsVisioDocWhenInvokingWithValidPathToExistingVisioDoc" {
			
			# Arrange
			
			# Act
			
			# Assert
			$visioDoc | Should Not Be $null;
		}
		
		AfterEach {
			# $null = Close-Visio $visioDoc;
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
