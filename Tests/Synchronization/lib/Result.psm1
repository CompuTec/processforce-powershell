using module .\TestResult.psm1;
class Result {
	[System.Collections.Generic.List[TestResult]] $TestResults;
	
	Result() {
		$this.TestResults = New-Object 'System.Collections.Generic.List[TestResult]';
	}

	AddTestResult([string]$TestName, [bool]$SuccessDI_BOMA, [bool]$SuccessXML_BOMA, [bool]$SuccessDI_BOMD, [bool]$SuccessXML_BOMD, [string]$ExceptionDI, [string]$ExceptionXML){
		$TestResult = New-Object 'TestResult';
		$TestResult.TestName = $TestName;
		$TestResult.SuccessDI_BOM_A = $SuccessDI_BOMA;
		$TestResult.SuccessDI_BOM_D = $SuccessDI_BOMD;
		$TestResult.SuccessXML_BOM_A = $SuccessXML_BOMA;
		$TestResult.SuccessXML_BOM_D = $SuccessXML_BOMD;
		$TestResult.ExceptionDI = $ExceptionDI;
		$TestResult.ExceptionXML = $ExceptionXML;
		$this.TestResults.Add($TestResult);
	}
}