using module .\TestResult.psm1;
class Result {
	[System.Collections.Generic.List[TestResult]] $TestResults;
	
	Result() {
		$this.TestResults = New-Object 'System.Collections.Generic.List[TestResult]';
	}

	AddTestResult([string]$TestName, [bool]$SuccessDI, [bool]$SuccessXML, [string]$ExceptionDI, [string]$ExceptionXML){
		$TestResult = New-Object 'TestResult';
		$TestResult.TestName = $TestName;
		$TestResult.SuccessDI = $SuccessDI;
		$TestResult.SuccessXML = $SuccessXML;
		$TestResult.ExceptionDI = $ExceptionDI;
		$TestResult.ExceptionXML = $ExceptionXML;
		$this.TestResults.Add($TestResult);
	}
}