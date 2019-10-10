class TestResult {
	[string] $TestName;
	[bool] $SuccessDI = $false;
	[bool] $SuccessXML = $false;
	[string] $ExceptionDI;
	[string] $ExceptionXML;
}