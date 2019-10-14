class TestResult {
	[string] $TestName;
	[bool] $SuccessDI_BOM_A = $false;
	[bool] $SuccessXML_BOM_A = $false;
	[bool] $SuccessDI_BOM_D = $false;
	[bool] $SuccessXML_BOM_D = $false;
	[string] $ExceptionDI;
	[string] $ExceptionXML;
}