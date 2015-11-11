(function () {
	var tenantName = 'contoso';
	
	var officeAddin = angular.module('officeAddin');
	officeAddin.constant('appId', '54306a7c-fce1-4899-8f9d-8b21c9937c69');
	officeAddin.constant('sharePointUrl', 'https://fancycoder.sharepoint.com');
	officeAddin.constant('proxyHackUrl', 'https://fancycoder.sharepoint.com/sites/Dev/SitePages/DelveClassifierProxy.aspx');
})();