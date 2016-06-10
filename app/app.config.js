(function () {
	var tenantName = 'contoso';
	
	var officeAddin = angular.module('officeAddin');
	officeAddin.constant('appId', '51578174-15c1-41dd-ac7e-a04ce892d59c');
	officeAddin.constant('sharePointUrl', 'https://alexepam-my.sharepoint.com');
	officeAddin.constant('proxyPageUrl', '/Style%20Library/Auth.aspx');
	officeAddin.constant('azureOrigin', 'https://alex-epam.azurewebsites.net');
})();