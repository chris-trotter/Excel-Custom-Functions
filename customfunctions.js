const endPoint = 'https://excel-ch-api.azurewebsites.net';

function httpRequest(url, callback) {
		var xhr = new XMLHttpRequest();

		// The last parameter must be set to true to make an asynchronous request
		xhr.open('GET', url, true);
		
		xhr.setRequestHeader('Content-type', 'application/json');
		xhr.setRequestHeader('Accept', 'application/json');

		xhr.onload = function() {
			if (xhr.status >= 200 && xhr.status < 300) {
				callback(xhr.response);
			} else {
				callback(xhr.status + ' error');
			}
		};
		xhr.send();
}

function NAME(company_number) {
	// waits 1 second before returning the result
	
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/company-name/' + company_number, function(response){
			resolve(response);
		});
	});
}

function REGADDRESS(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/company-address/' + company_number, function(response){
			resolve(response);
		});
	});
}

function LIQUIDATED(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/has-been-liquidated/' + company_number, function(response){
			resolve(response);
		});
	});
}

function SICCODES(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/sic-code/' + company_number, function(response){
			resolve(response);
		});
	});
}

function ACCOUNTINGDATE(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/accounting-reference-date/' + company_number, function(response){
			resolve(response);
		});
	});
}

function OVERDUESTATUS(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/overdue-accounts/' + company_number, function(response){
			resolve(response);
		});
	});
}

function INCORPORATIONDATE(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/incorporation-date/' + company_number, function(response){
			resolve(response);
		});
	});
}

function COMPANYSTATUS(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/company-status/' + company_number, function(response){
			resolve(response);
		});
	});
}

function DIRECTORS(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/company-active-directors/' + company_number, function(response){
			resolve(response);
		});
	});
}

function SIGCONTROL(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/company-significant-control/' + company_number, function(response){
			resolve(response);
		});
	});
}

function LASTMEMBERSLIST(company_number) {
	return new OfficeExtension.Promise(function(resolve) {
		httpRequest(endPoint + '/last-full-members-list/' + company_number, function(response){
			resolve(response);
		});
	});
}
