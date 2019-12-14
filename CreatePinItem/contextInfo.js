var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


exports.getReqDigest = function getReqDigest(context, tenant, resource) {
	var authorityHostUrl = 'https://login.microsoftonline.com';
	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync(__dirname + '/devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);

	var accessToken = '';

	// return 1;

	return new Promise((resolve, reject) => {
		return authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint,
			function (err, tokenResponse) {
				if (err) {
					context.log('well that didn\'t work: ' + err.stack);
					context.done();
					return;
				}
				resolve(tokenResponse.accessToken);

				// return tokenResponse.accessToken;
			});
	});

}
