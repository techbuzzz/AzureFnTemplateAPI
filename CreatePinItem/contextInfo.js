var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


exports.getReqDigest = async function getReqDigest(tenant,resource)
{
    var authorityHostUrl = 'https://login.microsoftonline.com';
    var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync(__dirname + '/devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);
    
   return await authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
		if (err) {
			context.log('well that didn\'t work: ' + err.stack);
			context.done();
			return;
		}
		context.log(tokenResponse);

		return tokenResponse.accessToken;
});
}