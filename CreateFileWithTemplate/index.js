// Create file in library
var request = require("request");
var adal = require("adal-node");
var fs = require("fs");

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var reqDigest = '';
    var siteUrl ='';
    var templateFileUrl='';
    var sourceFileUrl='';
	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
        tenant = req.body.tenant;
        reqDigest =req.body.ReqDigest;
        siteUrl=req.body.SiteUrl;
        templateFileUrl=req.body.TemplateFileUrl;
        sourceFileUrl=req.body.SourceFileUrl;
	}

	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync(__dirname + '/devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);

	authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
		if (err) {
			context.log('well that didn\'t work: ' + err.stack);
			context.done();
			return;
		}
		context.log(tokenResponse);

		var accesstoken = tokenResponse.accessToken;

	/*	var options = {
			method: "POST",
			async: false,
			uri: resource + "/_api/contextinfo",
			headers: {
				'Authorization': 'Bearer ' + accesstoken,
				'Accept': 'application/json; odata=verbose',
				'Content-Type': 'application/json; odata=verbose'
			}
		};

		context.log(options);
		request(options, function (error, res, body) {
			context.log(error);
			reqDigest = JSON.parse(body).d.GetContextWebInformation.FormDigestValue;
        }); */
        
	});
	
	var authContextCreate = new adal.AuthenticationContext(authorityUrl);

	authContextCreate.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
		if (err) {
			context.log('well that didn\'t work: ' + err.stack);
			context.done();
			return;
		}
		
		context.log(tokenResponse);
		var accesstoken = tokenResponse.accessToken;
        
		var options = {
			method: "POST",
			uri: siteUrl + "/_api/web/getfilebyserverrelativeurl('" + templateFileUrl + "')/copyTo('" + sourceFileUrl + "')",
			headers: {
				'Authorization': 'Bearer ' + accesstoken,
				'Accept': 'application/json; odata=verbose',
				'Content-Type': 'application/json; odata=verbose',
				'X-RequestDigest': reqDigest,
				'X-HTTP-Method': 'POST'
			}
		};

		context.log(options);
		request(options, function (error, res, body) {
			context.log(error);
			context.log(body);
			context.res = {
				body: body || ''
			};
			context.done();
		});
	});
};