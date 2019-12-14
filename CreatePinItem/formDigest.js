var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var ctx = require("./contextInfo");

exports.formDigest = function formDigest(context, tenant, resource) {
	var authorityHostUrl = 'https://login.microsoftonline.com';
	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync(__dirname + '/devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);

    var formDigest = '';
    ctx.getReqDigest(context, tenant, resource).then(accessToken => {
        console.log(accessToken);
        return new Promise
    });

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
                
                var accesstoken = tokenResponse.accessToken;

                var options = {
                    method: "POST",
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
                    context.log(body);
                    context.res = {
                        body: body || ''
                    };
                    context.done();
                });
			});
	});

}
