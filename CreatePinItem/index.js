// CreateListItem
var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var ctx = require("./contextInfo")

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var reqDigest = '';
	var documentLibrary = '';
	var documentLibraryUrl = '';
	var siteUrl = '';
	var title = "New";
	var pinnedType = '';
	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		reqDigest = req.body.ReqDigest;
		documentLibrary = req.body.DocumentLibrary;
		documentLibraryUrl = req.body.DocumentLibraryUrl;
		siteUrl = req.body.SiteUrl;
		pinnedType = req.body.PinnedType;
	}

	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync(__dirname + '/devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];
	var accessToken = '';
	
	ctx.getReqDigest(context, tenant, resource).then(accessToken => {
		console.log(accessToken);
		var itemProperties = JSON.stringify({
			__metadata: {
				type: "SP.Data.DocsNodePinnedLocationsListItem"
			},
			Title: title,
			PinnedType: pinnedType,
			DocumentLibrary: documentLibrary,
			DocumentLibraryURL: {
				'__metadata': {
					'type': 'SP.FieldUrlValue'
				},
				'Description': 'Library Url',
				'Url': documentLibraryUrl
			},
			SiteURL: {
				'__metadata': {
					'type': 'SP.FieldUrlValue'
				},
				'Description': 'Site Url',
				'Url': siteUrl
			}
		});


		var options = {
			method: "POST",
			async: false,
			uri: resource + "/_api/web/lists/getbytitle('DocsNodePinnedLocations')/items",
			body: itemProperties,
			headers: {
				'Authorization': 'Bearer ' + accessToken,
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