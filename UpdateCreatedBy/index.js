// CreateListItem
var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var ctx = require("./updateAuthor");

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var reqDigest = '';
    var documentLibrary ='';
    //var documentLibraryUrl='';
    var siteUrl ='';
    var itemID;
    var userID;
    var listItemEntityType = "SP.Data.DocsNodePinnedLocationsListItem";
    if(req.body && req.body.ListItemEntityType )
    {
        listItemEntityType=req.body.ListItemEntityType;
    }
	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
        tenant = req.body.tenant;
        reqDigest =req.body.ReqDigest;
        documentLibrary=req.body.DocumentLibrary;
        itemID=req.body.ItemID;
        siteUrl=req.body.SiteUrl;
        userID=req.body.UserID;
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

		var getSuccessMsg= "";

		ctx.updateAuthor(context, tenant, resource,userID).then(result => {
			getSuccessMsg=result.success;
			context.res = {
				body: getSuccessMsg || ''
			};
			context.done();
		});

	});
};