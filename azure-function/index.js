
var sp_pnp_js_1 = require("sp-pnp-js");
var HardwareRequestListTitle = "Hardware Requests";
var HardwareRequestService = (function () {
    function HardwareRequestService(web) {
        this.web = web;
    }
    HardwareRequestService.createAppOnly = function (siteUrl, clientId, clientSecret) {
        sp_pnp_js_1.default.setup({
            fetchClientFactory: function () {
                return new sp_pnp_js_1.NodeFetchClient(siteUrl, clientId, clientSecret);
            }
        });
        return new HardwareRequestService(new sp_pnp_js_1.Web(siteUrl));
    };
    HardwareRequestService.createForCurrentWeb = function () {
        return new HardwareRequestService(sp_pnp_js_1.default.sp.web);
    };
    HardwareRequestService.prototype.submitRequest = function (request) {
        console.log(request);
        return this.web.lists.getByTitle(HardwareRequestListTitle).items.add({
            Title: request.title || ("Request " + new Date().toUTCString()),
            HW_HardwareType: request.type,
            HW_Remark: request.remark,
            HW_Quantity: request.quantity
        });
    };
    return HardwareRequestService;
}());


module.exports = function (context, req) {

    context.log('Sending app-only request to Hardware Requests list.');

    var siteUrl = "https://<yourtenant>.sharepoint.com/sites/<yoursite>";
    var clientId = "<your client id>;
    var clientSecret = "<your client secret>";

    // Instantiate the service in app-only
    var service = HardwareRequestService.createAppOnly(siteUrl, clientId, clientSecret);

    // The body is the request to create
    var request = req.body && JSON.parse(req.body);

    context.log(request);

    if (!request || !request.type) {
        context.res = {
            headers: {
                "Content-Type": "application/json",
                "Access-Control-Allow-Credentials": "true",
                "Access-Control-Allow-Origin": "https://<tenant>.sharepoint.com",
                "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
                "Access-Control-Allow-Headers": "Content-Type, Set-Cookie",
                "Access-Control-Max-Age": "86400"
            },

            body: { status: "Properly authenticated" }
        };

        context.log("No BODY");
        context.log(context.res);
        context.done();
        return;
    }

    context.log("BODY available, submitting the request");
    // Submit the request
    service.submitRequest(request)
        .then(function () {
            context.log("Request successfully submitted");
            context.res = {
                headers: {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Credentials": "true",
                    "Access-Control-Allow-Origin": "https://<tenant>.sharepoint.com",
                    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
                    "Access-Control-Allow-Headers": "Content-Type, Set-Cookie",
                    "Access-Control-Max-Age": "86400"
                },

                body: { status: "succeeded" }
            };
            context.log(context.res);
            context.done();
        }).catch(function (error) {
            context.log(error);
            context.log("Request cannot be submitter");
            context.res = {
                headers: {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Credentials": "true",
                    "Access-Control-Allow-Origin": "https://<tenant>.sharepoint.com",
                    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
                    "Access-Control-Allow-Headers": "Content-Type, Set-Cookie",
                    "Access-Control-Max-Age": "86400"
                },
                status: 400,
                body: { status: "failed" }
            };

            context.log(context.res);
            context.done();
        });
};