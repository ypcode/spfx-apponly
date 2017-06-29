import pnp, { Web, NodeFetchClient } from "sp-pnp-js";
import { IHardwareRequest } from "../model/IHardwareRequest";

const HardwareRequestListTitle = "Hardware Requests";

export class HardwareRequestService {

    private web: Web;

    constructor (web: Web) {
        this.web = web;
    }

    public static createAppOnly(siteUrl: string, clientId: string, clientSecret: string) : HardwareRequestService {
        pnp.setup({
            fetchClientFactory: () => {
                return new NodeFetchClient(siteUrl, clientId, clientSecret);
            }
        });

        return new HardwareRequestService(new Web(siteUrl));
    }

    public static createForCurrentWeb() : HardwareRequestService {
        return new HardwareRequestService(pnp.sp.web);
    }

    public submitRequest(request: IHardwareRequest) : Promise<any> {
        console.log(request);
        return this.web.lists.getByTitle(HardwareRequestListTitle).items.add({
            Title: request.title || ("Request " + new Date().toUTCString()),
            HW_HardwareType: request.type,
            HW_Remark: request.remark,
            HW_Quantity: request.quantity
        });
    }
}