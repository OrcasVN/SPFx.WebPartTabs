import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from '@pnp/sp'
import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";

export default class WebPartTabsServices {
    private wpContext: WebPartContext;
    private readonly _spfi: SPFI;

    constructor(wpContext: WebPartContext) {
        this.wpContext = wpContext;
        this._spfi = spfi().using(SPFx(wpContext));
    }

    public async getWebParts() {
        try {
            const partDefs = await this._spfi.web.getClientsideWebParts();
            const options = []
            partDefs.forEach(item => {
                const manifest = JSON.parse(item.Manifest)
                options.push({
                    key: manifest.id, text: item.Name, value: manifest.id, id: item.Id
                })
            })

            console.log(options)
            return options
        } catch (error) {
            throw new Error(error);
        }
    }
}