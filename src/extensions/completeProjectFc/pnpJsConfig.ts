import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { ISPFXContext, spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/fields";


let sp: SPFI;

export const getSP = (context?: FormCustomizerContext): SPFI => {
    if (!sp && context !== null) {
        sp = spfi().using(SPFx(context as ISPFXContext))
    }
    return sp;
}