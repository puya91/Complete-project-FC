
import { FormDisplayMode, Guid } from "@microsoft/sp-core-library";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { IListDataRequest } from "../models/IListDataRequest";

export interface ICompleteProjectFcProps {
    context: FormCustomizerContext;
    displayMode: FormDisplayMode;
    onSave: () => void;
    onClose: () => void;
    
    listGuid: Guid;
    itemID: number | undefined;
    listItem: IListDataRequest;
    businessListGuid: string;
}