import { IFilePickerResult } from "@pnp/spfx-controls-react";

export interface IUpdateStates {
    riskTitle: string | undefined;
    selectedUsers: string[];
    business: string | undefined;
    country: string | undefined;
    riskDate: Date | undefined;
    notes: string | undefined;
    state: string;
    containsDocuments: string;
    riskReport: IFilePickerResult[] | undefined;
}