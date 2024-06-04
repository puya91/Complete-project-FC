import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { SPFI} from "@pnp/sp";
import { IItem } from "@pnp/sp/items/types";
import { getSP } from "../pnpJsConfig";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/sites";
import "@pnp/sp/attachments";
import { IFilePickerResult } from "@pnp/spfx-controls-react";
import { IUpdateStates } from "../models/IUpdateStates";
import { IDocumentData } from "../models/IDocumentData";

    let sp: SPFI;

    const attachDocument = async (_sp: SPFI, itemId: number, documentPeriodListId: string, fileItem: IFilePickerResult): Promise<void> => {
        try {
            const result = await fileItem.downloadFileContent();
            const item: IItem = _sp.web.lists.getById(documentPeriodListId).items.getById(itemId);
            await item.attachmentFiles.add(fileItem.fileName, result);
        } catch (error) {
            console.error('Error attaching document:', error);
        }
    }

    export const initializeSpObject = (context: FormCustomizerContext): void =>{
        sp = getSP(context)
    }

    export const getDocumentsUrl = async (subfolderTitle: string): Promise<IDocumentData[] | undefined> => {
        try {
            // Items of RiskEventDocumentsList
            const documentsListId = '59eed830-55d8-4736-92ad-4244ef1a2eec';
            const documentsList = sp.web.lists.getById(documentsListId);
            const items = await documentsList.items.filter(`substringof('${subfolderTitle}_',Title)`)();

            const urlsPromises = items.map(async (item) => {

                const attachments = await documentsList.items.getById(item.Id).attachmentFiles();
                
                return attachments.map(attachment => ({
                    name: attachment.FileName,
                    url: `${window.location.origin}${attachment.ServerRelativeUrl}`
                }));
            });
    
            const urlsArrays = await Promise.all(urlsPromises);
            const urls = urlsArrays.flat(); // Flatten the array of arrays into a single array 
            return urls;
        } catch (error) {
            console.error("Error fetching file URL:", error);
        }
    }

    export const onRejectOrCloseSubmit = async (stateValue: string, elementId: string): Promise<void> => {
        const requestListGuid = '48db0c6b-7b64-499f-8e4a-035499aef8f2';

        try {
            await sp.web.lists.getById(requestListGuid).items.getById(parseInt(elementId, 10)).update({
                State: stateValue
            });
            
            console.log("Item updated successfully");
        } catch (error) {
            console.error("Error updating item: ", error);
        }
    }

    export const updateListItems = async (context: FormCustomizerContext, states: IUpdateStates, elementId: string): Promise<void> => {
        const requestListGuid = '48db0c6b-7b64-499f-8e4a-035499aef8f2';
        const documentsListGuid = '59eed830-55d8-4736-92ad-4244ef1a2eec';

        try {
            await sp.web.lists.getById(requestListGuid).items.getById(parseInt(elementId, 10)).update({
                RiskTitle: states.riskTitle,
                Business: states.business,
                Country: states.country,
                RiskDate: states.riskDate,
                AssignedToPeopleId: states.selectedUsers,
                AdditionalNotes: states.notes,
                ContainsDocuments: states.containsDocuments,
                State: states.state,
            })

            if (states.riskReport !== undefined) {
                const elementTitle = `RiskEventRequestsId-${elementId}`;
                const documentsListPath = `${context.pageContext.web.absoluteUrl}/Lists/RiskEventDocumentsList`;
                
                // Check if the subfolder exists
                let folderExists = false;
                try {
                    const subfolder = await sp.web.lists.getById(documentsListGuid).items.filter(`Title eq '${elementTitle}'`)();
                    if (subfolder.length === 1) {
                        folderExists = true;
                    }
                } catch (error) {
                    console.error('Error checking for folder existence:', error);
                    throw error; 
                }

                // If the folder does not exist, create it
                if (!folderExists) {
                    await sp.web.lists.getById(documentsListGuid).rootFolder.addSubFolderUsingPath(elementTitle);
                }
        
                for (const element of states.riskReport) {
                    try {
                        const result = await sp.web.lists.getById(documentsListGuid).addValidateUpdateItemUsingPath([
                            {
                                FieldName: "Title",
                                FieldValue: `${elementTitle}_${element.fileName}`,
                            }
                        ], `${documentsListPath}/${elementTitle}`);
                        
                        await attachDocument(sp, Number(result[1].FieldValue), documentsListGuid, element);
                    } catch (error) {
                        console.error('Error processing risk report:', error);
                    }
                }
            }
        }
        catch (error) {
            console.error('Error adding items to the list:', error);
        }
    }

    export const deleteFiles = async (documentsSubStringTitle: string): Promise<void> => {
        const documentsListGuid = '59eed830-55d8-4736-92ad-4244ef1a2eec';

        try {
            const subfolderToDelete = await sp.web.lists.getById(documentsListGuid).items.filter(`Title eq '${documentsSubStringTitle}'`)();
    
            await sp.web.lists.getById(documentsListGuid).items.getById(subfolderToDelete[0].Id).delete();
            console.log(`Subfolder ${documentsSubStringTitle} and its documents are deleted successfully.`);
        } catch (error) {
            console.error("Error deleting documents: ", error);
        }
    }