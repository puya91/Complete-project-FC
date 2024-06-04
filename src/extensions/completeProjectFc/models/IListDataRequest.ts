
export interface IListDataRequest {
    Title: string,
    ID: string,
    Created: Date,
    AuthorId: number,
    
    RiskTitle: string | undefined,
    Business: string | undefined,
    Country: string | undefined,
    RiskDate: Date | undefined,
    AdditionalNotes: string | undefined,
    State: string,
    ContainsDocuments: string,
    AssignedToPeopleId: number[],
}

