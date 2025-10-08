export interface IApprovalRequestListItem {
  Title: string;
  Comments?: string;
  ApproversId: number[];
  RequestorId: number;
  SitecollectionURL: string;
  ItemIDs: string;
}