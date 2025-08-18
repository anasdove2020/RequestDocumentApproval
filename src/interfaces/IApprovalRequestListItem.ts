export interface IApprovalRequestListItem {
  Title: string;
  Comments?: string;
  ApproverId: number[];
  RequestorId: number;
  SitecollectionURL: string;
  ItemIDs: string;
}