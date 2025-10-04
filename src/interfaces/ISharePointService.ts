/* eslint-disable @typescript-eslint/no-explicit-any */
import { IApprovalRequest } from "./IRequestApprovalModalProps";

export interface ISharePointService {
  setMainUrl: (mainSiteUrl: string, mainRequestApprovalUrl: string) => void;
  submitApprovalRequest(approvalRequest: IApprovalRequest): Promise<any>;
  getApprovalRequests(): Promise<any[]>;
  getUsers(): Promise<any[]>;
}