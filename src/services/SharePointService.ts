/* eslint-disable @typescript-eslint/no-explicit-any */
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/graph/users";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { AadTokenProviderFactory } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { SPFI, spfi, SPFx as spSPFx } from "@pnp/sp/presets/all";
import { Logger, LogLevel } from "@pnp/logging";
import { graphfi, GraphFI, SPFx as gSPFx } from "@pnp/graph";
import { IApprovalRequest } from "../interfaces/IRequestApprovalModalProps";
import { IApprovalRequestListItem } from "../interfaces/IApprovalRequestListItem";
import { ISharePointService } from "../interfaces/ISharePointService";
import { DOCUMENT_STATUS, LIST_NAME } from "../utils/constants";

export default class SharePointService implements ISharePointService {
  public static readonly serviceKey: ServiceKey<ISharePointService> = 
    ServiceKey.create<ISharePointService>("RequestApproval.SharePointService", SharePointService);

  private _pageContext!: PageContext;
  private _sp: SPFI;
  private _graph: GraphFI;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(async () => {
      const aadTokenProviderFactory = serviceScope.consume(AadTokenProviderFactory.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._sp = spfi().using(spSPFx({ pageContext: this._pageContext }));
      this._graph = graphfi().using(gSPFx({ aadTokenProviderFactory }));

      Logger.log({
        message: `SharePointService initialized for ${this._pageContext.user.displayName}`,
        level: LogLevel.Verbose,
      });
    });
  }
  
  public async getUsers(): Promise<any[]> {
    const users = await this._graph.users();
    return users;
  }
  
  public async submitApprovalRequest(approvalRequest: IApprovalRequest): Promise<any> {
    try {
      const currentUser = this._pageContext.user;
      const requestorUser = await this._sp.web.ensureUser(currentUser.loginName);
      const approverIds: number[] = [];

      if (approvalRequest.approvers.length > 0) {
        for (const approver of approvalRequest.approvers) {
          const approverUser = await this._sp.web.ensureUser(approver);
          approverIds.push(approverUser.Id);
        }
      }

      const listItemData: IApprovalRequestListItem = {
        Title: `Shared Documents`,
        ApproverId: approverIds,
        RequestorId: requestorUser.Id,
        SitecollectionURL: "sites/Sandpit",
        ItemIDs: approvalRequest.files.map(item => String(item.id)).join(";"),
        Comments:
          approvalRequest.reason ||
          `Request for ${approvalRequest.files.length} file(s). Priority: ${
            approvalRequest.priority
          }. ${
            approvalRequest.selfApproval
              ? "Self-approved."
              : "Pending approval."
          }`,
      };

      const result = await this._sp.web.lists
        .getByTitle(LIST_NAME.APPROVAL_REQUEST)
        .items.add(listItemData);

      const itemData = result.data || result;
      const itemId = itemData?.Id || itemData?.ID || "Unknown";

      Logger.log({
        message: `✅ Approval request submitted successfully with ID: ${itemId}`,
        level: LogLevel.Info,
      });

      await this.updateSharedDocument(approvalRequest, currentUser.displayName);

      return itemData;
    } catch (error) {
      Logger.log({
        message: `❌ Error submitting approval request: ${error.message}`,
        level: LogLevel.Error,
      });
      throw error;
    }
  }
  
  public async getApprovalRequests(): Promise<any[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle(LIST_NAME.APPROVAL_REQUEST)
        .items.select("Id", "Title", "Comments", "Created", "Author/Title")
        .expand("Author")()
        .catch((error) => {
          Logger.log({
            message: `❌ Error getting approval requests: ${error.message}`,
            level: LogLevel.Error,
          });
          throw error;
        });

      Logger.log({
        message: `✅ Retrieved ${items.length} approval requests`,
        level: LogLevel.Info,
      });

      return items;
    } catch (error) {
      Logger.log({
        message: `❌ Error getting approval requests: ${error.message}`,
        level: LogLevel.Error,
      });
      throw error;
    }
  }

  private async updateSharedDocument(approvalRequest: IApprovalRequest, username: string): Promise<any> {
    const spIds = approvalRequest.files.map(item => String(item.id));
    const today = new Date();
    const formattedDate = `${today.getDate().toString().padStart(2, "0")}/${(today.getMonth() + 1).toString().padStart(2, "0")}/${today.getFullYear()}`;

    const newHistory = `${username} - ${formattedDate} - ${approvalRequest.selfApproval
        ? "Submitted document for self approval."
        : "Submitted document for approval."}`;

    for (const spId of spIds) {
      const item = await this._sp.web.lists
        .getByTitle(LIST_NAME.SHARED_DOCUMENT)
        .items.getById(Number(spId))
        .select("History")();

      const prevHistory: string = item.History || "";

      const updatedHistory = prevHistory ? `${prevHistory}\n${newHistory}` : newHistory;
        
      await this._sp.web.lists
        .getByTitle(LIST_NAME.SHARED_DOCUMENT)
        .items.getById(Number(spId))
        .update({
          Approval_x0020_Status: approvalRequest.selfApproval ? DOCUMENT_STATUS.AUTO_APPROVED : DOCUMENT_STATUS.WAITING_FOR_APPROVAL,
          Approval_x0020_History: updatedHistory
        });
    }
  }
}
