/* eslint-disable @typescript-eslint/no-explicit-any */
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { AadTokenProviderFactory } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { SPFI, spfi, SPFx as spSPFx } from "@pnp/sp/presets/all";
import { Logger, LogLevel } from "@pnp/logging";
import { IApprovalRequest } from "../../model/IRequestApprovalModalProps";

import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/graph/users";
import { graphfi, GraphFI, SPFx as gSPFx } from "@pnp/graph";

export interface IApprovalRequestListItem {
  // Map to actual SharePoint List internal field names
  Title: string; // Maps to "Approval Title" column (SharePoint uses Title as internal name)
  Comments?: string; // Maps to "Comments" column (multi-line text)
  ApproverId: any;
  RequestorId: number;
  SitecollectionURL: string;
  ItemIDs: string;
  // Note: Approver and Attachments columns are skipped for now
}

export interface ISharePointListService {
  submitApprovalRequest(approvalRequest: IApprovalRequest): Promise<any>;
  getApprovalRequests(): Promise<any[]>;
  getUsers(): Promise<any[]>;
}

export default class SharePointListService implements ISharePointListService {
  public static readonly serviceKey: ServiceKey<ISharePointListService> =
    ServiceKey.create<ISharePointListService>(
      "RequestApproval.SharePointListService",
      SharePointListService
    );

  private static readonly LIST_NAME = "Approval Request List";
  private _pageContext!: PageContext;
  private _sp: SPFI;
  private _graph: GraphFI;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(async () => {
      const aadTokenProviderFactory = serviceScope.consume(AadTokenProviderFactory.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);

      // Initialize PnPjs with SPFx context
      this._sp = spfi().using(spSPFx({ pageContext: this._pageContext }));
      this._graph = graphfi().using(gSPFx({ aadTokenProviderFactory }));

      Logger.log({
        message: `SharePointListService initialized for ${this._pageContext.user.displayName}`,
        level: LogLevel.Verbose,
      });
    });
  }

  /**
   * Get Users
   * @returns 
   */
   public async getUsers(): Promise<any[]> {
    const users = await this._graph.users();
    return users;
  }

  /**
   * Submit an approval request to the SharePoint List
   * @param approvalRequest - The approval request data
   * @returns Promise resolving to the created list item
   */
  public async submitApprovalRequest(
    approvalRequest: IApprovalRequest
  ): Promise<any> {
    try {
      // Get the current user
      const currentUser = this._pageContext.user;

      const requestorUser = await this._sp.web.ensureUser(currentUser.loginName);

      const approverIds: number[] = [];

      if (approvalRequest.approvers.length > 0) {
        for (const approver of approvalRequest.approvers) {
          const approverUser = await this._sp.web.ensureUser(approver);
          approverIds.push(approverUser.Id);
        }
      }

      // Prepare the list item data (using SharePoint internal field names)
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

      // Use PnPjs - clean and simple, no headers needed
      const result = await this._sp.web.lists
        .getByTitle(SharePointListService.LIST_NAME)
        .items.add(listItemData);

      // PnPjs returns the item data directly in result.data, but let's handle both cases
      const itemData = result.data || result;
      const itemId = itemData?.Id || itemData?.ID || "Unknown";

      Logger.log({
        message: `✅ Approval request submitted successfully with ID: ${itemId}`,
        level: LogLevel.Info,
      });

      console.log("🔍 PnPjs result structure:", result);
      return itemData;
    } catch (error) {
      Logger.log({
        message: `❌ Error submitting approval request: ${error.message}`,
        level: LogLevel.Error,
      });
      console.log(error.message);
      throw error;
    }
  }

  /**
   * Get all approval requests from the SharePoint List
   * @returns Promise resolving to array of approval requests
   */
  public async getApprovalRequests(): Promise<any[]> {
    try {
      // Use PnPjs - clean and simple
      const items = await this._sp.web.lists
        .getByTitle(SharePointListService.LIST_NAME)
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
}
