import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { RequestApprovalModal } from "../components/RequestApprovalModal";
import {
  ISelectedFile,
  IApprovalRequest,
} from "../../model/IRequestApprovalModalProps";
import SharePointListService, {
  ISharePointListService,
} from "../../services/SharePointAPI/SharePointListService";
import SuccessDialog from "../components/RequestApprovalModal/SuccessDialog";
import WarningDialog from "../components/RequestApprovalModal/WarningDialog";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IRequestDocumentApprovalCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "RequestDocumentApprovalCommandSet";

// export interface IFileProperties {
//   fileLeafRef?: string;
//   fileUrl?: string;
//   fileIcon?: string;
//   TenantUrl?: string;
//   DriveId?: string;

export default class RequestDocumentApprovalCommandSet extends BaseListViewCommandSet<IRequestDocumentApprovalCommandSetProperties> {
  private _sharePointListService: ISharePointListService;
  private _modalContainer: HTMLDivElement | null = null;

  public async onInit(): Promise<void> {
    console.log("🚀 RequestDocumentApprovalCommandSet initialized");

    // Initialize SharePointListService using ServiceKey pattern
    this._sharePointListService = this.context.serviceScope.consume(
      SharePointListService.serviceKey
    );

    // Create modal container
    this._modalContainer = document.createElement("div");
    document.body.appendChild(this._modalContainer);

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand("REQUEST_APPROVAL_COMMAND");
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    return Promise.resolve();
  }

  public onDispose(): void {
    // Clean up modal container
    if (this._modalContainer) {
      // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
      ReactDOM.unmountComponentAtNode(this._modalContainer);
      document.body.removeChild(this._modalContainer);
      this._modalContainer = null;
    }
    super.onDispose();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "REQUEST_APPROVAL_COMMAND":
        this._handleRequestApproval(event);
        break;

      // case 'COMMAND_2':
      //   Dialog.alert(`Clicked ${strings.Command2}`).catch(() => {
      //     /* handle error */
      //   });
      //   break;
      default:
        throw new Error("Unknown command");
    }
  }

  private _handleRequestApproval(
    event: IListViewCommandSetExecuteEventParameters
  ): void {
    const selectedItems = this.context.listView.selectedRows;

    if (!selectedItems || selectedItems.length === 0) {
      Dialog.alert(
        "Please select at least one document to request approval."
      ).catch(() => {
        /* handle error */
      });
      return;
    }

    // Get selected file information
    const selectedFiles: ISelectedFile[] = selectedItems.map((item) => {
      // Check if item is a folder by examining ContentType or FSObjType
      const contentType = item.getValueByName("ContentType");
      const fsobjType = item.getValueByName("FSObjType");
      const isFolder =
        fsobjType === "1" ||
        contentType === "Folder" ||
        (contentType && contentType.indexOf("Folder") !== -1);

      return {
        name:
          item.getValueByName("FileLeafRef") || item.getValueByName("Title"),
        id: item.getValueByName("ID"),
        serverRelativeUrl: item.getValueByName("FileRef"),
        size: item.getValueByName("File_x0020_Size"),
        modified: item.getValueByName("Modified"),
        modifiedBy:
          item.getValueByName("Author")?.title || item.getValueByName("Author"),
        isFolder: isFolder,
      };
    });

    console.log("🚀 Selected files for approval:", selectedFiles);

    // Show custom React modal
    this._showApprovalModal(selectedFiles);
  }

  private _showApprovalModal(selectedFiles: ISelectedFile[]): void {
    if (!this._modalContainer) return;

    const modalElement = React.createElement(RequestApprovalModal, {
      isOpen: true,
      selectedFiles: selectedFiles,
      context: this.context,
      onDismiss: () => {
        this._hideApprovalModal();
      },
      onSubmit: async (approvalRequest: IApprovalRequest) => {
        await this._processApprovalRequest(approvalRequest);
        this._hideApprovalModal();
      },
    });

    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDOM.render(modalElement, this._modalContainer);
  }

  private _hideApprovalModal(): void {
    if (!this._modalContainer) return;
    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDOM.unmountComponentAtNode(this._modalContainer);
  }

  private async _processApprovalRequest(
    approvalRequest: IApprovalRequest
  ): Promise<void> {
    console.log("🔄 Processing approval request:", approvalRequest);

    try {
      // Submit the approval request to SharePoint List
      const listItem = await this._sharePointListService.submitApprovalRequest(
        approvalRequest
      );

      console.log("✅ Approval request saved to SharePoint List:", listItem);

      // Show success message - simple format for demo
      // const statusText = approvalRequest.selfApproval
      //   ? "Self-Approved"
      //   : "Pending Approval";
      // Dialog.alert(
      //   `✅ Approval request submitted successfully! Status: ${statusText}. ${approvalRequest.files.length} file(s) processed and saved to SharePoint List.`
      // ).catch(() => {
      //   /* handle error */
      // });
      const dialog = new SuccessDialog(
        `Your approval request has been submitted successfully.`
      );
      dialog.show().catch(() => {
        /* handle error */
      });
    } catch (error) {
      console.error("❌ Failed to process approval request:", error);

      const errorMessage =
        error instanceof Error ? error.message : String(error);
      Dialog.alert(
        `Failed to submit approval request to SharePoint: ${errorMessage}` +
          `Please check the SharePoint List exists and you have permissions to write to it.`
      ).catch(() => {
        /* handle error */
      });
      throw error; // Re-throw so modal can handle the error
    }
  }

  //public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {

  //}

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");

    const compareOneCommand: Command = this.tryGetCommand("REQUEST_APPROVAL_COMMAND");
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.

      //Add logic to exclude folders from selection and triggering approvals

      //Logic to display "request approval" command only when 1 to 5 documents are selected in librayry
      const selectedCount = this.context.listView.selectedRows?.length || 0;

      if (selectedCount >= 1 && selectedCount <= 5) {
        // Show command if between 1 and 5 items
        compareOneCommand.visible = true;
      } else {
        compareOneCommand.visible = false;
  
        if (selectedCount > 5) {
          const dialog = new WarningDialog(
            "You have selected more than 5 items.\n\nPlease choose up to 5 items for request approval."
          );
          dialog.show().catch(() => {
            /* handle error */
          });
        }
      }

      // switch (this.context.listView.selectedRows?.length) {
      //   case 1:
      //     compareOneCommand.visible = true;
      //     break;

      //   case 2:
      //     compareOneCommand.visible = true;
      //     break;
      //   case 3:
      //     compareOneCommand.visible = true;
      //     break;
      //   case 4:
      //     compareOneCommand.visible = true;
      //     break;
      //   case 5:
      //     compareOneCommand.visible = true;
      //     break;
      //   default:
      //     compareOneCommand.visible = false;
      //     break;
      // }
      // if (this.context.listView.selectedRows?.length === 1 || this.context.listView.selectedRows?.length <= 6) {
      //   compareOneCommand.visible = true;
      // }
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  };
}
