import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export interface ISelectedFile {
  id: string;
  name: string;
  serverRelativeUrl: string;
  size?: number;
  modified?: Date;
  modifiedBy?: string;
  isFolder?: boolean;
  contentType: string;
}

export interface IRequestApprovalModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmit: (approvalRequest: IApprovalRequest) => Promise<void>;
  selectedFiles: ISelectedFile[];
  context: ListViewCommandSetContext;
}

export interface IApprovalRequest {
  files: ISelectedFile[];
  reason: string;
  approvers: string[];
  priority: 'Low' | 'Medium' | 'High';
  dueDate?: Date;
  selfApproval: boolean; // New field for "Do you want to approve this yourself?"
}
