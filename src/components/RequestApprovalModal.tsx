/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import {
  Modal,
  PrimaryButton,
  DefaultButton,
  TextField,
  Stack,
  Text,
  MessageBar,
  MessageBarType,
  ChoiceGroup,
  IChoiceGroupOption,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { NormalPeoplePicker } from "office-ui-fabric-react/lib/Pickers";
import { IPersonaProps } from "office-ui-fabric-react/lib/Persona";
import {
  IRequestApprovalModalProps,
  IApprovalRequest,
} from "../interfaces/IRequestApprovalModalProps";
import { SelectedFilesList } from "./SelectedFilesList";
import { ISharePointService } from "../interfaces/ISharePointService";
import SharePointService from "../services/SharePointService";

const selfApprovalOptions: IChoiceGroupOption[] = [
  { key: "true", text: "Yes" },
  { key: "false", text: "No" },
];

export const RequestApprovalModal: React.FC<IRequestApprovalModalProps> = ({
  isOpen,
  selectedFiles,
  onDismiss,
  onSubmit,
  context,
}) => {
  const [selfApproval, setSelfApproval] = React.useState<boolean | undefined>(
    undefined
  );
  const [selectedApprovers, setSelectedApprovers] = React.useState<
    IPersonaProps[]
  >([]);
  const [comments, setComments] = React.useState<string>("");
  const [authorComments, setAuthorComments] = React.useState<string>("");
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(
    undefined
  );
  const [cachedUsers, setCachedUsers] = React.useState<any[]>([]);
  const [isLoadingUsers, setIsLoadingUsers] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string | null>(null);

  React.useEffect(() => {
    const loadUsers = async (): Promise<any> => {
      if (isOpen && cachedUsers.length === 0 && !isLoadingUsers) {
        setIsLoadingUsers(true);
        setLoadError(null);
        try {
          const spService: ISharePointService = context.serviceScope.consume(
            SharePointService.serviceKey
          );
          const users = await spService.getUsers();
          setCachedUsers(users);
        } catch (err: any) {
          setLoadError(
            "Failed to load users from SharePoint. You can still type manually."
          );
        } finally {
          setIsLoadingUsers(false);
        }
      }
    };
    loadUsers().catch(() => {
      /* handle error */
    });
  }, [isOpen, context, cachedUsers.length, isLoadingUsers]);

  const resetForm = React.useCallback((): void => {
    setSelfApproval(undefined);
    setSelectedApprovers([]);
    setComments("");
    setAuthorComments("");
    setIsSubmitting(false);
    setErrorMessage(undefined);
  }, []);

  const handleDismiss = React.useCallback((): void => {
    resetForm();
    onDismiss();
  }, [resetForm, onDismiss]);

  const handleSelfApprovalChange = React.useCallback(
    (
      _?: React.FormEvent<HTMLElement | HTMLInputElement>,
      option?: IChoiceGroupOption
    ): void => {
      if (option) {
        setSelfApproval(option.key === "true");
        if (errorMessage) setErrorMessage(undefined);
      }
    },
    [errorMessage]
  );

  const onResolveSuggestions = React.useCallback(
    async (
      filterText: string,
      currentPersonas?: IPersonaProps[]
    ): Promise<IPersonaProps[]> => {
      if (!filterText || filterText.length < 2) return [];

      // Debounce 300ms
      await new Promise((resolve) => setTimeout(resolve, 300));

      const selectedIds = new Set(
        (currentPersonas || []).map((p) => p.id ?? p.secondaryText)
      );

      const filtered = cachedUsers
        .filter(
          (user) =>
            (user.displayName
              ?.toLowerCase()
              .includes(filterText.toLowerCase()) ||
              user.mail?.toLowerCase().includes(filterText.toLowerCase())) &&
            !selectedIds.has(user.mail)
        )
        .slice(0, 10)
        .map(
          (user) =>
            ({
              text: user.displayName,
              secondaryText: user.mail,
              key: user.mail,
            } as IPersonaProps)
        );

      return filtered;
    },
    [cachedUsers]
  );

  const onPeoplePickerChange = React.useCallback(
    (items?: IPersonaProps[]): void => {
      setSelectedApprovers(items || []);
      if (errorMessage) setErrorMessage(undefined);
    },
    [errorMessage]
  );

  const handleCommentsChange = React.useCallback(
    (
      _: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setComments(newValue || "");
      if (errorMessage) setErrorMessage(undefined);
    },
    [errorMessage]
  );

  const handleAuthorCommentsChange = React.useCallback(
    (
      _: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setAuthorComments(newValue || "");
      if (errorMessage) setErrorMessage(undefined);
    },
    [errorMessage]
  );

  const handleSubmit = React.useCallback(async (): Promise<void> => {
    if (selfApproval === undefined) {
      setErrorMessage(
        "Please select whether you want to approve this yourself."
      );
      return;
    }

    if (selfApproval === false) {
      if (selectedApprovers.length === 0) {
        setErrorMessage("Please select at least one approver.");
        return;
      }

      if (!comments.trim()) {
        setErrorMessage("Please provide comments for the approval request.");
        return;
      }
    }

    setIsSubmitting(true);
    setErrorMessage(undefined);
    try {
      const approvalRequest: IApprovalRequest = {
        files: selectedFiles,
        reason: selfApproval === true ? authorComments.trim() : comments.trim(),
        authorComments: selfApproval === true ? authorComments.trim() : "",
        approvers:
          selfApproval === true
            ? []
            : selectedApprovers
                .map((a) => a.secondaryText || "")
                .filter((text) => text.length > 0),
        priority: "Medium", // Default priority since dropdown removed
        selfApproval: selfApproval === true,
      };

      await onSubmit(approvalRequest);
      resetForm();
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);

      let friendlyMessage = "";

      if (error.message?.includes("locked")) {
        friendlyMessage = `Failed to submit approval request because the file is currently locked for editing. Please close the file and try again.`;
      } else if (errorMessage.includes("does not exist")) {
        const match = errorMessage.match(/Column '([^']+)' does not exist/i);
        const columnName = match
          ? match[1].replace(/_x0020_/g, " ")
          : "Unknown";
        friendlyMessage = `The column "${columnName}" does not exist in the target SharePoint list. Please make sure this column is created before submitting the request.`;
      } else {
        friendlyMessage = `Failed to submit approval request to SharePoint. Please ensure the SharePoint list exists and you have permission to update it.`;
      }

      setErrorMessage(friendlyMessage);
      setIsSubmitting(false);
    }
  }, [
    selfApproval,
    selectedApprovers,
    comments,
    authorComments,
    selectedFiles,
    onSubmit,
    resetForm,
  ]);

  if (!isOpen) return null;

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={handleDismiss}
      isBlocking={false}
      containerClassName="ms-modalExample-container"
    >
      <div style={{ padding: "24px", minWidth: "600px", maxWidth: "800px" }}>
        <Stack tokens={{ childrenGap: 20 }}>
          {/* Header */}
          <Stack
            horizontal
            horizontalAlign="space-between"
            verticalAlign="center"
          >
            <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
              Request Approval
            </Text>
            <DefaultButton
              text="✕"
              onClick={handleDismiss}
              styles={{ root: { minWidth: "auto", padding: "4px 8px" } }}
            />
          </Stack>

          {/* Error Message */}
          {errorMessage && (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
            >
              {errorMessage}
            </MessageBar>
          )}

          {/* Main Content - Two Column Layout */}
          <Stack horizontal tokens={{ childrenGap: 24 }}>
            {/* Left Column - Form Fields */}
            <Stack styles={{ root: { flex: 1 } }} tokens={{ childrenGap: 16 }}>
              {/* Self Approval Radio Buttons */}
              <ChoiceGroup
                label="Do you want to approve this yourself?"
                selectedKey={
                  selfApproval === undefined ? undefined : String(selfApproval)
                }
                options={selfApprovalOptions}
                onChange={handleSelfApprovalChange}
                required
              />

              {selfApproval === true && (
                <>
                  <TextField
                    label="Author comments"
                    placeholder="Please provide comments..."
                    multiline
                    rows={4}
                    value={authorComments}
                    onChange={handleAuthorCommentsChange}
                  />
                </>
              )}

              {selfApproval === false && (
                /* If NO - Show approvers and comments fields */
                <>
                  <Stack tokens={{ childrenGap: 5 }}>
                    <Text
                      variant="medium"
                      styles={{
                        root: {
                          fontWeight: 600,
                          color: "#323130",
                          marginBottom: "4px",
                        },
                      }}
                    >
                      Select Approver(s){" "}
                      <span style={{ color: "#a4262c" }}>*</span>
                    </Text>

                    {isLoadingUsers ? (
                      <Spinner
                        size={SpinnerSize.small}
                        label="Loading users..."
                      />
                    ) : (
                      <NormalPeoplePicker
                        onResolveSuggestions={(filterText, _) =>
                          onResolveSuggestions(filterText, selectedApprovers)
                        }
                        onChange={onPeoplePickerChange}
                        getTextFromItem={(p: IPersonaProps) => p.text || ""}
                        pickerSuggestionsProps={{
                          suggestionsHeaderText: "Suggested People",
                          noResultsFoundText: "No results found",
                          loadingText: "Loading...",
                          searchingText: "Searching...",
                        }}
                        key="normal"
                        removeButtonAriaLabel="Remove"
                        inputProps={{
                          placeholder: "Type to search for approvers",
                          "aria-label": "Select Approver(s)",
                        }}
                        resolveDelay={500}
                      />
                    )}

                    {loadError && (
                      <MessageBar
                        messageBarType={MessageBarType.warning}
                        isMultiline={false}
                      >
                        {loadError}
                      </MessageBar>
                    )}
                  </Stack>

                  <TextField
                    label="Comments"
                    placeholder="Please provide comments for the approval request..."
                    multiline
                    rows={4}
                    value={comments}
                    onChange={handleCommentsChange}
                    required
                  />
                </>
              )}
            </Stack>

            {/* Right Column - Selected Files */}
            <Stack styles={{ root: { flex: 1 } }}>
              <SelectedFilesList selectedFiles={selectedFiles} />
            </Stack>
          </Stack>

          {/* Action Buttons */}
          <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 10 }}>
            <DefaultButton
              text="Cancel"
              onClick={handleDismiss}
              disabled={isSubmitting}
            />
            <PrimaryButton
              text={isSubmitting ? "Submitting..." : "Submit Request"}
              onClick={handleSubmit}
              disabled={
                isSubmitting ||
                selfApproval === undefined ||
                (selfApproval === false &&
                  (selectedApprovers.length === 0 || comments.trim() === ""))
              }
            />
          </Stack>
        </Stack>
      </div>
    </Modal>
  );
};
