import * as React from "react";
import { Modal, PrimaryButton, DefaultButton, TextField, Stack, Text, MessageBar, MessageBarType, ChoiceGroup, IChoiceGroupOption} from "@fluentui/react";
import { NormalPeoplePicker } from "office-ui-fabric-react/lib/Pickers";
import { IPersonaProps } from "office-ui-fabric-react/lib/Persona";
import { IRequestApprovalModalProps, IApprovalRequest } from "../interfaces/IRequestApprovalModalProps";
import { SelectedFilesList } from "./SelectedFilesList";
import { ISharePointService } from "../interfaces/ISharePointService";
import { UserSearchService } from "../services/UserSearchService";
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
  context
}) => {
  const [selfApproval, setSelfApproval] = React.useState<boolean | undefined>(undefined);
  const [selectedApprovers, setSelectedApprovers] = React.useState<IPersonaProps[]>([]);
  const [comments, setComments] = React.useState<string>("");
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const resetForm = React.useCallback((): void => {
    setSelfApproval(undefined);
    setSelectedApprovers([]);
    setComments("");
    setIsSubmitting(false);
    setErrorMessage(undefined);
  }, []);

  const handleDismiss = React.useCallback((): void => {
    resetForm();
    onDismiss();
  }, [resetForm, onDismiss]);

  const handleSelfApprovalChange = React.useCallback((_?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void => {
      if (option) {
        setSelfApproval(option.key === "true");
        if (errorMessage) setErrorMessage(undefined);
      }
    },
    [errorMessage]
  );

  const onResolveSuggestions = React.useCallback(
    async (filterText: string, currentPersonas?: IPersonaProps[]): Promise<IPersonaProps[]> => {
      if (filterText.length >= 2) {
        try {
          const spService: ISharePointService = context.serviceScope.consume(SharePointService.serviceKey);
          const spUsers = await spService.getUsers();
          const users = await UserSearchService.searchUsers(spUsers, filterText, 10);
          const selectedIds = new Set((currentPersonas || []).map((p) => p.id ?? p.secondaryText));

          return users.filter((user) => !selectedIds.has(user.mail)).map(
            (user) => ({
              text: user.displayName,
              secondaryText: user.mail,
              key: user.mail,
            } as IPersonaProps));
        } catch (error) {
          return [];
        }
      }
      return [];
    },
    []
  );

  const onPeoplePickerChange = React.useCallback((items?: IPersonaProps[]): void => {
      setSelectedApprovers(items || []);
      if (errorMessage) setErrorMessage(undefined);
    }, [errorMessage]);

  const handleCommentsChange = React.useCallback((_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
      setComments(newValue || "");
      if (errorMessage) setErrorMessage(undefined);
    }, [errorMessage]);

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
        reason: selfApproval === true ? "Self-approved" : comments.trim(),
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
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      setErrorMessage(`Failed to submit approval request: ${errorMessage}`);
      setIsSubmitting(false);
    }
  }, [
    selfApproval,
    selectedApprovers,
    comments,
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
              isMultiline={false}
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
                    <NormalPeoplePicker
                      onResolveSuggestions={(filterText, _) =>
                        onResolveSuggestions(filterText, selectedApprovers)
                      }
                      onChange={onPeoplePickerChange}
                      getTextFromItem={(persona: IPersonaProps) => persona.text || ""}
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
                      // itemLimit={1}
                      resolveDelay={800}
                    />
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
