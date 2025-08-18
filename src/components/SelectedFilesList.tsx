import * as React from "react";
import { Stack, Text, TooltipHost } from "@fluentui/react";
import { ISelectedFile } from "../interfaces/IRequestApprovalModalProps";
import FileIcon from "./FileIcon";

export interface ISelectedFilesListProps {
  selectedFiles: ISelectedFile[];
}

export const SelectedFilesList: React.FC<ISelectedFilesListProps> = ({
  selectedFiles,
}) => {
  const renderFileItem = React.useCallback(
    (file: ISelectedFile, index: number): JSX.Element => {
      return (
        <div
          key={index}
          style={{
            display: "flex",
            alignItems: "center",
            gap: "10px",
            padding: "8px 0",
            borderBottom:
              index < selectedFiles.length - 1 ? "1px solid #f3f2f1" : "none",
          }}
        >
          {FileIcon({ fileName: file.name, isFolder: file.isFolder })}
          <TooltipHost content={file.name}>
            <div
              style={{
                fontWeight: 600,
                maxWidth: "250px",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
              }}
            >
              {file.name}
            </div>
          </TooltipHost>
        </div>
      );
    },
    [selectedFiles.length]
  );

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
        Selected Files ({selectedFiles.length})
      </Text>
      <div
        style={{
          maxHeight: "200px",
          overflowY: "auto",
          border: "1px solid #edebe9",
          borderRadius: "4px",
          padding: "12px",
          backgroundColor: "#fafafa",
        }}
      >
        {selectedFiles.map((file, index) => renderFileItem(file, index))}
      </div>
    </Stack>
  );
};
