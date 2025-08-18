import * as React from "react";
import { Icon, IIconProps } from "@fluentui/react";

interface IFileIconProps extends Omit<IIconProps, "iconName"> {
  fileName: string;
  isFolder?: boolean;
}

export default function FileIcon({
  fileName,
  isFolder,
  styles,
  className,
  ...iconProps
}: IFileIconProps): React.ReactNode {
  const getFileIconData = React.useCallback(
    (
      fileName: string,
      isFolder?: boolean
    ): { iconName: string; color: string } => {
      // If it's a folder, return folder icon
      if (isFolder) {
        return { iconName: "FabricFolder", color: "#ffb900" }; // SharePoint folder yellow
      }

      const extension = fileName.split(".").pop()?.toLowerCase() || "";

      // Document types
      if (["doc", "docx"].indexOf(extension) !== -1)
        return { iconName: "WordDocument", color: "#2b579a" }; // Word blue
      if (["xls", "xlsx"].indexOf(extension) !== -1)
        return { iconName: "ExcelDocument", color: "#217346" }; // Excel green
      if (["ppt", "pptx"].indexOf(extension) !== -1)
        return { iconName: "PowerPointDocument", color: "#d24726" }; // PowerPoint orange
      if (["pdf"].indexOf(extension) !== -1)
        return { iconName: "PDF", color: "#dc3545" }; // PDF red

      // Text and code files
      if (["txt"].indexOf(extension) !== -1)
        return { iconName: "TextDocument", color: "#6c757d" }; // Gray
      if (["rtf"].indexOf(extension) !== -1)
        return { iconName: "RichTextDocument", color: "#6c757d" }; // Gray
      if (["csv"].indexOf(extension) !== -1)
        return { iconName: "Table", color: "#217346" }; // Green like Excel

      // Image files
      if (["jpg", "jpeg", "png", "gif", "bmp", "svg"].indexOf(extension) !== -1)
        return { iconName: "FileImage", color: "#e83e8c" }; // Pink/magenta

      // Video files
      if (["mp4", "avi", "mov", "wmv", "flv", "webm"].indexOf(extension) !== -1)
        return { iconName: "Video", color: "#6f42c1" }; // Purple

      // Audio files
      if (["mp3", "wav", "flac", "aac", "ogg"].indexOf(extension) !== -1)
        return { iconName: "MusicNote", color: "#fd7e14" }; // Orange

      // Archive files
      if (["zip", "rar", "7z", "tar", "gz"].indexOf(extension) !== -1)
        return { iconName: "ZipFolder", color: "#ffc107" }; // Yellow/amber

      // Code files
      if (
        [
          "js",
          "ts",
          "jsx",
          "tsx",
          "html",
          "css",
          "scss",
          "json",
          "xml",
        ].indexOf(extension) !== -1
      )
        return { iconName: "Code", color: "#20c997" }; // Teal

      // Default for unknown types
      return { iconName: "Page", color: "#0078d4" }; // Default blue
    },
    []
  );

  const { iconName, color } = getFileIconData(fileName, isFolder);

  // Default styles that can be overridden
  const defaultStyles = {
    root: { fontSize: 16, color: color },
  };

  // Merge default styles with any provided styles
  const mergedStyles = React.useMemo(() => {
    if (!styles) return defaultStyles;

    // If styles is a function, we can't merge it easily, so just use it
    if (typeof styles === "function") return styles;

    // If styles is an object, merge with defaults
    return {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      root: { ...defaultStyles.root, ...(styles as any).root },
    };
  }, [styles, defaultStyles]);

  return (
    <Icon
      {...iconProps}
      iconName={iconName}
      styles={mergedStyles}
      className={className}
    />
  );
}
