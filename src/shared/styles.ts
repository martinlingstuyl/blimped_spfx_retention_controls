import { IDialogFooterStyles } from "@fluentui/react/lib/Dialog";
import { IMessageBarStyles } from "@fluentui/react/lib/MessageBar";
import { IStackItemStyles, IStackTokens } from "@fluentui/react/lib/Stack";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";

export const messageBarStyles: IMessageBarStyles = {
  root: {
    marginBottom: "10px",
  },
};

export const stackItemStyles: IStackItemStyles = {
  root: {
    alignItems: "center",
    display: "flex",
    width: "250px",
  },
};

export const dialogFooterStyles: IDialogFooterStyles = {
  action: { },
  actions: {},
  actionsRight: {
    justifyContent: "space-between", 
    display: "flex"
  },
};

export const stackTokens: IStackTokens = {
  childrenGap: 5,
};

export const iconClass = mergeStyles({
  fontSize: 14,
  height: 14,
  width: 14,
  margin: "0 0 0 0",
  cursor: "pointer",
});

export const classNames = mergeStyleSets({
  green: [{ color: "darkgreen" }, iconClass],
  red: [{ color: "indianred" }, iconClass],
  blue: [{ color: "#28a8ea" }, iconClass],
  dark: [{ }, iconClass],
});
