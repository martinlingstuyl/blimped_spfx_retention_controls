import * as React from "react";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { messageBarStyles } from "../../../../shared/styles";
import { INotification } from "../../../../shared/interfaces/INotification";

interface NotificationBarProps {
  notification?: INotification;
  isServedFromLocalhost?: boolean;
}

export const NotificationBar: React.FC<NotificationBarProps> = ({ 
  notification, 
  isServedFromLocalhost 
}) => {
  return (
    <>
      {isServedFromLocalhost && (
        <div style={{ marginBottom: 20 }}>
          <MessageBar messageBarType={MessageBarType.success}>
            Served from localhost
          </MessageBar>
        </div>
      )}
      {notification && (
        <MessageBar styles={messageBarStyles} messageBarType={notification.notificationType}>
          {notification.message}
        </MessageBar>
      )}
    </>
  );
};