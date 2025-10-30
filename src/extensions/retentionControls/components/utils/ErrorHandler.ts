import { MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Log } from "@microsoft/sp-core-library";
import { LOG_SOURCE } from "../../RetentionControlsCommandSet";
import { INotification } from "../../../../shared/interfaces/INotification";

export type ErrorSeverity = 'error' | 'warning' | 'info';

export interface ErrorContext {
  operation: string;
  itemName?: string;
  itemId?: number;
  additionalData?: Record<string, unknown>;
}

export class ErrorHandler {
  private static instance: ErrorHandler;
  private onNotificationChange?: (notification: INotification | undefined) => void;

  private constructor() {}

  public static getInstance(): ErrorHandler {
    if (!ErrorHandler.instance) {
      ErrorHandler.instance = new ErrorHandler();
    }
    return ErrorHandler.instance;
  }

  public setNotificationCallback(callback: (notification: INotification | undefined) => void): void {
    this.onNotificationChange = callback;
  }

  public handleError(
    error: Error | string,
    context: ErrorContext,
    severity: ErrorSeverity = 'error'
  ): void {
    const errorMessage = typeof error === 'string' ? error : error.message;
    const contextualMessage = this.buildContextualMessage(errorMessage, context);
    
    // Log the error with full context
    Log.error(LOG_SOURCE, new Error(contextualMessage + JSON.stringify(context)));

    // Notify UI if callback is set
    if (this.onNotificationChange) {
      this.onNotificationChange({
        message: contextualMessage,
        notificationType: this.severityToMessageBarType(severity)
      });
    }
  }

  public handleSuccess(message: string): void {
    if (this.onNotificationChange) {
      this.onNotificationChange({
        message,
        notificationType: MessageBarType.success
      });
    }
  }

  public clearNotification(): void {
    if (this.onNotificationChange) {
      this.onNotificationChange(undefined);
    }
  }

  private buildContextualMessage(errorMessage: string, context: ErrorContext): string {
    const parts = [errorMessage];
    
    if (context.itemName) {
      parts.unshift(`Error with item '${context.itemName}':`);
    } else if (context.operation) {
      parts.unshift(`Error during ${context.operation}:`);
    }

    return parts.join(' ');
  }

  private severityToMessageBarType(severity: ErrorSeverity): MessageBarType {
    switch (severity) {
      case 'error':
        return MessageBarType.error;
      case 'warning':
        return MessageBarType.warning;
      case 'info':
        return MessageBarType.info;
      default:
        return MessageBarType.error;
    }
  }
}

// Convenience functions for common error scenarios
export const errorHandler = ErrorHandler.getInstance();

export const handleApiError = (error: Error | string, operation: string): void => {
  errorHandler.handleError(error, { operation });
};

export const handleItemError = (error: Error | string, itemName: string, operation: string): void => {
  errorHandler.handleError(error, { operation, itemName });
};

export const handleSuccess = (message: string): void => {
  errorHandler.handleSuccess(message);
};

export const clearNotification = (): void => {
  errorHandler.clearNotification();
};