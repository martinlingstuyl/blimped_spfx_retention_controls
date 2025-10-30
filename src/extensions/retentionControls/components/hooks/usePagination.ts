import * as React from "react";
import { IButtonProps } from "@fluentui/react/lib/Button";
import { format } from "@fluentui/react/lib/Utilities";
import * as strings from "RetentionControlsCommandSetStrings";
import { IPaginationHook } from "../types/LibraryViewTypes";

interface UsePaginationProps {
  totalPages: number;
  pageNumber: number;
  onPageChange: (page: number) => void;
}

export const usePagination = ({ totalPages, pageNumber, onPageChange }: UsePaginationProps): IPaginationHook => {
  const paginationButtons: IButtonProps[] = React.useMemo(() => {
    if (totalPages <= 1) {
      return [];
    }

    return [
      {
        iconProps: { iconName: "ChevronLeft" },
        onClick: () => onPageChange(pageNumber - 1),
        disabled: pageNumber === 1,
        title: pageNumber === 1 ? strings.IsFirstPage : format(strings.ToPage, pageNumber - 1)
      },
      {
        iconProps: { iconName: "ChevronRight" },
        onClick: () => onPageChange(pageNumber + 1),
        disabled: pageNumber === totalPages,
        title: pageNumber === totalPages ? strings.IsLastPage : format(strings.ToPage, pageNumber + 1)
      }
    ];
  }, [totalPages, pageNumber, onPageChange]);

  return { paginationButtons };
};