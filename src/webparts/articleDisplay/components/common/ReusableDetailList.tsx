import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DetailsHeader,
} from "@fluentui/react/lib/DetailsList";
import { mergeStyles } from "@fluentui/react";
import "./styles.css";
import { IColumnInfo } from "../PagesList/PagesService";

const customBodyClass = mergeStyles({
  row: {
    selectors: {
      ":nth-child(odd), .row :nth-child(odd)": {
        backgroundColor: "#efefef", // Background color for odd rows
      },
      ":nth-child(even), .row :nth-child(even)": {
        backgroundColor: "white", // Background color for even rows
      },
    },
  },
  overflowY: "auto",
  maxHeight: "700px",
});

// Define custom header styles
const customHeaderClass = mergeStyles({
  backgroundColor: "#efefef", // Custom background color
  color: "white", // Custom text color
  paddingTop: 0,
  paddingBottom: 0,
  header: {
    backgroundColor: "#0078d4", // Custom header background color
    borderBottom: "1px solid #ccc",
  },
});

export interface IReusableDetailListcomponents {
  columns: (
    columns: IColumnInfo[],
    onColumnClick: any,
    sortBy: string,
    isDecending: boolean,
    setShowFilter: (column: IColumn, columnType: string) => void
  ) => IColumn[];
  columnInfos: IColumnInfo[];
  setShowFilter: (column: IColumn, columnType: string) => void;
  items: any[];
  sortPages: (column: IColumn, isAscending: boolean) => void;
  sortBy: string;
  siteUrl: string;
  isDecending: boolean;
}

export class ReusableDetailList extends React.Component<
  IReusableDetailListcomponents,
  {}
> {
  constructor(components: IReusableDetailListcomponents) {
    super(components);
  }

  componentDidMount() {
    window.addEventListener("contentLoaded", this.handleContentLoaded);
  }

  componentWillUnmount() {
    window.removeEventListener("contentLoaded", this.handleContentLoaded);
  }

  componentDidUpdate(prevcomponents: IReusableDetailListcomponents) {
    if (prevcomponents.items !== this.props.items) {
      this.forceUpdate();
      window.dispatchEvent(new Event("contentLoaded"));
    }
  }

  handleContentLoaded = () => {
    const navSection: HTMLElement | null =
      document.querySelector(".custom-nav");
    const detailSection: HTMLElement | null =
      document.querySelector(".detail-display");

    function adjustNavHeight() {
      if (navSection && detailSection) {
        const detailHeight = detailSection.offsetHeight;
        const minHeight = 500; // Minimum height in pixels
        navSection.style.height = `${Math.max(detailHeight, minHeight)}px`;
      }
    }

    adjustNavHeight();
    window.addEventListener("resize", adjustNavHeight);
  };
  _onRenderDetailsHeader = (components: any) => {
    if (!components) {
      return null;
    }

    // Apply custom styles to the header
    return (
      <DetailsHeader
        {...components}
        styles={{
          root: customHeaderClass, // Apply custom styles
        }}
      />
    );
  };

  public render() {
    const {
      columnInfos,
      columns,
      items,
      sortPages,
      sortBy,
      isDecending,
      setShowFilter,
    } = this.props;

    return (
      <div>
        <DetailsList
          styles={{
            root: customBodyClass,
          }}
          items={items}
          compact={true}
          columns={columns(
            columnInfos,
            sortPages,
            sortBy,
            isDecending,
            setShowFilter
          )}
          selectionMode={SelectionMode.none}
          getKey={this._getKey}
          setKey="none"
          layoutMode={DetailsListLayoutMode.fixedColumns}
          isHeaderVisible={true}
          onRenderDetailsHeader={this._onRenderDetailsHeader}
          onItemInvoked={this._onItemInvoked}
          className="detailList"
        />
      </div>
    );
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _onItemInvoked = (item: any): void => {
    window.open(`${this.props.siteUrl}${item.FileRef}`, "_blank");
  };
}
