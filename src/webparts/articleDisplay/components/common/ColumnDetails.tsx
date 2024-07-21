import { IColumn, Icon } from "@fluentui/react";
import * as React from "react";

const CellRender = (props: {
  columnName: string;
  columnType: string;
  item: any;
}) => {
  const { columnName, item, columnType } = props;

  switch (columnType) {
    case "Text":
      if (columnName === "Title") {
        return (
          <a href={item.FileRef} target="_blank">
            {item.Title}
          </a>
        );
      }
      return <div>{item[columnName]}</div>;
    case "DateTime":
      const date = new Date(item[columnName]);

      const optionsDate: any = {
        year: "numeric",
        month: "short",
        day: "numeric",
      };
      const formattedDate = date.toLocaleDateString("en-US", optionsDate);

      const optionsTime: any = {
        hour: "numeric",
        minute: "numeric",
        hour12: true,
      };
      const formattedTime = date.toLocaleTimeString("en-US", optionsTime);

      const formattedDateTime = `${formattedDate} ${formattedTime}`;
      return <div>{formattedDateTime}</div>;
    case "TaxonomyFieldTypeMulti":
      const categories = item.TaxCatchAll.map(
        (category: any) => category.Term
      ).join(", ");
      return <div>{categories}</div>;
    case "Number":
      return <div>{item[columnName]}</div>;
    case "User":
      return <div>{item[columnName].Title}</div>;
    default:
      return <div>{item[columnName]}</div>;
  }
};
const HeaderRender = (
  column: IColumn,
  columnType: string,
  onColumnClick: (column: IColumn) => void,
  setShowFilter: (column: IColumn, columnType: string) => void
): JSX.Element => {
  return (
    <div
      style={{
        display: "flex",
        alignItems: "start",
        justifyContent: "space-between",
        width: "100%", // Adjust padding as needed
        boxSizing: "border-box",
      }}
    >
      <span
        onClick={() => {
          if (column.fieldName !== "Categories0") {
            onColumnClick(column);
          }
        }}
        style={{
          flex: 1,
          cursor: "pointer",
        }}
      >
        {column.name}
      </span>

      <Icon
        iconName="Filter"
        onClick={() => setShowFilter(column, columnType)}
        style={{ cursor: "pointer" }}
      />
    </div>
  );
};

export { CellRender, HeaderRender };
