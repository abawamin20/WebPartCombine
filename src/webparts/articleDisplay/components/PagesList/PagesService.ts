import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/fields";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { CellRender } from "../common/ColumnDetails";
import { getColumnMaxWidth, getColumnMinWidth } from "../utils/columnUtils";
import { ConstructedFilter } from "./PanelComponent";

export interface ITerm {
  Id: string;
  Name: string;
  parentId: string;
  Children?: ITerm[];
}

export interface TermSet {
  setId: string;
  terms: ITerm[];
}

export interface FilterDetail {
  filterColumn: string;
  filterColumnType: string;
  values: string[];
}

export interface IColumnInfo {
  InternalName: string;
  DisplayName: string;
  MinWidth: number;
  ColumnType: string;
  MaxWidth: number;
  OnRender?: (items: any) => JSX.Element;
}
class PagesService {
  private _sp: SPFI;

  constructor(private context: WebPartContext) {
    this._sp = spfi().using(SPFx(this.context));
  }
  /**
   * Fetch distinct values for a given column from a list of items.
   * @param {string} columnName - The name of the column to fetch distinct values for.
   * @param {any[]} values - The list of items to extract distinct values from.
   * @returns {Promise<string[] | ConstructedFilter[]>} - A promise that resolves to an array of distinct values.
   */
  getDistinctValues = async (
    columnName: string,
    columnType: string,
    values: any
  ): Promise<(string | ConstructedFilter)[]> => {
    try {
      const items = values; // The list of items to fetch distinct values from.

      // Extract distinct values from the column
      const distinctValues: (string | ConstructedFilter)[] = [];
      const seenValues = new Set<string | ConstructedFilter>(); // A set to keep track of seen values to avoid duplicates.

      items.forEach((item: any) => {
        switch (columnType) {
          case "TaxonomyFieldTypeMulti":
            if (item.TaxCatchAll && item.TaxCatchAll.length > 0) {
              // Extract distinct values from the TaxCatchAll column
              item.TaxCatchAll.forEach((category: any) => {
                const uniqueValue = category.Term;
                if (!seenValues.has(uniqueValue)) {
                  seenValues.add(uniqueValue);
                  distinctValues.push(uniqueValue);
                }
              });
            }
            break;
          case "DateTime":
            let uniqueDateValue = item[columnName]; // The value of the column for the current item.
            // Handle ISO date strings by extracting only the date part
            uniqueDateValue = new Date(uniqueDateValue)
              .toISOString()
              .split("T")[0];

            if (!seenValues.has(uniqueDateValue)) {
              seenValues.add(uniqueDateValue);
              distinctValues.push(uniqueDateValue);
            }
            break;
          case "User":
            const userValue = item[columnName];
            if (
              userValue &&
              userValue.Title &&
              !seenValues.has(userValue.Title)
            ) {
              seenValues.add(userValue.Title);
              const user: ConstructedFilter = {
                text: userValue.Title,
                value: userValue.Id,
              };
              distinctValues.push(user);
            }
            break;
          case "Number":
            const uniqueNumberValue = item[columnName]; // The value of the column for the current item.
            console.log(uniqueNumberValue);
            if (!seenValues.has(uniqueNumberValue)) {
              seenValues.add(uniqueNumberValue);
              distinctValues.push(uniqueNumberValue);
            }
            break;
          default:
            const uniqueValue = item[columnName]; // The value of the column for the current item.
            if (!seenValues.has(uniqueValue)) {
              seenValues.add(uniqueValue);
              distinctValues.push(uniqueValue);
            }
            break;
        }
      });

      return distinctValues;
    } catch (error) {
      console.error("Error fetching distinct values:", error);
      throw error;
    }
  };

  /**
   * Retrieves a page of filtered Site Pages items.
   *
   * @param viewId The selected view id
   * @param pageNumber The page number to retrieve (1-indexed).
   * @param pageSize The number of items to retrieve per page. Defaults to 10.
   * @param orderBy The column to sort the items by. Defaults to "Created".
   * @param isAscending Whether to sort in ascending or descending order. Defaults to true.
   * @param folderPath The folder path to search in. Defaults to "" (root of the site).
   * @param searchText Text to search for in the Title, Article ID, or Modified columns.
   * @param filters An array of FilterDetail objects to apply to the query.
   * @returns A promise that resolves with an array of items.
   */
  getFilteredPages = async (
    viewId: string,
    pageNumber: number,
    pageSize: number = 10,
    orderBy: string = "Created",
    isAscending: boolean = true,
    folderPath: string = "",
    searchText: string = "",
    filters: FilterDetail[]
  ) => {
    try {
      const skip = (pageNumber - 1) * pageSize;
      const list = this._sp.web.lists.getByTitle("Site Pages");
      const fields = await list.views.getById(viewId).fields();
      /**
       * Generates a filter query string based on the provided filters.
       *
       */
      let filterQuery = `startswith(FileDirRef, '${folderPath}') and FSObjType eq 0${
        searchText
          ? ` and (substringof('${searchText}', Title) or Article_x0020_ID eq '${searchText}' or substringof('${searchText}', Modified))`
          : ""
      }`;

      // Append filter conditions based on the provided filters.
      filters.forEach((filter) => {
        if (filter.values.length > 0) {
          switch (filter.filterColumnType) {
            case "TaxonomyFieldTypeMulti":
              // Append category filters to the filter query.
              const categoryFilters = filter.values
                .map((value) => `TaxCatchAll/Term eq '${value}'`)
                .join(" or ");
              filterQuery += ` and (${categoryFilters})`;
              break;

            case "DateTime":
              // Append date filters to the filter query.
              const dateFilters = filter.values
                .map((value) => {
                  // Generate start and end dates for the filter.
                  const startDate = new Date(value);
                  const endDate = new Date(value);
                  endDate.setDate(endDate.getDate() + 1); // Include the entire day

                  return `${
                    filter.filterColumn
                  } ge datetime'${startDate.toISOString()}' and ${
                    filter.filterColumn
                  } lt datetime'${endDate.toISOString()}'`;
                })
                .join(" or ");
              filterQuery += ` and (${dateFilters})`;
              break;
            case "User":
              // Append user filters to the filter query.
              const userFilters = filter.values
                .map(
                  (value) =>
                    `${filter.filterColumn}/Id eq '${value}' or Editor/Id eq '${value}'`
                )
                .join(" or ");
              filterQuery += ` and (${userFilters})`;
              break;
            default:
              // Append column filters to the filter query.
              const columnFilters = filter.values
                .map((value) => `${filter.filterColumn} eq '${value}'`)
                .join(" or ");
              filterQuery += ` and (${columnFilters})`;
              break;
          }
        }
      });

      /**
       * Retrieves the items from the SharePoint list based on the provided filter query,
       * selects specific columns, expands the TaxCatchAll field, applies pagination,
       * and orders the results.
       *
       */
      const pages: any[] = await list.items
        .filter(filterQuery)
        .select(
          // Select the required columns
          "Title",
          "Description",
          "FileLeafRef",
          "FileRef",
          "Modified",
          "Id",
          "TaxCatchAll/Term",
          "Article_x0020_ID",
          "Author/Title",
          "Editor/Title",
          "Author/Id",
          "Editor/Id",
          ...fields.Items
        )
        .expand("TaxCatchAll", "Author", "Editor") // Expand the TaxCatchAll field to get the Term value
        .skip(skip) // Apply pagination by skipping the specified number of items
        .orderBy(orderBy, isAscending)(); // Order the results based on the specified column and sort order

      return pages;
    } catch (error) {
      console.error("Error fetching filtered pages:", error);
      throw new Error("Error fetching filtered pages");
    }
  };

  /**
   * Retrieves the columns for a specified view in the SharePoint list.
   */
  public async getColumns(viewId: string): Promise<IColumnInfo[]> {
    const fields = await this._sp.web.lists
      .getByTitle("Site Pages")
      .views.getById(viewId)
      .fields();

    // Fetching detailed field information to get both internal and display names
    const fieldDetailsPromises = fields.Items.map((field: any) =>
      this._sp.web.lists
        .getByTitle("Site Pages")
        .fields.getByInternalNameOrTitle(field)()
    );

    const fieldDetails = await Promise.all(fieldDetailsPromises);

    return fieldDetails.map((field: any) => ({
      InternalName: field.InternalName,
      DisplayName: field.Title,
      ColumnType: field.TypeAsString,
      MinWidth: getColumnMinWidth(field.InternalName),
      MaxWidth: getColumnMaxWidth(field.InternalName),
      OnRender: (item: any) =>
        CellRender({
          columnName: field.InternalName,
          columnType: field.TypeAsString,
          item,
        }),
    }));
  }
}

export default PagesService;
