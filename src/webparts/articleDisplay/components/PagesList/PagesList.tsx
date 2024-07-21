import * as React from "react";
import { ReusableDetailList } from "../common/ReusableDetailList";
import PagesService, { FilterDetail, IColumnInfo } from "./PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PagesColumns } from "./PagesColumns";
import { DefaultButton, IColumn } from "@fluentui/react";
import { makeStyles, useId, Input } from "@fluentui/react-components";
import styles from "./pages.module.scss";
import "./pages.css";
import { FilterPanelComponent } from "./PanelComponent";

export interface IPagesListProps {
  context: WebPartContext;
  catagory: string;
  selectedViewId: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    gap: "2px",
    maxWidth: "400px",
    alignItems: "center",
  },
});

const PagesList = (props: IPagesListProps) => {
  // Destructure the props
  const { context, catagory, selectedViewId } = props;

  /**
   * State variables for the component.
   */

  // Options for the page size dropdown
  const [pageSizeOption] = React.useState<number[]>([
    10, 15, 20, 40, 60, 80, 100,
  ]);

  const [columnInfos, setColumnInfos] = React.useState<IColumnInfo[]>([]);

  // The search text for filtering pages
  const [searchText, setSearchText] = React.useState<string>(""); // Initially set to empty string

  // The list of pages
  const [pages, setPages] = React.useState<any[]>([]); // Initially set to empty array

  // The initial list of pages
  const [initialPages, setInitialPages] = React.useState<any[]>([]); // Initially set to empty array

  // The paginated list of pages
  const [paginatedPages, setPaginatedPages] = React.useState<any[]>([]); // Initially set to empty array

  // The column to sort by
  const [sortBy, setSortBy] = React.useState<string>(""); // Initially set to empty string

  // The current page number
  const [currentPageNumber, setCurrentPageNumber] = React.useState<number>(1); // Initially set to 1

  // The total number of pages
  const [totalPages, setTotalPages] = React.useState<number>(1); // Initially set to 1

  // The number of items to display per page
  const [pageSize, setPageSize] = React.useState<number>(10); // Initially set to 10

  // The total number of items
  const [totalItems, setTotalItems] = React.useState<number>(0); // Initially set to 0

  // The sorting order
  const [isDecending, setIsDecending] = React.useState<boolean>(false); // Initially set to false

  // Whether to show the filter panel
  const [showFilter, setShowFilter] = React.useState<boolean>(false); // Initially set to false

  // The column to filter by
  const [filterColumn, setFilterColumn] = React.useState<string>(""); // Initially set to empty string
  // The type of column to filter by
  const [filterColumnType, setFilterColumnType] = React.useState<string>(""); // Initially set to empty string

  // The filter details
  const [filterDetails, setFilterDetails] = React.useState<FilterDetail[]>([]); // Initially set to empty array

  // Create an instance of the PagesService class
  const pagesService = new PagesService(context);

  // Get a unique id for the input field
  const inputId = useId("input");

  // Get the styles for the input field
  const inputStyles = useStyles();

  /**
   * Resets the filters by clearing the checked items and
   * calling the applyFilters function with an empty filter detail.
   */
  const resetFilters = () => {
    // Clear the filter details
    setFilterDetails([]);

    // Clear the search text
    setSearchText("");

    // Call the fetchPages function with the default arguments
    fetchPages(1, pageSize, "Created", true, "", catagory, []);
  };

  /**
   * Fetches the paginated pages based on the given parameters.
   *
   * @param {number} [page=1] - The page number to fetch. Defaults to 1.
   * @param {number} [pageSizeAmount=pageSize] - The number of items per page. Defaults to the `pageSize` state variable.
   * @param {string} [sortBy="Created"] - The column to sort by. Defaults to "Created".
   * @param {boolean} [isSortedDescending=isDecending] - Whether to sort in descending order. Defaults to the `isDecending` state variable.
   * @param {string} [searchText=""] - The search text to filter by. Defaults to an empty string.
   * @param {string} [category=catagory] - The category to filter by. Defaults to the `catagory` state variable.
   * @param {FilterDetail[]} filterDetails - The filter details to apply.
   *
   * @return {Promise<void>} - A promise that resolves when the paginated pages are fetched.
   */
  const fetchPages = (
    page = 1,
    pageSizeAmount = pageSize,
    sortBy = "Created",
    isSortedDescending = isDecending,
    searchText = "",
    category = catagory,
    filterDetails: FilterDetail[]
  ) => {
    const url = `${context.pageContext.web.serverRelativeUrl}/SitePages/${category}`;

    return (
      pagesService
        // Call the pagesService to fetch the filtered pages
        .getFilteredPages(
          selectedViewId,
          page,
          pageSizeAmount,
          sortBy,
          isSortedDescending,
          url,
          searchText,
          filterDetails
        )
        .then((res) => {
          // Set the total number of items in the response
          setTotalItems(res.length);

          // Calculate the total number of pages based on the page size
          const totalPages = Math.ceil(res.length / pageSizeAmount);

          // If there are no pages, set the total number of pages to 1
          if (totalPages === 0) {
            setTotalPages(1);
          } else {
            // Otherwise, set the total number of pages based on the page size
            setTotalPages(Math.ceil(res.length / pageSizeAmount));
          }

          // Slice the response to get the paginated pages for the current page number
          setPaginatedPages(res.slice(0, pageSizeAmount));

          // Set the state with all the pages
          setPages(res);

          // Return the response
          return res;
        })
    );
  };

  /**
   * Fetches the pages from the given path and filter categories
   * and updates the state with the initial pages
   * @param path - The path to the SitePages library
   */
  const getPages = async (path: string): Promise<void> => {
    // Get the initial pages from the API
    const initialPagesFromApi = await fetchPages(
      1,
      pageSize,
      "Created",
      true,
      searchText,
      path,
      filterDetails
    );

    // Update the state with the initial pages
    setInitialPages(initialPagesFromApi);
  };

  /**
   * Applies the given filter details to filter the pages
   *
   * @param {FilterDetail} filterDetail - The filter detail object containing the filter details
   */
  const applyFilters = (filterDetail: FilterDetail): void => {
    /**
     * Updates the current filter details state with the new filter detail,
     * or removes the filter detail if the values array is empty.
     *
     */

    // Initialize an empty array to store the current filter details
    let currentFilters: FilterDetail[] = [];

    // Check if the filterDetail values array has any elements
    if (filterDetail.values.length > 0) {
      // If it does, update or add filter detail for the specified column

      // Update the current filter details by filtering out any existing filter details
      // for the same column, and adding the new filter detail
      currentFilters = [
        ...filterDetails.filter(
          (item) => item.filterColumn !== filterDetail.filterColumn
        ),
        filterDetail,
      ];
    } else {
      // If the filterDetail values array is empty, remove any existing filter detail for the column
      currentFilters = filterDetails.filter(
        (item) => item.filterColumn !== filterDetail.filterColumn
      );
    }

    // Update the filter details state with the new filter details
    setFilterDetails(currentFilters);

    // Check if the filterDetail values array has any elements
    if (filterDetail.values.length > 0) {
      currentFilters = [
        ...filterDetails.filter((item) => item.filterColumn !== filterColumn),
        // Add the new filter detail with the specified filter column and values
        { filterColumn, filterColumnType, values: filterDetail.values },
      ];
    } else {
      currentFilters = filterDetails.filter(
        (item) => item.filterColumn !== filterColumn
      );
    }

    setFilterDetails(currentFilters); // Update filter details state
    fetchPages(
      1, // Page number
      pageSize, // Page size
      "Created", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      currentFilters // Updated filter details
    );
  };

  /**
   * Sort the pages list based on the specified column.
   *
   * @param {IColumn} column - The column to sort by.
   */
  const sortPages = (column: IColumn) => {
    // Set the sort by column state
    setSortBy(column.fieldName as string);

    // If the column is the same as the current sort by column, toggle the sort order
    if (column.fieldName === sortBy) {
      setIsDecending(!isDecending);
    } else {
      // Otherwise, set the sort order to descending
      setIsDecending(true);
    }

    // Fetch the pages list with the new sort criteria
    fetchPages(
      1, // Page number
      pageSize, // Page size
      column.fieldName, // Sorting criteria
      column.isSortedDescending, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      filterDetails // Filter details
    );
  };

  const handlePageChange = (page: number, pageSizeChanged = pageSize) => {
    // Ensure page is an integer
    const currentPage = Math.ceil(page);

    // Update current page number state
    setCurrentPageNumber(currentPage);

    // Calculate slice indices for pagination
    const startIndex = (currentPage - 1) * pageSizeChanged;

    const endIndex = currentPage * pageSizeChanged;

    // Slice the 'pages' array to get the current page of data
    const paginated = pages.slice(startIndex, endIndex);

    setTotalPages(Math.ceil(totalItems / pageSizeChanged));
    // Update paginated pages state
    setPaginatedPages(paginated);
  };

  /**
   * Handles the search functionality by fetching pages with specified parameters.
   */
  const handleSearch = () => {
    fetchPages(
      1, // Page number
      pageSize, // Page size
      "Created", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category
      filterDetails // Filter details
    );
  };

  /**
   * Navigates to the first page of paginated data.
   */
  const goToFirstPage = () => handlePageChange(1);

  /**
   * Navigates to the last page of paginated data.
   */
  const goToLastPage = () => handlePageChange(totalPages);

  /**
   * Navigates to the previous page of paginated data.
   * If the current page is the first page, no action is taken.
   */
  const goToPreviousPage = () =>
    handlePageChange(Math.max(currentPageNumber - 1, 1));

  /**
   * Navigates to the next page of paginated data.
   * If the current page is the last page, no action is taken.
   */
  const goToNextPage = () =>
    handlePageChange(Math.min(currentPageNumber + 1, totalPages));

  /**
   * Handles the input change event.
   * Parses the input value to an integer and calls handlePageChange with the parsed value.
   * If the input is not a number, calls handlePageChange with 0.
   *
   * @param e - The event object
   */
  const handleInputChange = (e: any) => {
    const inputValue = e.target.value;

    if (!isNaN(inputValue)) {
      const page = parseInt(inputValue, 10);
      handlePageChange(page);
    } else {
      handlePageChange(0);
    }
  };

  /**
   * Handles the change event of the page size dropdown.
   *
   * This function is triggered when the user selects a new page size from the dropdown.
   * It updates the page size state and calls the `handlePageChange` function to update
   * the paginated data.
   *
   * @function handlePageSizeChange
   * @memberof PagesList
   *
   * @param {any} e - The event object.
   * @return {void}
   */
  const handlePageSizeChange = (e: any) => {
    // Update the page size state
    setPageSize(e.target.value);
    // Handle the page change with the new page size
    handlePageChange(1, e.target.value);
  };

  /**
   * Dismisses the filter panel.
   * Sets the showFilter state to false.
   *
   * @function dismissPanel
   * @memberof PagesList
   * @returns {void}
   */
  const dismissPanel = (): void => {
    setShowFilter(false);
  };

  const getColumns = async (selectedViewId: string) => {
    const columns = await pagesService.getColumns(selectedViewId);
    setColumnInfos(columns);
  };

  React.useEffect(() => {
    getPages(catagory);
  }, [catagory]);

  React.useEffect(() => {
    console.log(402, selectedViewId);
    getColumns(selectedViewId);
  }, [selectedViewId]);

  return (
    <div className="w-pageSize0 detail-display">
      {showFilter && (
        <FilterPanelComponent
          isOpen={showFilter}
          headerText="Filter Articles"
          applyFilters={applyFilters}
          dismissPanel={dismissPanel}
          selectedItems={
            filterDetails.filter(
              (item) => item.filterColumn === filterColumn
            )[0] || { filterColumn: "", values: [] }
          }
          columnName={filterColumn}
          columnType={filterColumnType}
          pagesService={pagesService}
          data={initialPages}
        />
      )}
      <div className={`${styles.top}`}>
        <div
          className={`${styles["first-section"]} d-flex justify-content-between align-items-end py-2 px-2`}
        >
          <span className={`fs-4 ${styles["knowledgeText"]}`}>
            {catagory && <span className="">{catagory}</span>}
          </span>
          <div className={`${inputStyles.root} d-flex align-items-center me-2`}>
            <Input
              id={inputId}
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleSearch();
                }
              }}
              placeholder="Search"
            />
          </div>
        </div>

        <div
          className={`d-flex justify-content-between align-items-center fs-5 px-2 my-2`}
        >
          <span>Articles /</span>
          {totalItems > 0 ? (
            <div className="d-flex align-items-center">
              {filterDetails && filterDetails.length > 0 && (
                <DefaultButton
                  onClick={() => {
                    resetFilters();
                  }}
                >
                  Clear
                </DefaultButton>
              )}
              <span className="ms-2 fs-6">Results ({totalItems})</span>
            </div>
          ) : (
            <span className="fs-6">No articles to display</span>
          )}
        </div>
      </div>

      <ReusableDetailList
        items={paginatedPages}
        columns={PagesColumns}
        columnInfos={columnInfos}
        setShowFilter={(column: IColumn, columnType: string) => {
          setShowFilter(!showFilter);
          column.data;
          setFilterColumn(column.fieldName as string);
          setFilterColumnType(columnType);
        }}
        sortPages={sortPages}
        sortBy={sortBy}
        siteUrl={window.location.origin}
        isDecending={isDecending}
      />
      <div className="d-flex justify-content-end">
        <div
          className="d-flex align-items-center my-1"
          style={{
            fontSize: "13px",
          }}
        >
          <div className="d-flex align-items-center me-3">
            <span className={`me-2 ${styles.blueText}`}>Items / Page </span>
            <select
              className="form-select"
              value={pageSize}
              onChange={handlePageSizeChange}
              name="pageSize"
              style={{
                width: 80,
                height: 35,
              }}
            >
              {pageSizeOption.map((pageSize) => {
                return (
                  <option key={pageSize} value={pageSize}>
                    {pageSize}
                  </option>
                );
              })}
            </select>
          </div>
          <span className={`me-2 ${styles.blueText}`}>Page</span>
          <input
            type="text"
            value={currentPageNumber}
            onChange={handleInputChange}
            className="form-control"
            style={{
              width: 40,
              height: 35,
            }}
          />
          <span className="fs-6 mx-2">of {totalPages}</span>
          <span
            onClick={goToFirstPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber === 1 && styles.disabledPagination
            }`}
          >
            <i className="fa fa-step-backward" aria-hidden="true"></i>
          </span>
          <span
            onClick={goToPreviousPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber === 1 && styles.disabledPagination
            }`}
          >
            <i className="fa fa-caret-left" aria-hidden="true"></i>
          </span>
          <span
            onClick={goToNextPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber >= totalPages ? styles.disabledPagination : ""
            }`}
          >
            <i className="fa fa-caret-right" aria-hidden="true"></i>
          </span>
          <span
            onClick={goToLastPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber >= totalPages ? styles.disabledPagination : ""
            }`}
          >
            <i className="fa fa-step-forward" aria-hidden="true"></i>
          </span>
        </div>
      </div>
    </div>
  );
};

export default PagesList;
