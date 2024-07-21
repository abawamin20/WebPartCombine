import * as React from "react";
import { SPComponentLoader } from "@microsoft/sp-loader";

import type { IArticleDisplayProps } from "./IArticleDisplayProps";
import TermSetList from "./SideNavigation/TermSet";
import PagesList from "./PagesList/PagesList";

export default class ArticleDisplay extends React.Component<
  IArticleDisplayProps,
  {
    catagory: string;
  }
> {
  constructor(props: IArticleDisplayProps) {
    super(props);
    this.state = {
      catagory: "",
    };
  }
  componentDidMount(): void {
    const cssURL =
      "https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css";
    SPComponentLoader.loadCss(cssURL);
    SPComponentLoader.loadCss(
      "https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css"
    );
  }
  public render(): React.ReactElement<IArticleDisplayProps> {
    const { context, groupId, setNames, selectedView } = this.props;

    return (
      <div className="row">
        <div className="col-3">
          <TermSetList
            context={context}
            groupId={groupId}
            setNames={setNames}
            updateCatagory={(value: string) =>
              this.setState({ catagory: value })
            }
          />
        </div>
        <div className="col-9">
          <PagesList
            context={context}
            catagory={this.state.catagory}
            selectedViewId={selectedView}
          />
        </div>
      </div>
    );
  }
}
