import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { Tree, TreeItem, TreeItemLayout } from "@fluentui/react-components";
import styles from "./TermSet.module.scss";

interface Term {
  Id: string;
  Name: string;
  HierarchyLevel: number;
  SetId: string;
  MainParentId?: string | null;
  ParentName?: string | null;
  Children?: Term[];
}

interface TermSet {
  setId: string;
  setName: string;
  terms: Term[];
}

interface TermSetListProps {
  context: WebPartContext;
  groupId: string;
  setNames: string[];
  updateCatagory: (value: string | null) => void;
}

const TermSetList: React.FC<TermSetListProps> = (props: TermSetListProps) => {
  const [termSets, setTermSets] = React.useState<TermSet[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [selectedTermId, setSelectedTermId] = React.useState<string>("");

  const { groupId, setNames, updateCatagory } = props;

  React.useEffect(() => {
    if (setNames && setNames.length > 0) {
      const fetchTerms = async (
        setId: string,
        parentTermId?: string,
        parentName?: string,
        mainParentId?: string
      ): Promise<Term[]> => {
        const termsUrl = parentTermId
          ? `${props.context.pageContext.web.absoluteUrl}/_api/v2.1/termStore/termSets('${setId}')/terms('${parentTermId}')/getlegacychildren`
          : `${props.context.pageContext.web.absoluteUrl}/_api/v2.1/termStore/termSets('${setId}')/getlegacychildren`;

        try {
          const response = await props.context.spHttpClient.get(
            termsUrl,
            SPHttpClient.configurations.v1
          );
          if (!response.ok) {
            throw new Error("Failed to fetch terms");
          }
          const termsData = await response.json();

          const terms = await Promise.all(
            termsData.value.map(async (term: any) => {
              const children: Term[] =
                term.childrenCount > 0
                  ? await fetchTerms(
                      setId,
                      term.id,
                      parentName ||
                        (term.labels.length > 0 ? term.labels[0].name : ""),
                      mainParentId || term.id
                    )
                  : [];

              return {
                Id: term.id,
                Name: term.labels.length > 0 ? term.labels[0].name : "",
                SetId: setId,
                MainParentId: mainParentId || term.id,
                ParentName:
                  parentName ||
                  (term.labels.length > 0 ? term.labels[0].name : ""),
                Children: children,
                HierarchyLevel: parentTermId ? 2 : 1,
              };
            })
          );

          return terms;
        } catch (error) {
          console.error(`Error fetching terms for set ${setId}:`, error);
          return [];
        }
      };

      const fetchData = async () => {
        try {
          const termSets: TermSet[] = [];

          // Assume group ID is known, replace with actual group ID

          for (const setName of setNames) {
            const encodedSetName = encodeURIComponent(setName);
            const setsApiUrl = `${props.context.pageContext.web.absoluteUrl}/_api/v2.1/termStore/termgroups('${groupId}')/termsets?$filter=localizedNames/any(n:n/name eq '${encodedSetName}')&$select=id,localizedNames`;

            const setsResponse = await props.context.spHttpClient.get(
              setsApiUrl,
              SPHttpClient.configurations.v1
            );

            if (!setsResponse.ok) {
              throw new Error("Failed to fetch term sets");
            }

            const setsData = await setsResponse.json();

            for (const set of setsData.value) {
              try {
                const terms = await fetchTerms(set.id);
                termSets.push({
                  setId: set.id,
                  setName: set.localizedNames[0].name,
                  terms: terms,
                });
              } catch (error) {
                console.error(`Error fetching terms for set ${set.id}:`, error);
              }
            }
          }

          setTermSets(termSets);
        } catch (error) {
          console.error("Error fetching term sets:", error);
          setTermSets([]); // Use empty array if term sets cannot be fetched
        } finally {
          setIsLoading(false);
        }
      };

      fetchData();
    } else {
      setTermSets([]);
    }
  }, [props.context, setNames, groupId]);

  const handleTermClick = (term: Term) => {
    setSelectedTermId(term.Id);
    if (term.ParentName) updateCatagory(term.ParentName);
  };

  const renderTreeItems = (terms: Term[]) => {
    return terms.map((term) => (
      <TreeItem
        key={term.Id}
        itemType={term.Children && term.Children.length > 0 ? "branch" : "leaf"}
      >
        <TreeItemLayout
          onClick={() => handleTermClick(term)}
          className={`${
            selectedTermId === term.Id ? styles.selectedItem : ""
          } ${styles.term_item}`}
        >
          {term.Name}
        </TreeItemLayout>
        {term.Children && term.Children.length > 0 && (
          <Tree>{renderTreeItems(term.Children)}</Tree>
        )}
      </TreeItem>
    ));
  };

  const renderTermSets = (sets: TermSet[]) => {
    return sets.map((set) => renderTreeItems(set.terms));
  };
  return (
    <div>
      {isLoading ? (
        <p>Loading term sets...</p>
      ) : (
        <div className="custom-nav">
          <div className={`${styles.knowledgeSetText}`}>Knowledge Bases</div>
          <div className={`${styles.termSet} `}>
            <Tree aria-label="Term Sets Tree">{renderTermSets(termSets)}</Tree>
          </div>
        </div>
      )}
    </div>
  );
};

export default TermSetList;
