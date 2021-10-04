import * as React from "react";
import "office-ui-fabric-react/dist/css/fabric.css";
import { ITrainingExerciseProps } from "./ITrainingExerciseProps";
import { ITrainingExerciseState } from "./ITrainingExerciseState";
import SharepointService from "../../../services/Sharepoint/SharepointService";
import {
  SearchBox,
  PrimaryButton,
  Panel,
  PanelType,
} from "office-ui-fabric-react";

export default class TrainingExercise extends React.Component<
  ITrainingExerciseProps,
  ITrainingExerciseState
> {
  constructor(props: ITrainingExerciseProps) {
    super(props);
    this.search = this.search.bind(this);
    this.state = {
      items: [],
      searchInputValue: "",
      filtered: [],
      showPanel: false,
      dismissPanel: false,
      listOfPeople: [],
      filteredListOfPeople: [],
      hasResults: true,
      loading: false,
      department: "",
    };
  }

  public render(): React.ReactElement<ITrainingExerciseProps> {
    return (
      <div className="ms-Grid">
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-md10">
            <SearchBox
              className="react-search-box"
              placeholder="Search for departmen by name..."
              onChange={(event) =>
                this.setState({ searchInputValue: event.target.value })
              }
              value={this.state.loading ? "" : this.state.searchInputValue}
            ></SearchBox>
          </div>
          <div className="ms-Grid-col ms-u-md2">
            <PrimaryButton
              className="Primary"
              text={this.state.loading ? "Loading" : "Search"}
              onClick={this.search.bind(this.state.searchInputValue)}
              disabled={this.state.loading}
            ></PrimaryButton>
          </div>
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col">
            {!this.state.hasResults
              ? "No matching departments"
              : this.state.filtered.map((x) => {
                  return (
                    <>
                      <div key={x.Id.toString()}>
                        <p>
                          <strong>Department: </strong>
                          {x.Title}
                        </p>
                        <p>
                          <strong>Responsible person: </strong>
                          {x.Responsible_x0020_Person.Title}
                        </p>
                      </div>
                      <PrimaryButton
                        text="Details"
                        onClick={this._showPanel.bind(
                          x.Responsible_x0020_Person.Title,
                          x.Title
                        )}
                      ></PrimaryButton>
                      <Panel
                        headerText="Details"
                        type={PanelType.medium}
                        isOpen={this.state.showPanel}
                        onDismiss={this._hidePanel}
                        closeButtonAriaLabel="Close"
                      >
                        <p>
                          <strong>"Department name: "</strong>
                          {x.Title}
                        </p>
                        <p>
                          <strong>"Responsible person: "</strong>
                          {x.Responsible_x0020_Person.Title}
                        </p>
                        <strong>
                          "List of all people from the Department"
                        </strong>
                        {this.state.filteredListOfPeople.map((x) => {
                          return <p>{"Name: " + x.Title}</p>;
                        })}
                      </Panel>
                    </>
                  );
                })}
          </div>
        </div>
      </div>
    );
  }

  private async search(): Promise<void> {
    this.setState({ loading: true });
    await SharepointService.getListItems(
      "ab066fe5-1c54-443f-9fce-e5beae412a93",
      ["Responsible_x0020_Person"],
      ["Title", "Responsible_x0020_Person/Title", "Id"]
    ).then((x) => {
      this.setState({ items: x.value });
    });
    this.filterResult();
  }

  private filterResult(): void {
    if (!this.state.searchInputValue) {
      this.setState({ filtered: [], loading: false, hasResults: false });
      return;
    }
    this.setState({ department: this.state.searchInputValue });
    const filt = this.state.items.filter((item) =>
      item.Title.toLowerCase().includes(this.state.searchInputValue)
    );

    this.setState({
      filtered: filt,
      searchInputValue: "",
      loading: false,
      hasResults: true,
    });
    if (!filt.length) {
      this.setState({ hasResults: false });
    }
  }

  private _showPanel = () => {
    this.setState({ showPanel: true });
    this.getAllPeopleByDepartmentName();
  };

  private _hidePanel = () => {
    this.setState({ showPanel: false });
  };
  private async getAllPeopleByDepartmentName() {
    await SharepointService.getListItems(
      "32413ea7-de4c-4df5-bca3-97deb0058c54",
      [],
      ["Title", "Department"]
    ).then((x) => {
      this.setState({ listOfPeople: x.value });
      this.filterByDepartment();
    });
  }

  private filterByDepartment(): void {
    const filteredPeople = this.state.listOfPeople.filter((p) =>
      p.Department.toLowerCase().includes(this.state.department)
    );
    this.setState({ filteredListOfPeople: filteredPeople });
  }
}
