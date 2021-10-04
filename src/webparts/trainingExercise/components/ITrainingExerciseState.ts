import { IListItems } from "../../../services/Sharepoint/IListItems";

export interface ITrainingExerciseState {
  items: IListItems[];
  searchInputValue: string;
  filtered: IListItems[];
  showPanel: boolean;
  dismissPanel: boolean;
  listOfPeople: IListItems[];
  filteredListOfPeople: IListItems[];
  hasResults: boolean;
  loading:boolean;
  department:string;
}
