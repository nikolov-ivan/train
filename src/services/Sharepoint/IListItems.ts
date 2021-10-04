export interface IListItems {
  Title: string;
  Department: string;
  Responsible_x0020_Person: {
    Title: string;
  };
  Id: Int16Array;
}

export interface IListItemsCollection {
  value: IListItems[];
}
