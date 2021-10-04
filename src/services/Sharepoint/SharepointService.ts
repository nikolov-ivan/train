import { EnvironmentType } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IListFieldCollection } from "./IListField";
import { IListItemsCollection } from "./IListItems";

export class SharepointServiceManager {
  public context: WebPartContext;
  public enviromentType: EnvironmentType;

  public setup(context: WebPartContext, enviromentType: EnvironmentType): void {
    this.context = context;
    this.enviromentType = enviromentType;
  }

  public get(relativeEndPointUrl: string): Promise<any> {
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}${relativeEndPointUrl}`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        return response.json();
      })
      .catch((error) => {
        return Promise.reject(error);
      });
  }

  public getListItems(
    listId: string,
    expandFields?: string[],
    selectedFields?: string[]
  ): Promise<IListItemsCollection> {
    return this.get(
      `/_api/lists/getbyid('${listId}')/items${
        expandFields ? `?$expand=${expandFields.join(",")}` : ""
      }${selectedFields ? `&$select=${selectedFields.join(",")}` : ""}`
    );
  }

  public getListFields(
    listId: string,
    showHiddenFields: boolean = false
  ): Promise<IListFieldCollection> {
    return this.get(
      `/_api/lists/getbyid('${listId}')/fields${
        !showHiddenFields ? "?$filter=Hidden eq false" : ""
      }`
    );
  }
}

const SharepointService = new SharepointServiceManager();

export default SharepointService;
