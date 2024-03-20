import { override } from '@microsoft/decorators';
// import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
// import { Dialog } from '@microsoft/sp-dialog';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
// import * as strings from 'SiteOwnerExtensionApplicationCustomizerStrings';
import './custom.css'
const LOG_SOURCE: string = 'SiteOwnerExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISiteOwnerExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SiteOwnerExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISiteOwnerExtensionApplicationCustomizerProperties> {
  public rootSiteCollectionUrl: string = "";
  public ownersToExclude: string[] = ['adm', 'admin', 'workflow'];


  @override
  public async onInit(): Promise<void> {
    let owners = await this.getSiteOwners();
    let urlParts = this.context.pageContext.site.absoluteUrl.split('/');
    this.rootSiteCollectionUrl = urlParts.slice(0, 3).join('/');

    //console.log(this.rootSiteCollectionUrl);
    this.renderSiteOwners(owners);
    return Promise.resolve();
  }

  private showHideSiteOwners(hideSiteOwners: boolean): void {
    if (hideSiteOwners) {
      const element: HTMLDivElement = document.querySelector('.ownerDivContainer');
      if (element) {
        element.style.display = 'none';
      }

      const element2: HTMLButtonElement = document.querySelector('.siteOwnerButton');
      if (element2) {
        element2.style.display = 'block';
      }
    } else {
      const element: HTMLDivElement = document.querySelector('.ownerDivContainer');
      if (element) {
        element.style.display = 'block';
      }

      const element2: HTMLButtonElement = document.querySelector('.siteOwnerButton');
      if (element2) {
        element2.style.display = 'none';
      }
    }
  }

  private renderSiteOwners(owners) {
    let siteOwnerButton: HTMLButtonElement = document.createElement('button');
    siteOwnerButton.textContent = 'Seitenbesitzer';
    siteOwnerButton.classList.add('siteOwnerButton');
    siteOwnerButton.addEventListener("click", () => this.showHideSiteOwners(false));

    let body = document.querySelector('body') as HTMLElement | null;
    let siteOwnerDivContainer: HTMLDivElement = document.createElement('div');
    siteOwnerDivContainer.classList.add('ownerDivContainer');
    const siteTitle: string = this.context.pageContext.web.title;
    const headingContainer: HTMLDivElement = document.createElement('div');
    headingContainer.classList.add('headingContainer');
    const heading: HTMLDivElement = document.createElement('div');
    heading.classList.add('heading');
    heading.innerHTML = 'Besitzer von ' + siteTitle;
    headingContainer.appendChild(heading);
    const closeButton: HTMLDivElement = document.createElement('div');
    closeButton.classList.add('closeButton');
    closeButton.addEventListener("click", () => this.showHideSiteOwners(true));
    headingContainer.appendChild(closeButton);
    siteOwnerDivContainer.appendChild(headingContainer);
    let ownerDivContainer: HTMLDivElement = document.createElement('div');
    ownerDivContainer.classList.add('ownerDivContainer2');


    for (let i: number = 0; i < owners.length; i++) {
      let excludeOwner: boolean = false;
      for (let n: number = 0; n < this.ownersToExclude.length; n++) {
        if (owners[i].Email.includes(this.ownersToExclude[n])) {
          excludeOwner = true;
        }
      }
      if (!excludeOwner) {
        let ownerItem: HTMLDivElement = document.createElement('div');
        ownerItem.classList.add('ownerItem');
        let profilPicture: HTMLImageElement = document.createElement('img');
        profilPicture.classList.add('profileImg');
        profilPicture.src = this.rootSiteCollectionUrl + "/_layouts/15/userphoto.aspx?size=L&accountname=" + owners[i].Email
        let displayName: HTMLDivElement = document.createElement('div');
        displayName.innerHTML = owners[i].Title;
        displayName.classList.add('displayName');
        ownerItem.appendChild(profilPicture);
        ownerItem.appendChild(displayName);
        ownerDivContainer.appendChild(ownerItem);
      }
    }
    siteOwnerDivContainer.appendChild(ownerDivContainer);
    body.appendChild(siteOwnerButton);
    body.appendChild(siteOwnerDivContainer);
  }

  private async getSiteOwners(): Promise<any[]> {
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups?$select=Title,Id`;
    const response = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
    const groups = await response.json();

    const ownerGroups = groups.value.filter(group => group.Title.toLowerCase().includes('owner') || group.Title.toLowerCase().includes('besitzer'));
    let owners = [];

    for (let group of ownerGroups) {
      const usersUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getbyid(${group.Id})/users?$select=Title,Email`;
      const usersResponse = await this.context.spHttpClient.get(usersUrl, SPHttpClient.configurations.v1);
      const usersJson = await usersResponse.json();
      owners = owners.concat(usersJson.value);
    }
    return owners
  }
}
