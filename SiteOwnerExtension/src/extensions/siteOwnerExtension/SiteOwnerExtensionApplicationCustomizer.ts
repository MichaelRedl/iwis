import { override } from '@microsoft/decorators';
// import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
// import { Dialog } from '@microsoft/sp-dialog';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { MSGraphClient } from '@microsoft/sp-http';
// import * as strings from 'SiteOwnerExtensionApplicationCustomizerStrings';
import './custom.css'
import { app, initialize } from "@microsoft/teams-js";
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
  public ownersToExclude: string[] = ['adm', 'admin', 'workflow', 'System', 'Admin', 'Workflow'];
  public extensionAlreadyRendered: boolean = false;

  @override
  public async onInit(): Promise<void> {
    if (sessionStorage.getItem('showExtension') === null) {
      sessionStorage.setItem('showExtension', JSON.stringify(true));
    }



    await this.context.application.navigatedEvent.add(this, this.onNavigated);
    return Promise.resolve();
  }

  private async onNavigated() {
    await this.deleteElementsByClassName('ownerDivContainer');
    await this.deleteElementsByClassName('siteOwnerButton');
    let owners = await this.getSiteOwners();
    if (window.navigator.userAgent.indexOf('Teams') !== -1) {
      // Application runs inside TEAMS
    } else {
      for (let i: number = 0; i < owners.length; i++) {
        let dispName: string = owners[i][0];
        dispName = dispName.toLocaleLowerCase();
        if (dispName.indexOf("besitzer") !== -1 || dispName.indexOf("owner") !== -1) {
          let groupOwners = await this.getMembersOfO365GroupByEmail(owners[i][1]);
          owners.splice(i, 1);
          for (let n: number = 0; n < groupOwners.length; n++) {
            owners.splice(i, 0, groupOwners[n])
          }

        }
      }

      const emailTracker: { [email: string]: boolean } = {};

      // Filter the array, keeping only the first occurrence of each email
      const filteredOwners = owners.filter(([name, email]) => {
        if (emailTracker[email]) {
          return false;
        } else {
          emailTracker[email] = true;
          return true;
        }
      });


      owners = filteredOwners;






      let urlParts = this.context.pageContext.site.absoluteUrl.split('/');
      this.rootSiteCollectionUrl = urlParts.slice(0, 3).join('/');

      let languageID = await this.getSiteLanguage();
      let tempUrl: string = `${this.context.pageContext.web.absoluteUrl}`;
      // alert(window.navigator.userAgent);
      let inIframe: boolean = await this.inIframe();
      if (tempUrl.indexOf('sharepoint') !== -1 && inIframe === false) {
        this.renderSiteOwners(owners, languageID);
      }
    }
  }


  private async deleteElementsByClassName(className: string) {
    const elements = document.querySelectorAll(`.${className}`);
    elements.forEach((element) => {
      element.parentNode.removeChild(element);
    });
  }
  private showHideSiteOwners(hideSiteOwners: boolean): void {
    if (hideSiteOwners) {
      sessionStorage.setItem('showExtension', JSON.stringify(false));
      const elements = document.querySelectorAll('.ownerDivContainer');
      if (elements) {
        // element.style.display = 'none';
        for (let i: number = 0; i < elements.length; i++) {
          (elements[i] as HTMLElement).style.display = 'none';
        }
      }

      const elements2 = document.querySelectorAll('.siteOwnerButton');
      if (elements2) {
        for (let i: number = 0; i < elements2.length; i++) {
          (elements2[i] as HTMLButtonElement).style.display = 'block';
        }
      }
    } else {
      sessionStorage.setItem('showExtension', JSON.stringify(true));
      const elements = document.querySelectorAll('.ownerDivContainer');
      if (elements) {
        // element.style.display = 'none';
        for (let i: number = 0; i < elements.length; i++) {
          (elements[i] as HTMLElement).style.display = 'block';
        }
      }

      const elements2 = document.querySelectorAll('.siteOwnerButton');
      if (elements2) {
        for (let i: number = 0; i < elements2.length; i++) {
          (elements2[i] as HTMLButtonElement).style.display = 'none';
        }
      }
    }
  }

  private renderSiteOwners(owners, langaugeID) {
    const showExtension = JSON.parse(sessionStorage.getItem('showExtension'));
    let siteOwnerButton: HTMLButtonElement = document.createElement('button');
    if (langaugeID === 1031) {
      siteOwnerButton.textContent = 'Seitenbesitzer';
    } else {
      siteOwnerButton.textContent = 'Siteowner';
    }
    siteOwnerButton.classList.add('siteOwnerButton');
    if(showExtension){
      siteOwnerButton.style.display = 'none';
    } else {
      siteOwnerButton.style.display = 'block';
    }
    siteOwnerButton.addEventListener("click", () => this.showHideSiteOwners(false));

    // let body: HTMLElement =  document.querySelector('.SPPageChrome');
    let body = document.querySelector('body') as HTMLElement | null;
    let siteOwnerDivContainer: HTMLDivElement = document.createElement('div');
    siteOwnerDivContainer.classList.add('ownerDivContainer');
    if(showExtension){
      siteOwnerDivContainer.style.display = 'block';
    } else {
      siteOwnerDivContainer.style.display = 'none';
    }
    const siteTitle: string = this.context.pageContext.web.title;
    const headingContainer: HTMLDivElement = document.createElement('div');
    headingContainer.classList.add('headingContainer');
    const heading: HTMLDivElement = document.createElement('div');
    heading.classList.add('heading');
    if (langaugeID === 1031) {
      heading.innerHTML = 'Seitenbesitzer';
    } else {
      heading.innerHTML = 'Siteowner';
    }
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
        if (owners[i][1].includes(this.ownersToExclude[n]) || owners[i][0].includes(this.ownersToExclude[n])) {
          excludeOwner = true;
        }
      }
      if (!excludeOwner) {
        let ownerItem: HTMLDivElement = document.createElement('div');
        ownerItem.classList.add('ownerItem');
        let profilPicture: HTMLImageElement = document.createElement('img');
        profilPicture.classList.add('profileImg');
        profilPicture.src = this.rootSiteCollectionUrl + "/_layouts/15/userphoto.aspx?size=L&accountname=" + owners[i][1]
        let displayName: HTMLAnchorElement = document.createElement('a');
        displayName.innerHTML = owners[i][0];
        displayName.href = 'mailto:' + owners[i][1];
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

    const roleAssignResponse = await this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/roleassignments?$expand=RoleDefinitionBindings`, SPHttpClient.configurations.v1);
    const roleAssignData = await roleAssignResponse.json();
    // const fullControlGroups = roleAssignData.value.filter(ra => ra.RoleDefinitionBindings.some(rdb => rdb.Id === 1073741829));

    let ownerGroups = [];
    let stopLoop: boolean = false;
    for (let n: number = 0; n < groups.value.length; n++) {
      if (stopLoop === false) {
        if (groups.value[n].Title.toLowerCase().includes('besitzer') || groups.value[n].Title.toLowerCase().includes('owner')) {
          ownerGroups = groups.value.filter(group => group.Id === groups.value[n].Id);
          stopLoop = true;
        }
      }
    }



    /*  for(let i: number = 0; i < fullControlGroups.length; i++){
        for(let n: number = 0; n < groups.value.length; n++){
          if(fullControlGroups[i].PrincipalId === groups.value[n].Id){
            if(groups.value[n].Title.toLowerCase().includes('besitzer') || groups.value[n].Title.toLowerCase().includes('owner')){
              ownerGroups = groups.value.filter(group => group.Id === groups.value[n].Id);
            }
          }
        }
      }*/


    // const ownerGroups = groups.value.filter(group => group.Title.toLowerCase().includes('owner') || group.Title.toLowerCase().includes('besitzer'));

    //   ownerGroups = groups.value.filter(group => group.Id === 3);


    /* let url: string = this.context.pageContext.web.absoluteUrl;
     const parts = url.split('/');
     const urlName = parts.pop();
     
     if(ownerGroups.length === 0){
       ownerGroups = groups.value.filter(group => group.Title.toLowerCase().includes('owner') || group.Title.toLowerCase().includes('besitzer'));
       ownerGroups = ownerGroups.filter(group => group.Title.toLowerCase().includes(' '+urlName.toLowerCase()))
     }
 */
    //let ownerGroups = [];


    let owners = [];

    for (let group of ownerGroups) {
      const usersUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getbyid(${group.Id})/users?$select=Title,Email`;
      const usersResponse = await this.context.spHttpClient.get(usersUrl, SPHttpClient.configurations.v1);
      const usersJson = await usersResponse.json();
      owners = owners.concat(usersJson.value);
    }



    let ownerNamesAndEmails: string[][] = [];
    for (let i: number = 0; i < owners.length; i++) {
      let name: string = owners[i].Title;
      let mail: string = owners[i].Email;
      let arrayy: string[] = [name, mail];
      ownerNamesAndEmails.push(arrayy);
    }
    return ownerNamesAndEmails

  }


  private async getMembersOfO365GroupByEmail(groupEmail: string): Promise<any[]> {
    try {
      const client = await this.context.msGraphClientFactory.getClient();
      const groupsRes = await client.api(`/groups?$filter=mail eq '${groupEmail}'`).get();
      if (groupsRes && groupsRes.value && groupsRes.value.length > 0) {
        const groupId = groupsRes.value[0].id;
        const membersRes = await client.api(`/groups/${groupId}/owners`).get();

        let groupOwners: string[][] = [];
        for (let i: number = 0; i < membersRes.value.length; i++) {
          let excludeOwner: boolean = false;
          for (let n: number = 0; n < this.ownersToExclude.length; n++) {
            if (membersRes.value[i].displayName.includes(this.ownersToExclude[n]) || membersRes.value[i].mail.includes(this.ownersToExclude[n])) {
              excludeOwner = true;
            }
          }
          if (!excludeOwner) {
            let tempArr: string[] = [membersRes.value[i].displayName, membersRes.value[i].mail];
            groupOwners.push(tempArr);
          }
        }


        return groupOwners; // Assuming the members are in the `value` property
      } else {
        console.log("No group found with that email.");
        return [];
      }
    } catch (error) {
      console.error("Error fetching group or members:", error);
      return [];
    }
  }

  private async inIframe(): Promise<any> {
    try {
      return window.self !== window.top;
    } catch (e) {
      return true;
    }
  }

  private async getSiteLanguage(): Promise<void> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/regionalSettings", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json) => {
        let lcid = json.LocaleId;
        return lcid;
        // You can return this language value or use it as needed
      })
      .catch((error) => {
        console.error("Error fetching site language:", error);
      });
  }
}
