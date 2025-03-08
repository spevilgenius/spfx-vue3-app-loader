/* eslint-disable no-prototype-builtins */
/* eslint-disable promise/param-names */
/* eslint-disable @typescript-eslint/no-this-alias */
/* eslint-disable dot-notation */
/* eslint-disable @microsoft/spfx/import-requires-chunk-name */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle, PropertyPaneDropdown, IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import axios from "axios";

import { SPPermission } from "@microsoft/sp-page-context";
import { DisplayMode } from "@microsoft/sp-core-library";
import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/sites";
import "@pnp/sp/site-users/web";
import "@pnp/graph/mail";
import "@pnp/graph/users";
import "@pnp/graph/calendars";

export interface IVueAppLoaderWebPartProps {
  description: string;
  title: string;
  externalWidget?: string; // name of the widget
  cssFile: string; // css file of the widget
  jsFile: string; // javascript file of the widget
  widgetns: string; // namespace exported by the widget
  widgetsettings: string; // A JSON string that holds the widget settings
  removePadding: boolean;
  editmode: boolean;
  siteurl: string;
  weburl: string;
  siteid: string;
  webid: string;
  instanceId: string;
  fullcontrol: boolean;
  displayname: string;
  loginname: string;
  email: string;
  siteadmin: boolean;
}

interface VueAppControl {
  render: (
    container: Element,
    props: any
  ) => {
    unmount: () => void;
    update: (newProps: any) => void;
  };
}

export default class VueAppLoaderWebPart extends BaseClientSideWebPart<IVueAppLoaderWebPartProps> {
  private basews: string = "Widget Settings Placeholder";
  private widgets: IPropertyPaneDropdownOption[];
  public widgetlibraryurl: string = "https://legodan.sharepoint.com/sites/LegoTeam/_api/web/lists/getbytitle('ScriptWebParts')/items?$select=*";
  public widgetinfos = new Array<any>();
  private styleloaded: boolean = false;
  private scriptloaded: boolean = false;
  private vueInstance: { update: (newProps: any) => void; unmount: () => void } | null = null;
  private loadingPromise: Promise<any> | null = null;
  private appID: string;
  private sp: SPFI;
  private graph: GraphFI;

  private isPropertyPaneOpen: boolean = false;

  protected async onInit(): Promise<void> {
    this.appID = "APP_" + this.context.instanceId;
    this.properties.widgetsettings = this.properties.widgetsettings && this.properties.widgetsettings.length > 0 ? this.properties.widgetsettings : this.basews;
    this.properties.fullcontrol = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    this.properties.displayname = this.context.pageContext.user.displayName;
    this.properties.loginname = this.context.pageContext.user.loginName;
    this.properties.email = this.context.pageContext.user.email;
    this.properties.siteadmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
    this.properties.instanceId = this.context.instanceId;
    this.properties.siteurl = this.context.pageContext.site.absoluteUrl;
    this.properties.weburl = this.context.pageContext.web.absoluteUrl;
    this.properties.siteid = this.context.pageContext.site.id.toString();
    this.properties.webid = this.context.pageContext.web.id.toString();
    this.sp = spfi().using(spSPFx(this.context));
    this.graph = graphfi().using(graphSPFx(this.context));
    this.loadingPromise = this.loadExternalResources();
    return super.onInit();
  }

  private async loadExternalResources(): Promise<void> {
    if (this.properties.jsFile && this.properties.jsFile.length > 0) {
      if (this.properties.widgetns && this.properties.widgetns.length > 0) {
        try {
          const timeoutms = 30000;
          const timeoutPromise = new Promise<never>((_, reject) => {
            setTimeout(() => reject(new Error("Timeout: Initialization took to long")), timeoutms);
          });
          await this.loadStylesheet(this.properties.cssFile)
            .then(async () => {
              this.styleloaded = true;
            })
            .catch((error) => console.error("Style load error: " + error));
          await Promise.race([
            SPComponentLoader.loadScript(this.properties.jsFile, {
              globalExportsName: this.properties.widgetns,
            }).then(() => {
              this.scriptloaded = true;
            }),
            timeoutPromise,
          ]);
          this.loadingPromise = null;
          return;
        } catch (error) {
          console.error("ERROR LOADING RESOURCES: ", error);
          this.loadingPromise = null;
          throw error;
        }
      }
    }
  }

  public async render(): Promise<void> {
    if (this.properties.jsFile && this.properties.jsFile.length > 0) {
      if (this.properties.widgetns && this.properties.widgetns.length > 0) {
        this.domElement.innerHTML = `
          <div id="SLWP_${this.properties.instanceId}">
            <div id="loadingspinner" style="text-align: center; padding: 20px;">
              <svg width="40" height="40" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                <circle cx="12" cy="12" r="0" fill="#0078d4">
                  <animate attributeName="r" values="0; 8; 0" dur="1.5s" repeatCount="indefinite" begin="0" />
                  <animate attributeName="opacity" values="1; .5; 1" dur="1.5s" repeatCount="indefinite" begin="0" />
                </circle>
              </svg>
              <p>Loading Component</p>
            </div>
            <div id="${this.appID}">
          </div>
        `;

        if (this.loadingPromise) {
          await this.loadingPromise
            .then(() => {
              return this.initializeApp();
            })
            .then(() => {
              const loadingelement = document.getElementById("loadingspinner");
              if (loadingelement) {
                loadingelement.style.display = "none";
              }
            })
            .catch((error) => {
              console.error("FAILED TO LOAD RESOURCES AND INITIALIZE VUE APP:", error);
            });
        } else if (this.styleloaded && this.scriptloaded) {
          await this.initializeApp()
            .then(() => {
              const loadingelement = document.getElementById("loadingspinner");
              if (loadingelement) {
                loadingelement.style.display = "none";
              }
            })
            .catch((error) => {
              console.error("RESOURCES LOADED BUT FAILED TO INITIALIZE VUE APP:", error);
            });
        } else {
          // No promise exists and resources are not loaded so start loading again
          this.loadingPromise = this.loadExternalResources();
          await this.loadingPromise
            .then(() => {
              return this.initializeApp();
            })
            .then(() => {
              const loadingelement = document.getElementById("loadingspinner");
              if (loadingelement) {
                loadingelement.style.display = "none";
              }
            })
            .catch((error) => {
              console.error("FAILED TO LOAD RESOURCES AND INITIALIZE VUE APP:", error);
            });
        }

        if (this.displayMode === DisplayMode.Edit) {
          this.properties.editmode = true;
        }
        if (this.displayMode === DisplayMode.Read) {
          this.properties.editmode = false;
        }
      }
    } else {
      this.domElement.innerHTML = "NOTHING TO RENDER UNTIL A PROPER WIDGET IS SELECTED.";
    }
  }

  private async initializeApp(): Promise<void> {
    const timeoutPromise = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error("Timeout: Initialization took to long")), 10000);
    });
    try {
      // Racing the timeout
      await Promise.race([this.initAppImplementation(), timeoutPromise]);

      if (!window["Vue"] || !window[this.properties.widgetns]) {
        console.error("initializeApp (PROMISE) - CHECK A");
        return;
      }

      const appContainer = document.getElementById(this.appID) as Element;
      if (!appContainer) {
        console.error("initializeApp (PROMISE) - CHECK B");
        return;
      }
    } catch (error) {
      console.error("initializeApp (PROMISE) - CATCH ERROR INITIALIZING APP: " + error);
      throw error;
    }
  }

  private async initAppImplementation(): Promise<void> {
    const appContainer = document.getElementById(this.appID) as Element;
    const vueLibrary: VueAppControl = window[this.properties.widgetns];

    const pnp = {
      sp: this.sp,
      graph: this.graph,
      testSP: async () => {
        try {
          // sp connection
          const user = await this.sp.web.currentUser();
          console.log("SP Test Successful - Current User:", user);
          return true;
        } catch (error) {
          console.error("SP Test Failed:", error);
        }
      },
      testGRAPH: async () => {
        try {
          // graph connection
          const me = await this.graph.me();
          console.log("GRAPH Test Successful - Me:", me);
          return true;
        } catch (error) {
          console.error("GRAPH Test Failed:", error);
        }
      },
    };

    if (vueLibrary) {
      this.vueInstance = vueLibrary.render(appContainer, {
        pnp: pnp,
        editmode: this.displayMode === DisplayMode.Edit ? true : false,
        isEditing: this.isPropertyPaneOpen,
        title: this.properties.title,
        siteurl: this.properties.siteurl,
        weburl: this.properties.weburl,
        siteadmin: this.properties.siteadmin,
        displayname: this.properties.displayname,
        instanceid: this.properties.instanceId,
        email: this.properties.email,
        loginname: this.properties.loginname,
        fullcontrol: this.properties.fullcontrol,
        siteid: this.properties.siteid,
        webid: this.properties.webid,
        widgetsettings: this.properties.widgetsettings,
        onPropertyChanged: (propertyName: string, newValue: any) => {
          if (propertyName in this.properties) {
            this.properties[propertyName] = newValue;
            requestAnimationFrame(() => {
              if (this.context && this.context.propertyPane) {
                this.context.propertyPane.refresh();
              }
            });
          }
        },
      });

      const loadingelement = document.getElementById("loadingspinner");
      if (loadingelement) {
        loadingelement.style.display = "none";
      }
    } else {
      console.error("SOMEHOW THE LIBRARY IS STILL NOT WORKING");
    }
  }

  public loadStylesheet(url: string): Promise<void> {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`link[href="${url}"]`)) {
        resolve();
        return;
      }

      const link = document.createElement("link");
      link.type = "text/css";
      link.href = url;
      link.rel = "stylesheet";
      link.onload = () => resolve();
      link.onerror = (e) => reject(new Error(`Failed to load stylesheet: ${url}`));

      document.head.appendChild(link);
    });
  }

  private async getWidgets(): Promise<IPropertyPaneDropdownOption[]> {
    const k = new Array<IPropertyPaneDropdownOption>();
    const response = await axios.get(this.widgetlibraryurl, {
      headers: {
        accept: "application/json;odata=verbose",
      },
    });
    // console.log('GET WIDGETS RESPONSE: ' + response)
    // Ensure the current user is a site admin. This allows the user to test development applications.
    this.properties.siteadmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
    // This expects the following fields to be in the list: Title, Status, JSfile, CSSFile, WidgetNS, Details
    // Some fields may not be needed depending on what you are loading but Title is required. WidgetNS, JSFile, and CSSFile are required to load vue widgets
    // These fields define the widget and the exported namespace [WidgetNS] from your vue app. Of course you can change this around depending on your needs
    // We are using two arrays because the property pane dropdown [IPropertyPaneDropdownOption] does not support the extra columns
    // so we load the objects into the dropdown properties and use the widgetinfos array to store the other columns
    const j = response.data.d.results;
    for (let i = 0; i < j.length; i++) {
      const status = j[i].Status;
      const widgetns = j[i].WidgetNS;
      if (status === "Production") {
        k.push({
          index: i,
          key: j[i].Title,
          text: j[i].Title,
        });
        this.widgetinfos.push({
          index: i,
          title: j[i].Title,
          details: j[i].Details,
          status: status,
          widgetns: widgetns,
          jsFile: j[i].JSFile,
          cssFile: j[i].CSSFile,
        });
      } else {
        if (this.properties.siteadmin === true) {
          k.push({
            index: i,
            key: j[i].Title,
            text: j[i].Title,
          });
          this.widgetinfos.push({
            index: i,
            title: j[i].Title,
            details: j[i].Details,
            status: status,
            widgetns: widgetns,
            jsFile: j[i].JSFile,
            cssFile: j[i].CSSFile,
          });
        }
      }
    }
    return k;
  }

  protected onDispose(): void {
    if (this.vueInstance) {
      this.vueInstance.unmount();
      this.vueInstance = null;
    }
    this.styleloaded = false;
    this.scriptloaded = false;
    this.loadingPromise = null;
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.isPropertyPaneOpen = true;
    this.properties.siteadmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
    const wopts: IPropertyPaneDropdownOption[] = await this.getWidgets();
    this.widgets = wopts;
    this.context.propertyPane.refresh();
    if (this.vueInstance) {
      this.vueInstance.update({
        isEditing: true,
        editmode: true,
      });
    }
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === "externalWidget") {
      for (let i = 0; i < this.widgetinfos.length; i++) {
        if (this.widgetinfos[i].title === newValue) {
          this.properties.jsFile = this.widgetinfos[i].jsFile;
          this.properties.cssFile = this.widgetinfos[i].cssFile;
          this.properties.widgetns = this.widgetinfos[i].widgetns;
          this.context.propertyPane.refresh();
        }
      }
    }
    if (this.vueInstance) {
      this.vueInstance.update({
        [propertyPath]: newValue,
      });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Widget Title",
                  value: this.properties.title,
                }),
                PropertyPaneToggle("removePadding", {
                  label: "Remove top/bottom padding of web part container",
                  checked: this.properties.removePadding,
                  onText: "Remove padding",
                  offText: "Keep padding",
                }),
                PropertyPaneDropdown("externalWidget", {
                  label: "Select Script Widget",
                  options: this.widgets,
                  selectedKey: this.properties.externalWidget,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
