import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// UPGRADE: PnP v4 Imports
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import * as strings from 'PowerBrandingApplicationCustomizerStrings';
import AdminPanel from './AdminPanel';

const LOG_SOURCE: string = 'PowerBrandingApplicationCustomizer';
const CACHE_KEY: string = "PowerBranding_CUSTOMIZER_RULES";
const LOADER_ID: string = "Power-Branding-loader-overlay";

export interface IPowerBrandingApplicationCustomizerProperties {
  listName: string;
  adminKey: string;
}

export default class PowerBrandingApplicationCustomizer
  extends BaseApplicationCustomizer<IPowerBrandingApplicationCustomizerProperties> {

  private sp!: SPFI; // UPGRADE: Store SPFI context globally

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // 1. Show Loader Immediately from Cache
    this.checkAndApplyCachedLoader();

    // UPGRADE: Initialize PnP v4
    this.sp = spfi().using(SPFx(this.context));

    // 2. Apply cached rules
    this.applyCachedRules();

    // 3. Fetch fresh rules and apply them
    await this.refreshRulesFromList();

    // 4. Hide Loader
    setTimeout(() => {
      this.hideLoader();
    }, 1000);

    // Run admin panel logic in background
    void this.checkAdminAccess();
  }

  private async checkAdminAccess(): Promise<void> {
    const now = new Date();
    const month = ("0" + (now.getMonth() + 1)).slice(-2);
    const year = now.getFullYear();
    const dynamicKey = `SecureAdmin${month}${year}`;

    let providedKey: string | null = null;
    if (window.location.search && window.location.search.length > 1) {
      const query = window.location.search.substring(1);
      const vars = query.split('&');
      for (let i = 0; i < vars.length; i++) {
        const pair = vars[i].split('=');
        if (decodeURIComponent(pair[0]).toLowerCase() === 'poweradmin') {
          providedKey = pair.length > 1 ? decodeURIComponent(pair[1]) : "";
          break;
        }
      }
    }

    if (providedKey && providedKey === dynamicKey) {
      const isAdmin = await this.checkIfUserIsAdmin();
      if (isAdmin) {
        this.renderAdminPanel();
      }
    }
  }

  private async checkIfUserIsAdmin(): Promise<boolean> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser`;
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const user = await response.json();
        return user.IsSiteAdmin === true;
      }
    } catch (error) {
      console.error("Error checking admin status", error);
    }
    return false;
  }

  private applyCachedRules(): void {
    const cachedData = localStorage.getItem(CACHE_KEY);
    if (cachedData) {
      try {
        const rules = JSON.parse(cachedData);
        this.applyRules(rules);
      } catch (e) { console.error(e); }
    }
  }

  private async refreshRulesFromList(): Promise<void> {
    const listTitle = this.properties.listName || "PowerBranding";
    try {
      // UPGRADE: PnP v4 Execution with ()
      const activeRules = await this.sp.web.lists.getByTitle(listTitle).items
        .filter("IsActive eq 1")();

      localStorage.setItem(CACHE_KEY, JSON.stringify(activeRules));
      this.applyRules(activeRules);
    } catch (error) {
      console.log("Power Branding Error: List not ready or error fetching.");
    }
  }

  private applyRules(rules: any[]): void {
    let cssString = "";
    rules.forEach(rule => {
      const important = rule.IsOverride ? " !important" : "";
      if (rule.ActionType === "Hide") {
        cssString += `${rule.Selector} { display: none${important}; } `;
      } else if (rule.ActionType === "Style") {
        cssString += `${rule.Selector} { ${rule.CSSValue}${important}; } `;
      } else if (rule.ActionType === "Script") {
        this.executeScript(rule.CSSValue);
      }
    });

    if (cssString) {
      const styleId = "power-branding-styles";
      let styleElement = document.getElementById(styleId) as HTMLStyleElement;
      if (!styleElement) {
        styleElement = document.createElement('style');
        styleElement.id = styleId;
        styleElement.type = 'text/css';
        document.head.appendChild(styleElement);
      }
      styleElement.innerHTML = cssString;
    }
  }
  private executeScript(scriptContent: string): void {
    try {
      // eslint-disable-next-line no-new-func
      const dynamicFunction = new Function('context', scriptContent);
      dynamicFunction(this.context);
    } catch (e) {
      console.error("Power Branding Script Error:", e);
    }
  }

  private renderAdminPanel(): void {
    const hostId = 'power-branding-admin-panel-hostV2';
    let hostElement = document.getElementById(hostId);
    if (!hostElement) {
      hostElement = document.createElement('div');
      hostElement.id = hostId;
      document.body.appendChild(hostElement);
    }
    const element = React.createElement(
      AdminPanel as any,
      {
        isOpen: true,
        listName: this.properties.listName || "PowerBranding",
        sp: this.sp, // UPGRADE: Passing SPFI instance
        onClose: () => {
          if (hostElement) ReactDom.unmountComponentAtNode(hostElement);
        }
      }
    );
    ReactDom.render(element, hostElement);
  }

  // --- LOADER LOGIC ---
  private checkAndApplyCachedLoader(): void {
    const cachedData = localStorage.getItem(CACHE_KEY);
    if (cachedData) {
      try {
        const rules = JSON.parse(cachedData);
        const loaderRule = rules.find((r: any) => r.ActionType === "Loader" && r.IsActive);
        if (loaderRule) {
          this.renderLoader(loaderRule.CSSValue);
        }
      } catch (e) { console.error("Loader Cache Error", e); }
    }
  }

  private renderLoader(htmlContent: string): void {
    if (document.getElementById(LOADER_ID)) return;

    const loaderOverlay = document.createElement('div');
    loaderOverlay.id = LOADER_ID;
    loaderOverlay.style.cssText = `
      position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
      background: #ffffff; z-index: 999999; display: flex; 
      justify-content: center; align-items: center;
    `;
    loaderOverlay.innerHTML = htmlContent;
    document.body.appendChild(loaderOverlay);
  }

  private hideLoader(): void {
    const loader = document.getElementById(LOADER_ID);
    if (loader) {
      loader.style.transition = "opacity 0.5s ease";
      loader.style.opacity = "0";
      setTimeout(() => {
        if (loader.parentNode) loader.parentNode.removeChild(loader);
      }, 500);
    }
  }
}