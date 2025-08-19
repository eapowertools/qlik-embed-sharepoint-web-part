import { Version } from "@microsoft/sp-core-library";
import {
	type IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";

import styles from "./QlikEmbedWebPart.module.scss";
import * as strings from "QlikEmbedWebPartStrings";

export interface IQlikEmbedWebPartProps {
	tenantURL: string;
	clientID: string;
	appID: string;
	objectID: string;
}
// some change to be deleted later
export default class QlikEmbedWebPart extends BaseClientSideWebPart<IQlikEmbedWebPartProps> {
	private _isDarkTheme: boolean = false;
	// @ts-expect-error: This is used in onInit(), but TS doesn't pick up the usage.
	private _environmentMessage: string = "";
	private _sectionTagValue: string = "";
	private _redirectURI: string = "";
	private _allowedRegions: string[] = ["us", "eu", "de", "uk", "se", "sg", "ap", "jp", "in", "ae"];

	public render(): void {
		// access current DOM by using 'this.domElement'

		if (this._redirectURI === "") {
			this._redirectURI =
				this.context.pageContext.site.absoluteUrl + this.context.pageContext.site.serverRequestPath;
		}
		console.log("Redirect URI: " + this._redirectURI);

		let hasValidConfig: boolean = false;
		let configError: boolean = false;
		let configErrorMessage: string = "";
		if (this._sectionTagValue === "") {
			this._sectionTagValue = `${styles.qlikEmbed}${
				!!this.context.sdks.microsoftTeams ? styles.teams : ""
			}`;
		}

		// clear object
		const sectionToRemove = document.getElementById(this._sectionTagValue);
		if (sectionToRemove !== null) {
			sectionToRemove.remove();
		}

		// create new section for chart/message
		const sectionTag: HTMLElement = document.createElement("section");
		sectionTag.classList.add(this._sectionTagValue);
		sectionTag.id = this._sectionTagValue;

		// VALIDATION
		if (this.properties.tenantURL !== "" && this.properties.tenantURL !== undefined) {
			const tenantValidation: string[] = this.properties.tenantURL.split(".");
			if (tenantValidation[0] === "") {
				configError = true;
				configErrorMessage = `Tenant "${this.properties.tenantURL}" has no tenant name.`;
			}
			if (this._allowedRegions.indexOf(tenantValidation[1]) === -1) {
				configError = true;
				configErrorMessage = `Tenant "${this.properties.tenantURL}" has an invalid region.`;
			}

			// tenant is valid, check client ID
			if (!configError) {
				// at this point i need to validate the oauth client.
				hasValidConfig = false;

				// validate App ID here:
				// if valid, set "hasValidConfig" true
				// if not valid and not empty, set configError true
				// and set configErrorMessage with a message saying something useful.
				// put in code below here.
			}
		}

		if (hasValidConfig) {
			const scriptTag: HTMLScriptElement = document.createElement("script");
			scriptTag.setAttribute("crossorigin", "anonymous");
			scriptTag.setAttribute("type", "application/javascript");
			scriptTag.setAttribute(
				"src",
				"https://cdn.jsdelivr.net/npm/@qlik/embed-web-components@1/dist/index.min.js"
			);
			scriptTag.setAttribute("data-host", `${this.properties.tenantURL}`);
			scriptTag.setAttribute("data-client-id", `${this.properties.clientID}`);
			scriptTag.setAttribute("data-redirect-uri", `${this._redirectURI}`);
			scriptTag.setAttribute("data-auto-redirect", "true");
			scriptTag.setAttribute("data-access-token-storage", "session");

			const embedDiv: HTMLDivElement = document.createElement("div");
			embedDiv.classList.add(`${styles.qlikChart}`);
			const embedTag: HTMLElement = document.createElement("qlik-embed");
			embedTag.classList.add(`${styles.qlikChart}`);
			embedTag.setAttribute("ui", "analytics/chart");
			embedTag.setAttribute("app-id", `${this.properties.appID}`);
			embedTag.setAttribute("object-id", `${this.properties.objectID}`);
			embedDiv.appendChild(embedTag);

			sectionTag.appendChild(scriptTag);
			sectionTag.appendChild(embedDiv);
		} else {
			const sectionHeaderDiv: HTMLDivElement = document.createElement("div");
			sectionHeaderDiv.classList.add(`${styles.welcome}`);

			if (!configError) {
				sectionHeaderDiv.innerHTML = `<img alt="" src="${
					this._isDarkTheme ? require("./assets/qlikLogo.png") : require("./assets/qlikLogo.png")
				}" class="${styles.welcomeImage}" />
				<p>Use sharepoint to configure this object to embed a Qlik chart.</p>
				`;
			} else {
				sectionHeaderDiv.innerHTML = `<img alt="" src="${
					this._isDarkTheme ? require("./assets/qlikLogo.png") : require("./assets/qlikLogo.png")
				}" class="${styles.welcomeImage}" />
				<p>Error configuring chart:</p>
				<p class="${styles.chartError}">${configErrorMessage}</p>
				`;
			}

			sectionTag.appendChild(sectionHeaderDiv);
		}

		this.domElement.appendChild(sectionTag);
	}

	protected onInit(): Promise<void> {
		return this._getEnvironmentMessage().then((message) => {
			this._environmentMessage = message;
		});
	}

	private _getEnvironmentMessage(): Promise<string> {
		if (!!this.context.sdks.microsoftTeams) {
			// running in Teams, office.com or Outlook
			return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
				let environmentMessage: string = "";
				switch (context.app.host.name) {
					case "Office": // running in Office
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentOffice
							: strings.AppOfficeEnvironment;
						break;
					case "Outlook": // running in Outlook
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentOutlook
							: strings.AppOutlookEnvironment;
						break;
					case "Teams": // running in Teams
					case "TeamsModern":
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentTeams
							: strings.AppTeamsTabEnvironment;
						break;
					default:
						environmentMessage = strings.UnknownEnvironment;
				}

				return environmentMessage;
			});
		}

		return Promise.resolve(
			this.context.isServedFromLocalhost
				? strings.AppLocalEnvironmentSharePoint
				: strings.AppSharePointEnvironment
		);
	}

	protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
		if (!currentTheme) {
			return;
		}

		this._isDarkTheme = !!currentTheme.isInverted;
		const { semanticColors } = currentTheme;

		if (semanticColors) {
			this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
			this.domElement.style.setProperty("--link", semanticColors.link || null);
			this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
		}
	}

	protected get dataVersion(): Version {
		return Version.parse("1.0");
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupName: strings.TenantConfigGroupName,
							groupFields: [
								PropertyPaneTextField("tenantURL", {
									label: strings.tenantURLFieldLabel,
								}),
								PropertyPaneTextField("clientID", {
									label: strings.clientIDFieldLabel,
								}),
							],
						},
						{
							groupName: strings.ObjectConfigGroupName,
							groupFields: [
								PropertyPaneTextField("appID", {
									label: strings.appIDFieldLabel,
								}),
								PropertyPaneTextField("objectID", {
									label: strings.objectIDFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}
