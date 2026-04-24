import { Version } from "@microsoft/sp-core-library";
import {
	type IPropertyPaneConfiguration,
	PropertyPaneDropdown,
	PropertyPaneTextField,
	PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { setDefaultHostConfig } from "@qlik/api/auth";
import { getAssistants, type Assistant } from "@qlik/api/assistants";
import { getItems, type ItemResultResponseBody } from "@qlik/api/items";
import { openAppSession, type GenericObjectEntry, type SheetListItem } from "@qlik/api/qix";

import styles from "./QlikEmbedWebPart.module.scss";
import {
	PropertyPaneSearchableDropdown,
	type ISearchableDropdownOption,
} from "./propertyPane/PropertyPaneSearchableDropdown";
import * as strings from "QlikEmbedWebPartStrings";

type ContentType = "app" | "sheet" | "chart" | "assistant";
type ChartSize = "small" | "medium" | "large" | "xlarge" | "custom";
type DropdownOption = ISearchableDropdownOption;
type PropertyPaneField =
	| ReturnType<typeof PropertyPaneDropdown>
	| ReturnType<typeof PropertyPaneTextField>
	| ReturnType<typeof PropertyPaneToggle>
	| ReturnType<typeof PropertyPaneSearchableDropdown>;
type PropertyPaneAccessor = {
	refresh?: () => void;
};
type PagedResponse<T> = {
	data?: {
		data?: T[];
	};
	next?: () => Promise<PagedResponse<T>>;
};
type EmbedConfig = {
	ui: string;
	attributes: Array<readonly [string, string | undefined]>;
};

const DEFAULT_CONTENT_TYPE: ContentType = "sheet";
const DEFAULT_CHART_SIZE: ChartSize = "medium";
const APP_CONTENT_TYPES: ReadonlySet<ContentType> = new Set(["app", "sheet", "chart"]);
const VALID_CONTENT_TYPES: ReadonlySet<ContentType> = new Set([
	"app",
	"sheet",
	"chart",
	"assistant",
]);
const VALID_CHART_SIZES: ReadonlySet<ChartSize> = new Set([
	"small",
	"medium",
	"large",
	"xlarge",
	"custom",
]);
const CONTENT_TYPE_OPTIONS: ReadonlyArray<DropdownOption> = [
	{ key: "app", text: "App" },
	{ key: "sheet", text: "Sheet" },
	{ key: "chart", text: "Chart" },
	{ key: "assistant", text: "Assistant" },
];
const CHART_SIZE_OPTIONS: ReadonlyArray<DropdownOption> = [
	{ key: "small", text: "Small" },
	{ key: "medium", text: "Medium" },
	{ key: "large", text: "Large" },
	{ key: "xlarge", text: "XLarge" },
	{ key: "custom", text: "Custom" },
];
const DEFAULT_WELCOME_MESSAGE = "Use SharePoint to configure this web part to embed Qlik content.";
const NO_SHEET_OPTION: DropdownOption = {
	key: "",
	text: "No sheet (open app landing page)",
};
const QLIK_EMBED_SCRIPT_SOURCE =
	"https://cdn.jsdelivr.net/npm/@qlik/embed-web-components@1/dist/index.min.js";

export interface IQlikEmbedWebPartProps {
	tenant: string;
	clientID: string;
	appID: string;
	selectedChartSize: ChartSize | string;
	selectedContentType?: ContentType;
	sheetId?: string;
	chartId?: string;
	assistantId?: string;
	assistantLegacy?: boolean;
	useClassicSheetUi?: boolean;
	useClassicChartUi?: boolean;
	customChartHeight?: string;
}

export default class QlikEmbedWebPart extends BaseClientSideWebPart<IQlikEmbedWebPartProps> {
	// @ts-expect-error: This is used in onInit(), but TS doesn't pick up the usage.
	private _environmentMessage: string = "";
	private _redirectURI: string = "";
	private readonly _allowedRegions: ReadonlySet<string> = new Set([
		"us",
		"eu",
		"de",
		"uk",
		"se",
		"sg",
		"ap",
		"jp",
		"in",
		"ae",
	]);

	private _appOptions: DropdownOption[] = [];
	private _sheetOptions: DropdownOption[] = [];
	private _chartOptions: DropdownOption[] = [];
	private _assistantOptions: DropdownOption[] = [];

	private _appsCacheKey: string = "";
	private _sheetsCacheKey: string = "";
	private _chartsCacheKey: string = "";
	private _assistantsCacheKey: string = "";

	private _appsRequestId: number = 0;
	private _sheetsRequestId: number = 0;
	private _chartsRequestId: number = 0;
	private _assistantsRequestId: number = 0;

	private _isLoadingApps: boolean = false;
	private _isLoadingSheets: boolean = false;
	private _isLoadingCharts: boolean = false;
	private _isLoadingAssistants: boolean = false;

	private _appsLoadError: string = "";
	private _sheetsLoadError: string = "";
	private _chartsLoadError: string = "";
	private _assistantsLoadError: string = "";

	public async render(): Promise<void> {
		this._ensurePropertyDefaults();
		this._ensureRedirectUri();

		const validationErrors: string[] = this._getValidationErrors();
		const selectionGuidance = validationErrors.length === 0 ? this._getSelectionGuidanceMessage() : undefined;
		const sectionTag = this._createSection();

		if (validationErrors.length === 0 && !selectionGuidance) {
			sectionTag.appendChild(this._createEmbedPreview());
		} else {
			sectionTag.appendChild(
				this._createWelcomeState(selectionGuidance || DEFAULT_WELCOME_MESSAGE, validationErrors)
			);
		}

		this.domElement.replaceChildren(sectionTag);
	}

	protected onInit(): Promise<void> {
		this._ensurePropertyDefaults();
		this._ensureRedirectUri();

		return this._getEnvironmentMessage().then((message) => {
			this._environmentMessage = message;
		});
	}

	protected async onPropertyPaneConfigurationStart(): Promise<void> {
		this._ensurePropertyDefaults();
		await this._loadPropertyPaneData(true);
	}

	protected onPropertyPaneFieldChanged(
		propertyPath: string,
		oldValue: string | boolean | undefined,
		newValue: string | boolean | undefined
	): void {
		super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

		if (oldValue === newValue) {
			return;
		}

		this._ensurePropertyDefaults();
		let shouldReloadPropertyPaneData = false;

		switch (propertyPath) {
			case "tenant":
			case "clientID":
				this._clearAppData();
				this._clearSheetData();
				this._clearChartData();
				this._clearAssistantData();
				this.properties.appID = "";
				this.properties.sheetId = "";
				this.properties.chartId = "";
				this.properties.assistantId = "";
				shouldReloadPropertyPaneData = true;
				break;
			case "selectedContentType":
				this._handleContentTypeChange(newValue as ContentType | undefined);
				shouldReloadPropertyPaneData = true;
				break;
			case "appID":
				this._clearSheetData();
				this._clearChartData();
				this.properties.sheetId = "";
				this.properties.chartId = "";
				shouldReloadPropertyPaneData = true;
				break;
			case "sheetId":
				this._clearChartData();
				this.properties.chartId = "";
				shouldReloadPropertyPaneData = true;
				break;
			default:
				break;
		}

		this._refreshPropertyPane();
		this.render().catch(() => undefined);
		if (shouldReloadPropertyPaneData) {
			this._loadPropertyPaneData().catch(() => undefined);
		}
	}

	private _createSection(): HTMLElement {
		const sectionTag = document.createElement("section");
		sectionTag.classList.add(styles.qlikEmbed);
		if (!!this.context.sdks.microsoftTeams) {
			sectionTag.classList.add(styles.teams);
		}

		return sectionTag;
	}

	private _createEmbedPreview(): DocumentFragment {
		const preview = document.createDocumentFragment();
		const embedDiv = document.createElement("div");
		const embedTag = document.createElement("qlik-embed");
		const embedConfig = this._getEmbedConfig();

		this._applyChartSize(embedDiv);
		embedTag.classList.add(styles.qlikChart);
		embedTag.setAttribute("ui", embedConfig.ui);
		this._setElementAttributes(embedTag, embedConfig.attributes);

		embedDiv.appendChild(embedTag);
		preview.appendChild(this._createEmbedScriptTag());
		preview.appendChild(embedDiv);

		return preview;
	}

	private _createEmbedScriptTag(): HTMLScriptElement {
		const scriptTag = document.createElement("script");
		scriptTag.setAttribute("crossorigin", "anonymous");
		scriptTag.setAttribute("type", "application/javascript");
		scriptTag.setAttribute("src", QLIK_EMBED_SCRIPT_SOURCE);
		scriptTag.setAttribute("data-host", this.properties.tenant.trim());
		scriptTag.setAttribute("data-client-id", this.properties.clientID.trim());
		scriptTag.setAttribute("data-redirect-uri", this._redirectURI);
		scriptTag.setAttribute("data-auto-redirect", "true");
		scriptTag.setAttribute("data-access-token-storage", "session");

		return scriptTag;
	}

	private _getEmbedConfig(): EmbedConfig {
		const appId = this._getConfiguredValue(this.properties.appID);
		const sheetId = this._getConfiguredValue(this.properties.sheetId);
		const chartId = this._getConfiguredValue(this.properties.chartId);
		const assistantId = this._getConfiguredValue(this.properties.assistantId);

		switch (this.properties.selectedContentType as ContentType) {
			case "app":
				return {
					ui: "classic/app",
					attributes: [
						["app-id", appId],
						["sheet-id", sheetId],
					],
				};
			case "sheet":
				return {
					ui: this.properties.useClassicSheetUi ? "classic/app" : "analytics/sheet",
					attributes: [
						["app-id", appId],
						["sheet-id", sheetId],
					],
				};
			case "chart":
				return {
					ui: this.properties.useClassicChartUi ? "classic/chart" : "analytics/chart",
					attributes: [
						["app-id", appId],
						["object-id", chartId],
					],
				};
			case "assistant":
				return {
					ui: this.properties.assistantLegacy ? "ai/assistant" : "ai/agentic-assistant",
					attributes: [["assistant-id", assistantId]],
				};
			default:
				return {
					ui: "classic/app",
					attributes: [["app-id", appId]],
				};
		}
	}

	private _setElementAttributes(
		element: HTMLElement,
		attributes: Array<readonly [string, string | undefined]>
	): void {
		for (let index = 0; index < attributes.length; index++) {
			const [attribute, value] = attributes[index];
			if (value) {
				element.setAttribute(attribute, value);
			}
		}
	}

	private _createTextBlock(
		tagName: "p" | "span",
		text: string,
		className?: string
	): HTMLParagraphElement | HTMLSpanElement {
		const element = document.createElement(tagName);
		element.textContent = text;
		if (className) {
			element.className = className;
		}

		return element;
	}

	private _createWelcomeState(message: string, validationErrors: string[]): HTMLDivElement {
		const sectionHeaderDiv: HTMLDivElement = document.createElement("div");
		sectionHeaderDiv.classList.add(styles.welcome);
		const logoImage = document.createElement("img");
		logoImage.alt = "";
		logoImage.src = require("./assets/qlikLogo.png");
		logoImage.className = styles.welcomeImage;

		sectionHeaderDiv.appendChild(logoImage);

		if (validationErrors.length > 0) {
			sectionHeaderDiv.appendChild(this._createTextBlock("p", "Error configuring embed:"));
			for (let index = 0; index < validationErrors.length; index++) {
				sectionHeaderDiv.appendChild(
					this._createTextBlock("p", validationErrors[index], styles.chartError)
				);
			}
			return sectionHeaderDiv;
		}

		sectionHeaderDiv.appendChild(this._createTextBlock("p", message));
		return sectionHeaderDiv;
	}

	private _ensurePropertyDefaults(): void {
		if (!VALID_CONTENT_TYPES.has(this.properties.selectedContentType as ContentType)) {
			this.properties.selectedContentType = DEFAULT_CONTENT_TYPE;
		}

		if (!VALID_CHART_SIZES.has(this.properties.selectedChartSize as ChartSize)) {
			this.properties.selectedChartSize = DEFAULT_CHART_SIZE;
		}

		if (typeof this.properties.useClassicSheetUi === "undefined") {
			this.properties.useClassicSheetUi = true;
		}

		if (typeof this.properties.useClassicChartUi === "undefined") {
			this.properties.useClassicChartUi = false;
		}

		if (typeof this.properties.assistantLegacy === "undefined") {
			this.properties.assistantLegacy = false;
		}

		if (!this.properties.customChartHeight) {
			this.properties.customChartHeight = "";
		}
	}

	private _ensureRedirectUri(): void {
		if (this._redirectURI === "") {
			this._redirectURI =
				this.context.pageContext.site.absoluteUrl + this.context.pageContext.site.serverRequestPath;
		}
	}

	private _configureQlikAuth(): void {
		setDefaultHostConfig({
			authType: "oauth2",
			host: this.properties.tenant.trim(),
			clientId: this.properties.clientID.trim(),
			redirectUri: this._redirectURI,
			accessTokenStorage: "session",
			autoRedirect: true,
		});
	}

	private async _loadPropertyPaneData(force: boolean = false): Promise<void> {
		const selectedContentType = this.properties.selectedContentType as ContentType;
		const tenantError = this._validateTenant(this.properties.tenant);
		const clientError = this._validateClientId(this.properties.clientID);

		if (tenantError || clientError) {
			return;
		}

		if (this._usesAppSelections(selectedContentType)) {
			await this._loadApps(force);
		}

		if (this._usesAppSelections(selectedContentType) && this.properties.appID) {
			await this._loadSheets(force);
		}

		if (selectedContentType === "chart" && this.properties.appID && this.properties.sheetId) {
			await this._loadCharts(force);
		}

		if (selectedContentType === "assistant") {
			await this._loadAssistants(force);
		}
	}

	private async _loadApps(force: boolean = false): Promise<void> {
		const cacheKey = `${this.properties.tenant.trim()}|${this.properties.clientID.trim()}`;
		if (!force && this._appsCacheKey === cacheKey && !this._appsLoadError) {
			return;
		}

		const requestId = ++this._appsRequestId;
		this._isLoadingApps = true;
		this._appsLoadError = "";
		this._refreshPropertyPane();

		try {
			this._configureQlikAuth();
			const apps = await this._collectPagedData<ItemResultResponseBody>(async () =>
				getItems({
					resourceType: "app",
					limit: 100,
				})
			);

			if (requestId !== this._appsRequestId) {
				return;
			}

			this._appOptions = apps
				.filter((app) => !!app.resourceId)
				.map((app) => ({
					key: app.resourceId || "",
					text: this._formatOptionLabel(app.name, "Untitled App", app.resourceId || ""),
				}))
				.sort(this._sortDropdownOptions);
			this._appsCacheKey = cacheKey;

			if (this.properties.appID && !this._hasOption(this._appOptions, this.properties.appID)) {
				this.properties.appID = "";
				this.properties.sheetId = "";
				this.properties.chartId = "";
				this._clearSheetData();
				this._clearChartData();
			}
		} catch (error) {
			if (requestId !== this._appsRequestId) {
				return;
			}

			this._appOptions = [];
			this._appsCacheKey = "";
			this._appsLoadError = this._getErrorMessage("Failed to load apps.", error);
		} finally {
			if (requestId === this._appsRequestId) {
				this._isLoadingApps = false;
				this._refreshPropertyPane();
				this.render().catch(() => undefined);
			}
		}
	}

	private async _loadSheets(force: boolean = false): Promise<void> {
		const cacheKey = this.properties.appID;
		if (!force && this._sheetsCacheKey === cacheKey && !this._sheetsLoadError) {
			return;
		}

		const requestId = ++this._sheetsRequestId;
		const originalSheetId = this.properties.sheetId || "";
		const originalChartId = this.properties.chartId || "";
		this._isLoadingSheets = true;
		this._sheetsLoadError = "";
		this._refreshPropertyPane();

		this._configureQlikAuth();
		const appSession = openAppSession({ appId: this.properties.appID });

		try {
			const doc = await appSession.getDoc();
			const sheets = await doc.getSheetList();

			if (requestId !== this._sheetsRequestId) {
				return;
			}

			this._sheetOptions = sheets
				.map((sheet) => ({
					key: sheet.qInfo?.qId || "",
					text: this._formatOptionLabel(
						this._getSheetTitle(sheet),
						"Unnamed Sheet",
						sheet.qInfo?.qId || ""
					),
				}))
				.filter((sheet) => sheet.key !== "")
				.sort(this._sortDropdownOptions);
			this._sheetsCacheKey = cacheKey;

			if (this.properties.sheetId && !this._hasOption(this._sheetOptions, this.properties.sheetId)) {
				this.properties.sheetId = "";
				this.properties.chartId = "";
				this._clearChartData();
			}
		} catch (error) {
			if (requestId !== this._sheetsRequestId) {
				return;
			}

			this._sheetOptions = [];
			this._sheetsCacheKey = "";
			this._sheetsLoadError = this._getErrorMessage("Failed to load sheets.", error);
		} finally {
			await appSession.close().catch(() => undefined);

			if (requestId === this._sheetsRequestId) {
				this._isLoadingSheets = false;
				this._refreshPropertyPane();
				if (
					this.properties.selectedContentType !== "app" ||
					originalSheetId !== (this.properties.sheetId || "") ||
					originalChartId !== (this.properties.chartId || "")
				) {
					this.render().catch(() => undefined);
				}
			}
		}
	}

	private async _loadCharts(force: boolean = false): Promise<void> {
		const cacheKey = `${this.properties.appID}|${this.properties.sheetId || ""}`;
		if (!force && this._chartsCacheKey === cacheKey && !this._chartsLoadError) {
			return;
		}

		const requestId = ++this._chartsRequestId;
		this._isLoadingCharts = true;
		this._chartsLoadError = "";
		this._refreshPropertyPane();

		this._configureQlikAuth();
		const appSession = openAppSession({ appId: this.properties.appID });

		try {
			const doc = await appSession.getDoc();
			const sheetObject = await doc.getObject(this.properties.sheetId || "");
			const chartTree = (await sheetObject.getFullPropertyTree()) as GenericObjectEntry;
			const chartOptions = this._getChartOptions(chartTree);

			if (requestId !== this._chartsRequestId) {
				return;
			}

			this._chartOptions = chartOptions.sort(this._sortDropdownOptions);
			this._chartsCacheKey = cacheKey;

			if (this.properties.chartId && !this._hasOption(this._chartOptions, this.properties.chartId)) {
				this.properties.chartId = "";
			}
		} catch (error) {
			if (requestId !== this._chartsRequestId) {
				return;
			}

			this._chartOptions = [];
			this._chartsCacheKey = "";
			this._chartsLoadError = this._getErrorMessage("Failed to load charts.", error);
		} finally {
			await appSession.close().catch(() => undefined);

			if (requestId === this._chartsRequestId) {
				this._isLoadingCharts = false;
				this._refreshPropertyPane();
				this.render().catch(() => undefined);
			}
		}
	}

	private async _loadAssistants(force: boolean = false): Promise<void> {
		const cacheKey = `${this.properties.tenant.trim()}|${this.properties.clientID.trim()}`;
		if (!force && this._assistantsCacheKey === cacheKey && !this._assistantsLoadError) {
			return;
		}

		const requestId = ++this._assistantsRequestId;
		this._isLoadingAssistants = true;
		this._assistantsLoadError = "";
		this._refreshPropertyPane();

		try {
			this._configureQlikAuth();
			const assistants = await this._collectPagedData<Assistant>(async () =>
				getAssistants({
					limit: 100,
				})
			);

			if (requestId !== this._assistantsRequestId) {
				return;
			}

			this._assistantOptions = assistants
				.map((assistant) => ({
					key: assistant.id,
					text: this._formatOptionLabel(
						assistant.title || assistant.name,
						"Untitled Assistant",
						assistant.id
					),
				}))
				.sort(this._sortDropdownOptions);
			this._assistantsCacheKey = cacheKey;

			if (
				this.properties.assistantId &&
				!this._hasOption(this._assistantOptions, this.properties.assistantId)
			) {
				this.properties.assistantId = "";
			}

			this._applyDefaultAssistantSelection();
		} catch (error) {
			if (requestId !== this._assistantsRequestId) {
				return;
			}

			this._assistantOptions = [];
			this._assistantsCacheKey = "";
			this._assistantsLoadError = this._getErrorMessage("Failed to load assistants.", error);
		} finally {
			if (requestId === this._assistantsRequestId) {
				this._isLoadingAssistants = false;
				this._refreshPropertyPane();
				this.render().catch(() => undefined);
			}
		}
	}

	private _getChartOptions(chartTree: GenericObjectEntry): DropdownOption[] {
		if (!chartTree.qChildren || chartTree.qChildren.length === 0) {
			return [];
		}

		return chartTree.qChildren
			.map((child) => {
				const qProperty = child.qProperty as
					| {
						title?: unknown;
						qMetaDef?: {
							title?: unknown;
						};
						qInfo?: {
							qId?: string;
							qType?: string;
						};
					}
					| undefined;
				const chartId = qProperty?.qInfo?.qId || "";
				const chartType = qProperty?.qInfo?.qType || "object";
				const chartTitle =
					this._getConfiguredValue(this._getTextValue(qProperty?.qMetaDef?.title)) ||
					this._getConfiguredValue(this._getTextValue(qProperty?.title)) ||
					chartType;

				return {
					key: chartId,
					text: this._formatOptionLabel(chartTitle, chartType, chartType),
				};
			})
			.filter((chart) => chart.key !== "");
	}

	private _getValidationErrors(): string[] {
		const errors: string[] = [];
		const tenantError = this._validateTenant(this.properties.tenant);
		if (tenantError) {
			errors.push(tenantError);
		}

		const clientError = this._validateClientId(this.properties.clientID);
		if (clientError) {
			errors.push(clientError);
		}

		return errors;
	}

	private _getSelectionGuidanceMessage(): string | undefined {
		if (
			this.properties.selectedChartSize === "custom" &&
			!this._hasConfiguredValue(this.properties.customChartHeight)
		) {
			return "Enter a custom height to preview the embed.";
		}

		const selectedContentType = this.properties.selectedContentType as ContentType;
		switch (selectedContentType) {
			case "app":
				return this._hasConfiguredValue(this.properties.appID)
					? undefined
					: "Select an app to preview the embed.";
			case "sheet":
				if (!this._hasConfiguredValue(this.properties.appID)) {
					return "Select an app before choosing a sheet.";
				}
				return this._hasConfiguredValue(this.properties.sheetId)
					? undefined
					: "Select a sheet to preview the embed.";
			case "chart":
				if (!this._hasConfiguredValue(this.properties.appID)) {
					return "Select an app before choosing a chart.";
				}
				if (!this._hasConfiguredValue(this.properties.sheetId)) {
					return "Select a sheet before choosing a chart.";
				}
				return this._hasConfiguredValue(this.properties.chartId)
					? undefined
					: "Select a chart to preview the embed.";
			case "assistant":
				return this._hasConfiguredValue(this.properties.assistantId)
					? undefined
					: "Select an assistant to preview the embed.";
			default:
				return undefined;
		}
	}

	private _validateTenant(tenant: string | undefined): string | undefined {
		const value = tenant ? tenant.trim() : "";
		if (value === "") {
			return 'Tenant field format should be: "tenantName.region.qlikcloud.com".';
		}

		if (/^https?:\/\//i.test(value)) {
			return `Tenant "${value}" should not include a protocol.`;
		}

		if (value.indexOf("/") !== -1) {
			return `Tenant "${value}" should not include a path or trailing slash.`;
		}

		const tenantMatch = /^([a-z0-9-]+)\.([a-z0-9-]+)\.qlikcloud\.com$/i.exec(value);
		if (!tenantMatch) {
			return `Tenant field format should be: "tenantName.region.qlikcloud.com".`;
		}

		if (!this._allowedRegions.has(tenantMatch[2].toLowerCase())) {
			return `Tenant "${value}" has an invalid region.`;
		}

		return undefined;
	}

	private _validateClientId(clientID: string | undefined): string | undefined {
		const value = clientID ? clientID.trim() : "";
		if (value === "") {
			return "Client ID is required.";
		}

		if (!/^[A-Fa-f0-9]{32}$/.test(value)) {
			return `The client ID provided: "${value}" is invalid.`;
		}

		return undefined;
	}

	private _applyChartSize(embedDiv: HTMLDivElement): void {
		const selectedChartSize = this.properties.selectedChartSize as ChartSize;
		switch (selectedChartSize) {
			case "small":
				embedDiv.classList.add(styles.qlikChartSmall);
				return;
			case "medium":
				embedDiv.classList.add(styles.qlikChartMedium);
				return;
			case "large":
				embedDiv.classList.add(styles.qlikChartLarge);
				return;
			case "xlarge":
				embedDiv.classList.add(styles.qlikChartXLarge);
				return;
			case "custom":
				embedDiv.style.width = "100%";
				embedDiv.style.height = this.properties.customChartHeight || "";
				return;
			default:
				return;
		}
	}

	private _handleContentTypeChange(nextContentType: ContentType | undefined): void {
		switch (nextContentType) {
			case "app":
			case "sheet":
				this.properties.chartId = "";
				this.properties.assistantId = "";
				this._clearChartData();
				break;
			case "chart":
				this.properties.assistantId = "";
				break;
			case "assistant":
				this.properties.sheetId = "";
				this.properties.chartId = "";
				this._clearSheetData();
				this._clearChartData();
				this._applyDefaultAssistantSelection();
				break;
			default:
				break;
		}
	}

	private _clearAppData(): void {
		this._appOptions = [];
		this._appsCacheKey = "";
		this._appsLoadError = "";
	}

	private _clearSheetData(): void {
		this._sheetOptions = [];
		this._sheetsCacheKey = "";
		this._sheetsLoadError = "";
	}

	private _clearChartData(): void {
		this._chartOptions = [];
		this._chartsCacheKey = "";
		this._chartsLoadError = "";
	}

	private _clearAssistantData(): void {
		this._assistantOptions = [];
		this._assistantsCacheKey = "";
		this._assistantsLoadError = "";
	}

	private _applyDefaultAssistantSelection(): boolean {
		if (this.properties.selectedContentType !== "assistant") {
			return false;
		}

		if (this.properties.assistantId && this._hasOption(this._assistantOptions, this.properties.assistantId)) {
			return false;
		}

		if (this._assistantOptions.length === 0) {
			return false;
		}

		this.properties.assistantId = this._assistantOptions[0].key;
		return true;
	}

	private _getSheetTitle(sheet: SheetListItem): string {
		const qMeta = sheet.qMeta as { title?: string } | undefined;
		return qMeta?.title || "Unnamed Sheet";
	}

	private _hasConfiguredValue(value: string | undefined): boolean {
		return typeof value === "string" && value.trim() !== "";
	}

	private _getConfiguredValue(value: string | undefined): string | undefined {
		if (!this._hasConfiguredValue(value)) {
			return undefined;
		}

		return value?.trim();
	}

	private _getTextValue(value: unknown): string | undefined {
		if (typeof value === "string") {
			return value.trim() || undefined;
		}

		if (typeof value === "number" || typeof value === "boolean") {
			return String(value);
		}

		if (!value || typeof value !== "object") {
			return undefined;
		}

		const candidateValue = value as {
			qExpr?: unknown;
			qStringExpression?: unknown;
			qText?: unknown;
			title?: unknown;
			value?: unknown;
		};

		return (
			this._getTextValue(candidateValue.qText) ||
			this._getTextValue(candidateValue.title) ||
			this._getTextValue(candidateValue.value) ||
			this._getTextValue(candidateValue.qExpr) ||
			this._getTextValue(candidateValue.qStringExpression)
		);
	}

	private _usesAppSelections(contentType: ContentType): boolean {
		return APP_CONTENT_TYPES.has(contentType);
	}

	private async _collectPagedData<T>(
		loadFirstPage: () => Promise<PagedResponse<T>>
	): Promise<T[]> {
		let response = await loadFirstPage();
		const items = response.data?.data ? response.data.data.slice() : [];

		while (response.next) {
			response = await response.next();
			if (response.data?.data) {
				items.push(...response.data.data);
			}
		}

		return items;
	}

	private _formatOptionLabel(
		label: string | undefined,
		fallbackLabel: string,
		detail: string
	): string {
		const baseLabel = label?.trim() || fallbackLabel;
		return detail ? `${baseLabel} (${detail})` : baseLabel;
	}

	private _getErrorMessage(prefix: string, error: unknown): string {
		if (error instanceof Error && error.message) {
			return `${prefix} ${error.message}`;
		}

		return prefix;
	}

	private _refreshPropertyPane(): void {
		const propertyPane = this.context.propertyPane as PropertyPaneAccessor | undefined;
		propertyPane?.refresh?.();
	}

	private _hasOption(options: DropdownOption[], key: string): boolean {
		return options.some((option) => option.key === key);
	}

	private _sortDropdownOptions(a: DropdownOption, b: DropdownOption): number {
		return a.text.localeCompare(b.text, undefined, { sensitivity: "base" });
	}

	private _getEnvironmentMessage(): Promise<string> {
		if (!!this.context.sdks.microsoftTeams) {
			return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
				let environmentMessage: string = "";
				switch (context.app.host.name) {
					case "Office":
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentOffice
							: strings.AppOfficeEnvironment;
						break;
					case "Outlook":
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentOutlook
							: strings.AppOutlookEnvironment;
						break;
					case "Teams":
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

		const theme = currentTheme as IReadonlyTheme & {
			semanticColors?: {
				bodyText?: string;
				link?: string;
				linkHovered?: string;
			};
		};

		const { semanticColors } = theme;

		if (semanticColors) {
			this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
			this.domElement.style.setProperty("--link", semanticColors.link || null);
			this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
		}
	}

	protected get dataVersion(): Version {
		return Version.parse("1.0");
	}

	private _getObjectConfigurationFields(
		selectedContentType: ContentType,
		canLoadTenantData: boolean
	): PropertyPaneField[] {
		const objectFields: PropertyPaneField[] = [
			PropertyPaneDropdown("selectedContentType", {
				label: strings.contentTypeFieldLabel,
				options: [...CONTENT_TYPE_OPTIONS],
				selectedKey: this.properties.selectedContentType,
			}),
		];

		if (this._usesAppSelections(selectedContentType)) {
			objectFields.push(
				this._createAppField(canLoadTenantData),
				this._createSheetField(selectedContentType)
			);
		}

		if (selectedContentType === "sheet") {
			objectFields.push(
				PropertyPaneToggle("useClassicSheetUi", {
					label: strings.useClassicSheetUiFieldLabel,
					checked: !!this.properties.useClassicSheetUi,
				})
			);
		}

		if (selectedContentType === "chart") {
			objectFields.push(
				this._createChartField(),
				PropertyPaneToggle("useClassicChartUi", {
					label: strings.useClassicChartUiFieldLabel,
					checked: !!this.properties.useClassicChartUi,
				})
			);
		}

		if (selectedContentType === "assistant") {
			objectFields.push(
				this._createAssistantField(canLoadTenantData),
				PropertyPaneToggle("assistantLegacy", {
					label: strings.assistantLegacyFieldLabel,
					checked: !!this.properties.assistantLegacy,
				})
			);
		}

		objectFields.push(...this._getChartSizeFields());
		return objectFields;
	}

	private _createAppField(
		canLoadTenantData: boolean
	): ReturnType<typeof PropertyPaneSearchableDropdown> {
		return PropertyPaneSearchableDropdown({
			targetProperty: "appID",
			label: strings.appIDFieldLabel,
			options: this._appOptions,
			selectedKey: this.properties.appID || undefined,
			disabled:
				!canLoadTenantData || this._isLoadingApps || !!this._appsLoadError || this._appOptions.length === 0,
			placeholder: !canLoadTenantData
				? "Enter a valid tenant host and client ID first."
				: this._isLoadingApps
					? "Loading apps..."
					: this._appOptions.length === 0
						? "No apps available."
						: "Search apps...",
			errorMessage: this._appsLoadError || undefined,
		});
	}

	private _createSheetField(
		selectedContentType: ContentType
	): ReturnType<typeof PropertyPaneSearchableDropdown> {
		const isAppContent = selectedContentType === "app";
		return PropertyPaneSearchableDropdown({
			targetProperty: "sheetId",
			label: isAppContent ? strings.optionalSheetFieldLabel : strings.sheetFieldLabel,
			options: isAppContent ? [NO_SHEET_OPTION, ...this._sheetOptions] : this._sheetOptions,
			selectedKey: this.properties.sheetId || undefined,
			disabled:
				!this.properties.appID ||
				this._isLoadingSheets ||
				!!this._sheetsLoadError ||
				this._sheetOptions.length === 0,
			placeholder: !this.properties.appID
				? "Select an app first."
				: this._isLoadingSheets
					? "Loading sheets..."
					: this._sheetOptions.length === 0
						? "No sheets available for this app."
						: isAppContent
							? "Search sheets to optionally open one..."
							: "Search sheets...",
			errorMessage: this._sheetsLoadError || undefined,
		});
	}

	private _createChartField(): ReturnType<typeof PropertyPaneSearchableDropdown> {
		return PropertyPaneSearchableDropdown({
			targetProperty: "chartId",
			label: strings.chartFieldLabel,
			options: this._chartOptions,
			selectedKey: this.properties.chartId || undefined,
			disabled:
				!this.properties.sheetId ||
				this._isLoadingCharts ||
				!!this._chartsLoadError ||
				this._chartOptions.length === 0,
			placeholder: !this.properties.sheetId
				? "Select a sheet first."
				: this._isLoadingCharts
					? "Loading charts..."
					: this._chartOptions.length === 0
						? "No charts available for this sheet."
						: "Search charts...",
			errorMessage: this._chartsLoadError || undefined,
		});
	}

	private _createAssistantField(
		canLoadTenantData: boolean
	): ReturnType<typeof PropertyPaneSearchableDropdown> {
		return PropertyPaneSearchableDropdown({
			targetProperty: "assistantId",
			label: strings.assistantFieldLabel,
			options: this._assistantOptions,
			selectedKey: this.properties.assistantId || undefined,
			disabled:
				!canLoadTenantData ||
				this._isLoadingAssistants ||
				!!this._assistantsLoadError ||
				this._assistantOptions.length === 0,
			placeholder: !canLoadTenantData
				? "Enter a valid tenant host and client ID first."
				: this._isLoadingAssistants
					? "Loading assistants..."
					: this._assistantOptions.length === 0
						? "No assistants available."
						: "Search assistants...",
			errorMessage: this._assistantsLoadError || undefined,
		});
	}

	private _getChartSizeFields(): PropertyPaneField[] {
		const chartSizeFields: PropertyPaneField[] = [
			PropertyPaneDropdown("selectedChartSize", {
				label: strings.chartSizeFieldLabel,
				options: [...CHART_SIZE_OPTIONS],
				selectedKey: this.properties.selectedChartSize,
			}),
		];

		if (this.properties.selectedChartSize === "custom") {
			chartSizeFields.push(
				PropertyPaneTextField("customChartHeight", {
					label: strings.customChartHeightFieldLabel,
				})
			);
		}

		return chartSizeFields;
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		this._ensurePropertyDefaults();

		const selectedContentType = this.properties.selectedContentType as ContentType;
		const canLoadTenantData =
			!this._validateTenant(this.properties.tenant) && !this._validateClientId(this.properties.clientID);

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
								PropertyPaneTextField("tenant", {
									label: strings.tenantFieldLabel,
								}),
								PropertyPaneTextField("clientID", {
									label: strings.clientIDFieldLabel,
								}),
							],
						},
						{
							groupName: strings.ObjectConfigGroupName,
							groupFields: this._getObjectConfigurationFields(
								selectedContentType,
								canLoadTenantData
							),
						},
					],
				},
			],
		};
	}
}
