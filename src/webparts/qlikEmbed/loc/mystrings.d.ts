declare interface IQlikEmbedWebPartStrings {
	PropertyPaneDescription: string;
	TenantConfigGroupName: string;
	ObjectConfigGroupName: string;
	tenantFieldLabel: string;
	clientIDFieldLabel: string;
	appIDFieldLabel: string;
	contentTypeFieldLabel: string;
	sheetFieldLabel: string;
	optionalSheetFieldLabel: string;
	chartFieldLabel: string;
	assistantFieldLabel: string;
	assistantLegacyFieldLabel: string;
	useClassicSheetUiFieldLabel: string;
	useClassicChartUiFieldLabel: string;
	chartSizeFieldLabel: string;
	customChartHeightFieldLabel: string;
	AppLocalEnvironmentSharePoint: string;
	AppLocalEnvironmentTeams: string;
	AppLocalEnvironmentOffice: string;
	AppLocalEnvironmentOutlook: string;
	AppSharePointEnvironment: string;
	AppTeamsTabEnvironment: string;
	AppOfficeEnvironment: string;
	AppOutlookEnvironment: string;
	UnknownEnvironment: string;
}

declare module "QlikEmbedWebPartStrings" {
	const strings: IQlikEmbedWebPartStrings;
	export = strings;
}
