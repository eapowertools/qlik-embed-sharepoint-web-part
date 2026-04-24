import * as React from "react";
import * as ReactDOM from "react-dom";
import {
	PropertyPaneFieldType,
	type IPropertyPaneCustomFieldProps,
	type IPropertyPaneField,
} from "@microsoft/sp-property-pane";

export interface ISearchableDropdownOption {
	key: string;
	text: string;
	disabled?: boolean;
}

export interface IPropertyPaneSearchableDropdownProps {
	targetProperty: string;
	label: string;
	options: ISearchableDropdownOption[];
	selectedKey?: string;
	disabled?: boolean;
	placeholder?: string;
	errorMessage?: string;
}

interface ISearchableDropdownCustomFieldProps
	extends IPropertyPaneCustomFieldProps,
		IPropertyPaneSearchableDropdownProps {}

type ChangeCallback = (targetProperty?: string, newValue?: unknown, isValidEntry?: boolean) => void;

interface IReactDropdownProps extends IPropertyPaneSearchableDropdownProps {
	onChange: ChangeCallback;
}

interface IMenuPosition {
	left: number;
	top: number;
	width: number;
	maxHeight: number;
}

const containerStyle: React.CSSProperties = {
	display: "flex",
	flexDirection: "column",
	gap: "8px",
	marginBottom: "16px",
};

const labelStyle: React.CSSProperties = {
	fontSize: "14px",
	fontWeight: 600,
	color: "#323130",
};

const wrapperStyle: React.CSSProperties = {
	position: "relative",
};

const inputBaseStyle: React.CSSProperties = {
	boxSizing: "border-box",
	width: "100%",
	padding: "8px 58px 8px 10px",
	border: "1px solid #c8c6c4",
	borderRadius: "2px",
	backgroundColor: "#ffffff",
	color: "#323130",
	fontSize: "14px",
	lineHeight: 1.4,
};

const inputDisabledStyle: React.CSSProperties = {
	...inputBaseStyle,
	backgroundColor: "#f3f2f1",
	color: "#a19f9d",
	cursor: "not-allowed",
};

const inputErrorStyle: React.CSSProperties = {
	borderColor: "#a4262c",
};

const iconButtonBaseStyle: React.CSSProperties = {
	position: "absolute",
	top: "50%",
	transform: "translateY(-50%)",
	display: "flex",
	alignItems: "center",
	justifyContent: "center",
	width: "24px",
	height: "24px",
	border: "none",
	background: "transparent",
	color: "#605e5c",
	cursor: "pointer",
	padding: 0,
};

const iconButtonDisabledStyle: React.CSSProperties = {
	...iconButtonBaseStyle,
	color: "#c8c6c4",
	cursor: "not-allowed",
};

const clearButtonHiddenStyle: React.CSSProperties = {
	...iconButtonBaseStyle,
	right: "30px",
	display: "none",
	fontSize: "16px",
	lineHeight: 1,
};

const clearButtonVisibleStyle: React.CSSProperties = {
	...clearButtonHiddenStyle,
	display: "flex",
};

const chevronButtonStyle: React.CSSProperties = {
	...iconButtonBaseStyle,
	right: "6px",
	fontSize: "12px",
	lineHeight: 1,
};

const menuStyleBase: React.CSSProperties = {
	position: "fixed",
	border: "1px solid #c8c6c4",
	borderRadius: "4px",
	backgroundColor: "#ffffff",
	boxShadow: "0 8px 24px rgba(0, 0, 0, 0.18)",
	overflowY: "auto",
	zIndex: 100000,
};

const optionBaseStyle: React.CSSProperties = {
	padding: "8px 10px",
	fontSize: "14px",
	lineHeight: 1.4,
	color: "#323130",
	cursor: "pointer",
	backgroundColor: "#ffffff",
};

const optionActiveStyle: React.CSSProperties = {
	...optionBaseStyle,
	backgroundColor: "#faf9f8",
};

const optionSelectedStyle: React.CSSProperties = {
	...optionBaseStyle,
	backgroundColor: "#f3f2f1",
	color: "#005a9e",
	fontWeight: 600,
};

const optionDisabledStyle: React.CSSProperties = {
	...optionBaseStyle,
	color: "#a19f9d",
	cursor: "not-allowed",
};

const emptyStateStyle: React.CSSProperties = {
	padding: "8px 10px",
	fontSize: "14px",
	lineHeight: 1.4,
	color: "#605e5c",
	backgroundColor: "#ffffff",
};

const errorMessageStyle: React.CSSProperties = {
	fontSize: "12px",
	lineHeight: 1.4,
	color: "#a4262c",
};

function sanitizeId(value: string): string {
	return value.replace(/[^a-zA-Z0-9_-]/g, "-");
}

function getOptionByKey(
	options: ISearchableDropdownOption[],
	key: string | undefined
): ISearchableDropdownOption | undefined {
	if (typeof key === "undefined") {
		return undefined;
	}

	for (let index = 0; index < options.length; index++) {
		if (options[index].key === key) {
			return options[index];
		}
	}

	return undefined;
}

function getFilteredOptions(
	options: ISearchableDropdownOption[],
	query: string,
	selectedOption: ISearchableDropdownOption | undefined
): ISearchableDropdownOption[] {
	const normalizedQuery = query.trim().toLowerCase();
	if (normalizedQuery === "") {
		return options;
	}

	if (selectedOption && normalizedQuery === selectedOption.text.toLowerCase()) {
		return options;
	}

	const filteredOptions: ISearchableDropdownOption[] = [];
	for (let index = 0; index < options.length; index++) {
		if (options[index].text.toLowerCase().indexOf(normalizedQuery) !== -1) {
			filteredOptions.push(options[index]);
		}
	}

	return filteredOptions;
}

function getMenuPosition(wrapper: HTMLElement): IMenuPosition {
	const rect = wrapper.getBoundingClientRect();
	const viewportPadding = 12;
	const availableBelow = window.innerHeight - rect.bottom - viewportPadding;
	const availableAbove = rect.top - viewportPadding;
	const shouldOpenAbove = availableBelow < 140 && availableAbove > availableBelow;
	const maxHeight = Math.max(120, Math.min(220, shouldOpenAbove ? availableAbove : availableBelow));

	return {
		left: Math.round(rect.left),
		top: Math.round(shouldOpenAbove ? rect.top - maxHeight - 4 : rect.bottom + 4),
		width: Math.round(rect.width),
		maxHeight: Math.round(maxHeight),
	};
}

function getOptionStyle(
	option: ISearchableDropdownOption,
	isSelected: boolean,
	isActive: boolean
): React.CSSProperties {
	if (option.disabled) {
		return optionDisabledStyle;
	}

	if (isSelected) {
		return optionSelectedStyle;
	}

	if (isActive) {
		return optionActiveStyle;
	}

	return optionBaseStyle;
}

const SearchableDropdownField: React.FC<IReactDropdownProps> = (props) => {
	const fieldId = sanitizeId(props.targetProperty);
	const inputId = `${fieldId}-input`;
	const listboxId = `${fieldId}-listbox`;
	const errorId = `${fieldId}-error`;
	const wrapperRef = React.useRef<HTMLDivElement | null>(null);
	const inputRef = React.useRef<HTMLInputElement | null>(null);
	const listRef = React.useRef<HTMLDivElement | null>(null);
	const portalContainerRef = React.useRef<HTMLDivElement | null>(null);
	const selectedOption = getOptionByKey(props.options, props.selectedKey);
	const selectedOptionText = selectedOption ? selectedOption.text : "";
	const [query, setQuery] = React.useState<string>(selectedOption ? selectedOption.text : "");
	const [isOpen, setIsOpen] = React.useState<boolean>(false);
	const [activeIndex, setActiveIndex] = React.useState<number>(-1);
	const [menuPosition, setMenuPosition] = React.useState<IMenuPosition | null>(null);

	if (portalContainerRef.current === null && typeof document !== "undefined") {
		portalContainerRef.current = document.createElement("div");
	}

	const disabled = !!props.disabled || props.options.length === 0;
	const errorMessage = props.errorMessage ? props.errorMessage.trim() : "";
	const filteredOptions = getFilteredOptions(props.options, query, selectedOption);
	const activeDescendantId =
		isOpen && activeIndex >= 0 && activeIndex < filteredOptions.length
			? `${fieldId}-option-${activeIndex}`
			: undefined;

	React.useEffect(() => {
		if (!isOpen) {
			setQuery(selectedOptionText);
		}
	}, [isOpen, props.selectedKey, props.options, selectedOptionText]);

	React.useEffect(() => {
		const portalContainer = portalContainerRef.current;
		if (!portalContainer) {
			return;
		}

		document.body.appendChild(portalContainer);
		return () => {
			if (portalContainer.parentNode) {
				portalContainer.parentNode.removeChild(portalContainer);
			}
		};
	}, []);

	React.useEffect(() => {
		if (!isOpen || disabled || !wrapperRef.current) {
			return;
		}

		const updatePosition = (): void => {
			if (wrapperRef.current) {
				setMenuPosition(getMenuPosition(wrapperRef.current));
			}
		};

		const handleDocumentMouseDown = (event: MouseEvent): void => {
			const target = event.target as Node | null;
			if (
				target &&
				((wrapperRef.current && wrapperRef.current.contains(target)) ||
					(listRef.current && listRef.current.contains(target)))
			) {
				return;
			}

			setIsOpen(false);
			setActiveIndex(-1);
			setQuery(selectedOptionText);
		};

		updatePosition();
		document.addEventListener("mousedown", handleDocumentMouseDown, true);
		window.addEventListener("resize", updatePosition);
		document.addEventListener("scroll", updatePosition, true);

		return () => {
			document.removeEventListener("mousedown", handleDocumentMouseDown, true);
			window.removeEventListener("resize", updatePosition);
			document.removeEventListener("scroll", updatePosition, true);
		};
	}, [isOpen, disabled, selectedOptionText]);

	const openList = (): void => {
		if (disabled) {
			return;
		}

		if (wrapperRef.current) {
			setMenuPosition(getMenuPosition(wrapperRef.current));
		}

		setIsOpen(true);
		if (filteredOptions.length > 0) {
			setActiveIndex(0);
		} else {
			setActiveIndex(-1);
		}
	};

	const closeList = (resetQuery: boolean): void => {
		setIsOpen(false);
		setActiveIndex(-1);
		if (resetQuery) {
			setQuery(selectedOptionText);
		}
	};

	const handleSelect = (option: ISearchableDropdownOption): void => {
		setQuery(option.text);
		props.onChange(props.targetProperty, option.key, true);
		setIsOpen(false);
		setActiveIndex(-1);
	};

	const handleClear = (): void => {
		if (disabled) {
			return;
		}

		setQuery("");
		props.onChange(props.targetProperty, "", true);
		setIsOpen(true);
		setActiveIndex(props.options.length > 0 ? 0 : -1);
		if (inputRef.current) {
			inputRef.current.focus();
		}
	};

	const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
		const nextValue = event.target.value;
		setQuery(nextValue);
		if (nextValue === "" && typeof props.selectedKey !== "undefined" && props.selectedKey !== "") {
			props.onChange(props.targetProperty, "", true);
		}
		setIsOpen(true);
		setActiveIndex(0);
	};

	const handleInputKeyDown = (event: React.KeyboardEvent<HTMLInputElement>): void => {
		if (event.key === "ArrowDown") {
			event.preventDefault();
			if (!isOpen) {
				openList();
				return;
			}
			if (filteredOptions.length > 0) {
				setActiveIndex(Math.min(activeIndex + 1, filteredOptions.length - 1));
			}
			return;
		}

		if (event.key === "ArrowUp") {
			event.preventDefault();
			if (!isOpen) {
				openList();
				return;
			}
			if (filteredOptions.length > 0) {
				setActiveIndex(Math.max(activeIndex - 1, 0));
			}
			return;
		}

		if (event.key === "Enter") {
			if (isOpen && activeIndex >= 0 && activeIndex < filteredOptions.length) {
				event.preventDefault();
				handleSelect(filteredOptions[activeIndex]);
			}
			return;
		}

		if (event.key === "Escape") {
			event.preventDefault();
			closeList(true);
			return;
		}

		if (event.key === "Tab") {
			closeList(true);
		}
	};

	const handleInputBlur = (): void => {
		window.setTimeout(() => {
			const activeElement = document.activeElement;
			if (
				activeElement &&
				((wrapperRef.current && wrapperRef.current.contains(activeElement)) ||
					(listRef.current && listRef.current.contains(activeElement)))
			) {
				return;
			}
			closeList(true);
		}, 0);
	};

	const inputStyle = errorMessage
		? { ...(disabled ? inputDisabledStyle : inputBaseStyle), ...inputErrorStyle }
		: disabled
			? inputDisabledStyle
			: inputBaseStyle;

	const showClearButton = !disabled && query !== "";
	const toggleButtonStyle = disabled ? { ...chevronButtonStyle, ...iconButtonDisabledStyle } : chevronButtonStyle;

	const menu = isOpen && menuPosition && portalContainerRef.current
		? ReactDOM.createPortal(
				<div
					ref={listRef}
					id={listboxId}
					role="listbox"
					style={{
						...menuStyleBase,
						left: `${menuPosition.left}px`,
						top: `${menuPosition.top}px`,
						width: `${menuPosition.width}px`,
						maxHeight: `${menuPosition.maxHeight}px`,
					}}
				>
						{filteredOptions.length === 0 ? (
							<div style={emptyStateStyle}>No matches found.</div>
						) : (
							filteredOptions.map((option, index) => {
								const isSelected = !!selectedOption && option.key === selectedOption.key;
								const isActive = index === activeIndex;
								return (
									<div
										key={option.key || `option-${index}`}
										id={`${fieldId}-option-${index}`}
										role="option"
										aria-selected={isSelected}
										aria-disabled={!!option.disabled}
										style={getOptionStyle(option, isSelected, isActive)}
										onMouseDown={(event) => {
											event.preventDefault();
											if (!option.disabled) {
												handleSelect(option);
											}
										}}
										onMouseEnter={() => {
											if (!option.disabled) {
												setActiveIndex(index);
											}
										}}
									>
										{option.text}
								</div>
							);
						})
					)}
				</div>,
				portalContainerRef.current
			)
		: null;

	return (
		<>
			<div style={containerStyle}>
				<label htmlFor={inputId} style={labelStyle}>
					{props.label}
				</label>
				<div ref={wrapperRef} style={wrapperStyle}>
					<input
						ref={inputRef}
						id={inputId}
						type="text"
						role="combobox"
						aria-autocomplete="list"
						aria-activedescendant={activeDescendantId}
						aria-expanded={isOpen}
						aria-controls={listboxId}
						aria-describedby={errorMessage ? errorId : undefined}
						aria-haspopup="listbox"
						aria-invalid={!!errorMessage}
						value={query}
						placeholder={props.placeholder || `Search ${props.label.toLowerCase()}...`}
						disabled={disabled}
						autoComplete="off"
						style={inputStyle}
						onFocus={openList}
						onClick={() => {
							if (!isOpen) {
								openList();
							}
						}}
						onChange={handleInputChange}
						onKeyDown={handleInputKeyDown}
						onBlur={handleInputBlur}
					/>
					<button
						type="button"
						aria-label={`Clear ${props.label}`}
						disabled={disabled}
						style={showClearButton ? clearButtonVisibleStyle : clearButtonHiddenStyle}
						onMouseDown={(event) => {
							event.preventDefault();
							handleClear();
						}}
					>
						×
					</button>
					<button
						type="button"
						aria-label={`Toggle ${props.label} options`}
						disabled={disabled}
						style={toggleButtonStyle}
						onMouseDown={(event) => {
							event.preventDefault();
							if (disabled) {
								return;
							}
							if (isOpen) {
								closeList(false);
							} else {
								openList();
								if (inputRef.current) {
									inputRef.current.focus();
								}
							}
						}}
						>
						{isOpen ? "▲" : "▼"}
					</button>
				</div>
				{errorMessage ? (
					<span id={errorId} role="alert" style={errorMessageStyle}>
						{errorMessage}
					</span>
				) : null}
			</div>
			{menu}
		</>
	);
};

export function PropertyPaneSearchableDropdown(
	props: IPropertyPaneSearchableDropdownProps
): IPropertyPaneField<ISearchableDropdownCustomFieldProps> {
	return {
		type: PropertyPaneFieldType.Custom,
		targetProperty: props.targetProperty,
		properties: {
			...props,
			key: `searchableDropdown-${props.targetProperty}`,
			onRender: (
				domElement: HTMLElement,
				_context?: unknown,
				changeCallback?: ChangeCallback
			): void => {
				ReactDOM.render(
					<SearchableDropdownField
						{...props}
						onChange={changeCallback || (() => undefined)}
					/>,
					domElement
				);
			},
			onDispose: (domElement: HTMLElement): void => {
				ReactDOM.unmountComponentAtNode(domElement);
			},
		},
	};
}
