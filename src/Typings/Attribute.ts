
import { LookupType } from "./CommonTypes";
import { Control } from "./Control";
import { ExecutionContext } from "./ExecutionContext";
import { FormCollection } from "./FormContext";

/**
 * @author Thomas Carlsen
 */


/**
 * See the **formContext.data.entity.attributes** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes).
 * 
 * _@description_: Columns (aka Attributes or Fields) contain data in the model-driven apps form or grids. 
 * Use the formContext.data.entity.attributes collection or the formContext.getAttribute shortcut method to access a collection of columns. 
 */
export interface Attribute {

    /**
     * See the **Controls-collection for attributes** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/controls-collection).
     * 
     * Use the Controls collection to access controls associated with columns/fields/attributes.
     */
    controls: FormCollection<Control>;

    /**
     * See the **getAttributeType** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getattributetype).
     * 
     * @return String value that represents the type of the column/attribute/field.
     */
    getAttributeType(): AttributeType;


    /**
     * See the **addOnChange** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/addonchange).
     * 
     * _@description_: Sets a function to be called when the OnChange event occurs. 
     * The `execution context` is automatically passed as the first parameter to this function.
     * @param functionRef Function name or an anonymous function.
     */
     addOnChange(functionRef: () => void): void; // TODO - is it possible to tell users that type of the first arguement of their function is ExecutionContext if they have any argument at all 

    /**
     * See the **addOnChange** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/addonchange).
     * 
     * _@description_: Sets a function to be called when the OnChange event occurs. 
     * The `execution context` is automatically passed as the first parameter to this function.
     * @param functionRef Function name or an anonymous function.
     */
    addOnChange(functionRef: (exe: ExecutionContext, ...rest: any) => void): void; // TODO - is it possible to tell users that type of the first arguement of their function is ExecutionContext if they have any argument at all 

    /**
     * See the **fireOnChange** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/fireOnChange).
     * 
     * _@description_: Causes the OnChange event to occur on the column so that any script associated to that event can execute.
     */
    fireOnChange(): void;

    /**
     * See the **getFormat** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getFormat).
     * 
     *  _@description_: This method will return one of the following string values or null:
     * - date
     * - datetime
     * - duration
     * - email
     * - language
     * - none
     * - phone
     * - text
     * - textarea
     * - tickersymbol
     * - timezone
     * - url
     * @return String value that represents formatting options for the column.
     */
    getFormat(): AttributeFormatOption | null;

    /**
     * See the **getIsDirty** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getIsDirty).
     * 
     * _@description_: Returns a boolean value indicating if there are unsaved changes to the column value. 
     * An unsaved change to a column value means the client value is different from the last known committed value retrieved from Dataverse by the client from runtime.
     * @return True if there are unsaved changes, otherwise false.
     */
    getIsDirty(): boolean;

    /**
     * See the **getName** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getName).
     * 
     * @return String representing the logical name of the column
     */    
    getName(): string;

    /**
     * See the **getParent** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getParent).
     * 
     * @return The `formContext.data.entity` object that is the parent to all the columns.
     */
    getParent(): {};

    /**
     * See the **getRequiredLevel** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getRequiredLevel).
     * 
     * @return A string value indicating whether a value for the column is required or recommended. "none", "required" or "recommended".
     */
    getRequiredLevel(): RequirementLevel;

    /**
     * See the **getSubmitMode** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getSubmitMode).
     * 
     * @return A string indicating when data from the column will be submitted when the record is saved. "always", "never" or "dirty".
     */
    getSubmitMode(): SubmitMode;

    /**
     * See the **getUserPrivilege** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getUserPrivilege).
     * 
     * @return An object with three boolean properties corresponding to privileges indicating if the user can create, read or update data values for a column. This function is intended for use when Field Level Security modifies a user’s privileges for a particular column.
     */
    getUserPrivilege(): {canRead: boolean; canUpdate: boolean; canCreate: boolean};

    /**
     * See the **getValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getValue).
     * 
     * @return Retrieves the data value for a column. Please look at the doc for further info.
     */
    getValue(): Boolean | Date | Number | LookupType[] | String | Number[] | null;

    /**
     * See the **isValid** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/isValid).
     * 
     * @return A boolean value to indicate whether the value of a column is valid.
     */
    isValid(): boolean;

    /**
     * See the **removeOnChange** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/removeOnChange).
     * 
     * _@description_: Removes a function from the OnChange event handler for a column.
     * @param functionRef Function name or an anonymous function.
     */
    removeOnChange(functionRef: Function): void;

    /**
     * See the **setRequiredLevel** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/setRequiredLevel).
     * 
     * _@description_: Sets whether data is required or recommended for the column before the record can be saved.
     * @param requirementLevel {@link RequirementLevel}.
     */
    setRequiredLevel(requirementLevel: RequirementLevel): void;
    // setRequiredLevel(requirementLevel: RequirementLevel): void;

    /**
     * See the **setSubmitMode** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/setSubmitMode).
     * 
     * _@description_: Sets whether data from the column will be submitted when the record is saved.
     * @param submitMode {@link SubmitMode}.
     */
    setSubmitMode(submitMode: SubmitMode): void;

    /**
     * See the **setValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/setValue).
     * 
     * _Note_: Updating a column using setValue will not cause the OnChange event handlers to run. 
     * If you want the OnChange event handlers to run you must use fireOnChange in addition to setValue.
     * 
     * _@description_: Sets the data value for a column.
     * @param value Depends on the type of column. Please look at the doc for further info.
     */
    setValue(value: Boolean | Date | Number | LookupType[] | String | Number[] | null): void; // TODO: We could overload instead of or'ing the types. This way we can have finer descriptions.

    /**
     * See the **setIsValid** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/setIsValid).
     * 
     * _@description_: Sets a value for a column to determine whether it is valid or invalid with a message.
     * @param bool Specify false to set the column value to invalid and true to set the value to valid.
     * @param message The message to display.
     */
    setIsValid(bool: boolean, message?: string): void;
}

export interface BooleanAttribute extends Attribute {
    /**
     * See the **getInitialValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getinitialvalue).
     * 
     * _@description_: Returns a value that represents the value set for a Yes/No, Choice or Choices column when the form is opened.
     * @return The initial value for the column.
     */
    getInitialValue(): Number;
}

export interface LookupAttribute extends Attribute {
    /**
     * See the **getIsPartyList** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getIsPartyList).
     * 
     * _@description_: Returns a boolean value indicating if there are unsaved changes to the column value. 
     * An unsaved change to a column value means the client value is different from the last known committed value retrieved from Dataverse by the client from runtime.
     * @return A string indicating when data from the column will be submitted when the record is saved.
     */
    getIsPartyList(): boolean;

    // TODO
    /**
     * See the **getValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getValue).
     * 
     * @return Retrieves the data value for a column or null if the column doesn't have a value. {@link LookupType} also know as EntityReference.
     */
    getValue(): LookupType[] | null;
}

export interface CommonOptionSetAttribute extends Attribute {
    /**
     * See the **getInitialValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getinitialvalue).
     * 
     * _@description_: Returns a Boolean value indicating whether the lookup represents a partylist lookup. 
     * Partylist lookups allow for multiple records to be set, such as the To: column for an email table record.
     * @return True if the lookup column is a partylist, otherwise false.
     */
    getInitialValue(): Number;
    
    /**
     * See the **getOption** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getOption).
     * 
     * _@description_: Returns an option object with the value matching the argument (label or enumeration value) passed to the method.
     * @value Number (enumeration value of the option).
     * @return The logical name of the column.
     */
    getOption(value: number): Option;

    /**
     * See the **getOptions** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getOptions).
     * 
     * @return An array of option objects representing valid options for a column.
     */
    getOptions(): Option[];

    /**
     * See the **getText** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getText).
     * 
     * @return Returns a string value of the text for the currently selected option for a choice or choices column.
     */
    getText(): string;
}

// AKA Optionset
export interface ChoiceAttribute extends CommonOptionSetAttribute { 
    /**
     * See the **getSelectedOption** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getSelectedOption).
     * 
     * @return The option object - {@link Option}
     */
    getSelectedOption(): Option; 

    // TODO:
    /**
     * See the **getValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getValue).
     * 
     * @return Retrieves the data value for a column or null if the column doesn't have a value.
     */
    getValue(): number | null;
}

// AKA Multi-select Optionset
export interface ChoicesAttribute extends CommonOptionSetAttribute {
    /**
     * See the **getSelectedOption** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getSelectedOption).
     * 
     * @return An array of option objects selected in choices - {@link Option}
     */
    getSelectedOption(): Option[]; 

    // TODO:
    /**
     * See the **getValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getValue).
     * 
     * @return Retrieves the data value for a column or null if the column doesn't have a value.
     */
    getValue(): number[] | null;
}

// Number column type (decimal, double, integer, money)
export interface NumberAttribute extends Attribute {
    /**
     * See the **getMax** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getMax).
     * 
     * @return A number indicating the maximum allowed value for a column
     */
    getMax(): number;

    /**
     * See the **getMin** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getMin).
     * 
     * @return A number indicating the minimum allowed value for a column.
     */
    getMin(): number;

    /**
     * See the **getPrecision** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getPrecision).
     * 
     * @return The number of digits allowed to the right of the decimal point.
     */
    getPrecision(): number;

    /**
     * See the **setPrecision** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/setPrecision).
     * 
     * _@description_: Sets the number of digits allowed to the right of the decimal point.
     * @value Number of digits allowed to the right of the decimal point.
     */
    setPrecision(value: number): void;
}

export interface StringAttribute extends Attribute {
    /**
     * See the **getMaxLength** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getMaxLength).
     * 
     * _Note_: The email form description column is a memo column, but it does not have a getMaxLength method.
     * @return A number indicating the maximum length of a string or memo column.
     */
    getMaxLength(): number;

    /**
     * See the **getValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/getValue).
     * 
     * @return Retrieves the data value for a column.
     */
    getValue(): string | null;

    // TODO rest for specific function for this interface
    //setValue(); // TODO
}


// ENUMS



export enum AttributeType {
    boolean = "boolean",
    datetime = "datetime",
    decimal = "decimal",
    double = "double",
    integer = "integer",
    lookup = "lookup",
    memo = "memo",
    money = "money",
    choices = "choices",
    choice = "choice", // Looks like it is still outputting "optionset"
    string = "string",
}

export enum AttributeFormatOption {
    date = "date",
    datetime = "datetime",
    duration = "duration",
    email = "email",
    language = "language",
    none = "none",
    phone = "phone",
    text = "text",
    textarea = "textarea",
    tickersymbol = "tickersymbol",
    timezone = "timezone",
    url = "url",
}

export type RequirementLevel = "none" | "required" | "recommended";

// export enum RequirementLevel {
//     none = "none", 
//     required = "required",
//     recommended = "recommended",
// }

/**
 * - always: The data is always sent with a save.
 * - never: The data is never sent with a save. When this value is used, the column(s) in the form for this column cannot be edited.
 * - dirty: Default behavior. The data is sent with the save when it has changed.
 */
export enum SubmitMode {
    always = "always", 
    never = "never", 
    dirty = "dirty",
}





// Other types

export interface Option {
    text: string;
    value: number;
}