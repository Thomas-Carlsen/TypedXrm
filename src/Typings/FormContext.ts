import { Attribute } from "./Attribute";
import { LookupType } from "./CommonTypes";


// Docs: https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/clientapi-form-context
export interface FormContext {
    data: FormContextData;

    // https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-ui
    ui: {}; // TODO
    
    // Shortcuts

    /**
     * This is a shortcut for the `formContext.data.entity.attributes`. 
     * Read more about it in the documentaion [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes). 
     * @return All attributes on the form
     */
    getAttribute(): Attribute[];
    getAttribute(attributeName: string): Attribute;
    getAttribute(index: number): Attribute;
}

// Docs: https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data
/**
 * Provides properties and methods to work with the data on a form, including table data and data in the business process flow control.
 */
export interface FormContextData {
    /**
     * Collection of non-table data on the form. Items in this collection are of the same type as the column collection, but they are not columns of the form table.
     */
    attributes: FormCollection<Attribute>;
    
    
    /**
     * See the **formcontext.data.entity** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity).
     * 
     * Provides properties and methods to retrieve information specific to the record displayed on the page, the save method, and a collection of all the columns included in the form. 
     * Column data is limited to columns represented on the form.
     */
    entity: FormContextEntity;
    
    
    /**
     * See the **formcontext.data.process** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-process).
     * 
     * Provides events, methods, and objects to interact with the business process flow data on a form. 
     * See formContext.ui.process (Client API reference) for methods to interact with the business process flow control on the form.
     */
    process: FormContextProcess; //TODO
}

// TODO Make below interfaces
/**
 * General functionality for a form collection.
 * Collection elements can be:
 * - Attributes
 * - Controls
 * - Sections
 * - Tabs
 * - QuickForms
 * - Navigation Items
 * - FormSelector Items
 * - Process Stages
 * - Process Steps
 */
export interface FormCollection<ElementType> {
    forEach(element: ElementType): any; // TODO - what type do the collection elements have and what do the foreach return if something at all?
    // To access a column within the collection, you pass either the name (string) or the index value (number) of the column as an argument to the method
    // - But that is only for columns - I will use this as a base until an example of a more complex collection element appears
    get(): ElementType[]; 
    get(elementName: string): ElementType; // TODO - know that this is used with a string to get attributes - formContext.getAttribute("bla") is a shorthand of this function - use below as guidance
    get(index: number): ElementType;
    getLength(): number;
}


/**
 * See the **formcontext.data.entity** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity).
 * 
 * Provides properties and methods to retrieve information specific to the record displayed on the page, the save method, and a collection of all the columns included in the form. 
 * Column data is limited to columns represented on the form.
 */
export interface FormContextEntity {
    /**
     * See the **attributes on form** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes).
     * 
     * Provides access to each table column that is available on the form. 
     * Only those columns added to the form are available. 
     * Use the `formContext.data.entity.attributes` collection or the `formContext.getAttribute` shortcut method to access a collection of columns.
     */
    attributes: FormCollection<Attribute>;

    /**
     * See the **addOnSave** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/addOnSave).
     * 
     * _@description_: Adds a function to be called when the OnSave event is triggered. The event occurs before the save occurs, giving the handler an option to cancel the save operation.
     * @param myFunction The function to be executed when the record is saved. The function will be added to the bottom of the event handler pipeline. The execution context is automatically passed as the first parameter to the function. See Execution context for more information.
     */
    addOnSave(myFunction: Function): void;

    /**
     * See the **addOnPostSave** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/events/postsave).
     * 
     * _@description_: PostSave event occurs after the OnSave event is complete. This event is used to support or execute custom logic using web resources to perform after Save actions when the save event is successful or failed due to server errors.
     * @param myFunction The function to add to the PostSave event. The execution context is automatically passed as the first parameter to this function.
     */
    addOnPostSave(myFunction: Function): void; 
    
    /**
     * See the **getDataXml** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/getDataXml).
     * 
     * @return A string representing the XML that will be sent to the server when the record is saved. Only data in columns that have changed or have their submit mode set to "always" are sent to the server.
     */
    getDataXml(): string;  
    
    /**
    * See the **getEntityName** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/getEntityName).
    * 
    * @return A string representing the logical name of the table for the record.
    */
    getEntityName(): string; 

    /**
    * See the **getEntityReference** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/getEntityReference).
    * 
    * @return A lookup ({@link LookupType}) value that references the record. 
    */
    getEntityReference(): LookupType; 

    /**
    * See the **getId** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/getId).
    * 
    * @return A string representing the GUID value for the record.
    */
    getId(): string; 

    /**
    * See the **getIsDirty** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/getIsDirty).
    * 
    * _@description_: Gets a boolean value indicating whether any columns in the form have been modified.
    * @return True if any columns in the form have been changed; false otherwise.
    */
    getIsDirty(): boolean; 
    
    /**
    * See the **getPrimaryAttributeValue** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/getPrimaryAttributeValue).
    * 
    * _@description_: Gets a string for the value of the primary column of the table.
    * @return The name of the table.
    */
    getPrimaryAttributeValue(): string; 
    
    /**
    * See the **isValid** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/isValid).
    * 
    * _@description_: Gets a boolean value indicating whether all of the table data is valid.
    * @return True if all of the table data is valid; false otherwise.
    */
    isValid(): boolean; 
    
    /**
    * See the **removeOnSave** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/removeOnSave).
    * 
    * _@description_: Gets a boolean value indicating whether all of the table data is valid.
    * @param myFunction The function to be removed for the OnSave event.
    */
    removeOnSave(myFunction: Function): void;

    /**
     * See the **removeOnPostSave** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/controls/removeonpostsave).
     * 
     * _@description_: Removes the event handler from the PostSave event.
     * @param myFunction The function to be removed from the PostSave event.
     */
    removeOnPostSave(myFunction: Function): void;

    /**
     * See the **save** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-entity/save).
     * 
     * @deprecated This method is deprecated and we recommend to use the [formContext.data.save](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data/save) method.
     * 
     * _@description_: Saves the record synchronously with the options to close the form or open a new form after the save is completed.
     * 
     * @param saveOption (Optional) Specify options for saving the record. If no parameter is included in the method, the record will simply be saved. This is the equivalent of using the Save command. You can specify one of the following values:
        - **saveandclose**: This is the equivalent of using the Save and Close command.
        - **saveandnew**: This is the equivalent of the using the Save and New command.
     */
    save(saveOption?: string): void;
}


/**
 * See the **formcontext.data.process** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-process).
 * 
 * Provides events, methods, and objects to interact with the business process flow data on a form. 
 * See formContext.ui.process (Client API reference) for methods to interact with the business process flow control on the form.
 */
export interface FormContextProcess {

    /**
     * See the **addOnPreProcessStatusChange** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-process/eventhandlers/addonpreprocessstatuschange).
     * 
     * _@description_: Adds a function as an event handler for the [OnPreProcessStatusChange](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/events/onpreprocessstatuschange) event so that it will be called **before** the business process flow status changes.
     * 
     * _Note_: This client API is only supported on the Unified Client. The legacy web client does not support this client API.
     * 
     * @param myFunction The function to be executed when the business process flow status changes. The function will be added to the start of the event handler pipeline. The execution context is automatically passed as the first parameter to the function. See Execution context for more information. 
     * 
     * You should use a reference to a named function rather than an anonymous function if you may later want to remove the event handler.
     */
    addOnPreProcessStatusChange(myFunction: Function): void; 

    /**
     * See the **removeOnPreProcessStatusChange** documentation [here](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-data-process/eventhandlers/removeOnPreProcessStatusChange).
     * 
     * _@description_: Removes an event handler from the [OnPreProcessStatusChange](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/events/onpreprocessstatuschange) event.
     * 
     * @param myFunction The function to be removed from the [OnPreProcessStatusChange](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/events/onpreprocessstatuschange) event.
     * 
     */
    removeOnPreProcessStatusChange(myFunction: Function): void; 


    addOnProcessStatusChange(): {}; //TODO
    removeOnProcessStatusChange(): {}; //TODO

    addOnStageChange(): {}; //TODO
    removeOnStageChange(): {}; //TODO

    addOnStageSelected(): {}; //TODO
    removeOnStageSelected(): {}; //TODO

    getActiveProcess(): {}; //TODO
    setActiveProcess(): {}; //TODO

    getId(): {}; //TODO
    getName(): {}; //TODO
    getStages(): {}; //TODO
    isRendered(): {}; //TODO

    getProcessInstances(): {}; //TODO
    setActiveProcessInstance(): {}; //TODO

    getInstanceId(): {}; //TODO
    getInstanceName(): {}; //TODO
    getStatus(): {}; //TODO
    setStatus(): {}; //TODO

    getActiveStage(): {}; //TODO
    setActiveStage(): {}; //TODO

    moveNext(): {}; //TODO
    movePrevious(): {}; //TODO

    getActivePath(): {}; //TODO
    getEnabledProcesses(): {}; //TODO
    getSelectedStage(): {}; //TODO
}