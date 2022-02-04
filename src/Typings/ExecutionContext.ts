
import { Attribute } from "./Attribute";
import { FormContext } from "./FormContext";


// Docs: https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/execution-context
/**
 * ExecutionContext
 */
export interface ExecutionContext {
    getDepth(): {}; // TODO
    getEventArgs(): {}; // TODO
    getEventSource(): {}; // TODO
    getFormContext(): FormContext; 
    getSharedVariable(): {}; // TODO
    setSharedVariable(): {}; // TODO
}