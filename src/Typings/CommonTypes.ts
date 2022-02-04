

/**
 * _NOTE_: Certain lookups allow for multiple records to be associated in a lookup, such as the To: column for an email table record. 
 * Therefore, all lookup data values use an array of lookup objects – even when the lookup column does not support more than one record reference to be added.
 * 
 * - entityType: String. The name of the table displayed in the lookup.
 * - id: String: The string representation of the GUID value for the record displayed in the lookup.
 * - name: String: The text representing the record to be displayed in the lookup.
 */
export interface LookupType {
    entityType: string;
    id: string;
    name: string;
}