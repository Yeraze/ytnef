void fprintProperty(TNEFStruct TNEF, FILE *FPTR, DWORD PROPTYPE, DWORD PROPID, char TEXT[]) {
    variableLength *vl;
    if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PROPTYPE, PROPID))) != MAPI_UNDEFINED) { 
        if (vl->size > 0)  
            if ((vl->size == 1) && (vl->data[0] == 0)) {
            } else { 
                fprintf(FPTR, TEXT, vl->data); 
            } 
    }
}

void fprintUserProp(TNEFStruct TNEF, FILE *FPTR, DWORD PROPTYPE, DWORD PROPID, char TEXT[]) {
    variableLength *vl;
    if ((vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PROPTYPE, PROPID))) != MAPI_UNDEFINED) { 
        if (vl->size > 0)  
            if ((vl->size == 1) && (vl->data[0] == 0)) {
            } else { 
                fprintf(FPTR, TEXT, vl->data); 
            } 
    }
}

void quotedfprint(FILE *FPTR, variableLength *VL) {
    int index;

    for (index=0;index<VL->size-1; index++) { 
        if (VL->data[index] == '\n') { 
            fprintf(FPTR, "=0A"); 
        } else if (VL->data[index] == '\r') { 
        } else { 
            fprintf(FPTR, "%c", VL->data[index]); 
        } 
    }
}

void Cstylefprint(FILE *FPTR, variableLength *VL) {
    int index;

    for (index=0;index<VL->size-1; index++) { 
        if (VL->data[index] == '\n') { 
            fprintf(FPTR, "\\n"); 
        } else if (VL->data[index] == '\r') { 
            // Print nothing.
        } else if (VL->data[index] == ';') {
            fprintf(FPTR, "\\;");
        } else if (VL->data[index] == ',') {
            fprintf(FPTR, "\\,");
        } else if (VL->data[index] == '\\') {
            fprintf(FPTR, "\\");
        } else { 
            fprintf(FPTR, "%c", VL->data[index]); 
        } 
    }
}

