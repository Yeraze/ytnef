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

void PrintRTF(FILE *fptr, variableLength *VL) {
    int index;
    char *byte;
    int brace_ct;
    int key;

    key = 0;
    brace_ct = 0;

    for(index = 0, byte=VL->data; index < VL->size; index++, byte++) {
        if (*byte == '}') {
            brace_ct--;
            key = 0;
            continue;
        }
        if (*byte == '{') {
            brace_ct++;
            continue;
        }
        if (*byte == '\\') {
            key = 1;
        }
        if (isspace(*byte)) {
            key = 0;
        }
        if ((brace_ct == 1) && (key == 0)) {
            if (*byte == '\n') { 
                fprintf(fptr, "\\n"); 
            } else if (*byte == '\r') { 
                // Print nothing.
            } else if (*byte == ';') {
                fprintf(fptr, "\\;");
            } else if (*byte == ',') {
                fprintf(fptr, "\\,");
            } else if (*byte == '\\') {
                fprintf(fptr, "\\");
            } else { 
                fprintf(fptr, "%c", *byte );
            } 
        }
    }
    fprintf(fptr, "\n");

}

