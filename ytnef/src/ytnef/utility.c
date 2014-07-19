/*
*    Yerase's TNEF Stream Reader
*    Copyright (C) 2003  Randall E. Hand
*
*    This program is free software; you can redistribute it and/or modify
*    it under the terms of the GNU General Public License as published by
*    the Free Software Foundation; either version 2 of the License, or
*    (at your option) any later version.
*
*    This program is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*    GNU General Public License for more details.
*
*    You should have received a copy of the GNU General Public License
*    along with this program; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*
*    You can contact me at randall.hand@gmail.com for questions or assistance
*/
void fprintProperty(TNEFStruct TNEF, FILE *FPTR, DWORD PROPTYPE, DWORD PROPID, char TEXT[]) {
    variableLength *vl;
    if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PROPTYPE, PROPID))) != MAPI_UNDEFINED) { 
        if (vl->size > 0) {
            if ((vl->size == 1) && (vl->data[0] == 0)) {

            } else { 
                fprintf(FPTR, TEXT, vl->data); 
            }
        }
    }
}

void fprintUserProp(TNEFStruct TNEF, FILE *FPTR, DWORD PROPTYPE, DWORD PROPID, char TEXT[]) {
    variableLength *vl;
    if ((vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PROPTYPE, PROPID))) != MAPI_UNDEFINED) { 
        if (vl->size > 0) {
            if ((vl->size == 1) && (vl->data[0] == 0)) {
            } else { 
                fprintf(FPTR, TEXT, vl->data); 
            } 
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

