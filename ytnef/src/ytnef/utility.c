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
#include <ctype.h>
#include "settings.h"
void fprintProperty(TNEFStruct TNEF, FILE *FPTR, DWORD PROPTYPE, DWORD PROPID,
                    char TEXT[]) {
  variableLength *vl;
  if ((vl = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PROPTYPE,
                             PROPID))) != MAPI_UNDEFINED) {
    if (vl->size > 0) {
      if ((vl->size == 1) && (vl->data[0] == 0)) {

      } else {
        fprintf(FPTR, TEXT, vl->data);
      }
    }
  }
}

void fprintUserProp(TNEFStruct TNEF, FILE *FPTR, DWORD PROPTYPE, DWORD PROPID,
                    char TEXT[]) {
  variableLength *vl;
  if ((vl = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PROPTYPE,
                             PROPID))) != MAPI_UNDEFINED) {
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

  for (index = 0; index < VL->size - 1; index++) {
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

  for (index = 0; index < strlen((char*)VL->data); index++) {
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

void ConsumeHyperlink(FILE *fptr, char** str) {
  char *link = NULL;
  char *text = NULL;
  char *bgn = *str;
  char *end = NULL;
  bgn = strchr(bgn, '"');
  bgn++;
  end = strchr(bgn, '"');
  unsigned int size = end - bgn;
  link = calloc(size + 1, 1);
  strncpy(link, bgn, size);
  link[size] = '\0';
  bgn = strstr(end, "\\fldrslt");
  int bSlash = 1;
  for (;;) {
    bgn++;
    if (*bgn == '\\') {
      bSlash = 1;
      continue;
    } else if (*bgn == ' ') {
      bSlash = 0;
      continue;
    } else if (*bgn == '{') {
      bSlash = 0;
      continue;
    } else if (*bgn == '}') {
      break;
    }
	  if (!bSlash) {
      break;
    }
  }
  if (*bgn != '}') {
    end = strchr(bgn, '}');
    size = end - bgn;
    text = calloc(size + 1, 1);
    strncpy(text, bgn, size);
    text[size] = '\0';
  }
  if ((text != NULL) && (strcmp(link, text) != 0)) {
    fprintf(fptr, "%s <%s>", text, link );
  } else {
    fprintf(fptr, "%s", link );
  }
  free(text);
  free(link);
  *str = end;
  return;
}

void PrintRTF(FILE *fptr, variableLength *VL) {
  char *byte = 0;
  int bSlash = 0;
  byte = strstr((char*)VL->data, "\\pard");
  if (!byte)
    return;
  for (; *byte != '\0'; byte++) {
    if (*byte == '\\') {
      bSlash = 1;
      continue;
    } else if (*byte == ' ') {
      if (bSlash) {
        bSlash = 0;
        continue;
      }
    } else if (*byte == '{') {
      bSlash = 0;
      continue;
    } else if (*byte == '}') {
      bSlash = 0;
      continue;
    } else if (*byte == 'l' && bSlash) {
      if(strncmp(byte, "listtext", 8) == 0) {
        fprintf(fptr, "*\t");
        continue;
      } else
      if(strncmp(byte, "lsdlocked", 9) == 0) {
        byte += 9;
        if(isdigit(*byte)){
          byte = strchr(byte, ';');
          continue;
        }
      }
    } else if (*byte == 'd' && bSlash) {
      if(strncmp(byte, "datastore", 9) == 0) {
        byte = strchr(byte, '}');
        continue;
      }
    } else if (*byte == '~' && bSlash) {
      fprintf(fptr, " ");
      bSlash = 0;
      continue;
    }

    if (!bSlash) {
      if (*byte == 'H') {
        if(strncmp(byte, "HYPERLINK", 9) == 0) {
          ConsumeHyperlink(fptr, &byte);
          continue;
        }
      } else if (*byte == '\n') { 
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
      continue;
    }
  }
  fprintf(fptr, "\n");
  return;
}
