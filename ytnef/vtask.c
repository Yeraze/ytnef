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
#include "settings.h"
void SaveVTask(TNEFStruct TNEF) {
  variableLength *vl;
  variableLength *filename;
  int index;
  char ifilename[MAX_FILENAME_SIZE];
  char *charptr, *charptr2;
  dtr thedate;
  FILE *fptr;
  DDWORD *ddword_ptr;

  vl = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8,
                        PR_CONVERSATION_TOPIC));

  if (vl == MAPI_UNDEFINED) {
    return;
  }

  index = strlen((char*)vl->data);
  while (vl->data[index] == ' ')
    vl->data[index--] = 0;

  snprintf(ifilename, MAX_FILENAME_SIZE, "%s.vcf", vl->data);
  SanitizeFilename(ifilename);
  if (filepath) {
    char temp[MAX_FILENAME_SIZE];
    memcpy(temp, ifilename, MAX_FILENAME_SIZE);
    snprintf(ifilename, MAX_FILENAME_SIZE, "%s/%s", filepath, temp);
  }
  printf("%s\n", ifilename);

  if (savefiles == 0) return;

  if ((fptr = fopen(ifilename, "wb")) == NULL) {
    printf("Error writing file to disk!");
  } else {
    fprintf(fptr, "BEGIN:VCALENDAR\n");
    fprintf(fptr, PRODID);
    fprintf(fptr, "VERSION:2.0\n");
    fprintf(fptr, "METHOD:PUBLISH\n");
    filename = NULL;

    fprintf(fptr, "BEGIN:VTODO\n");
    if (TNEF.messageID[0] != 0) {
      fprintf(fptr, "UID:%s\n", TNEF.messageID);
    }
    filename = MAPIFindUserProp(&(TNEF.MapiProperties), \
                                PROP_TAG(PT_STRING8, 0x8122));
    if (filename != MAPI_UNDEFINED) {
      fprintf(fptr, "ORGANIZER:%s\n", filename->data);
    }


    if ((filename = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8,
                                     PR_DISPLAY_TO))) != MAPI_UNDEFINED) {
      filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8,
                                  0x811f));
    }
    if ((filename != MAPI_UNDEFINED) && (filename->size > 1)) {
      charptr = (char*)filename->data - 1;
      charptr2 = (char *)strstr((char*)(charptr + 1), ";");
      while (charptr != NULL) {
        charptr++;
        charptr2 = (char *)strstr((char *)charptr, ";");
        if (charptr2 != NULL) {
          *charptr2 = 0;
        }
        while (*charptr == ' ')
          charptr++;
        fprintf(fptr, "ATTENDEE;CN=%s;ROLE=REQ-PARTICIPANT:%s\n", charptr, charptr);
        charptr = charptr2;
      }
    }

    if (TNEF.subject.size > 0) {
      fprintf(fptr, "SUMMARY:");
      Cstylefprint(fptr, &(TNEF.subject));
      fprintf(fptr, "\n");
    }

    if (TNEF.body.size > 0) {
      fprintf(fptr, "DESCRIPTION:");
      Cstylefprint(fptr, &(TNEF.body));
      fprintf(fptr, "\n");
    }

    filename = MAPIFindProperty(&(TNEF.MapiProperties), \
                                PROP_TAG(PT_SYSTIME, PR_CREATION_TIME));
    if (filename != MAPI_UNDEFINED) {
      fprintf(fptr, "DTSTAMP:");
      MAPISysTimetoDTR(filename->data, &thedate);
      fprintf(fptr, "%04i%02i%02iT%02i%02i%02iZ\n",
              thedate.wYear, thedate.wMonth, thedate.wDay,
              thedate.wHour, thedate.wMinute, thedate.wSecond);
    }

    filename = MAPIFindUserProp(&(TNEF.MapiProperties), \
                                PROP_TAG(PT_SYSTIME, 0x8517));
    if (filename != MAPI_UNDEFINED) {
      fprintf(fptr, "DUE:");
      MAPISysTimetoDTR(filename->data, &thedate);
      fprintf(fptr, "%04i%02i%02iT%02i%02i%02iZ\n",
              thedate.wYear, thedate.wMonth, thedate.wDay,
              thedate.wHour, thedate.wMinute, thedate.wSecond);
    }
    filename = MAPIFindProperty(&(TNEF.MapiProperties), \
                                PROP_TAG(PT_SYSTIME, PR_LAST_MODIFICATION_TIME));
    if (filename != MAPI_UNDEFINED) {
      fprintf(fptr, "LAST-MODIFIED:");
      MAPISysTimetoDTR(filename->data, &thedate);
      fprintf(fptr, "%04i%02i%02iT%02i%02i%02iZ\n",
              thedate.wYear, thedate.wMonth, thedate.wDay,
              thedate.wHour, thedate.wMinute, thedate.wSecond);
    }
    // Class
    filename = MAPIFindUserProp(&(TNEF.MapiProperties), \
                                PROP_TAG(PT_BOOLEAN, 0x8506));
    if (filename != MAPI_UNDEFINED) {
      ddword_ptr = (DDWORD *)filename->data;
      fprintf(fptr, "CLASS:");
      if (*ddword_ptr == 1) {
        fprintf(fptr, "PRIVATE\n");
      } else {
        fprintf(fptr, "PUBLIC\n");
      }
    }
    fprintf(fptr, "END:VTODO\n");
    fprintf(fptr, "END:VCALENDAR\n");
    fclose(fptr);
  }

}

