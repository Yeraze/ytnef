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
unsigned char GetRruleCount(unsigned char a, unsigned char b) {
  return ((a << 8) | b);
}

unsigned char GetRruleMonthNum(unsigned char a, unsigned char b) {
  switch (a) {
    case 0x00:
      switch (b) {
        case 0x00:
          // Jan
          return (1);
        case 0xA3:
          // May
          return (5);
        case 0xAE:
          // Nov
          return (11);
      }
      break;
    case 0x60:
      switch (b) {
        case 0xAE:
          // Feb
          return (2);
        case 0x51:
          // Jun
          return (6);
      }
      break;
    case 0xE0:
      switch (b) {
        case 0x4B:
          // Mar
          return (3);
        case 0x56:
          // Sep
          return (9);
      }
      break;
    case 0x40:
      switch (b) {
        case 0xFA:
          // Apr
          return (4);
      }
      break;
    case 0x20:
      if (b == 0xFA) {
        // Jul
        return (7);
      }
      break;
    case 0x80:
      if (b == 0xA8) {
        // Aug
        return (8);
      }
      break;
    case 0xA0:
      if (b == 0xFF) {
        // Oct
        return (10);
      }
      break;
    case 0xC0:
      if (b == 0x56) {
        return (12);
      }
  }

  // Error
  return (0);
}

char *GetRruleDayname(unsigned char a) {
  static char daystring[25];

  *daystring = 0;

  if (a & 0x01) {
    strcat(daystring, "SU,");
  }
  if (a & 0x02) {
    strcat(daystring, "MO,");
  }
  if (a & 0x04) {
    strcat(daystring, "TU,");
  }
  if (a & 0x08) {
    strcat(daystring, "WE,");
  }
  if (a & 0x10) {
    strcat(daystring, "TH,");
  }
  if (a & 0x20) {
    strcat(daystring, "FR,");
  }
  if (a & 0x40) {
    strcat(daystring, "SA,");
  }

  if (strlen(daystring)) {
    daystring[strlen(daystring) - 1] = 0;
  }

  return (daystring);
}

void PrintRrule(FILE *fptr, char *recur_data, int size, TNEFStruct TNEF) {
  variableLength *filename;

  if (size < 0x1F) {
    return;
  }

  fprintf(fptr, "RRULE:FREQ=");

  if (recur_data[0x04] == 0x0A) {
    fprintf(fptr, "DAILY");

    if (recur_data[0x16] == 0x23 || recur_data[0x16] == 0x22 ||
        recur_data[0x16] == 0x21) {
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_I2, 0x0011))) != MAPI_UNDEFINED) {
        fprintf(fptr, ";INTERVAL=%d", *(filename->data));
      }
      if (recur_data[0x16] == 0x22 || recur_data[0x16] == 0x21) {
        fprintf(fptr, ";COUNT=%d",
                GetRruleCount(recur_data[0x1B], recur_data[0x1A]));
      }
    } else if (recur_data[0x16] == 0x3E) {
      fprintf(fptr, ";BYDAY=MO,TU,WE,TH,FR");
      if (recur_data[0x1A] == 0x22 || recur_data[0x1A] == 0x21) {
        fprintf(fptr, ";COUNT=%d",
                GetRruleCount(recur_data[0x1F], recur_data[0x1E]));
      }
    }
  } else if (recur_data[0x04] == 0x0B) {
    fprintf(fptr, "WEEKLY;INTERVAL=%d;BYDAY=%s",
            recur_data[0x0E], GetRruleDayname(recur_data[0x16]));
    if (recur_data[0x1A] == 0x22 || recur_data[0x1A] == 0x21) {
      fprintf(fptr, ";COUNT=%d",
              GetRruleCount(recur_data[0x1F], recur_data[0x1E]));
    }
  } else if (recur_data[0x04] == 0x0C) {
    fprintf(fptr, "MONTHLY");
    if (recur_data[0x06] == 0x02) {
      fprintf(fptr, ";INTERVAL=%d;BYMONTHDAY=%d", recur_data[0x0E],
              recur_data[0x16]);
      if (recur_data[0x1A] == 0x22 || recur_data[0x1A] == 0x21) {
        fprintf(fptr, ";COUNT=%d", GetRruleCount(recur_data[0x1F],
                recur_data[0x1E]));
      }
    } else if (recur_data[0x06] == 0x03) {
      fprintf(fptr, ";BYDAY=%s;BYSETPOS=%d;INTERVAL=%d",
              GetRruleDayname(recur_data[0x16]),
              recur_data[0x1A] == 0x05 ? -1 : recur_data[0x1A],
              recur_data[0x0E]);
      if (recur_data[0x1E] == 0x22 || recur_data[0x1E] == 0x21) {
        fprintf(fptr, ";COUNT=%d", GetRruleCount(recur_data[0x23],
                recur_data[0x22]));
      }
    }
  } else if (recur_data[0x04] == 0x0D) {
    fprintf(fptr, "YEARLY;BYMONTH=%d",
            GetRruleMonthNum(recur_data[0x0A], recur_data[0x0B]));
    if (recur_data[0x06] == 0x02) {
      fprintf(fptr, ";BYMONTHDAY=%d", recur_data[0x16]);
    } else if (recur_data[0x06] == 0x03) {
      fprintf(fptr, ";BYDAY=%s;BYSETPOS=%d",
              GetRruleDayname(recur_data[0x16]),
              recur_data[0x1A] == 0x05 ? -1 : recur_data[0x1A]);
    }
    if (recur_data[0x1E] == 0x22 || recur_data[0x1E] == 0x21) {
      fprintf(fptr, ";COUNT=%d", GetRruleCount(recur_data[0x23],
              recur_data[0x22]));
    }
  }
  fprintf(fptr, "\n");
}

void SaveVCalendar(TNEFStruct TNEF, int isMtgReq) {
  char ifilename[MAX_FILENAME_SIZE];
  variableLength *filename;
  char *charptr, *charptr2;
  FILE *fptr;
  int index;
  DDWORD *ddword_ptr;
  dtr thedate;

  if(isMtgReq) {
    CreateUniqueFilename(ifilename, MAX_FILENAME_SIZE, "MtgReq", "ics", filepath);
  } else {
    CreateUniqueFilename(ifilename, MAX_FILENAME_SIZE, "calendar", "ics", filepath);
  }

  printf("%s\n", ifilename);
  if (savefiles == 0)
    return;

  if ((fptr = fopen(ifilename, "wb")) == NULL) {
    printf("Error writing file to disk!");
  } else {
    fprintf(fptr, "BEGIN:VCALENDAR\n");
    if (TNEF.messageClass[0] != 0) {
      charptr2 = TNEF.messageClass;
      charptr = charptr2;
      while (*charptr != 0) {
        if (*charptr == '.') {
          charptr2 = charptr;
        }
        charptr++;
      }
      if (strcmp((char*)charptr2, ".MtgCncl") == 0) {
        fprintf(fptr, "METHOD:CANCEL\n");
      } else {
        fprintf(fptr, "METHOD:REQUEST\n");
      }
    } else {
      fprintf(fptr, "METHOD:REQUEST\n");
    }
    fprintf(fptr, PRODID);
    fprintf(fptr, "VERSION:2.0\n");
    fprintf(fptr, "BEGIN:VEVENT\n");

    // UID
    // After alot of comparisons, I'm reasonably sure this is totally
    // wrong.  But it's not really necessary.
    //
    // I think it only exists to connect future modification entries to
    // this entry. so as long as it's incorrectly interpreted the same way
    // every time, it should be ok :)
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_BINARY, 0x3))) == MAPI_UNDEFINED) {
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_BINARY, 0x23))) == MAPI_UNDEFINED) {
        filename = NULL;
      }
    }
    if (filename != NULL) {
      fprintf(fptr, "UID:");
      for (index = 0; index < filename->size; index++) {
        fprintf(fptr, "%02X", (unsigned char)filename->data[index]);
      }
      fprintf(fptr, "\n");
    }

    // Sequence
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_LONG, 0x8201))) != MAPI_UNDEFINED) {
      ddword_ptr = (DDWORD *)filename->data;
      fprintf(fptr, "SEQUENCE:%i\n", (int) *ddword_ptr);
    }

    filename = NULL;
    if ((filename = MAPIFindProperty(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_BINARY, PR_SENDER_SEARCH_KEY)))
        == MAPI_UNDEFINED) {
      if ((filename = MAPIFindProperty(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_UNICODE, PR_SENT_REPRESENTING_EMAIL_ADDRESS)))
          == MAPI_UNDEFINED) {
        filename = NULL;
      }
    }
    if (filename != NULL) {
      charptr = (char*)filename->data;
      charptr2 = (char*)strstr((char*)charptr, ":");
      if (charptr2 == NULL)
        charptr2 = charptr;
      else
        charptr2++;
      fprintf(fptr, "ORGANIZER;CN=\"%s\":MAILTO:%s\n",
              charptr2, charptr2);
    }

    // Required Attendees
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_STRING8, 0x823b))) == MAPI_UNDEFINED) {
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_UNICODE, 0x823b))) == MAPI_UNDEFINED) {
        filename = NULL;
      }
    }
    if (filename != NULL) {
      // We have a list of required participants, so
      // write them out.
      if (strlen((char*)filename->data) > 1) {
        charptr = (char*)filename->data - 1;
        charptr2 = (char*)strstr((char*)(charptr + 1), ";");
        while (charptr != NULL) {
          charptr++;
          charptr2 = (char *)strstr((char*)charptr, ";");
          if (charptr2 != NULL) {
            *charptr2 = 0;
          }
          while (*charptr == ' ')
            charptr++;
          fprintf(fptr, "ATTENDEE;PARTSTAT=NEEDS-ACTION;");
          fprintf(fptr, "ROLE=REQ-PARTICIPANT;RSVP=TRUE;");
          fprintf(fptr, "CN=\"%s\":MAILTO:%s\n",
                  charptr, charptr);
          charptr = charptr2;
        }
      }
      // Optional attendees
      filename = NULL;
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_STRING8, 0x823c))) == MAPI_UNDEFINED) {
        if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                         PROP_TAG(PT_UNICODE, 0x823c))) == MAPI_UNDEFINED) {
          filename = NULL;
        }
      }
      if (filename != NULL) {
        // The list of optional participants
        if (strlen((char*)filename->data) > 1) {
          charptr = (char*)filename->data - 1;
          charptr2 = (char *)strstr((char*)(charptr + 1), ";");
          while (charptr != NULL) {
            charptr++;
            charptr2 = (char *)strstr((char*)charptr, ";");
            if (charptr2 != NULL) {
              *charptr2 = 0;
            }
            while (*charptr == ' ')
              charptr++;
            fprintf(fptr, "ATTENDEE;PARTSTAT=NEEDS-ACTION;");
            fprintf(fptr, "ROLE=OPT-PARTICIPANT;RSVP=TRUE;");
            fprintf(fptr, "CN=\"%s\":MAILTO:%s\n",
                    charptr, charptr);
            charptr = charptr2;
          }
        }
      }
    } else {
      filename = NULL;
      if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_STRING8, 0x8238))) == MAPI_UNDEFINED) {
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_UNICODE, 0x8238))) == MAPI_UNDEFINED) {
          filename = NULL;
        }
      }
      if (filename != NULL) {
        if (strlen((char*)filename->data) > 1) {
          charptr = (char*)filename->data - 1;
          charptr2 = (char *)strstr((char*)(charptr + 1), ";");
          while (charptr != NULL) {
            charptr++;
            charptr2 = (char *)strstr((char*)charptr, ";");
            if (charptr2 != NULL) {
              *charptr2 = 0;
            }
            while (*charptr == ' ')
              charptr++;
            fprintf(fptr, "ATTENDEE;PARTSTAT=NEEDS-ACTION;");
            fprintf(fptr, "ROLE=REQ-PARTICIPANT;RSVP=TRUE;");
            fprintf(fptr, "CN=\"%s\":MAILTO:%s\n", charptr, charptr);
            charptr = charptr2;
          }
        }
      }
    }
    // Summary
    filename = NULL;
    if ((filename = MAPIFindProperty(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_STRING8, PR_CONVERSATION_TOPIC))) == MAPI_UNDEFINED) {
      if ((filename = MAPIFindProperty(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_UNICODE, PR_CONVERSATION_TOPIC))) == MAPI_UNDEFINED) {
        filename = NULL;
      }
    }
    if (filename != NULL) {
      fprintf(fptr, "SUMMARY:");
      Cstylefprint(fptr, filename);
      fprintf(fptr, "\n");
    }

    // Description
    if ((filename = MAPIFindProperty(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_BINARY, PR_RTF_COMPRESSED)))
        != MAPI_UNDEFINED) {
      variableLength buf;
      if ((buf.data = DecompressRTF(filename, &(buf.size))) != NULL) {
        fprintf(fptr, "DESCRIPTION:");
        PrintRTF(fptr, &buf);
        free(buf.data);
      }

    }

    // Location
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_STRING8, 0x0002))) == MAPI_UNDEFINED) {
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_STRING8, 0x8208))) == MAPI_UNDEFINED) {
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), 
                                       PROP_TAG(PT_UNICODE, 0x0002))) == MAPI_UNDEFINED) {
          if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), 
                                         PROP_TAG(PT_UNICODE, 0x8208))) == MAPI_UNDEFINED) {
            filename = NULL;
          }
        }
      }
    }
    if (filename != NULL) {
      fprintf(fptr, "LOCATION: %s\n", filename->data);
    }
    // Date Start
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_SYSTIME, 0x820d))) == MAPI_UNDEFINED) {
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_SYSTIME, 0x8516))) == MAPI_UNDEFINED) {
        filename = NULL;
      }
    }
    if (filename != NULL) {
      fprintf(fptr, "DTSTART:");
      MAPISysTimetoDTR(filename->data, &thedate);
      fprintf(fptr, "%04i%02i%02iT%02i%02i%02iZ\n",
              thedate.wYear, thedate.wMonth, thedate.wDay,
              thedate.wHour, thedate.wMinute, thedate.wSecond);
    }
    // Date End
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_SYSTIME, 0x820e))) == MAPI_UNDEFINED) {
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_SYSTIME, 0x8517))) == MAPI_UNDEFINED) {
        filename = NULL;
      }
    }
    if (filename != NULL) {
      fprintf(fptr, "DTEND:");
      MAPISysTimetoDTR(filename->data, &thedate);
      fprintf(fptr, "%04i%02i%02iT%02i%02i%02iZ\n",
              thedate.wYear, thedate.wMonth, thedate.wDay,
              thedate.wHour, thedate.wMinute, thedate.wSecond);
    }
    // Date Stamp
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_SYSTIME, 0x8202))) == MAPI_UNDEFINED) {
      if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_SYSTIME, 0x001a))) == MAPI_UNDEFINED) {
        filename = NULL;
      }
    }

    if (filename != NULL){
      fprintf(fptr, "DTSTAMP:");
      MAPISysTimetoDTR(filename->data, &thedate);
      fprintf(fptr, "%04i%02i%02iT%02i%02i%02iZ\n",
              thedate.wYear, thedate.wMonth, thedate.wDay,
              thedate.wHour, thedate.wMinute, thedate.wSecond);
    }
    // Class
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_BOOLEAN, 0x8506))) != MAPI_UNDEFINED) {
      ddword_ptr = (DDWORD *)filename->data;
      fprintf(fptr, "CLASS:");
      if (*ddword_ptr == 1) {
        fprintf(fptr, "PRIVATE\n");
      } else {
        fprintf(fptr, "PUBLIC\n");
      }
    }
    // Recurrence
    filename = NULL;
    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_BINARY, 0x8216))) != MAPI_UNDEFINED) {
      PrintRrule(fptr, (char*)filename->data, filename->size, TNEF);
    }

    // Wrap it up
    fprintf(fptr, "END:VEVENT\n");
    fprintf(fptr, "END:VCALENDAR\n");
    fclose(fptr);
  }
}

