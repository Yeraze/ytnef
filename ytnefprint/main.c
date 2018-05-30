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
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ytnef.h>
#include "config.h"

#define PRODID "PRODID:-//The Gauntlet//" PACKAGE_STRING "//EN\n"

TNEFStruct TNEF;
int verbose = 0;
void PrintTNEF(TNEFStruct TNEF);
void ProcessTNEF(TNEFStruct TNEF);
void SaveVCalendar(TNEFStruct TNEF);
void SaveVCard(TNEFStruct TNEF);
void SaveVTask(TNEFStruct TNEF);


void PrintHelp(void) {
  printf("Yerase TNEF Printer v");
  printf(VERSION);
  printf("\n");
  printf("\n");
  printf("  usage: ytnefprint <options> <filenames>\n");
  printf("\n");
  printf("Options:\n");
  printf("   -h   - Displays this help message\n");
  printf("   -v   - Verbose output (multiple -v's increase \n");
  printf("                   the level of output.\n");
  printf("\n");
  printf("Send bug reports to ");
  printf(PACKAGE_BUGREPORT);
  printf("\n");

}


int main(int argc, char **argv) {
  int index, i;

//    printf("Size of WORD is %i\n", sizeof(WORD));
//    printf("Size of DWORD is %i\n", sizeof(DWORD));
//    printf("Size of DDWORD is %i\n", sizeof(DDWORD));

  if (argc == 1) {
    printf("You must specify files to parse\n");
    PrintHelp();
    return -1;
  }

  for (i = 1; i < argc; i++) {
    if (argv[i][0] == '-') {
      switch (argv[i][1]) {
        case 'v': verbose++;
          break;
      }
      continue;
    }
    TNEFInitialize(&TNEF);
    TNEF.Debug = verbose;
    if (TNEFParseFile(argv[i], &TNEF) < 0) {
      printf("ERROR processing file\n");
      continue;
    }
    PrintTNEF(TNEF);
    TNEFFree(&TNEF);
  }
}

void PrintTNEF(TNEFStruct TNEF) {
  int index, i;
  int j, object;
  int count;
  FILE *fptr;
  char ifilename[1024];
  char *charptr, *charptr2;
  DDWORD ddword_tmp;
  int SaveFile;
  DDWORD *ddword_ptr;
  MAPIProps mapip;
  variableLength *filename;
  variableLength *filedata;
  Attachment *p;
  TNEFStruct emb_tnef;

  printf("---> In %s format\n", TNEF.version);
  if (TNEF.from.size > 0)
    printf("From: %s\n", TNEF.from.data);
  if (TNEF.subject.size > 0)
    printf("Subject: %s\n", TNEF.subject.data);
  if (TNEF.priority[0] != 0)
    printf("Message Priority: %s\n", TNEF.priority);
  if (TNEF.dateSent.wYear > 0) {
    printf("Date Sent: ");
    TNEFPrintDate(TNEF.dateSent);
    printf("\n");
  }
  if (TNEF.dateReceived.wYear > 0) {
    printf("Date Received: ");
    TNEFPrintDate(TNEF.dateReceived);
    printf("\n");
  }
  if (TNEF.messageStatus[0] != 0)
    printf("Message Status: %s\n", TNEF.messageStatus);
  if (TNEF.messageClass[0] != 0)  {
    printf("Message Class: %s\n", TNEF.messageClass);
    if (strcmp(TNEF.messageClass, "IPM.Contact") == 0) {
      printf("--> Found a contact card\n");
    }
    if (strcmp(TNEF.messageClass, "IPM.Task") == 0) {
      printf("--> Found a Task Entry\n");
    }
  }

  if (TNEF.OriginalMessageClass.size > 0)
    printf("Original Message Class: %s\n",
           TNEF.OriginalMessageClass.data);
  if (TNEF.messageID[0] != 0)
    printf("Message ID: %s\n", TNEF.messageID);
  if (TNEF.parentID[0] != 0)
    printf("Parent ID: %s\n", TNEF.parentID);
  if (TNEF.conversationID[0] != 0)
    printf("Conversation ID: %s\n", TNEF.conversationID);
  if (TNEF.DateStart.wYear > 0) {
    printf("Start Date: ");
    TNEFPrintDate(TNEF.DateStart);
    printf("\n");
  }
  if (TNEF.DateEnd.wYear > 0) {
    printf("End Date: ");
    TNEFPrintDate(TNEF.DateEnd);
    printf("\n");
  }
  if (TNEF.Owner.size > 0)
    printf("Owner: %s\n", TNEF.Owner.data);

  if (TNEF.Delegate.size > 0)
    printf("Delegate: %s\n", TNEF.Delegate.data);

  if (TNEF.AidOwner.size > 0)
    printf("Aid Owner: %s\n", TNEF.AidOwner.data);


  if (TNEF.body.size > 0)
    printf("-- Message Body (%i bytes) --\n%s\n-- End Body --\n",
           TNEF.body.size, TNEF.body.data);

  if (TNEF.MapiProperties.count > 0) {
    printf("    MAPI Properties: %u\n", TNEF.MapiProperties.count);
    MAPIPrint(&TNEF.MapiProperties);
  }

  if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                   PROP_TAG(PT_STRING8, 0x24))) != MAPI_UNDEFINED) {
    if (strcmp((char*)filename->data, "IPM.Appointment") == 0) {
      printf("--> Found an appointment entry\n");
    }
  }

  // Now Print file data
  p = TNEF.starting_attach.next;
  count = 0;
  while (p != NULL) {
    count++;
    printf("[%i] [", count);
    switch (p->RenderData.atyp) {
      case 0: printf("NULL      "); break;
      case 1: printf("File      "); break;
      case 2: printf("OLE Object"); break;
      case 3: printf("Picture   "); break;
      case 4: printf("Max       "); break;
      default: printf("Unknown   ");
    }
    printf("] ");
    if (p->Title.size > 0)
      printf("%s", p->Title.data);
    printf("\n");
    if (p->RenderData.dwFlags == 0x00000001)
      printf("     MAC Binary Encoding\n");

    if (p->TransportFilename.size > 0)
      printf("     Transported under the name %s\n",
             p->TransportFilename.data);

    if (p->Date.wYear > 0) {
      printf("    Date: ");
      TNEFPrintDate(p->Date);
      printf("\n");
    }
    if (p->CreateDate.wYear > 0) {
      printf("    Creation Date: ");
      TNEFPrintDate(p->CreateDate);
      printf("\n");
    }
    if (p->ModifyDate.wYear > 0) {
      printf("    Modified on: ");
      TNEFPrintDate(p->ModifyDate);
      printf("\n");
    }

    if (p->MAPI.count > 0) {
      printf("    MAPI Properties: %u\n", p->MAPI.count);
      if (verbose == 1) {
        MAPIPrint(&p->MAPI);
      }
    }

    if (p->FileData.size > 0) {
      printf("    Attachment Size:  %ib\n", p->FileData.size);

      if ((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(30,
                                       0x3707))) == MAPI_UNDEFINED) {
        if ((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(30,
                                         0x3001))) == MAPI_UNDEFINED) {
          filename = &(p->Title);
        }
      }


      object = 1;
      if ((filedata = MAPIFindProperty(&(p->MAPI),
                                       PROP_TAG(PT_OBJECT, PR_ATTACH_DATA_OBJ)))
          == MAPI_UNDEFINED) {
        if ((filedata = MAPIFindProperty(&(p->MAPI),
                                         PROP_TAG(PT_BINARY, PR_ATTACH_DATA_OBJ)))
            == MAPI_UNDEFINED) {
          filedata = &(p->FileData);
          object = 0;
        }
      }
      snprintf(ifilename, 1024, "%s", filename->data);
      for (i = 0; i < strlen(ifilename); i++)
        if (ifilename[i] == ' ')
          ifilename[i] = '_';
      printf("    File saves as [%s]\n", ifilename);
    }
    p = p->next;
  }
}

