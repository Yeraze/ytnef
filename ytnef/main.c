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
#include "settings.h"

#define PRODID "PRODID:-//The Gauntlet//" PACKAGE_STRING "//EN\n"

int verbose = 0;
int savefiles = 0;
int saveRTF = 0;
int saveHTML = 0;
int saveintermediate = 0;
int attachmentSize = 50;
char *filepath = NULL;
int ProcessTNEF(TNEFStruct TNEF);
void SaveVCalendar(TNEFStruct TNEF, int isMtgReq);
void SaveVCard(TNEFStruct TNEF);
void SaveVTask(TNEFStruct TNEF);


void PrintHelp(void) {
  printf("Yerase TNEF Exporter v");
  printf(VERSION);
  printf("\n");
  printf("\n");
  printf("  usage: ytnef [-+vhf] <filenames>\n");
  printf("\n");
  printf("   -/+v - Enables/Disables verbose output\n");
  printf("          Multiple -v's increase the level of output\n");
  printf("   -/+f - Enables/Disables saving of attachments\n");
  printf("   -/+F - Enables/Disables saving of the message body as RTF\n");
  printf("   -/+H - Enables/Disables saving of the message body as HTML\n");
  printf("   -/+a - Enables/Disables saving of intermediate files\n");
  printf("   -s <n> - Override the default Attachment Size (%i MB)\n", attachmentSize);
  printf("   -h   - Displays this help message\n");
  printf("\n");
  printf("Example:\n");
  printf("  ytnef -v winmail.dat\n");
  printf("     Parse with verbose output, don't save\n");
  printf("  ytnef -f . winmail.dat\n");
  printf("     Parse and save all attachments to local directory (.)\n");
  printf("  ytnef -F -f . winmail.dat\n");
  printf("     Parse and save all attachments to local directory (.)\n");
  printf("     Including saving the message text to a RTF file.\n\n");
  printf("Send bug reports to ");
  printf(PACKAGE_BUGREPORT);
  printf("\n");
}


int main(int argc, char **argv) {
  int index, i, errors=0;
  TNEFStruct TNEF;

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
        case 'a': saveintermediate = 1;
          break;
        case 'v': verbose++;
          break;
        case 'h': PrintHelp();
          return 0;
        case 'f': savefiles = 1;
          filepath = argv[i + 1];
          i++;
          break;
        case 's': 
          attachmentSize = atoi(argv[i+1]);
          i++;
          break;
        case 'F': saveRTF = 1;
          break;
        case 'H': saveHTML = 1;
          break;
        default:
          printf("Unknown option '%s'\n", argv[i]);
	  PrintHelp();	      
      }
      continue;

    }
    if (argv[i][0] == '+') {
      switch (argv[i][1]) {
        case 'a': saveintermediate = 0;
          break;
        case 'v': verbose--;
          break;
        case 'f': savefiles = 0;
          filepath = NULL;
          break;
        case 'F': saveRTF = 0;
          break;
        case 'H': saveHTML = 0;
          break;
        default:
          printf("Unknown option '%s'\n", argv[i]);
	  PrintHelp();	      
      }
      continue;

    }

    TNEFInitialize(&TNEF);
    TNEF.Debug = verbose;
    TNEF.attachmentSize = attachmentSize;
    if (TNEFParseFile(argv[i], &TNEF) == -1) {
      fprintf(stderr, "ERROR processing file\n");
      ++errors;
      continue;
    }
    errors += ProcessTNEF(TNEF);
    TNEFFree(&TNEF);
  }
  return -errors;
}

/* Return number of failures */
int ProcessTNEF(TNEFStruct TNEF) {
  char *astring;
  variableLength *filename;
  variableLength *filedata;
  Attachment *p;
  int RealAttachment;
  int object;
  char ifilename[MAX_FILENAME_SIZE+1];
  int i, count, failures = 0;
  int foundCal = 0;

  FILE *fptr;

// First see if this requires special processing.
// ie: it's a Contact Card, Task, or Meeting request (vCal/vCard)
  if (TNEF.messageClass[0] != 0)  {
    if (strcmp((char *)TNEF.messageClass, "IPM.Contact") == 0) {
      SaveVCard(TNEF);
    }
    if (strcmp((char *)TNEF.messageClass, "IPM.Task") == 0) {
      SaveVTask(TNEF);
    }
    if (strcmp(TNEF.messageClass, "IPM.Microsoft Schedule.MtgReq") == 0) {
      SaveVCalendar(TNEF, 1);
      foundCal = 1;
    }
    if (strcmp((char *)TNEF.messageClass, "IPM.Appointment") == 0) {
      SaveVCalendar(TNEF, 0);
      foundCal = 1;
    }
  }
  if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties),
                                   PROP_TAG(PT_STRING8, 0x24))) != MAPI_UNDEFINED) {
    if (strcmp((char*)filename->data, "IPM.Appointment") == 0) {
      // If it's "indicated" twice, we don't want to save 2 calendar
      // entries.
      if (foundCal == 0) {
        SaveVCalendar(TNEF, 0);
      }
    }
  }

  if (strcmp((char *)TNEF.messageClass, "IPM.Microsoft Mail.Note") == 0) {
    if ((saveRTF == 1) && (TNEF.subject.size > 0)) {
      // Description
      if ((filename = MAPIFindProperty(&(TNEF.MapiProperties),
                                       PROP_TAG(PT_BINARY, PR_RTF_COMPRESSED)))
          != MAPI_UNDEFINED) {
        int size;
        variableLength buf;
        if ((buf.data = DecompressRTF(filename, &(buf.size))) != NULL) {
          snprintf(ifilename, MAX_FILENAME_SIZE, "%s.rtf", TNEF.subject.data);
          SanitizeFilename(ifilename);
          if (filepath != NULL) {
            char tmp[MAX_FILENAME_SIZE+1];
            snprintf(tmp, MAX_FILENAME_SIZE, "%s/%s", filepath, ifilename);
            strcpy(ifilename, tmp);
          }

          if ((fptr = fopen(ifilename, "wb")) == NULL) {
            fprintf(stderr, "ERROR: Error writing %s to disk!\n", ifilename);
            ++failures;
          } else {
            fwrite(buf.data,
                   sizeof(BYTE),
                   buf.size,
                   fptr);
            fclose(fptr);
            printf("%s\n", ifilename);
          }
          free(buf.data);
        }
      }
    }

    if ((saveHTML == 1) && (TNEF.subject.size > 0)) {
      // filename seems to be the actual data rather than the filename
      if ((filename = MAPIFindProperty(&(TNEF.MapiProperties),
              PROP_TAG(PT_BINARY, PR_BODY_HTML)))
          != MAPI_UNDEFINED) {
        int size;
        variableLength buf;
        snprintf(ifilename, MAX_FILENAME_SIZE, "%s.html", TNEF.subject.data);
        SanitizeFilename(ifilename);
        if (filepath != NULL) {
          char tmp[MAX_FILENAME_SIZE+1];
          snprintf(tmp, MAX_FILENAME_SIZE, "%s/%s", filepath, ifilename);
          strcpy(ifilename, tmp);
        }

        if ((fptr = fopen(ifilename, "wb")) == NULL) {
          fprintf(stderr, "ERROR: Error writing %s to disk!\n", ifilename);
          ++failures;
        } else {
          fwrite(filename->data,
              sizeof(BYTE),
              filename->size,
              fptr);
          fclose(fptr);
          printf("%s\n", ifilename);
        }
      }
    }
  }

  if (strcmp(TNEF.messageClass, "IPM.StickyNote") == 0) {
    if ((saveRTF == 1) && (TNEF.subject.size > 0)) {
      // Description
      if ((filedata=MAPIFindProperty(&(TNEF.MapiProperties),
                                     PROP_TAG(PT_BINARY, PR_RTF_COMPRESSED)))
          != MAPI_UNDEFINED) {
        int size;
	      if ((filename=MAPIFindProperty(&(TNEF.MapiProperties),
                                    PROP_TAG(PT_STRING8, PR_CONVERSATION_TOPIC))) == MAPI_UNDEFINED) {
          if ((filename=MAPIFindProperty(&(TNEF.MapiProperties),
                                      PROP_TAG(PT_UNICODE, PR_CONVERSATION_TOPIC))) == MAPI_UNDEFINED) {
            filename = NULL;
          }
        }

        variableLength buf;
			  buf.data = DecompressRTF(filedata, &(buf.size));
        if (buf.data != NULL && filename != NULL) {
				  char fileNameBase[32];
				  strncpy(fileNameBase, (char*)filename->data, sizeof(fileNameBase) - 1);
				  
          SanitizeFilename(fileNameBase);

          CreateUniqueFilename(ifilename, MAX_FILENAME_SIZE, fileNameBase, "rtf", filepath);

          if ((fptr = fopen(ifilename, "wb"))==NULL) {
            fprintf(stderr, "ERROR: Error writing %s to disk!\n", ifilename);
            ++failures;
          } else {
            fwrite(buf.data,
                   sizeof(BYTE),
                   buf.size,
                   fptr);
            fclose(fptr);
            printf("%s\n", ifilename);
          }
          free(buf.data);
        }
      }
    }
  }

// Now process each attachment
  p = TNEF.starting_attach.next;
  count = 0;
  while (p != NULL) {
    count++;
    // Make sure it has a size.
    if (p->FileData.size > 0) {
      object = 1;


      // See if the contents are stored as "attached data"
      //  Inside the MAPI blocks.
      if ((filedata = MAPIFindProperty(&(p->MAPI),
                                       PROP_TAG(PT_OBJECT, PR_ATTACH_DATA_OBJ)))
          == MAPI_UNDEFINED) {
        if ((filedata = MAPIFindProperty(&(p->MAPI),
                                         PROP_TAG(PT_BINARY, PR_ATTACH_DATA_OBJ)))
            == MAPI_UNDEFINED) {
          // Nope, standard TNEF stuff.
          filedata = &(p->FileData);
          object = 0;
        }
      }
      // See if this is an embedded TNEF stream.
      RealAttachment = 1;
      if (object == 1) {
        // This is an "embedded object", so skip the
        // 16-byte identifier first.
        TNEFStruct emb_tnef;
        DWORD signature;
        memcpy(&signature, filedata->data + 16, sizeof(DWORD));
        if (TNEFCheckForSignature(signature) == 0) {
          // Has a TNEF signature, so process it.
          TNEFInitialize(&emb_tnef);
          emb_tnef.attachmentSize = attachmentSize;
          emb_tnef.Debug = TNEF.Debug;
          if (TNEFParseMemory(filedata->data + 16,
                              filedata->size - 16, &emb_tnef) != -1) {
            failures += ProcessTNEF(emb_tnef);
            RealAttachment = 0;
          }
          TNEFFree(&emb_tnef);
        }
      } else {
        TNEFStruct emb_tnef;
        DWORD signature;
        memcpy(&signature, filedata->data, sizeof(DWORD));
        if (TNEFCheckForSignature(signature) == 0) {
          // Has a TNEF signature, so process it.
          TNEFInitialize(&emb_tnef);
          emb_tnef.attachmentSize = attachmentSize;
          emb_tnef.Debug = TNEF.Debug;
          if (TNEFParseMemory(filedata->data,
                              filedata->size, &emb_tnef) != -1) {
            failures += ProcessTNEF(emb_tnef);
            RealAttachment = 0;
          }
          TNEFFree(&emb_tnef);
        }
      }
      if ((RealAttachment == 1) || (saveintermediate == 1)) {
        // Ok, it's not an embedded stream, so now we
        // process it.
        if ((filename = MAPIFindProperty(&(p->MAPI),
                                         PROP_TAG(30, 0x3707)))
            == MAPI_UNDEFINED) {
          if ((filename = MAPIFindProperty(&(p->MAPI),
                                           PROP_TAG(30, 0x3001)))
              == MAPI_UNDEFINED) {
            if ((filename = MAPIFindProperty(&(p->MAPI),
                                             PROP_TAG(30, 0x370C)))
                == MAPI_UNDEFINED) {
              filename = &(p->Title);
            }
          }
        }
        if (filename->size == 1) {
          filename = (variableLength *)malloc(sizeof(variableLength));
          filename->size = 20;
          filename->data = (BYTE *)malloc(20);
          snprintf((char*)filename->data, 19, "file_%03i.dat", count);
        }
        snprintf(ifilename, MAX_FILENAME_SIZE, "%.*s", filename->size, filename->data);
        for(i = 0; i < strlen(ifilename); i++) 
          if (ifilename[i] == ' ') 
            ifilename[i] = '_';
        if(filepath) {
          char tmp[MAX_FILENAME_SIZE];
          memcpy(tmp, ifilename, MAX_FILENAME_SIZE);
          snprintf(ifilename, MAX_FILENAME_SIZE, "%s/%s", filepath, tmp);
        }
        if (savefiles == 1) {
          if ((fptr = fopen(ifilename, "wb")) == NULL) {
            fprintf(stderr, "ERROR: Error writing %s to disk!\n", ifilename);
            ++failures;
          } else {
            if (object == 1) {
              fwrite(filedata->data + 16,
                     sizeof(BYTE),
                     filedata->size - 16,
                     fptr);
            } else {
              fwrite(filedata->data,
                     sizeof(BYTE),
                     filedata->size,
                     fptr);
            }
            fclose(fptr);
            printf("%s\n", ifilename);
          } // if we opened successfully
        } // if savefiles == 1
      } // if RealAttachment == 1
    } // if size>0
    p = p->next;
  } // while p!= null

  return failures;
}


#include "utility.c"
#include "vcal.c"
#include "vcard.c"
#include "vtask.c"
#include "settings.c"
