#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "tnef-types.h"
#include "tnef.h"
#include "mapi.h"
#include "mapidefs.h"
#include "config.h"

#define PRODID "PRODID:-//The Gauntlet//" PACKAGE_STRING "//EN\n"

TNEFStruct TNEF;
int verbose = 0;
int savefiles = 0;
int saveRTF = 0;
int listonly = 0;
int filenameonly = 0;
char *filepath = NULL;
void PrintTNEF(TNEFStruct TNEF);
void SaveVCalendar(TNEFStruct TNEF);
void SaveVCard(TNEFStruct TNEF);
void SaveVTask(TNEFStruct TNEF);


void PrintHelp(void) {
    printf("Yerase TNEF Exporter v");
            printf(VERSION);
            printf("\n");
    printf("\n");
    printf("  usage: ytnef [-+vhf] <filenames>\n");
    printf("\n");
    printf("   -l   - Enables List-Only mode (for tnefclean)\n");
    printf("   -L   - Enables List-Only mode (filenames only)\n");
    printf("   -/+v - Enables/Disables verbose printing of MAPI Properties\n");
    printf("   -/+f - Enables/Disables saving of attachments\n");
    printf("   -/+F - Enables/Disables saving of inline message text as\n");
    printf("          compressed RTF (not very *nix-friendly\n");
    printf("           (requires -f option)\n");
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


int main(int argc, char ** argv) {
    int index,i;

//    printf("Size of WORD is %i\n", sizeof(WORD));
//    printf("Size of DWORD is %i\n", sizeof(DWORD));
//    printf("Size of DDWORD is %i\n", sizeof(DDWORD));

    if (argc == 1) {
        printf("You must specify files to parse\n");
        PrintHelp();
        return -1;
    }
    
    for(i=1; i<argc; i++) {
        if (argv[i][0] == '-') {
            switch (argv[i][1]) {
                case 'v': verbose++;
                          break;
                case 'h': PrintHelp();
                          return;
                case 'f': savefiles = 1;
                          filepath = argv[i+1];
                          i++;
                          break;
                case 'l': listonly = 1;
                          break;
                case 'L': listonly = 1;
                          filenameonly = 1;
                          break;
                case 'F': saveRTF = 1;
                          break;
                default: 
                          printf("Unknown option '%s'\n", argv[i]);
            }
            continue;

        }
        if (argv[i][0] == '+') {
            switch (argv[i][1]) {
                case 'v': verbose--;
                          break;
                case 'f': savefiles = 0;
                          filepath = NULL;
                          break;
                case 'F': saveRTF = 0;
                          break;
                default: 
                          printf("Unknown option '%s'\n", argv[i]);
            }
            continue;

        }

        TNEFInitialize(&TNEF);
        TNEF.Debug = verbose;
        if (TNEFParseFile(argv[i], &TNEF) == -1) {
            printf(">>> ERROR processing file\n");
            continue;
        }
        if (listonly == 0) 
            printf("---> File %s\n", argv[i]);

        PrintTNEF(TNEF);
        TNEFFree(&TNEF);
    }
}





void PrintTNEF(TNEFStruct TNEF) {
    int index,i;
    int j, object;
    int count;
    FILE *fptr;
    char ifilename[256];
    char *charptr, *charptr2;
    DDWORD ddword_tmp;
    DDWORD *ddword_ptr;
    MAPIProps mapip;
    variableLength *filename;
    Attachment *p;
    TNEFStruct emb_tnef;

    if (listonly == 0) {
        printf("---> In %s format\n", TNEF.version);
        if (TNEF.from.size > 0) 
            printf("From: %s\n", TNEF.from.data);
        if (TNEF.subject.size > 0) 
            printf("Subject: %s\n", TNEF.subject.data);
        if (TNEF.priority[0] != 0) 
            printf("Message Priority: %s\n", TNEF.priority);
        if (TNEF.dateSent.wYear >0) {
            printf("Date Sent: ");
            TNEFPrintDate(TNEF.dateSent);
            printf("\n");
        }
        if (TNEF.dateReceived.wYear >0) {
            printf("Date Received: ");
            TNEFPrintDate(TNEF.dateReceived);
            printf("\n");
        }
        if (TNEF.messageStatus[0] != 0) 
            printf("Message Status: %s\n", TNEF.messageStatus);
    }
    if (TNEF.messageClass[0] != 0)  {
        if (listonly == 0) 
            printf("Message Class: %s\n", TNEF.messageClass);
        if (strcmp(TNEF.messageClass, "IPM.Contact") == 0) {
            if (listonly == 0) 
                printf("Found a contact card\n");
            if (savefiles == 1) 
                SaveVCard(TNEF);
        }
        if (strcmp(TNEF.messageClass, "IPM.Task") == 0) {
            if (listonly == 0) 
                printf("Found a Task Entry\n");
            if (savefiles == 1) 
                SaveVTask(TNEF);
        }
    }

    if (listonly == 0) {
        if (TNEF.OriginalMessageClass.size >0) 
            printf("Original Message Class: %s\n", 
                        TNEF.OriginalMessageClass.data);
        if (TNEF.messageID[0] != 0) 
            printf("Message ID: %s\n", TNEF.messageID);
        if (TNEF.parentID[0] != 0) 
            printf("Parent ID: %s\n", TNEF.parentID);
        if (TNEF.conversationID[0] != 0) 
            printf("Conversation ID: %s\n", TNEF.conversationID);
        if (TNEF.DateStart.wYear >0) {
            printf("Start Date: ");
            TNEFPrintDate(TNEF.DateStart);
            printf("\n");
        }
        if (TNEF.DateEnd.wYear > 0) {
            printf("End Date: ");
            TNEFPrintDate(TNEF.DateEnd);
            printf("\n");
        }
        if (TNEF.Owner.size > 0 ) 
            printf("Owner: %s\n", TNEF.Owner.data);

        if (TNEF.Delegate.size > 0) 
            printf("Delegate: %s\n", TNEF.Delegate.data);

        if (TNEF.AidOwner.size > 0) 
            printf("Aid Owner: %s\n", TNEF.AidOwner.data);


        if (TNEF.body.size>0) 
            printf("-- Message Body (%i bytes) --\n%s\n-- End Body --\n", 
                    TNEF.body.size, TNEF.body.data); 
    }
        
    if (TNEF.MapiProperties.count > 0) {
        if (listonly == 0) 
            printf("    MAPI Properties: %i\n", TNEF.MapiProperties.count);
        if ((filename = MAPIFindProperty(&(TNEF.MapiProperties), 
                   PROP_TAG(PT_BINARY,PR_RTF_COMPRESSED))) == MAPI_UNDEFINED) {

        } else if ((savefiles == 1) && (saveRTF == 1)) {
            if ((listonly == 1) && (filenameonly == 1)) 
                printf("message.rtf\n");

            if (filepath == NULL) {
                sprintf(ifilename, "message.rtf");
            } else {
                sprintf(ifilename, "%s/message.rtf", filepath);
            }
            if ((listonly == 1) && (filenameonly == 0)) 
                printf("%s\n", ifilename);
            if ((fptr = fopen(ifilename, "wb"))==NULL) {
                printf("Error writing file to disk!");
            } else {
                fwrite(filename->data+8, sizeof(BYTE), filename->size+8, fptr);
                fclose(fptr);
            }
        }
        if (verbose == 1) {
            MAPIPrint(&TNEF.MapiProperties);
        }
    }

    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties), 
                        PROP_TAG(PT_STRING8,0x24))) != MAPI_UNDEFINED) {
        if (strcmp(filename->data, "IPM.Appointment") == 0) {
            if (listonly == 0) 
                printf("Found an appointment entry\n");
            if (savefiles == 1) {
                SaveVCalendar(TNEF);
            }
        }
    }

    // Now Print file data
    p = TNEF.starting_attach.next;
    count = 0;
    while (p != NULL) {
        count++;
        if (listonly == 0) {
            printf("[%i] [", count);
            switch (p->RenderData.atyp) {
                case 0: printf("NULL      "); break;
                case 1: printf("File      "); break;
                case 2: printf("OLE Object"); break;
                case 3: printf("Picture   "); break;
                case 4: printf("Max       "); break;
                default:printf("Unknown   "); 
            }
            printf("] ");
            if (p->Title.size > 0) 
                printf("%s", p->Title.data);
            printf("\n");
            if (p->RenderData.dwFlags == 0x00000001) 
                printf("     MAC Binary Encoding\n");

            if (p->TransportFilename.size >0) 
                printf("     Transported under the name %s\n", 
                        p->TransportFilename.data);

            if (p->Date.wYear >0 ) {
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

            if (p->MAPI.count>0) {
                printf("    MAPI Properties: %i\n", p->MAPI.count);
                if (verbose == 1) {
                    MAPIPrint(&p->MAPI);
                }
            }
        }

        if (p->FileData.size > 0) {
            if (listonly == 0) 
                printf("    Attachment Size:  %ib\n", p->FileData.size);
            
            if ((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(30,0x3707))) == MAPI_UNDEFINED) {
                if ((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(30,0x3001))) == MAPI_UNDEFINED) {
                    filename = &(p->Title);
                }
            }
            if (listonly == 0) 
                printf("    File saves as [%s]\n", filename->data);

            if (savefiles == 1) {
                if ((listonly == 1) && (filenameonly == 1)) 
                    printf("%s\n", filename->data);
                if (filepath == NULL) {
                    sprintf(ifilename, "%s", filename->data);
                } else {
                    sprintf(ifilename, "%s/%s", filepath, filename->data);
                }
                for(i=0; i<strlen(ifilename); i++) 
                    if (ifilename[i] == ' ') 
                        ifilename[i] = '_';
                if ((listonly == 1) && (filenameonly == 0)) 
                    printf("%s\n", ifilename);
                if ((fptr = fopen(ifilename, "wb"))==NULL) {
                    printf("Error writing file to disk!");
                } else {
                    object = 1;           
                    if((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(PT_OBJECT, PR_ATTACH_DATA_OBJ))) == MAPI_UNDEFINED) {
                        if((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(PT_BINARY, PR_ATTACH_DATA_OBJ))) == MAPI_UNDEFINED) {
                            filename = &(p->FileData);
                            object = 0;
                        }
                    }
                    if (object == 1) {
                        fwrite(filename->data + 16, sizeof(BYTE), filename->size - 16, fptr);
                    } else {
                        fwrite(filename->data, sizeof(BYTE), filename->size, fptr);
                    }
                    fclose(fptr);
                    if (object == 1) {
                        if(listonly == 0) 
                            printf("Attempting to parse embedded TNEF stream\n");
                        TNEFInitialize(&emb_tnef);
                        if (TNEFParseFile(ifilename, &emb_tnef) == -1) {
                            if (listonly == 0) 
                                printf(">>> ERROR processing file\n");
                        } else {
                            if(listonly == 0) 
                                printf("---> File %s\n", ifilename );
                            PrintTNEF(emb_tnef);
                        }
                        TNEFFree(&emb_tnef);
                    }

                }
            }
        }
        p=p->next;
    }
}

#include "utility.c"
#include "vcal.c"
#include "vcard.c"
#include "vtask.c"



