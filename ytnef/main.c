#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "tnef-types.h"
#include "tnef.h"
#include "mapi.h"
#include "mapidefs.h"

#define VERSION "1.15"
#define PRODID "PRODID:-//The Gauntlet//yTNEF v1.15//EN\n"

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
    printf("     Including saving the message text to a RTF file.\n");

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
                case 'v': verbose = 1;
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
                case 'v': verbose = 0;
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

void fprintProperty(FILE *FPTR, DWORD PROPTYPE, DWORD PROPID, char TEXT[]) {
    variableLength *vl;
    if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PROPTYPE, PROPID))) != MAPI_UNDEFINED) { 
        if (vl->size > 0)  
            if ((vl->size == 1) && (vl->data[0] == 0)) {
            } else { 
                fprintf(FPTR, TEXT, vl->data); 
            } 
    }
}

void fprintUserProp(FILE *FPTR, DWORD PROPTYPE, DWORD PROPID, char TEXT[]) {
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
        } else if (VL->data[index] == ',') {
            fprintf(FPTR, "\\,");
        } else { 
            fprintf(FPTR, "%c", VL->data[index]); 
        } 
    }
}

void SaveVCalendar(TNEFStruct TNEF) {
    char ifilename[256];
    variableLength *filename;
    char *charptr, *charptr2;
    FILE *fptr;
    int index;
    DDWORD *ddword_ptr;
    DDWORD ddword_val;
    dtr thedate;
    int i;

    if (listonly == 0) 
        printf("-> Creating an icalendar attachment\n");
    if ((listonly == 1) && (filenameonly == 1)) 
        printf("calendar.vcf\n");
    if (filepath == NULL) {
        sprintf(ifilename, "calendar.vcf");
    } else {
        sprintf(ifilename, "%s/calendar.vcf", filepath);
    }
    if ((listonly == 1) && (filenameonly == 0)) 
        printf("%s\n", ifilename);

    if ((fptr = fopen(ifilename, "wb"))==NULL) {
            printf("Error writing file to disk!");
    } else {
        fprintf(fptr, "BEGIN:VCALENDAR\n");
        if (TNEF.messageClass[0] != 0) {
            charptr2=TNEF.messageClass;
            charptr=charptr2;
            while (*charptr != 0) {
                if (*charptr == '.') {
                    charptr2 = charptr;
                }
                charptr++;
            }
            if (strcmp(charptr2, ".MtgCncl") == 0) {
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
        filename = NULL;
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_BINARY, 0x3))) == MAPI_UNDEFINED) {
            if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_BINARY, 0x23))) == MAPI_UNDEFINED) {
                filename = NULL;
            }
        }
        if (filename!=NULL) {
            fprintf(fptr, "UID:");
            for(index=0;index<filename->size;index++) {
                fprintf(fptr,"%02x", (unsigned char)filename->data[index]);
            }
            fprintf(fptr,"\n");
        }

        // Sequence
        filename = NULL;
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_LONG, 0x8201))) != MAPI_UNDEFINED) {
            ddword_ptr = (DDWORD*)filename->data;
            fprintf(fptr, "SEQUENCE:%i\n", *ddword_ptr);
        }
        if ((filename=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_BINARY, PR_SENDER_SEARCH_KEY))) != MAPI_UNDEFINED) {
            charptr = filename->data;
            charptr2 = strstr(charptr, ":");
            if (charptr2 == NULL) 
                charptr2 = charptr;
            else
                charptr2++;
            fprintf(fptr, "ORGANIZER:MAILTO:%s\n", charptr2);
        }

        // Required Attendees
        if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x823b))) != MAPI_UNDEFINED) {
                // We have a list of required participants, so
                // write them out.
            if (filename->size > 1) {
                charptr = filename->data-1;
                charptr2=strstr(charptr+1, ";");
                while (charptr != NULL) {
                    charptr++;
                    charptr2 = strstr(charptr, ";");
                    if (charptr2 != NULL) {
                        *charptr2 = 0;
                    }
                    while (*charptr == ' ') 
                        charptr++;
                    fprintf(fptr, "ATTENDEE;CN=%s;ROLE=REQ-PARTICIPANT:%s\n", charptr, charptr);
                    charptr = charptr2;
                }
            }
            // Optional attendees
            if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x823c))) != MAPI_UNDEFINED) {
                    // The list of optional participants
                if (filename->size > 1) {
                    charptr = filename->data-1;
                    charptr2=strstr(charptr+1, ";");
                    while (charptr != NULL) {
                        charptr++;
                        charptr2 = strstr(charptr, ";");
                        if (charptr2 != NULL) {
                            *charptr2 = 0;
                        }
                        while (*charptr == ' ') 
                            charptr++;
                        fprintf(fptr, "ATTENDEE;CN=%s;ROLE=OPT-PARTICIPANT:%s\n", charptr, charptr);
                        charptr = charptr2;
                    }
                }
            }
        } else if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x8238))) != MAPI_UNDEFINED) {
            if (filename->size > 1) {
                charptr = filename->data-1;
                charptr2=strstr(charptr+1, ";");
                while (charptr != NULL) {
                    charptr++;
                    charptr2 = strstr(charptr, ";");
                    if (charptr2 != NULL) {
                        *charptr2 = 0;
                    }
                    while (*charptr == ' ') 
                        charptr++;
                    fprintf(fptr, "ATTENDEE;CN=%s;ROLE=REQ-PARTICIPANT:%s\n", charptr, charptr);
                    charptr = charptr2;
                }
            }

        }
        // Summary
        filename = NULL;
        if((filename=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_CONVERSATION_TOPIC)))!=MAPI_UNDEFINED) {
            fprintf(fptr, "SUMMARY:");
            Cstylefprint(fptr, filename);
            fprintf(fptr, "\n");
        }
        // Location
        filename = NULL;
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x0002))) == MAPI_UNDEFINED) {
            if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x8208))) == MAPI_UNDEFINED) {
                filename = NULL;
            }
        }
        if (filename != NULL) {
            fprintf(fptr,"LOCATION: %s\n", filename->data);
        }
        // Date Start
        filename = NULL;
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x820d))) == MAPI_UNDEFINED) {
            if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8516))) == MAPI_UNDEFINED) {
                filename=NULL;
            }
        }
        if (filename != NULL) {
            fprintf(fptr, "DTSTART:");
            MAPISysTimetoDTR(filename->data, &thedate);
            fprintf(fptr,"%04i%02i%02iT%02i%02i%02iZ\n", 
                    thedate.wYear, thedate.wMonth, thedate.wDay,
                    thedate.wHour, thedate.wMinute, thedate.wSecond);
        }
        // Date End
        filename = NULL;
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x820e))) == MAPI_UNDEFINED) {
            if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8517))) == MAPI_UNDEFINED) {
                filename=NULL;
            }
        }
        if (filename != NULL) {
            fprintf(fptr, "DTEND:");
            MAPISysTimetoDTR(filename->data, &thedate);
            fprintf(fptr,"%04i%02i%02iT%02i%02i%02iZ\n", 
                    thedate.wYear, thedate.wMonth, thedate.wDay,
                    thedate.wHour, thedate.wMinute, thedate.wSecond);
        }
        // Date Stamp
        filename = NULL;
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8202))) != MAPI_UNDEFINED) {
            fprintf(fptr, "CREATED:");
            MAPISysTimetoDTR(filename->data, &thedate);
            fprintf(fptr,"%04i%02i%02iT%02i%02i%02iZ\n", 
                    thedate.wYear, thedate.wMonth, thedate.wDay,
                    thedate.wHour, thedate.wMinute, thedate.wSecond);
        }
        // Class
        filename = NULL;
        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_BOOLEAN, 0x8506))) != MAPI_UNDEFINED) {
            ddword_ptr = (DDWORD*)filename->data;
            ddword_val = SwapDDWord((BYTE*)ddword_ptr);
            fprintf(fptr, "CLASS:" );
            if (*ddword_ptr == 1) {
                fprintf(fptr,"PRIVATE\n");
            } else {
                fprintf(fptr,"PUBLIC\n");
            }
        }

        // Wrap it up
        fprintf(fptr, "END:VEVENT\n");
        fprintf(fptr, "END:VCALENDAR\n");
        fclose(fptr);
    }
}


void SaveVCard(TNEFStruct TNEF) {
    char ifilename[512];
    FILE *fptr;
    variableLength *vl;
    variableLength *pobox, *street, *city, *state, *zip, *country;
    dtr thedate;
    int boolean,index,i;

    if (listonly == 0) 
        printf("-> Creating a vCard attachment: ");
    if ((vl = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_DISPLAY_NAME))) == MAPI_UNDEFINED) {
        if (listonly == 0) 
            printf("\n-> Contact has no name. Aborting\n");
        return;
    }
    
    if ((listonly == 1) && (filenameonly == 1)) 
        printf("%s.vcard\n", vl->data);
    if (filepath == NULL) {
        sprintf(ifilename, "%s.vcard", vl->data);
    } else {
        sprintf(ifilename, "%s/%s.vcard", filepath, vl->data);
    }
    for(i=0; i<strlen(ifilename); i++) 
        if (ifilename[i] == ' ') 
            ifilename[i] = '_';
    if ((listonly == 1) && (filenameonly == 0)) 
        printf("%s.vcard\n", vl->data);
    if (listonly == 0) 
        printf("%s\n", ifilename);
    if ((fptr = fopen(ifilename, "wb"))==NULL) {
            printf("Error writing file to disk!");
    } else {
        fprintf(fptr, "BEGIN:VCARD\n");
        fprintf(fptr, "VERSION:2.1\n");
        fprintf(fptr, "FN:%s\n", vl->data);

        fprintProperty(fptr, PT_STRING8, PR_NICKNAME, "NICKNAME:%s\n");
        fprintUserProp(fptr, PT_STRING8, 0x8554, "MAILER:Microsoft Outlook %s\n");
        fprintProperty(fptr, PT_STRING8, PR_SPOUSE_NAME, "X-EVOLUTION-SPOUSE:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_MANAGER_NAME, "X-EVOLUTION-MANAGER:%s\n");       
        fprintProperty(fptr, PT_STRING8, PR_ASSISTANT, "X-EVOLUTION-ASSISTANT:%s\n");

        // Organizational
        if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_COMPANY_NAME))) != MAPI_UNDEFINED) {
            if (vl->size > 0) {
                if ((vl->size == 1) && (vl->data[0] == 0)) {
                } else {
                    fprintf(fptr,"ORG:%s", vl->data);
                    if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_DEPARTMENT_NAME))) != MAPI_UNDEFINED) {
                        fprintf(fptr,";%s", vl->data);
                    }
                    fprintf(fptr, "\n");
                }
            }
        }

        fprintProperty(fptr, PT_STRING8, PR_OFFICE_LOCATION, "X-EVOLUTION-OFFICE:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_TITLE, "TITLE:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_PROFESSION, "ROLE:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_BODY, "NOTE:%s\n");
        if (TNEF.body.size > 0) {
            fprintf(fptr, "NOTE;QUOTED-PRINTABLE:");
            quotedfprint(fptr, &(TNEF.body));
            fprintf(fptr,"\n");
        }


        // Business Address
        boolean = 0;
        if ((pobox = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_POST_OFFICE_BOX))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((street = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_STREET_ADDRESS))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((city = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_LOCALITY))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((state = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_STATE_OR_PROVINCE))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((zip = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_POSTAL_CODE))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((country = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_COUNTRY))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if (boolean == 1) {
            fprintf(fptr, "ADR;QUOTED-PRINTABLE;WORK:");
            if (pobox != MAPI_UNDEFINED) {
                quotedfprint(fptr, pobox);
            }
            fprintf(fptr, ";;");
            if (street != MAPI_UNDEFINED) {
                quotedfprint(fptr, street);
            }
            fprintf(fptr, ";");
            if (city != MAPI_UNDEFINED) {
                quotedfprint(fptr, city);
            }
            fprintf(fptr, ";");
            if (state != MAPI_UNDEFINED) {
                quotedfprint(fptr, state);
            }
            fprintf(fptr, ";");
            if (zip != MAPI_UNDEFINED) {
                quotedfprint(fptr, zip);
            }
            fprintf(fptr, ";");
            if (country != MAPI_UNDEFINED) {
                quotedfprint(fptr, country);
            }
            fprintf(fptr,"\n");
            if ((vl = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x801b))) != MAPI_UNDEFINED) {
                fprintf(fptr, "LABEL;QUOTED-PRINTABLE;WORK:");
                quotedfprint(fptr, vl);
                fprintf(fptr,"\n");
            }
        }

        // Home Address
        boolean = 0;
        if ((pobox = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_HOME_ADDRESS_POST_OFFICE_BOX))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((street = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_HOME_ADDRESS_STREET))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((city = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_HOME_ADDRESS_CITY))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((state = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_HOME_ADDRESS_STATE_OR_PROVINCE))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((zip = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_HOME_ADDRESS_POSTAL_CODE))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((country = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_HOME_ADDRESS_COUNTRY))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if (boolean == 1) {
            fprintf(fptr, "ADR;QUOTED-PRINTABLE;HOME:");
            if (pobox != MAPI_UNDEFINED) {
                quotedfprint(fptr, pobox);
            }
            fprintf(fptr, ";;");
            if (street != MAPI_UNDEFINED) {
                quotedfprint(fptr, street);
            }
            fprintf(fptr, ";");
            if (city != MAPI_UNDEFINED) {
                quotedfprint(fptr, city);
            }
            fprintf(fptr, ";");
            if (state != MAPI_UNDEFINED) {
                quotedfprint(fptr, state);
            }
            fprintf(fptr, ";");
            if (zip != MAPI_UNDEFINED) {
                quotedfprint(fptr, zip);
            }
            fprintf(fptr, ";");
            if (country != MAPI_UNDEFINED) {
                quotedfprint(fptr, country);
            }
            fprintf(fptr,"\n");
            if ((vl = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x801a))) != MAPI_UNDEFINED) {
                fprintf(fptr, "LABEL;QUOTED-PRINTABLE;WORK:");
                quotedfprint(fptr, vl);
                fprintf(fptr,"\n");
            }
        }

        // Other Address
        boolean = 0;
        if ((pobox = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_OTHER_ADDRESS_POST_OFFICE_BOX))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((street = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_OTHER_ADDRESS_STREET))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((city = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_OTHER_ADDRESS_CITY))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((state = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_OTHER_ADDRESS_STATE_OR_PROVINCE))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((zip = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_OTHER_ADDRESS_POSTAL_CODE))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if ((country = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_OTHER_ADDRESS_COUNTRY))) != MAPI_UNDEFINED) {
            boolean = 1;
        }
        if (boolean == 1) {
            fprintf(fptr, "ADR;QUOTED-PRINTABLE;OTHER:");
            if (pobox != MAPI_UNDEFINED) {
                quotedfprint(fptr, pobox);
            }
            fprintf(fptr, ";;");
            if (street != MAPI_UNDEFINED) {
                quotedfprint(fptr, street);
            }
            fprintf(fptr, ";");
            if (city != MAPI_UNDEFINED) {
                quotedfprint(fptr, city);
            }
            fprintf(fptr, ";");
            if (state != MAPI_UNDEFINED) {
                quotedfprint(fptr, state);
            }
            fprintf(fptr, ";");
            if (zip != MAPI_UNDEFINED) {
                quotedfprint(fptr, zip);
            }
            fprintf(fptr, ";");
            if (country != MAPI_UNDEFINED) {
                quotedfprint(fptr, country);
            }
            fprintf(fptr,"\n");
        }


        fprintProperty(fptr, PT_STRING8, PR_CALLBACK_TELEPHONE_NUMBER, "TEL;X-EVOLUTION-CALLBACK:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_PRIMARY_TELEPHONE_NUMBER, "TEL;PREF:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_MOBILE_TELEPHONE_NUMBER, "TEL;CELL:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_RADIO_TELEPHONE_NUMBER, "TEL;X-EVOLUTION-RADIO:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_CAR_TELEPHONE_NUMBER, "TEL;CAR:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_OTHER_TELEPHONE_NUMBER, "TEL;VOICE:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_PAGER_TELEPHONE_NUMBER, "TEL;PAGER:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_TELEX_NUMBER, "TEL;X-EVOLUTION-TELEX:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_ISDN_NUMBER, "TEL;ISDN:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_HOME2_TELEPHONE_NUMBER, "TEL;HOME:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_TTYTDD_PHONE_NUMBER, "TEL;X-EVOLUTION-TTYTDD:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_HOME_TELEPHONE_NUMBER, "TEL;HOME;VOICE:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_ASSISTANT_TELEPHONE_NUMBER, "TEL;X-EVOLUTION-ASSISTANT:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_COMPANY_MAIN_PHONE_NUMBER, "TEL;WORK:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_BUSINESS2_TELEPHONE_NUMBER, "TEL;WORK;VOICE:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_PRIMARY_FAX_NUMBER, "TEL;PREF;FAX:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_BUSINESS_FAX_NUMBER, "TEL;WORK;FAX:%s\n");
        fprintProperty(fptr, PT_STRING8, PR_HOME_FAX_NUMBER, "TEL;HOME;FAX:%s\n");


        // Email addresses
        if ((vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x8083))) == MAPI_UNDEFINED) {
            vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x8084));
        }
        if (vl != MAPI_UNDEFINED) {
            if (vl->size > 0) 
                fprintf(fptr, "EMAIL:%s\n", vl->data);
        }
        if ((vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x8093))) == MAPI_UNDEFINED) {
            vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x8094));
        }
        if (vl != MAPI_UNDEFINED) {
            if (vl->size > 0) 
                fprintf(fptr, "EMAIL:%s\n", vl->data);
        }
        if ((vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x80a3))) == MAPI_UNDEFINED) {
            vl=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x80a4));
        }
        if (vl != MAPI_UNDEFINED) {
            if (vl->size > 0) 
                fprintf(fptr, "EMAIL:%s\n", vl->data);
        }

        fprintProperty(fptr, PT_STRING8, PR_BUSINESS_HOME_PAGE, "URL:%s\n");
        fprintUserProp(fptr, PT_STRING8, 0x80d8, "FBURL:%s\n");



        //Birthday
        if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, PR_BIRTHDAY))) != MAPI_UNDEFINED) {
            fprintf(fptr, "BDAY:");
            MAPISysTimetoDTR(vl->data, &thedate);
            fprintf(fptr, "%i-%02i-%02i\n", thedate.wYear, thedate.wMonth, thedate.wDay);
        }

        //Anniversary
        if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, PR_WEDDING_ANNIVERSARY))) != MAPI_UNDEFINED) {
            fprintf(fptr, "X-EVOLUTION-ANNIVERSARY:");
            MAPISysTimetoDTR(vl->data, &thedate);
            fprintf(fptr, "%i-%02i-%02i\n", thedate.wYear, thedate.wMonth, thedate.wDay);
        }
        fprintf(fptr, "END:VCARD\n");

    }
}

void SaveVTask(TNEFStruct TNEF) {
    variableLength *vl;
    variableLength *filename;
    int index,i;
    char ifilename[256];
    char *charptr, *charptr2;
    dtr thedate;
    FILE *fptr;
    DDWORD *ddword_ptr;
    DDWORD ddword_val;
    
    vl = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_CONVERSATION_TOPIC));

    if (vl == MAPI_UNDEFINED) {
        if (listonly == 0) 
            printf("No Conversation Topic\n");
        return;
    }
    if (listonly == 0) 
        printf("-> Creating a vCalendar Task attachment: ");

    index = strlen(vl->data);
    while (vl->data[index] == ' ') 
            vl->data[index--] = 0;
    if ((listonly == 1) && (filenameonly == 1)) 
        printf("%s.vcf\n", vl->data);

    if (filepath == NULL) {
        sprintf(ifilename, "%s.vcf", vl->data);
    } else {
        sprintf(ifilename, "%s/%s.vcf", filepath, vl->data);
    }
    for(i=0; i<strlen(ifilename); i++) 
        if (ifilename[i] == ' ') 
            ifilename[i] = '_';
    if ((listonly == 1) && (filenameonly == 0)) 
        printf("%s.vcf\n", vl->data);
    if (listonly == 0) 
        printf("%s\n", ifilename);
    if ((fptr = fopen(ifilename, "wb"))==NULL) {
            printf("Error writing file to disk!");
    } else {
        fprintf(fptr, "BEGIN:VCALENDAR\n");
        fprintf(fptr, PRODID);
        fprintf(fptr, "VERSION:2.0\n");
        fprintf(fptr, "METHOD:PUBLISH\n");
        filename = NULL;

        fprintf(fptr, "BEGIN:VTODO\n");
        if (TNEF.messageID[0] != 0) {
            fprintf(fptr,"UID:%s\n", TNEF.messageID);
        }
        filename = MAPIFindUserProp(&(TNEF.MapiProperties), \
                        PROP_TAG(PT_STRING8, 0x8122));
        if (filename != MAPI_UNDEFINED) {
            fprintf(fptr, "ORGANIZER:%s\n", filename->data);
        }
                 

        if ((filename = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_DISPLAY_TO))) != MAPI_UNDEFINED) {
            filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x811f));
        }
        if ((filename != MAPI_UNDEFINED) && (filename->size > 1)) {
            charptr = filename->data-1;
            charptr2=strstr(charptr+1, ";");
            while (charptr != NULL) {
                charptr++;
                charptr2 = strstr(charptr, ";");
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
            fprintf(fptr,"SUMMARY:");
            Cstylefprint(fptr,&(TNEF.subject));
            fprintf(fptr,"\n");
        }

        if (TNEF.body.size > 0) {
            fprintf(fptr,"DESCRIPTION:");
            Cstylefprint(fptr,&(TNEF.body));
            fprintf(fptr,"\n");
        }

        filename = MAPIFindProperty(&(TNEF.MapiProperties), \
                    PROP_TAG(PT_SYSTIME, PR_CREATION_TIME));
        if (filename != MAPI_UNDEFINED) {
            fprintf(fptr, "DTSTAMP:");
            MAPISysTimetoDTR(filename->data, &thedate);
            fprintf(fptr,"%04i%02i%02iT%02i%02i%02iZ\n", 
                    thedate.wYear, thedate.wMonth, thedate.wDay,
                    thedate.wHour, thedate.wMinute, thedate.wSecond);
        }

        filename = MAPIFindUserProp(&(TNEF.MapiProperties), \
                    PROP_TAG(PT_SYSTIME, 0x8517));
        if (filename != MAPI_UNDEFINED) {
            printf("Found Due\n");
            fprintf(fptr, "DUE:");
            MAPISysTimetoDTR(filename->data, &thedate);
            printf("Convertinf\n");
            fprintf(fptr,"%04i%02i%02iT%02i%02i%02iZ\n", 
                    thedate.wYear, thedate.wMonth, thedate.wDay,
                    thedate.wHour, thedate.wMinute, thedate.wSecond);
        }
        filename = MAPIFindProperty(&(TNEF.MapiProperties), \
                    PROP_TAG(PT_SYSTIME, PR_LAST_MODIFICATION_TIME));
        if (filename != MAPI_UNDEFINED) {
            fprintf(fptr, "LAST-MODIFIED:");
            MAPISysTimetoDTR(filename->data, &thedate);
            fprintf(fptr,"%04i%02i%02iT%02i%02i%02iZ\n", 
                    thedate.wYear, thedate.wMonth, thedate.wDay,
                    thedate.wHour, thedate.wMinute, thedate.wSecond);
        }
        // Class
        filename = MAPIFindUserProp(&(TNEF.MapiProperties), \
                        PROP_TAG(PT_BOOLEAN, 0x8506));
        if (filename != MAPI_UNDEFINED) {
            ddword_ptr = (DDWORD*)filename->data;
            ddword_val = SwapDDWord((BYTE*)ddword_ptr);
            fprintf(fptr, "CLASS:" );
            if (*ddword_ptr == 1) {
                fprintf(fptr,"PRIVATE\n");
            } else {
                fprintf(fptr,"PUBLIC\n");
            }
        }
        fprintf(fptr, "END:VTODO\n");
        fprintf(fptr, "END:VCALENDAR\n");
        fclose(fptr);
    }

}
