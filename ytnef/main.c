#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "tnef-types.h"
#include "tnef.h"
#include "mapi.h"
#include "mapidefs.h"

TNEFStruct TNEF;
int verbose = 0;
int savefiles = 0;
char *filepath = NULL;

void PrintTNEF(TNEFStruct TNEF);
void PrintHelp(void) {
    printf("Yerase TNEF Exporter v1.02\n");
    printf("\n");
    printf("  usage: ytnef [-+vhf] <filenames>\n");
    printf("\n");
    printf("   -/+v - Enables/Disables verbose printing of MAPI Properties\n");
    printf("   -/+f - Enables/Disables saving of attachments\n");
    printf("\n");
    printf("Example:\n");
    printf("  ytnef -v winmail.dat\n");
    printf("     Parse with verbose output, don't save\n");
    printf("  ytnef -f . winmail.dat\n");
    printf("     Parse and save all attachments to local directory (.)\n");
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
        
        printf("---> File %s\n", argv[i]);

        PrintTNEF(TNEF);
        TNEFFree(&TNEF);
    }
}





void PrintTNEF(TNEFStruct TNEF) {
    int index,i;
    int j;
    int count;
    FILE *fptr;
    char ifilename[256];
    char *charptr, *charptr2;
    DDWORD ddword_tmp;
    DDWORD *ddword_ptr;
    MAPIProps mapip;
    variableLength *filename;
    Attachment *p;
    dtr thedate;

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
    if (TNEF.messageClass[0] != 0) 
        printf("Message Class: %s\n", TNEF.messageClass);
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
    
    if (TNEF.MapiProperties.count > 0) {
        printf("    MAPI Properties: %i\n", TNEF.MapiProperties.count);
        if ((filename = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_BINARY,PR_RTF_COMPRESSED))) == (variableLength*)-1) {

        } else if (savefiles == 1) {
            if (filepath == NULL) {
                sprintf(ifilename, "message.rtf");
            } else {
                sprintf(ifilename, "%s/message.rtf", filepath);
            }
            if ((fptr = fopen(ifilename, "wb"))==NULL) {
                printf("Error writing file to disk!");
            } else {
                fwrite(filename->data, sizeof(BYTE), filename->size, fptr);
                fclose(fptr);
            }

        }
        if (verbose == 1) {
            MAPIPrint(&TNEF.MapiProperties);
        }
    }

    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8,0x24))) != (variableLength*)-1) {
        if (strcmp(filename->data, "IPM.Appointment") == 0) {
            printf("Found an appointment entry\n");
            if (savefiles == 1) {
                printf("-> Creating an icalendar attachment\n");
                if (filepath == NULL) {
                    sprintf(ifilename, "calendar.vcf");
                } else {
                    sprintf(ifilename, "%s/calendar.vcf", filepath, filename->data);
                }
                if ((fptr = fopen(ifilename, "wb"))==NULL) {
                        printf("Error writing file to disk!");
                } else {
                    fprintf(fptr, "BEGIN:VCALENDAR\n");
                    fprintf(fptr, "PRODID:-//The Gauntlet//Reader v1.0//EN\n");
                    fprintf(fptr, "VERSION:1.0\n");
                    fprintf(fptr, "BEGIN:VEVENT\n");

                    // Required Attendees
                    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x823b))) != (variableLength*)-1) {
                            // We have a list of required participants, so
                            // write them out.
                        if (filename->size > 1) {
                            charptr = filename->data-1;
                            charptr2=(strstr, charptr+1, ";");
                            while (charptr != NULL) {
                                charptr++;
                                charptr2 = strstr(charptr, ";");
                                if (charptr2 != NULL) {
                                    *charptr2 = 0;
                                }
                                fprintf(fptr, "ATTENDEE;ROLE=ATTENDEE;EXPECT=REQUIRE;%s\n", charptr);
                                charptr = charptr2;
                            }
                        }
                    }
                    // Optional attendees
                    if ((filename = MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x823c))) != (variableLength*)-1) {
                            // The list of optional participants
                        if (filename->size > 1) {
                            charptr = filename->data-1;
                            charptr2=(strstr, charptr+1, ";");
                            while (charptr != NULL) {
                                charptr++;
                                charptr2 = strstr(charptr, ";");
                                if (charptr2 != NULL) {
                                    *charptr2 = 0;
                                }
                                fprintf(fptr, "ATTENDEE;ROLE=ATTENDEE;EXPECT=REQUEST;%s\n", charptr);
                                charptr = charptr2;
                            }
                        }
                    }
                    // Summary
                    filename = NULL;
                    if((filename=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_CONVERSATION_TOPIC)))!=(variableLength*)-1) {
                        fprintf(fptr, "SUMMARY:%s\n", filename->data);
                    }
                    // Location
                    filename = NULL;
                    if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x0002))) == (variableLength*)-1) {
                        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, 0x8208))) == (variableLength*)-1) {
                            filename = NULL;
                        }
                    }
                    if (filename != NULL) {
                        fprintf(fptr,"LOCATION: %s\n", filename->data);
                    }
                    // UID
                    filename = NULL;
                    if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_BINARY, 0x3))) == (variableLength*)-1) {
                        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_BINARY, 0x23))) == (variableLength*)-1) {
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
                    // Date Start
                    filename = NULL;
                    if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8235))) == (variableLength*)-1) {
                        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x820d))) == (variableLength*)-1) {
                            if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8516))) == (variableLength*)-1) {
                                if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8502))) == (variableLength*)-1) {
                                    filename=NULL;
                                }
                            }
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
                    if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8236))) == (variableLength*)-1) {
                        if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x820e))) == (variableLength*)-1) {
                            if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8517))) == (variableLength*)-1) {
                                filename=NULL;
                            }
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
                    if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_SYSTIME, 0x8202))) != (variableLength*)-1) {
                        fprintf(fptr, "DCREATED:");
                        MAPISysTimetoDTR(filename->data, &thedate);
                        fprintf(fptr,"%04i%02i%02iT%02i%02i%02iZ\n", 
                                thedate.wYear, thedate.wMonth, thedate.wDay,
                                thedate.wHour, thedate.wMinute, thedate.wSecond);
                    }
                    // Sequence
                    filename = NULL;
                    if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_LONG, 0x8201))) != (variableLength*)-1) {
                        ddword_ptr = (DDWORD*)filename->data;
                        fprintf(fptr, "SEQUENCE:%i\n", *ddword_ptr);
                    }

                    // Class
                    filename = NULL;
                    if ((filename=MAPIFindUserProp(&(TNEF.MapiProperties), PROP_TAG(PT_BOOLEAN, 0x8506))) != (variableLength*)-1) {
                        ddword_ptr = (DDWORD*)filename->data;
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

        if (p->FileData.size > 0) {
            printf("    Attachment Size:  %ib\n", p->FileData.size);
            
            if ((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(30,0x3707))) == (variableLength*)-1) {
                if ((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(30,0x3001))) == (variableLength*)-1) {
                    filename = &(p->Title);
                }
            }
            printf("    File saves as [%s]\n", filename->data);

            if (savefiles == 1) {
                if (filepath == NULL) {
                    sprintf(ifilename, "%s", filename->data);
                } else {
                    sprintf(ifilename, "%s/%s", filepath, filename->data);
                }
                if ((fptr = fopen(ifilename, "wb"))==NULL) {
                        printf("Error writing file to disk!");
                } else {
                    if((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(PT_OBJECT, PR_ATTACH_DATA_OBJ))) == (variableLength*)-1) {
                        if((filename = MAPIFindProperty(&(p->MAPI), PROP_TAG(PT_BINARY, PR_ATTACH_DATA_OBJ))) == (variableLength*)-1) {
                            filename = &(p->FileData);
                        }
                    }
                    fwrite(filename->data, sizeof(BYTE), filename->size, fptr);
                    fclose(fptr);
                }
            }
        }
        p=p->next;
    }
}