#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include "tnef.h"
#include "mapi.h"
#include "mapidefs.h"
#include "mapitags.h"

#define MIN(x,y) (((x)<(y))?(x):(y))
void TNEFFillMapi(BYTE *data, DWORD size, MAPIProps *p);
void SetFlip(void);

int TNEFDefaultHandler STD_ARGLIST;
int TNEFAttachmentFilename STD_ARGLIST;
int TNEFAttachmentSave STD_ARGLIST;
int TNEFDetailedPrint STD_ARGLIST;
int TNEFHexBreakdown STD_ARGLIST;
int TNEFBody STD_ARGLIST;
int TNEFRendData STD_ARGLIST;
int TNEFDateHandler STD_ARGLIST;
int TNEFPriority  STD_ARGLIST;
int TNEFVersion  STD_ARGLIST;
int TNEFMapiProperties STD_ARGLIST;
int TNEFIcon STD_ARGLIST;
int TNEFSubjectHandler STD_ARGLIST;
int TNEFFromHandler STD_ARGLIST;
int TNEFRecipTable STD_ARGLIST;
int TNEFAttachmentMAPI STD_ARGLIST;
int TNEFSentFor STD_ARGLIST;
int TNEFMessageClass STD_ARGLIST;
int TNEFMessageID STD_ARGLIST;
int TNEFParentID STD_ARGLIST;
int TNEFOriginalMsgClass STD_ARGLIST;
int TNEFCodePage STD_ARGLIST;
int ByteOrder = -1;

BYTE *TNEFFileContents=NULL;
DWORD TNEFFileContentsSize;
BYTE *TNEFFileIcon=NULL;
DWORD TNEFFileIconSize;

TNEFHandler TNEFList[] = {
        {attNull,                    "Null",                        TNEFDefaultHandler},
        {attFrom,                    "From",                        TNEFFromHandler},
        {attSubject,                 "Subject",                     TNEFSubjectHandler},
        {attDateSent,                "Date Sent",                   TNEFDateHandler},
        {attDateRecd,                "Date Received",               TNEFDateHandler},
        {attMessageStatus,           "Message Status",              TNEFDefaultHandler},
        {attMessageClass,            "Message Class",               TNEFMessageClass},
        {attMessageID,               "Message ID",                  TNEFMessageID},
        {attParentID,                "Parent ID",                   TNEFParentID},
        {attConversationID,          "Conversation ID",             TNEFDefaultHandler},
        {attBody,                    "Body",                        TNEFBody},
        {attPriority,                "Priority",                    TNEFPriority},
        {attAttachData,              "Attach Data",                 TNEFAttachmentSave},
        {attAttachTitle,             "Attach Title",                TNEFAttachmentFilename},
        {attAttachMetaFile,          "Attach Meta-File",            TNEFIcon},
        {attAttachCreateDate,        "Attachment Create Date",      TNEFDateHandler},
        {attAttachModifyDate,        "Attachment Modify Date",      TNEFDateHandler},
        {attDateModified,            "Date Modified",               TNEFDateHandler},
        {attAttachTransportFilename, "Attachment Transport name",   TNEFDefaultHandler},
        {attAttachRenddata,          "Attachment Display info",     TNEFRendData},
        {attMAPIProps,               "MAPI Properties",             TNEFMapiProperties},
        {attRecipTable,              "Recip Table",                 TNEFRecipTable},
        {attAttachment,              "Attachment",                  TNEFAttachmentMAPI},
        {attTnefVersion,             "TNEF Version",                TNEFVersion},
        {attOemCodepage,             "OEM CodePage",                TNEFCodePage},
        {attOriginalMessageClass,    "Original Message Class",      TNEFOriginalMsgClass},
        {attOwner,                   "Owner",                       TNEFDefaultHandler},
        {attSentFor,                 "Sent For",                    TNEFSentFor},
        {attDelegate,                "Delegate",                    TNEFDefaultHandler},
        {attDateStart,               "Date Start",                  TNEFDateHandler},
        {attDateEnd,                 "Date End",                    TNEFDateHandler},
        {attAidOwner,                "Aid Owner",                   TNEFDefaultHandler},
        {attRequestRes,              "Request Response",            TNEFDefaultHandler} };


WORD SwapWord(BYTE *p) 
{
    WORD *word_ptr;
    BYTE bytes[2];
    if (ByteOrder == -1) {
        SetFlip();
    }

    if (ByteOrder == 0) {
        word_ptr = (WORD*)p;
        return *word_ptr;
    } else {
        bytes[0] = p[1];
        bytes[1] = p[0];
        word_ptr = (WORD*)&(bytes[0]);
        return *word_ptr;
    }
}

DWORD SwapDWord(BYTE *p)
{
    DWORD *dword_ptr;
    BYTE bytes[4];
    if (ByteOrder == -1) {
        SetFlip();
    }

    if (ByteOrder == 0) {
        dword_ptr = (DWORD*)p;
        return *dword_ptr;
    } else {
        bytes[0] = p[3];
        bytes[1] = p[2];
        bytes[2] = p[1];
        bytes[3] = p[0];
        dword_ptr = (DWORD*)&(bytes[0]);
        return *dword_ptr;
    }
}

DDWORD SwapDDWord(BYTE *p)
{
    DDWORD *ddword_ptr;
    BYTE bytes[8];
    if (ByteOrder == -1) {
        SetFlip();
    }

    if (ByteOrder == 0) {
        ddword_ptr = (DDWORD*)p;
        return *ddword_ptr;
    } else {
        bytes[0] = p[7];
        bytes[1] = p[6];
        bytes[2] = p[5];
        bytes[3] = p[4];
        bytes[4] = p[3];
        bytes[5] = p[2];
        bytes[6] = p[1];
        bytes[7] = p[0];
        ddword_ptr = (DDWORD*)&(bytes[0]);
        return *ddword_ptr;
    }
}


void SetFlip(void) {
    DWORD x = 0x04030201;
    int i;
    BYTE *p;

    p = (BYTE*)&x;
    for(i=0;i<sizeof(DWORD); i++) {
	    printf("%02x ", *(p+i));
    }
    printf(":");
    if (*p == 1) {
        printf("Detected Little-Endian architecture\n");
        ByteOrder = 0;
    } else {
        printf("Detected Big-Endian architecture\n");
        ByteOrder = 1;
    }
}

// -----------------------------------------------------------------------------
int TNEFDefaultHandler STD_ARGLIST {
    printf("%s: [%i] %s\n", TNEFList[id].name, size, data);
    return 0;
}

// -----------------------------------------------------------------------------
int TNEFCodePage STD_ARGLIST {
    TNEF->CodePage.size = size;
    TNEF->CodePage.data = calloc(size, sizeof(BYTE));
    memcpy(TNEF->CodePage.data, data, size);
    return 0;
}

// -----------------------------------------------------------------------------
int TNEFParentID STD_ARGLIST {
    memcpy(TNEF->parentID, data, MIN(size,sizeof(TNEF->parentID)));
    return 0;
}
// -----------------------------------------------------------------------------
int TNEFMessageID STD_ARGLIST {
    memcpy(TNEF->messageID, data, MIN(size,sizeof(TNEF->messageID)));
    return 0;
}
// -----------------------------------------------------------------------------
int TNEFBody STD_ARGLIST {
    TNEF->body.size = size;
    TNEF->body.data = calloc(size, sizeof(BYTE));
    memcpy(TNEF->body.data, data, size);
    return 0;
}
// -----------------------------------------------------------------------------
int TNEFOriginalMsgClass STD_ARGLIST {
    TNEF->OriginalMessageClass.size = size;
    TNEF->OriginalMessageClass.data = calloc(size, sizeof(BYTE));
    memcpy(TNEF->OriginalMessageClass.data, data, size);
    return 0;
}
// -----------------------------------------------------------------------------
int TNEFMessageClass STD_ARGLIST {
    memcpy(TNEF->messageClass, data, MIN(size,sizeof(TNEF->messageClass)));
    return 0;
}
// -----------------------------------------------------------------------------
int TNEFFromHandler STD_ARGLIST {
    TNEF->from.data = calloc(size, sizeof(BYTE));
    TNEF->from.size = size;
    memcpy(TNEF->from.data, data,size);
    return 0;
}
// -----------------------------------------------------------------------------
int TNEFSubjectHandler STD_ARGLIST {
    TNEF->subject.data = calloc(size, sizeof(BYTE));
    TNEF->subject.size = size;
    memcpy(TNEF->subject.data, data,size);
    return 0;
}

// -----------------------------------------------------------------------------
int TNEFRendData STD_ARGLIST {
    Attachment *p;
    // Find the last attachment.
    p = &(TNEF->starting_attach);
    while (p->next!=NULL) p=p->next;

    // Add a new one
    p->next = calloc(1,sizeof(Attachment));
    p=p->next;

    TNEFInitAttachment(p);
    
    memcpy(&(p->RenderData), data, sizeof(renddata));
    return 0;
}

// -----------------------------------------------------------------------------
int TNEFVersion STD_ARGLIST {
    WORD major;
    WORD minor;
    major = SwapWord(data+2);
    minor = SwapWord(data);
	   
    sprintf(TNEF->version, "TNEF%i.%i", major, minor);
    return 0;
}

// -----------------------------------------------------------------------------
int TNEFIcon STD_ARGLIST {
    Attachment *p;
    // Find the last attachment.
    p = &(TNEF->starting_attach);
    while (p->next!=NULL) p=p->next;

    p->IconData.size = size;
    p->IconData.data = calloc(size, sizeof(BYTE));
    memcpy(p->IconData.data, data, size);
}

// -----------------------------------------------------------------------------
int TNEFRecipTable STD_ARGLIST {
    DWORD count;
    BYTE *d;
    int current_row;
    int propcount;
    int current_prop;

    d = data;
    count = SwapDWord(d);
    d += 4;
//    printf("Recipient Table containing %u rows\n", count);

    return 0;

    for(current_row=0; current_row<count; current_row++) {
	propcount = SwapDWord(d);
        printf("> Row %i contains %i properties\n", current_row, propcount);
        d+=4;
        for(current_prop=0; current_prop<propcount; current_prop++) {


        }
    }
    return 0;
}
// -----------------------------------------------------------------------------
int TNEFAttachmentMAPI STD_ARGLIST {
    Attachment *p;
    // Find the last attachment.
    //
    p = &(TNEF->starting_attach);
    while (p->next!=NULL) p=p->next;
    TNEFFillMapi(data, size, &(p->MAPI));

    return 0;
}
// -----------------------------------------------------------------------------
int TNEFMapiProperties STD_ARGLIST {
    TNEFFillMapi(data, size, &(TNEF->MapiProperties));
    return 0;
}

void TNEFFillMapi(BYTE *data, DWORD size, MAPIProps *p) {
    int i,j;
    DWORD num;
    BYTE *d;
    MAPIProperty *mp;
    DWORD type;
    DWORD length;
    variableLength *vl;

    WORD temp_word;
    DWORD temp_dword;
    DDWORD temp_ddword;
    int count=-1;
    

    d = data;
    p->count = SwapDWord(data);
    d += 4;
    p->properties = calloc(p->count, sizeof(MAPIProperty));
    mp = p->properties;

    for(i=0; i<p->count; i++) {
        if (count == -1) {
	    mp->id = SwapDWord(d);
            d+=4;
            mp->custom = 0;
            mp->count = 1;
            mp->namedproperty = 0;
            length = -1;
            if (PROP_ID(mp->id) >= 0x8000) {
                // Read the GUID
                memcpy(&(mp->guid[0]), d, 16);
                d+=16;
    
                length = SwapDWord(d);
                d+=sizeof(DWORD);
                if (length > 0) {
                    mp->namedproperty = length;
                    mp->propnames = calloc(length, sizeof(variableLength));
                    while (length > 0) {
                        type = SwapDWord(d);
                        mp->propnames[length-1].data = calloc(type, sizeof(BYTE));
                        mp->propnames[length-1].size = type;
                        d+=4;
                        for(j=0; j<(type>>1); j++) {
                            mp->propnames[length-1].data[j] = d[j*2];
                        }
                        d += type + ((type % 4) ? (4 - type%4) : 0);
                        length--;
                    }
                } else {
                    // READ the type
                    type = SwapDWord(d);
                    d+=sizeof(DWORD);
                    mp->id = PROP_TAG(PROP_TYPE(mp->id), type);
                }
    
                mp->custom = 1;
            }
            
            //printf("Type id = %04x\n", PROP_TYPE(mp->id));
            if (PROP_TYPE(mp->id) & MV_FLAG) {
                mp->id = PROP_TAG(PROP_TYPE(mp->id) - MV_FLAG, PROP_ID(mp->id));
                mp->count = SwapDWord(d);
                d+=4;
                count = 0;
            }
            mp->data = calloc(mp->count, sizeof(variableLength));
            vl = mp->data;
        } else {
            i--;
            count++;
            vl = &(mp->data[count]);
            if (count == (mp->count-1)) {
                count = -1;
            }
        }

        switch (PROP_TYPE(mp->id)) {
            case PT_BINARY:
            case PT_OBJECT:
            case PT_STRING8:
            case PT_UNICODE:
                // First number of objects (assume 1 for now)
                if (count == -1) {
                    vl->size = SwapDWord(d);
                    d+=4;
                }

                // now size of object
                vl->size = SwapDWord(d);
                d+=4;

                // now actual object
                vl->data = calloc(vl->size, sizeof(BYTE));
                memcpy(vl->data, d, vl->size);

                // Make sure to read in a multiple of 4
                num = vl->size;
                d += num + ((num % 4) ? (4 - num%4) : 0);
                break;

            case PT_I2:
                // Read in 2 bytes, but proceed by 4 bytes
                vl->size = 2;
                vl->data = calloc(vl->size, sizeof(WORD));
                temp_word = SwapWord(d);
                memcpy(vl->data, &temp_word, vl->size);
                d += 4;
                break;
            case PT_BOOLEAN:
            case PT_LONG:
            case PT_R4:
            case PT_CURRENCY:
            case PT_APPTIME:
            case PT_ERROR:
                vl->size = 4;
                vl->data = calloc(vl->size, sizeof(BYTE));
                temp_dword = SwapDWord(d);
                memcpy(vl->data, &temp_dword, vl->size);
                d += 4;
                break;
            case PT_DOUBLE:
            case PT_I8:
            case PT_SYSTIME:
                vl->size = 8;
                vl->data = calloc(vl->size, sizeof(BYTE));
                temp_ddword = SwapDDWord(d);
                memcpy(vl->data, &temp_ddword, vl->size);
                d+=8;
                break;
        }
        if (count == (mp->count-1)) {
            count = -1;
        }
        if (count == -1) {
            mp++;
        }
    }
    if ((d-data) < size) {
        printf("ERROR DURING MAPI READ\n");
        printf("Read %i bytes, Expected %i bytes\n", (d-data), size);
        printf("%i bytes missing\n", size - (d-data));
    } else if ((d-data) > size){
        printf("ERROR DURING MAPI READ\n");
        printf("Read %i bytes, Expected %i bytes\n", (d-data), size);
        printf("%i bytes extra\n", (d-data)-size);
    }
    return;
}
// -----------------------------------------------------------------------------
int TNEFSentFor STD_ARGLIST {
    WORD name_length, addr_length;
    BYTE *d;

    d=data;

    while ((d-data)<size) {
        name_length = SwapWord(d);
        d+=sizeof(WORD);
        printf("Sent For : %s", d);
        d+=name_length;

        addr_length = SwapWord(d);
        d+=sizeof(WORD);
        printf("<%s>\n", d);
        d+=addr_length;
    }

}
// -----------------------------------------------------------------------------
int TNEFDateHandler STD_ARGLIST {
    dtr *Date;
    Attachment *p;
    WORD *tmp_src, *tmp_dst;
    int i;

    p = &(TNEF->starting_attach);
    switch (TNEFList[id].id) {
        case attDateSent: Date = &(TNEF->dateSent); break;
        case attDateRecd: Date = &(TNEF->dateReceived); break;
        case attDateModified: Date = &(TNEF->dateModified); break;
        case attDateStart: Date = &(TNEF->DateStart); break;
        case attDateEnd:  Date = &(TNEF->DateEnd); break;
        case attAttachCreateDate:
            while (p->next!=NULL) p=p->next;
            Date = &(p->CreateDate);
            break;
        case attAttachModifyDate:
            while (p->next!=NULL) p=p->next;
            Date = &(p->ModifyDate);
            break;
        default:
            printf("MISSING CASE\n");
            return -1;
    }

    tmp_src = (WORD*)data;
    tmp_dst = (WORD*)Date;
    for(i=0;i<sizeof(dtr)/sizeof(WORD);i++) {
        *tmp_dst++ = SwapWord((BYTE*)tmp_src++);
    }
    return 0;
}

void TNEFPrintDate(dtr Date) {
    char days[7][15] = {"Sunday", "Monday", "Tuesday", 
            "Wednesday", "Thursday", "Friday", "Saturday"};
    char months[12][15] = {"January", "February", "March", "April", "May",
            "June", "July", "August", "September", "October", "November",
            "December"};
    if (Date.wDayOfWeek < 7) 
        printf("%s ", days[Date.wDayOfWeek]);
    
    if ((Date.wMonth < 13) && (Date.wMonth>0)) 
        printf("%s ", months[Date.wMonth-1]);

    printf("%hu, %hu ", Date.wDay, Date.wYear);

    if (Date.wHour>12) 
        printf("%hu:%02hu:%02hu pm", (Date.wHour-12), 
                Date.wMinute, Date.wSecond);
    else if (Date.wHour == 12) 
        printf("%hu:%02hu:%02hu pm", (Date.wHour), 
                Date.wMinute, Date.wSecond);
    else
        printf("%hu:%02hu:%02hu am", Date.wHour, 
                Date.wMinute, Date.wSecond);
}
// -----------------------------------------------------------------------------
int TNEFHexBreakdown STD_ARGLIST {
    int i;
    printf("%s: [%i bytes] \n", TNEFList[id].name, size);

    for(i=0; i<size; i++) {
        printf("%02x ", data[i]);
        if ((i+1)%16 == 0) printf("\n");
    }
    printf("\n");
}
    
// -----------------------------------------------------------------------------
int TNEFDetailedPrint STD_ARGLIST {
    int i;
    printf("%s: [%i bytes] \n", TNEFList[id].name, size);

    for(i=0; i<size; i++) {
        printf("%c", data[i]);
    }
    printf("\n");
}

// -----------------------------------------------------------------------------
int TNEFAttachmentFilename STD_ARGLIST {
    Attachment *p;
    p = &(TNEF->starting_attach);
    while (p->next!=NULL) p=p->next;

    p->Title.size = size;
    p->Title.data = calloc(size, sizeof(BYTE));
    memcpy(p->Title.data, data, size);

    return 0;
}

// -----------------------------------------------------------------------------
int TNEFAttachmentSave STD_ARGLIST {
    Attachment *p;
    p = &(TNEF->starting_attach);
    while (p->next!=NULL) p=p->next;

    p->FileData.data = calloc(sizeof(unsigned char), size);
    p->FileData.size = size;

    memcpy(p->FileData.data, data, size);

    return 0;
}

// -----------------------------------------------------------------------------
int TNEFPriority STD_ARGLIST {
    DWORD value;

    value = SwapDWord(data);
    switch (value) {
        case 3:
            sprintf((TNEF->priority), "high");
            break;
        case 2:
            sprintf((TNEF->priority), "normal");
            break;
        case 1: 
            sprintf((TNEF->priority), "low");
            break;
        default: 
            sprintf((TNEF->priority), "N/A");
            break;
    }
    return 0;
}

// -----------------------------------------------------------------------------
int TNEFCheckForSignature(FILE *fptr) {
    DWORD signature = 0x223E9F78;
    DWORD sig;
    if (fread(&sig, sizeof(DWORD), 1, fptr) <1) {
        return -1;
    }

    sig = SwapDWord((BYTE*)&sig);

    if (signature == sig) {
        return 0;
    } else {
        return -1;
    }
}

// -----------------------------------------------------------------------------
int TNEFGetKey(FILE *fptr, WORD *key) {
    if (fread(key, sizeof(WORD), 1, fptr) <1) {
        return -1;
    }
    *key = SwapWord((BYTE*)&key);
    return 0;
}

// -----------------------------------------------------------------------------
int TNEFGetHeader(FILE *fptr, DWORD *type, DWORD *size, WORD *header) {
    BYTE component;
    unsigned char temp[8];
    int i;
    WORD wtemp;

    if (fread(&component, sizeof(BYTE), 1, fptr) <1) {
        return -1;
    }
    if (fread(type, sizeof(DWORD), 1, fptr) <1) {
        return -1;
    }
    if (fread(size, sizeof(DWORD), 1, fptr) <1) {
        return -1;
    }
    
    *header = component;
    *type = SwapDWord((BYTE*)type);
    *size = SwapDWord((BYTE*)size);
    for(i=0; i<8; i++) {
        wtemp = temp[i];
        *header = (*header + wtemp);
    }

    return 0;
}

// -----------------------------------------------------------------------------
int TNEFRawRead(FILE *fptr, BYTE *data, DWORD size, WORD *checksum) {
    WORD temp;
    int i;
    if (fread(data, sizeof(BYTE), size, fptr) < size) {
        return -1;
    }

    if (checksum != NULL) {
        *checksum = 0;
        for(i=0; i<size; i++) {
            temp = data[i];
            *checksum = (*checksum + temp);
        }
    }
    return 0;
}

#define INITVARLENGTH(x) (x).data = NULL; (x).size = 0;
#define INITDTR(x) (x).wYear=0; (x).wMonth=0; (x).wDay=0; \
                   (x).wHour=0; (x).wMinute=0; (x).wSecond=0; \
                   (x).wDayOfWeek=0;
#define INITSTR(x) memset((x), 0, sizeof(x));
void TNEFInitMapi(MAPIProps *p)
{
    p->count = 0;
    p->properties = NULL;
}

void TNEFInitAttachment(Attachment *p)
{
    INITDTR(p->Date);
    INITVARLENGTH(p->Title);
    INITVARLENGTH(p->MetaFile);
    INITDTR(p->CreateDate);
    INITDTR(p->ModifyDate);
    INITVARLENGTH(p->TransportFilename);
    INITVARLENGTH(p->FileData);
    INITVARLENGTH(p->IconData);
    memset(&(p->RenderData), 0, sizeof(renddata));
    TNEFInitMapi(&(p->MAPI));
    p->next = NULL;
}

void TNEFInitialize(TNEFStruct *TNEF)
{
    INITSTR(TNEF->version);
    INITVARLENGTH(TNEF->from);
    INITVARLENGTH(TNEF->subject);
    INITDTR(TNEF->dateSent);
    INITDTR(TNEF->dateReceived);

    INITSTR(TNEF->messageStatus);
    INITSTR(TNEF->messageClass);
    INITSTR(TNEF->messageID);
    INITSTR(TNEF->parentID);
    INITSTR(TNEF->conversationID);
    INITVARLENGTH(TNEF->body);
    INITSTR(TNEF->priority);
    TNEFInitAttachment(&(TNEF->starting_attach));
    INITDTR(TNEF->dateModified);
    TNEFInitMapi(&(TNEF->MapiProperties));
    INITVARLENGTH(TNEF->CodePage);
    INITVARLENGTH(TNEF->OriginalMessageClass);
    INITVARLENGTH(TNEF->Owner);
    INITVARLENGTH(TNEF->SentFor);
    INITVARLENGTH(TNEF->Delegate);
    INITDTR(TNEF->DateStart);
    INITDTR(TNEF->DateEnd);
    INITVARLENGTH(TNEF->AidOwner);
    TNEF->RequestRes=0;
}
#undef INITVARLENGTH
#undef INITDTR
#undef INITSTR

#define FREEVARLENGTH(x) if ((x).size > 0) { \
                            free((x).data); (x).size =0; }
void TNEFFree(TNEFStruct *TNEF) {
    Attachment *p, *store;

    FREEVARLENGTH(TNEF->from);
    FREEVARLENGTH(TNEF->subject);
    FREEVARLENGTH(TNEF->body);
    FREEVARLENGTH(TNEF->CodePage);
    FREEVARLENGTH(TNEF->OriginalMessageClass);
    FREEVARLENGTH(TNEF->Owner);
    FREEVARLENGTH(TNEF->SentFor);
    FREEVARLENGTH(TNEF->Delegate);
    FREEVARLENGTH(TNEF->AidOwner);
    TNEFFreeMapiProps(&(TNEF->MapiProperties));

    p = TNEF->starting_attach.next;
    while (p != NULL) {
        TNEFFreeAttachment(p);
        store = p->next;
        free(p);
        p=store;
    }
}

void TNEFFreeAttachment(Attachment *p)
{
    FREEVARLENGTH(p->Title);
    FREEVARLENGTH(p->MetaFile);
    FREEVARLENGTH(p->TransportFilename);
    FREEVARLENGTH(p->FileData);
    FREEVARLENGTH(p->IconData);
    TNEFFreeMapiProps(&(p->MAPI));
}

void TNEFFreeMapiProps(MAPIProps *p)
{
    int i,j;
    for(i=0; i<p->count; i++) {
        for(j=0; j<p->properties[i].count; j++) {
            FREEVARLENGTH(p->properties[i].data[j]);
        }
        free(p->properties[i].data);
    }
    free(p->properties);
    p->count = 0;
}
#undef FREEVARLENGTH

// -----------------------------------------------------------------------------
int TNEFParseFile(char *filename, TNEFStruct *TNEF) {
    FILE *fptr;
    WORD key;
    DWORD type;
    DWORD size;
    BYTE *data;
    WORD checksum, header_checksum;
    int i;

    printf("Attempting to parse %s...\n", filename);
    if ((fptr = fopen(filename, "rb")) == NULL) {
        printf("Unable to open file %s\n", filename);
        return -1;
    }
   
    if (TNEFCheckForSignature(fptr) == -1) {
        printf("Signature does not match. Not a TNEF file.\n");
        fclose(fptr);
        return -1;
    }

    if (TNEFGetKey(fptr, &key) == -1) {
        printf("Unable to retrieve key.\n");
        fclose(fptr);
        return -1;
    }

    while (TNEFGetHeader(fptr, &type, &size, &header_checksum) == 0) {
        if (size > 0) {
            data = calloc(size, sizeof(BYTE));
            if (TNEFRawRead(fptr, data, size, &header_checksum)==-1) {
                printf("Unable to read data.\n");
                fclose(fptr);
                free(data);
                return -1;
            }
            if (TNEFRawRead(fptr, (BYTE *)&checksum, 2, NULL) == -1) {
                printf("Unable to read checksum.\n");
                fclose(fptr);
                free(data);
                return -1;
            }
	    checksum = SwapWord((BYTE*)&checksum);
            if (checksum != header_checksum) {
                printf("CHECKSUM ERROR:\n");
                fclose(fptr);
                free(data);
                return -1;
            }
            for(i=0; i<(sizeof(TNEFList)/sizeof(TNEFHandler));i++) {
                if (TNEFList[i].id == type) {
                    if (TNEFList[i].handler != NULL) {
                        if (TNEFList[i].handler(TNEF, i, data, size) == -1) {
                            free(data);
                            fclose(fptr);
                            return -1;
                        }
                    } else {
                        printf("No handler for %s: %i bytes\n", 
                                TNEFList[i].name, size);
                    }
                }
            }

            free(data);
        }
    }

    fclose(fptr);
    return 0;

}

// ----------------------------------------------------------------------------

variableLength *MAPIFindUserProp(MAPIProps *p, unsigned int ID) 
{
    int i;
    if (p != NULL) {
        for(i=0;i<p->count; i++) {
            if ((p->properties[i].id == ID) && (p->properties[i].custom == 1)) {
                return (p->properties[i].data);
            }
        }
    }
    return (variableLength*)-1;
}

variableLength *MAPIFindProperty(MAPIProps *p, unsigned int ID)
{
    int i;
    if (p != NULL) {
        for(i=0;i<p->count; i++) {
            if ((p->properties[i].id == ID) && (p->properties[i].custom == 0)) {
                return (p->properties[i].data);
            }
        }
    }
    return (variableLength*)-1;
}

int MAPISysTimetoDTR(BYTE *data, dtr *thedate)
{
    DDWORD ddword_tmp;
    int startingdate = 0;
    int tmp_date;
    unsigned int months[] = {31,28,31,30,31,30,31,31,30,31,30,31};


    ddword_tmp = *((DDWORD*)data);
    ddword_tmp = ddword_tmp /10; // micro-s
    ddword_tmp /= 1000; // ms
    ddword_tmp /= 1000; // s

    thedate->wSecond = (ddword_tmp % 60);

    ddword_tmp /= 60; // seconds to minutes
    thedate->wMinute = (ddword_tmp % 60);

    ddword_tmp /= 60; //minutes to hours
    thedate->wHour = (ddword_tmp % 24);

    ddword_tmp /= 24; // Hours to days

    // Now calculate the year based on # of days
    thedate->wYear = 1601;
    startingdate = 1; 
    while(ddword_tmp >= 365) {
        thedate->wYear++;
        ddword_tmp = (ddword_tmp - 365);
        startingdate++;
        if ((thedate->wYear % 4) == 0) {
            if ((thedate->wYear % 100) == 0) {
                // if the year is 1700,1800,1900, etc, then it is only 
                // a leap year if exactly divisible by 400, not 4.
                if ((thedate->wYear % 400) == 0) {
                    ddword_tmp = (ddword_tmp - 1);
                    startingdate++;
                }
            }  else {
                ddword_tmp = (ddword_tmp - 1);
                startingdate++;
            }
        }
        startingdate %= 7;
    }

    // the remaining number is the day # in this year
    // So now calculate the Month, & Day of month
    if ((thedate->wYear % 4) == 0) {
        // 29 days in february in a leap year
        months[2] = 29;
    }

    tmp_date = (int)ddword_tmp;
    thedate->wDayOfWeek = (tmp_date + startingdate) % 7;
    thedate->wMonth = 0;

    while (tmp_date > months[thedate->wMonth]) {
        tmp_date -= months[thedate->wMonth];
        thedate->wMonth++;
    }
    thedate->wMonth++;
    thedate->wDay = tmp_date+1;
    return 0;
}

void MAPIPrint(MAPIProps *p)
{
    int j, i,index;
    DDWORD *ddword_ptr;
    dtr thedate;
    MAPIProperty *mapi;
    variableLength *mapidata;
    int found;
    for(j=0; j<p->count; j++) {
        mapi = &(p->properties[j]);
        printf("   #%i: Type: [", j);
        switch (PROP_TYPE(mapi->id)) {
            case PT_UNSPECIFIED:
                printf("  NONE   "); break;
            case PT_NULL:
                printf("  NULL   "); break;
            case PT_I2:
                printf("   I2    "); break;
            case PT_LONG:
                printf("  LONG   "); break;
            case PT_R4:
                printf("   R4    "); break;
            case PT_DOUBLE:
                printf(" DOUBLE  "); break;
            case PT_CURRENCY:
                printf("CURRENCY "); break;
            case PT_APPTIME:
                printf("APP TIME "); break;
            case PT_ERROR:
                printf("  ERROR  "); break;
            case PT_BOOLEAN:
                printf(" BOOLEAN "); break;
            case PT_OBJECT:
                printf(" OBJECT  "); break;
            case PT_I8:
                printf("   I8    "); break;
            case PT_STRING8:
                printf(" STRING8 "); break;
            case PT_UNICODE:
                printf(" UNICODE "); break;
            case PT_SYSTIME:
                printf("SYS TIME "); break;
            case PT_CLSID:
                printf("OLE GUID "); break;
            case PT_BINARY:
                printf(" BINARY  "); break;
            default:
                printf("<%x>", PROP_TYPE(mapi->id)); break;
        }
                
        printf("]  Code: [");
        if (mapi->custom == 1) {
            printf("UD:x%04x", PROP_ID(mapi->id));
        } else {
            found = 0;
            for(index=0; index<sizeof(MPList)/sizeof(MAPIPropertyTagList); index++) {
                if ((MPList[index].id == PROP_ID(mapi->id)) && (found == 0)) {
                    printf("%s", MPList[index].name);
                    found = 1;
                }
            }
            if (found == 0) {
                printf("0x%04x", PROP_ID(mapi->id));
            }
        }
        printf("]\n");
        if (mapi->namedproperty > 0) {
            for(i=0; i<mapi->namedproperty; i++) {
                printf("    Name: %s\n", mapi->propnames[i].data);
            }
        }
        for (i=0;i<mapi->count;i++) {
            mapidata = &(mapi->data[i]);
            if (mapi->count > 1) {
                printf("    [%i] ", i);
            } else {
                printf("    ");
            }
            printf("Size: %i", mapidata->size);
            switch (PROP_TYPE(mapi->id)) {
                case PT_SYSTIME:
                    MAPISysTimetoDTR(mapidata->data, &thedate);
                    printf("    Value: ");
                    TNEFPrintDate(thedate);
                    printf("\n");
                    break;
                case PT_LONG:
                    printf("    Value: %li\n", *(mapidata->data));
                    break;
                case PT_I2:
                    printf("    Value: %hi\n", *(mapidata->data));
                    break;
                case PT_BOOLEAN:
                    if (mapi->data->data[0]!=0) {
                        printf("    Value: True\n");
                    } else {
                        printf("    Value: False\n");
                    }
                    break;
                case PT_OBJECT:
                    printf("\n");
                    break;
                default:
                    printf("    Value: [%s]\n", mapidata->data);
            }
        }
    }
}

