#ifndef _TNEF_PROCS_H_
#define _TNEF_PROCS_H_

#include "tnef-types.h"
#include "mapi.h"
#include "mapidefs.h"
#define STD_ARGLIST (TNEFStruct *TNEF, int id, unsigned char *data, int size)
DWORD SwapDWord(BYTE *p);
WORD SwapWord(BYTE *p);


void TNEFInitMapi(MAPIProps *p);
void TNEFInitAttachment(Attachment *p);
void TNEFInitialize(TNEFStruct *TNEF);
void TNEFFree(TNEFStruct *TNEF);
void TNEFFreeAttachment(Attachment *p);
void TNEFFreeMapiProps(MAPIProps *p);
int TNEFCheckForSignature(DWORD sig);
int TNEFParseMemory(BYTE *memory, long size, TNEFStruct *TNEF);
int TNEFParseFile(char *filename, TNEFStruct *TNEF);
int TNEFParse(TNEFStruct *TNEF);
variableLength *MAPIFindUserProp(MAPIProps *p, unsigned int ID);
variableLength *MAPIFindProperty(MAPIProps *p, unsigned int ID);
int MAPISysTimetoDTR(BYTE *data, dtr *thedate);
void MAPIPrint(MAPIProps *p);
char* to_utf8(int len, char* buf);
WORD SwapWord(BYTE *p);
DWORD SwapDWord(BYTE *p);
DDWORD SwapDDWord(BYTE *p);
variableLength *MAPIFindUserProp(MAPIProps *p, unsigned int ID);
variableLength *MAPIFindProperty(MAPIProps *p, unsigned int ID);
unsigned char * DecompressRTF(variableLength *p, int *size);

/* ------------------------------------- */ 
/* TNEF Down-level Attributes/Properties */
/* ------------------------------------- */

#define atpTriples      ((WORD) 0x0000)
#define atpString       ((WORD) 0x0001)
#define atpText         ((WORD) 0x0002)
#define atpDate         ((WORD) 0x0003)
#define atpShort        ((WORD) 0x0004)
#define atpLong         ((WORD) 0x0005)
#define atpByte         ((WORD) 0x0006)
#define atpWord         ((WORD) 0x0007)
#define atpDword        ((WORD) 0x0008)
#define atpMax          ((WORD) 0x0009)

#define LVL_MESSAGE     ((BYTE) 0x01)
#define LVL_ATTACHMENT  ((BYTE) 0x02)

#define ATT_ID(_att)                ((WORD) ((_att) & 0x0000FFFF))
#define ATT_TYPE(_att)              ((WORD) (((_att) >> 16) & 0x0000FFFF))
#define ATT(_atp, _id)              ((((DWORD) (_atp)) << 16) | ((WORD) (_id)))

#define attNull                     ATT( 0,             0x0000)
#define attFrom                     ATT( atpTriples,    0x8000) /* PR_ORIGINATOR_RETURN_ADDRESS */
#define attSubject                  ATT( atpString,     0x8004) /* PR_SUBJECT */
#define attDateSent                 ATT( atpDate,       0x8005) /* PR_CLIENT_SUBMIT_TIME */
#define attDateRecd                 ATT( atpDate,       0x8006) /* PR_MESSAGE_DELIVERY_TIME */
#define attMessageStatus            ATT( atpByte,       0x8007) /* PR_MESSAGE_FLAGS */
#define attMessageClass             ATT( atpWord,       0x8008) /* PR_MESSAGE_CLASS */
#define attMessageID                ATT( atpString,     0x8009) /* PR_MESSAGE_ID */
#define attParentID                 ATT( atpString,     0x800A) /* PR_PARENT_ID */
#define attConversationID           ATT( atpString,     0x800B) /* PR_CONVERSATION_ID */
#define attBody                     ATT( atpText,       0x800C) /* PR_BODY */
#define attPriority                 ATT( atpShort,      0x800D) /* PR_IMPORTANCE */
#define attAttachData               ATT( atpByte,       0x800F) /* PR_ATTACH_DATA_xxx */
#define attAttachTitle              ATT( atpString,     0x8010) /* PR_ATTACH_FILENAME */
#define attAttachMetaFile           ATT( atpByte,       0x8011) /* PR_ATTACH_RENDERING */
#define attAttachCreateDate         ATT( atpDate,       0x8012) /* PR_CREATION_TIME */
#define attAttachModifyDate         ATT( atpDate,       0x8013) /* PR_LAST_MODIFICATION_TIME */
#define attDateModified             ATT( atpDate,       0x8020) /* PR_LAST_MODIFICATION_TIME */
#define attAttachTransportFilename  ATT( atpByte,       0x9001) /* PR_ATTACH_TRANSPORT_NAME */
#define attAttachRenddata           ATT( atpByte,       0x9002)
#define attMAPIProps                ATT( atpByte,       0x9003)
#define attRecipTable               ATT( atpByte,       0x9004) /* PR_MESSAGE_RECIPIENTS */
#define attAttachment               ATT( atpByte,       0x9005)
#define attTnefVersion              ATT( atpDword,      0x9006)
#define attOemCodepage              ATT( atpByte,       0x9007)
#define attOriginalMessageClass     ATT( atpWord,       0x0006) /* PR_ORIG_MESSAGE_CLASS */

#define attOwner                    ATT( atpByte,       0x0000) /* PR_RCVD_REPRESENTING_xxx  or
                                                                                                                                      PR_SENT_REPRESENTING_xxx */
#define attSentFor                  ATT( atpByte,       0x0001) /* PR_SENT_REPRESENTING_xxx */
#define attDelegate                 ATT( atpByte,       0x0002) /* PR_RCVD_REPRESENTING_xxx */
#define attDateStart                ATT( atpDate,       0x0006) /* PR_DATE_START */
#define attDateEnd                  ATT( atpDate,       0x0007) /* PR_DATE_END */
#define attAidOwner                 ATT( atpLong,       0x0008) /* PR_OWNER_APPT_ID */
#define attRequestRes               ATT( atpShort,      0x0009) /* PR_RESPONSE_REQUESTED */

typedef struct {
    DWORD id;
    char name[40];
    int (*handler) STD_ARGLIST;
} TNEFHandler;


#endif
