#ifndef _TNEF_TYPES_H_
#define _TNEF_TYPES_H_

#define BYTE unsigned char
#define WORD unsigned short int
#define DWORD unsigned int
#define ULONG unsigned int
#define DDWORD unsigned long long
typedef struct
{
    WORD    atyp;
    ULONG   ulPosition;
    WORD    dxWidth;
    WORD    dyHeight;
    DWORD   dwFlags;
} renddata;

/* ----------------------------------- */
/* TNEF Down-level Date/Time Structure */
/* ----------------------------------- */

typedef struct 
{
    WORD    wYear;
    WORD    wMonth;
    WORD    wDay;
    WORD    wHour;
    WORD    wMinute;
    WORD    wSecond;
    WORD    wDayOfWeek;
} dtr;

typedef struct
{
    char *data;
    int size;
} variableLength;

typedef struct {
    DWORD custom;
    BYTE guid[16];
    DWORD id;
    ULONG count;
    int namedproperty;
    variableLength *propnames;
    variableLength *data;
} MAPIProperty;

typedef struct MAPIProps {
    DWORD count;
    MAPIProperty *properties;
} MAPIProps;

typedef struct Attachment
{
    dtr Date;
    variableLength Title;
    variableLength MetaFile;
    dtr CreateDate;
    dtr ModifyDate;
    variableLength TransportFilename;
    renddata RenderData;
    MAPIProps MAPI;
    struct Attachment *next;
    variableLength FileData;
    variableLength IconData;
} Attachment;

typedef struct _TNEFIOStruct
{
    int (*InitProc) (struct _TNEFIOStruct *IO);
    int (*ReadProc) (struct _TNEFIOStruct *IO, int size, int count, void *dest);
    int (*CloseProc) (struct _TNEFIOStruct *IO);
    void *data;
} TNEFIOStruct;

typedef struct 
{
    char *filename;
    FILE *fptr;
    int Debug;
} TNEFFileInfo;

typedef struct 
{
    BYTE *dataStart;
    BYTE *ptr;
    long size;
    int Debug;
} TNEFMemInfo;

typedef struct 
{
    char version[10];
    variableLength from;
    variableLength subject;
    dtr dateSent;
    dtr dateReceived;
    char messageStatus[10];
    char messageClass[50];
    char messageID[50];
    char parentID[50];
    char conversationID[50];
    variableLength body;
    char priority[10];
    Attachment starting_attach;
    dtr dateModified;
    MAPIProps MapiProperties;
    variableLength CodePage;
    variableLength OriginalMessageClass;
    variableLength Owner;
    variableLength SentFor;
    variableLength Delegate;
    dtr DateStart;
    dtr DateEnd;
    variableLength AidOwner;
    int RequestRes;
    int Debug;
    TNEFIOStruct IO;
} TNEFStruct;

#endif
