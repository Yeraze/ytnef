void SaveVCard(TNEFStruct TNEF) {
    char ifilename[512];
    FILE *fptr;
    variableLength *vl;
    variableLength *pobox, *street, *city, *state, *zip, *country;
    dtr thedate;
    int boolean,index,i;

    if ((vl = MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_DISPLAY_NAME))) == MAPI_UNDEFINED) {
        if ((vl=MAPIFindProperty(&(TNEF.MapiProperties), PROP_TAG(PT_STRING8, PR_COMPANY_NAME))) == MAPI_UNDEFINED) {
            if (TNEF.subject.size > 0) {
                if (filepath == NULL) {
                    sprintf(ifilename, "%s.vcard", TNEF.subject.data);
                } else {
                    sprintf(ifilename, "%s/%s.vcard", filepath, TNEF.subject.data);
                }
            } else {
                if (filepath == NULL) {
                    sprintf(ifilename, "unknown.vcard");
                } else {
                    sprintf(ifilename, "%s/unknown.vcard", filepath);
                }
            }
        } else {
            if (filepath == NULL) {
                sprintf(ifilename, "%s.vcard", vl->data);
            } else {
                sprintf(ifilename, "%s/%s.vcard", filepath, vl->data);
            }
        }
    } else {
        if (filepath == NULL) {
            sprintf(ifilename, "%s.vcard", vl->data);
        } else {
            sprintf(ifilename, "%s/%s.vcard", filepath, vl->data);
        }
    }
    for(i=0; i<strlen(ifilename); i++) 
        if (ifilename[i] == ' ') 
            ifilename[i] = '_';
    printf("%s\n", ifilename);

    if (savefiles == 0) return;

    if ((fptr = fopen(ifilename, "wb"))==NULL) {
            printf("Error writing file to disk!");
    } else {
        fprintf(fptr, "BEGIN:VCARD\n");
        fprintf(fptr, "VERSION:2.1\n");
        if (vl != MAPI_UNDEFINED) {
            fprintf(fptr, "FN:%s\n", vl->data);
        }
        fprintProperty(TNEF, fptr, PT_STRING8, PR_NICKNAME, "NICKNAME:%s\n");
        fprintUserProp(TNEF, fptr, PT_STRING8, 0x8554, "MAILER:Microsoft Outlook %s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_SPOUSE_NAME, "X-EVOLUTION-SPOUSE:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_MANAGER_NAME, "X-EVOLUTION-MANAGER:%s\n");       
        fprintProperty(TNEF, fptr, PT_STRING8, PR_ASSISTANT, "X-EVOLUTION-ASSISTANT:%s\n");

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

        fprintProperty(TNEF, fptr, PT_STRING8, PR_OFFICE_LOCATION, "X-EVOLUTION-OFFICE:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_TITLE, "TITLE:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_PROFESSION, "ROLE:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_BODY, "NOTE:%s\n");
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


        fprintProperty(TNEF, fptr, PT_STRING8, PR_CALLBACK_TELEPHONE_NUMBER, "TEL;X-EVOLUTION-CALLBACK:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_PRIMARY_TELEPHONE_NUMBER, "TEL;PREF:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_MOBILE_TELEPHONE_NUMBER, "TEL;CELL:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_RADIO_TELEPHONE_NUMBER, "TEL;X-EVOLUTION-RADIO:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_CAR_TELEPHONE_NUMBER, "TEL;CAR:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_OTHER_TELEPHONE_NUMBER, "TEL;VOICE:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_PAGER_TELEPHONE_NUMBER, "TEL;PAGER:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_TELEX_NUMBER, "TEL;X-EVOLUTION-TELEX:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_ISDN_NUMBER, "TEL;ISDN:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_HOME2_TELEPHONE_NUMBER, "TEL;HOME:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_TTYTDD_PHONE_NUMBER, "TEL;X-EVOLUTION-TTYTDD:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_HOME_TELEPHONE_NUMBER, "TEL;HOME;VOICE:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_ASSISTANT_TELEPHONE_NUMBER, "TEL;X-EVOLUTION-ASSISTANT:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_COMPANY_MAIN_PHONE_NUMBER, "TEL;WORK:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_BUSINESS_TELEPHONE_NUMBER, "TEL;WORK:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_BUSINESS2_TELEPHONE_NUMBER, "TEL;WORK;VOICE:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_PRIMARY_FAX_NUMBER, "TEL;PREF;FAX:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_BUSINESS_FAX_NUMBER, "TEL;WORK;FAX:%s\n");
        fprintProperty(TNEF, fptr, PT_STRING8, PR_HOME_FAX_NUMBER, "TEL;HOME;FAX:%s\n");


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

        fprintProperty(TNEF, fptr, PT_STRING8, PR_BUSINESS_HOME_PAGE, "URL:%s\n");
        fprintUserProp(TNEF, fptr, PT_STRING8, 0x80d8, "FBURL:%s\n");



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

